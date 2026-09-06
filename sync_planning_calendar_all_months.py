import time
import zipfile
from datetime import datetime
from io import BytesIO

from graph_retry import install_retry_after_support, retry_delay_seconds as retry_wait_seconds
import sync_planning_calendar as calendar_sync
from sync_planning_layout import canonicalize_planning_layout, needs_schedule_reset
from sync_stock import DEST_PATH, GraphClient, get_access_token, is_retryable_graph_error


def resolve_plan_year(plan_month, now=None):
    now = now or datetime.now(calendar_sync.TIMEZONE)
    year = now.year
    delta = plan_month - now.month

    # Chọn năm gần nhất quanh thời điểm hiện tại để xử lý Dec -> Jan và Jan -> Dec.
    if delta <= -6:
        year += 1
    elif delta >= 6:
        year -= 1
    return year


def prepare_calendar_update_all_months(workbook_bytes, *, plan_year=None):
    reset_schedule = needs_schedule_reset(workbook_bytes)

    if plan_year is None:
        with zipfile.ZipFile(BytesIO(workbook_bytes), "r") as archive:
            selector = calendar_sync._read_selector_from_archive(archive)
        plan_month = calendar_sync.parse_plan_month(selector)
        plan_year = resolve_plan_year(plan_month)

    updated, info = _ORIGINAL_PREPARE(workbook_bytes, plan_year=plan_year)
    updated, layout_changes = canonicalize_planning_layout(
        updated,
        active_days=info["days"],
        reset_schedule=reset_schedule,
    )
    info = dict(info)
    info["changed_count"] += layout_changes
    info["layout_changes"] = layout_changes
    return updated, info


def run_calendar_with_retry(
    graph,
    drive_id,
    *,
    sleep_func=time.sleep,
    max_attempts=6,
    retry_delay_seconds=10,
):
    for attempt in range(1, max_attempts + 1):
        try:
            dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
            dest_bytes = graph.download_file(drive_id, dest_item["id"])
            updated_bytes, info = prepare_calendar_update_all_months(dest_bytes)

            print(
                f"[{calendar_sync.PLANNING_SHEET}] {calendar_sync.FC_SHEET}!{calendar_sync.FC_SELECTOR_CELL}="
                f"{info['selector']!r} -> {info['plan_month']:02d}/{info['plan_year']}; "
                f"dải ngày {info['start']}:{info['end']} ({info['days']} ngày)."
            )

            if info["changed_count"] == 0:
                print(
                    f"[{calendar_sync.PLANNING_SHEET}] Dải ngày đã đúng; không cần upload lại."
                )
                return info

            graph.upload_file(
                drive_id,
                dest_item["id"],
                updated_bytes,
                expected_etag=dest_item["eTag"],
            )
            print(
                f"[{calendar_sync.PLANNING_SHEET}] Đã cập nhật "
                f"{info['changed_count']} ô tiêu đề/layout ngày."
            )
            return info
        except Exception as exc:
            if not is_retryable_graph_error(exc) or attempt == max_attempts:
                raise
            delay = retry_wait_seconds(exc, retry_delay_seconds)
            print(
                f"[{calendar_sync.PLANNING_SHEET}] Graph tạm lỗi/file thay đổi; "
                f"tải snapshot mới và tính lại ({attempt}/{max_attempts}) sau {delay:g}s."
            )
            sleep_func(delay)

    raise RuntimeError("Không thể hoàn tất calendar sau các lần thử.")


_ORIGINAL_PREPARE = calendar_sync.prepare_calendar_update
calendar_sync.prepare_calendar_update = prepare_calendar_update_all_months


def main():
    install_retry_after_support()
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)
    run_calendar_with_retry(graph, drive_id)


if __name__ == "__main__":
    main()
