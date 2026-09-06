import time

from graph_retry import install_retry_after_support, retry_delay_seconds as retry_wait_seconds
from planning_schedule_report import (
    attach_output_hash,
    build_schedule_report,
    print_operational_report,
    save_schedule_report,
)
import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_debt as debt_priority
import sync_planning_schedule_priority_v8 as v8
from sync_planning_layout import canonicalize_planning_layout
from sync_stock import DEST_PATH, GraphClient, get_access_token, is_retryable_graph_error
from verify_planning_month import verify_workbook


PLANNING_SHEET = "Ke_hoach_SX"


def install_production_output_cleanup():
    """Install the production scheduler hooks and canonical output layout.

    V8 installs the full priority chain and may overwrite lower-level RGB/Galon
    hooks. Debt-aware priority must therefore be installed *after* V8 every
    time this installer is called. Repeated installer calls/retries reapply the
    debt-aware hooks instead of silently falling back to the old logic.
    """
    v8.install_priority_scheduler_v8()
    debt_priority.install_debt_aware_priority()

    original_patch = priority.base.patch_schedule_workbook

    def patch_schedule_without_extra_columns(workbook_bytes, headers, products, schedule):
        updated, changed = original_patch(
            workbook_bytes,
            headers,
            products,
            schedule,
        )
        updated, layout_changes = canonicalize_planning_layout(
            updated,
            active_days=len(headers),
            reset_schedule=False,
        )
        return updated, changed + layout_changes

    priority.base.patch_schedule_workbook = patch_schedule_without_extra_columns


def _print_scheduler_summary(info):
    for line, utilization in sorted(info["utilization"].items()):
        capacity = info["line_capacity"][line]
        mode = info["optimizer_meta"].get(line, {}).get("mode", "")
        setup = info["optimizer_meta"].get(line, {}).get("setup_shifts", 0)
        print(
            f"[Scheduler] {line}: mode={mode}; {capacity:g} ca/ngày; "
            f"setup {setup:g} ca; utilization tháng {utilization * 100:.1f}%."
        )


def run_scheduler_with_retry(
    graph,
    drive_id,
    *,
    sleep_func=time.sleep,
    max_attempts=6,
    retry_delay_seconds=10,
):
    """Read -> compute -> validate -> write as one retryable snapshot transaction."""
    for attempt in range(1, max_attempts + 1):
        try:
            dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
            dest_bytes = graph.download_file(drive_id, dest_item["id"])

            updated_bytes, info = priority.base.prepare_schedule_update(dest_bytes)
            report = build_schedule_report(
                dest_bytes,
                info,
                input_revision={
                    "item_id": dest_item.get("id"),
                    "etag": dest_item.get("eTag"),
                    "last_modified": dest_item.get("lastModifiedDateTime"),
                },
                algorithm="priority_v8",
            )
            report = attach_output_hash(report, updated_bytes)

            # Critical: validate the exact output bytes before any upload.
            verify_workbook(updated_bytes, schedule_report=report)
            save_schedule_report(report)
            _print_scheduler_summary(info)
            print_operational_report(report)

            if info["changed_count"] == 0:
                print(f"[{PLANNING_SHEET}] Lịch đã đúng; dry validation PASS, không cần upload lại.")
                return info, report

            graph.upload_file(
                drive_id,
                dest_item["id"],
                updated_bytes,
                expected_etag=dest_item["eTag"],
            )
            print(
                f"[{PLANNING_SHEET}] Đã validate trước publish và cập nhật "
                f"{info['changed_count']} ô lịch sản xuất."
            )
            return info, report
        except Exception as exc:
            if not is_retryable_graph_error(exc) or attempt == max_attempts:
                raise
            delay = retry_wait_seconds(exc, retry_delay_seconds)
            print(
                f"[{PLANNING_SHEET}] Graph tạm lỗi/file thay đổi; tải snapshot mới và "
                f"tính lại ({attempt}/{max_attempts}) sau {delay:g}s."
            )
            sleep_func(delay)

    raise RuntimeError("Không thể hoàn tất scheduler sau các lần thử.")


def main():
    install_retry_after_support()
    install_production_output_cleanup()
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)
    run_scheduler_with_retry(graph, drive_id)


if __name__ == "__main__":
    main()
