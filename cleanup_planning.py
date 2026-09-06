import time

from graph_retry import install_retry_after_support, retry_delay_seconds
from planning_cleanup import remove_sheet_formulas
from sync_stock import DEST_PATH, GraphClient, get_access_token, is_retryable_graph_error

PLANNING_SHEET = "Ke_hoach_SX"


def cleanup_with_retry(
    graph,
    drive_id,
    *,
    cleanup_func=remove_sheet_formulas,
    sleep_func=time.sleep,
    max_attempts=6,
    retry_delay_seconds_default=10,
    retry_delay_seconds=None,
):
    # Backward-compatible keyword used by older callers/tests.
    if retry_delay_seconds is not None:
        retry_delay_seconds_default = retry_delay_seconds

    for attempt in range(1, max_attempts + 1):
        try:
            # Read + compute + write are one retryable transaction. In
            # particular 412 must re-download and re-compute from the new ETag.
            dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
            dest_bytes = graph.download_file(drive_id, dest_item["id"])

            updated_bytes, removed = cleanup_func(
                dest_bytes,
                PLANNING_SHEET,
            )

            if removed == 0:
                print(f"[{PLANNING_SHEET}] Không có công thức cần loại bỏ.")
                return 0

            graph.upload_file(
                drive_id,
                dest_item["id"],
                updated_bytes,
                expected_etag=dest_item["eTag"],
            )
            return removed
        except Exception as exc:
            if not is_retryable_graph_error(exc) or attempt == max_attempts:
                raise

            delay = retry_delay_seconds(exc, retry_delay_seconds_default)
            print(
                f"[{PLANNING_SHEET}] Graph tạm lỗi/file thay đổi; "
                f"thử lại ({attempt}/{max_attempts}) sau {delay:g}s."
            )
            sleep_func(delay)

    raise RuntimeError("Không thể hoàn tất cleanup sau các lần thử.")


def main():
    install_retry_after_support()
    token = get_access_token()
    graph = GraphClient(token)

    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    removed = cleanup_with_retry(graph, drive_id)

    if removed:
        print(
            f"[{PLANNING_SHEET}] Đã loại bỏ {removed} công thức; "
            "các sheet khác giữ nguyên."
        )


if __name__ == "__main__":
    main()
