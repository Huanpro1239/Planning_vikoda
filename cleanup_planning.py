import time

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
    retry_delay_seconds=10,
):
    for attempt in range(1, max_attempts + 1):
        dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
        dest_bytes = graph.download_file(drive_id, dest_item["id"])

        updated_bytes, removed = cleanup_func(
            dest_bytes,
            PLANNING_SHEET,
        )

        if removed == 0:
            print(f"[{PLANNING_SHEET}] Không có công thức cần loại bỏ.")
            return 0

        try:
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

            print(
                f"[{PLANNING_SHEET}] File đang khóa/thay đổi; "
                f"thử lại ({attempt}/{max_attempts})."
            )
            sleep_func(retry_delay_seconds)

    raise RuntimeError("Không thể hoàn tất cleanup sau các lần thử.")


def main():
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
