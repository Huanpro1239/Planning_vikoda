from planning_cleanup import remove_sheet_formulas
from sync_stock import DEST_PATH, GraphClient, get_access_token

PLANNING_SHEET = "Ke_hoach_SX"


def main():
    token = get_access_token()
    graph = GraphClient(token)

    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
    dest_bytes = graph.download_file(drive_id, dest_item["id"])

    updated_bytes, removed = remove_sheet_formulas(
        dest_bytes,
        PLANNING_SHEET,
    )

    if removed == 0:
        print(f"[{PLANNING_SHEET}] Không có công thức cần loại bỏ.")
        return

    graph.upload_file(
        drive_id,
        dest_item["id"],
        updated_bytes,
        expected_etag=dest_item["eTag"],
    )

    print(
        f"[{PLANNING_SHEET}] Đã loại bỏ {removed} công thức; "
        "các sheet khác giữ nguyên."
    )


if __name__ == "__main__":
    main()
