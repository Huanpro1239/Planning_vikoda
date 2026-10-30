"""Chạy trên GitHub Actions: đọc toàn bộ Sell In + Danh mục từ SharePoint,
dựng lại Tong_Hop_Ban_Hang.xlsx và ghi đè lên SharePoint.

    python -X utf8 -m sell_in            # chạy thật
    python -X utf8 -m sell_in --dry-run  # chỉ tính, không ghi
"""
import argparse
import io
import sys
import time

from sharepoint.client import GraphClient, GraphRequestError, get_access_token

from .core import build_workbook, doc_danh_muc, tong_hop_nhieu_file

ROOT = "Planning App"
SELL_IN_DIR = f"{ROOT}/Bao cao ban hang/Sell In"
DANH_MUC_PATH = f"{ROOT}/Danh muc/Danh Muc San Pham.xlsx"
OUTPUT_PATH = f"{ROOT}/Clean Data/Data ban hang/Tong_Hop_Ban_Hang.xlsx"


def _liet_ke(graph, drive_id, folder):
    """Liệt kê đệ quy file .xlsm/.xlsx trong thư mục (bỏ file tạm ~$)."""
    out = []
    for item in graph.list_folder_children(drive_id, folder):
        path = f"{folder}/{item['name']}"
        if item.get("folder") is not None:
            out.extend(_liet_ke(graph, drive_id, path))
        elif item["name"].lower().endswith((".xlsm", ".xlsx")) and not item["name"].startswith("~$"):
            out.append((path, item))
    return sorted(out, key=lambda x: x[0])


def _ghi_dau_ra(graph, drive_id, content, lan_thu=5):
    for i in range(1, lan_thu + 1):
        try:
            try:
                item = graph.get_item_by_path(drive_id, OUTPUT_PATH)
            except GraphRequestError as exc:
                if exc.status_code != 404:
                    raise
                return graph.create_file_by_path(drive_id, OUTPUT_PATH, content, conflict_behavior="fail")
            return graph.upload_file(drive_id, item["id"], content, item["eTag"])
        except GraphRequestError as exc:
            # 423 = file đang bị khóa (mở trong Excel desktop), 412/409/5xx = thử lại
            if i == lan_thu or exc.status_code not in (409, 412, 423, 429, 500, 502, 503, 504):
                raise
            print(f"[WARN] Ghi file tổng lỗi {exc.status_code}, thử lại lần {i + 1} sau {20 * i}s")
            time.sleep(20 * i)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args(argv)

    graph = GraphClient(get_access_token())
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)
    print("Đã kết nối SharePoint Planning.")

    dm_item = graph.get_item_by_path(drive_id, DANH_MUC_PATH)
    dm = doc_danh_muc(io.BytesIO(graph.download_file(drive_id, dm_item["id"])))
    print(f"Danh mục: {len(dm['Vikoda'])} mã Vikoda, {len(dm['VKD'])} mã VKD")

    files = _liet_ke(graph, drive_id, SELL_IN_DIR)
    if not files:
        raise RuntimeError("Không tìm thấy file Sell In nào trong " + SELL_IN_DIR)
    sources = []
    for path, item in files:
        sources.append((item["name"], io.BytesIO(graph.download_file(drive_id, item["id"]))))
    data, log = tong_hop_nhieu_file(sources, dm)
    for r in log:
        thieu = f" | thiếu quy cách: {r[7]}" if r[7] else ""
        print(f"  {r[0]} -> {r[1]} {r[3]:02d}/{r[2]}: {r[6]} SP{thieu}")

    content, n_sp, n_thang = build_workbook(data, log)
    print(f"Tổng hợp: {len(files)} file, {n_sp} sản phẩm x {n_thang} tháng")
    if args.dry_run:
        print("[DRY-RUN] Không ghi SharePoint.")
        return 0
    _ghi_dau_ra(graph, drive_id, content)
    print(f"Đã ghi {OUTPUT_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
