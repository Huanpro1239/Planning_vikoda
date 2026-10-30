"""Chạy trên GitHub Actions: đọc toàn bộ Bao cao xuat kho + Danh mục từ SharePoint,
dựng lại Clean Data/Data Xuat kho/Tong_Hop_Xuat_Kho.xlsx và ghi đè lên SharePoint.

    python -X utf8 -m xuat_kho               # chạy thật
    python -X utf8 -m xuat_kho --if-changed  # chỉ chạy khi Xuất kho / Danh mục thay đổi
    python -X utf8 -m xuat_kho --dry-run     # chỉ tính, không ghi
"""
import argparse
import hashlib
import io
import sys
import time

from sell_in.__main__ import DANH_MUC_PATH, ROOT, _doc, _liet_ke
from sell_in.core import doc_danh_muc, doc_fingerprint
from sharepoint.client import GraphClient, GraphRequestError, get_access_token

from .core import build_workbook, tong_hop_nhieu_file

XUAT_KHO_DIR = f"{ROOT}/Bao cao xuat kho"
OUTPUT_PATH = f"{ROOT}/Clean Data/Data Xuat kho/Tong_Hop_Xuat_Kho.xlsx"


def _fingerprint(dm_item, files):
    parts = [f"{DANH_MUC_PATH}|{dm_item.get('eTag')}|{dm_item.get('size')}"]
    parts += [f"{p}|{it.get('eTag')}|{it.get('size')}" for p, it in files]
    return hashlib.sha256("\n".join(parts).encode("utf-8")).hexdigest()


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
            if i == lan_thu or exc.status_code not in (409, 412, 423, 429, 500, 502, 503, 504):
                raise
            print(f"[WARN] Ghi file tổng lỗi {exc.status_code}, thử lại lần {i + 1} sau {20 * i}s")
            time.sleep(20 * i)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--if-changed", action="store_true",
                    help="Bỏ qua nếu file Xuất kho và Danh mục không đổi so với lần ghi trước")
    args = ap.parse_args(argv)

    graph = GraphClient(get_access_token())
    site_id = _doc("site", graph.get_site_id)
    drive_id = _doc("drive", lambda: graph.get_default_drive_id(site_id))
    print("Đã kết nối SharePoint Planning.")

    dm_item = _doc("Danh mục", lambda: graph.get_item_by_path(drive_id, DANH_MUC_PATH))
    files = _liet_ke(graph, drive_id, XUAT_KHO_DIR)
    if not files:
        raise RuntimeError("Không tìm thấy file Xuất kho nào trong " + XUAT_KHO_DIR)
    fingerprint = _fingerprint(dm_item, files)

    if args.if_changed:
        cu = ""
        try:
            out_item = _doc("file tổng", lambda: graph.get_item_by_path(drive_id, OUTPUT_PATH))
            cu = doc_fingerprint(io.BytesIO(_doc("tải file tổng", lambda: graph.download_file(drive_id, out_item["id"]))))
        except GraphRequestError as exc:
            if exc.status_code != 404:
                raise
        if cu == fingerprint:
            print("Không có thay đổi Xuất kho / Danh mục -> bỏ qua.")
            return 0
        print("Phát hiện thay đổi -> dựng lại file tổng.")

    dm = doc_danh_muc(io.BytesIO(_doc("tải Danh mục", lambda: graph.download_file(drive_id, dm_item["id"]))))
    print(f"Danh mục: {len(dm['Vikoda'])} mã Vikoda, {len(dm['VKD'])} mã VKD")
    sources = []
    for path, item in files:
        data_bytes = _doc(f"tải {path}", lambda item=item: graph.download_file(drive_id, item["id"]))
        sources.append((path[len(XUAT_KHO_DIR) + 1:], io.BytesIO(data_bytes)))
    data, log = tong_hop_nhieu_file(sources, dm)
    for r in log:
        thieu = f" | thiếu quy cách: {r[8]}" if r[8] else ""
        print(f"  {r[0]} -> {r[1]} {r[3]:02d}/{r[2]}: {r[7]} SP, bỏ {r[6]} dòng chuyển ký gửi{thieu}")

    content, n_sp, n_thang = build_workbook(data, log, fingerprint)
    print(f"Tổng hợp: {len(files)} file, {n_sp} sản phẩm x {n_thang} tháng")
    if args.dry_run:
        print("[DRY-RUN] Không ghi SharePoint.")
        return 0
    _ghi_dau_ra(graph, drive_id, content)
    print(f"Đã ghi {OUTPUT_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
