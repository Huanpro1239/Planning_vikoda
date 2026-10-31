"""Chạy trên GitHub Actions: đọc NXT kế toán + Xuất kho + Danh mục từ SharePoint,
+ Tồn thực tế nhà máy, dựng lại Clean Data/Data NXT ke toan/Tong_Hop_NXT_Ke_Toan.xlsx và ghi đè lên SharePoint.

    python -X utf8 -m nxt_ke_toan               # chạy thật
    python -X utf8 -m nxt_ke_toan --if-changed  # chỉ chạy khi nguồn thay đổi
    python -X utf8 -m nxt_ke_toan --dry-run     # chỉ tính, không ghi
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
from .ky_gui import doc_luong_ky_gui, ton_ky_gui
from .nha_may import tong_hop_nha_may

NXT_DIR = f"{ROOT}/Bao cao ton kho/Ton ke toan"
XUAT_KHO_DIR = f"{ROOT}/Bao cao xuat kho"
NHA_MAY_DIR = f"{ROOT}/Bao cao ton kho/Ton thuc te"
OUTPUT_PATH = f"{ROOT}/Clean Data/Data NXT ke toan/Tong_Hop_NXT_Ke_Toan.xlsx"


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


def _tai(graph, drive_id, path, item):
    return io.BytesIO(_doc(f"tải {path}", lambda: graph.download_file(drive_id, item["id"])))


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--if-changed", action="store_true",
                    help="Bỏ qua nếu NXT kế toán, Xuất kho, Tồn nhà máy và Danh mục không đổi so với lần ghi trước")
    args = ap.parse_args(argv)

    graph = GraphClient(get_access_token())
    site_id = _doc("site", graph.get_site_id)
    drive_id = _doc("drive", lambda: graph.get_default_drive_id(site_id))
    print("Đã kết nối SharePoint Planning.")

    dm_item = _doc("Danh mục", lambda: graph.get_item_by_path(drive_id, DANH_MUC_PATH))
    nxt_files = _liet_ke(graph, drive_id, NXT_DIR)
    xk_files = _liet_ke(graph, drive_id, XUAT_KHO_DIR)
    try:
        nm_files = _liet_ke(graph, drive_id, NHA_MAY_DIR)
    except GraphRequestError as exc:
        if exc.status_code != 404:
            raise
        nm_files = []
    if not nxt_files:
        raise RuntimeError("Không tìm thấy file NXT kế toán nào trong " + NXT_DIR)
    fingerprint = _fingerprint(dm_item, nxt_files + xk_files + nm_files)

    if args.if_changed:
        cu = ""
        try:
            out_item = _doc("file tổng", lambda: graph.get_item_by_path(drive_id, OUTPUT_PATH))
            cu = doc_fingerprint(io.BytesIO(_doc("tải file tổng", lambda: graph.download_file(drive_id, out_item["id"]))))
        except GraphRequestError as exc:
            if exc.status_code != 404:
                raise
        if cu == fingerprint:
            print("Không có thay đổi NXT / Xuất kho / Danh mục -> bỏ qua.")
            return 0
        print("Phát hiện thay đổi -> dựng lại file tổng.")

    dm = doc_danh_muc(_tai(graph, drive_id, DANH_MUC_PATH, dm_item))
    print(f"Danh mục: {len(dm['Vikoda'])} mã Vikoda, {len(dm['VKD'])} mã VKD")
    sources = [(p[len(NXT_DIR) + 1:], _tai(graph, drive_id, p, it)) for p, it in nxt_files]
    data, log = tong_hop_nhieu_file(sources, dm)
    for r in log:
        thieu = f" | thiếu quy cách: {r[9]}" if r[9] else ""
        print(f"  {r[0]} -> {r[1]} {r[3]:02d}/{r[2]}: {r[6]} dòng, lệch cân đối {r[7]}{thieu}")

    flows = [doc_luong_ky_gui(_tai(graph, drive_id, p, it), p[len(XUAT_KHO_DIR) + 1:]) for p, it in xk_files]
    kg, nguon = ton_ky_gui(flows)
    print(f"Tồn ký gửi: {len(flows)} file Xuất kho, nguồn đầu kỳ: {nguon}")

    nha_may, nm_bo_qua = tong_hop_nha_may(
        [(p[len(NHA_MAY_DIR) + 1:], _tai(graph, drive_id, p, it)) for p, it in nm_files])
    print(f"Tồn nhà máy: {len(nha_may)} báo cáo "
          f"({', '.join(nha_may[k]['ngay'].strftime('%d/%m/%Y') for k in sorted(nha_may))}), "
          f"bỏ qua {len(nm_bo_qua)} file không đúng mẫu")

    content, n_sp, n_thang, n_canh_bao = build_workbook(data, log, fingerprint, kg, nguon, nha_may, nm_bo_qua)
    print(f"Tổng hợp: {len(nxt_files)} file NXT, {n_sp} sản phẩm (tồn thông minh) x {n_thang} tháng, "
          f"{n_canh_bao} cảnh báo")
    if args.dry_run:
        print("[DRY-RUN] Không ghi SharePoint.")
        return 0
    _ghi_dau_ra(graph, drive_id, content)
    print(f"Đã ghi {OUTPUT_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
