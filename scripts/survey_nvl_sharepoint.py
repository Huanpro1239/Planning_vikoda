"""Script to survey SharePoint, discover Kế hoạch mua hàng.xlsx path,
download real files, and execute dry-run with detailed reconciliation audit.
"""

import json
import os
import sys
from io import BytesIO
from pathlib import Path
from urllib.parse import unquote

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import load_workbook

from sync_nvl_stock import (
    NVLConfig,
    _safe_close_workbook,
    load_nvl_config,
    read_nvl_source_stock,
    reconcile_nvl_target,
    run_nvl_sync,
)
from sync_stock import GraphClient, GraphRequestError, HOSTNAME, SITE_PATH, get_access_token


def main():
    if not os.environ.get("MS_CLIENT_SECRET"):
        print("[SURVEY] Không có MS_CLIENT_SECRET trong môi trường. Bỏ qua survey online.")
        return

    print("[SURVEY] Bắt đầu kết nối Microsoft Graph...")
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)
    print(f"[SURVEY] Site ID: {site_id}, Drive ID: {drive_id}")

    # 1. Tìm kiếm file Kế hoạch mua hàng qua duyệt thư mục (tránh lỗi 500 của search API)
    target_item = None
    exact_target_path = ""
    candidate_items = []
    target_names = {"kế hoạch mua hàng.xlsx", "ke hoach mua hang.xlsx"}

    folders_to_explore = [
        "Tinh san xuat Mua hang 2027",
        "Ke hoach",
        "Data Mua Hang",
        "Data Ton NVL",
        "Kế hoạch cung ứng",
        "",  # root
    ]
    queue = list(folders_to_explore)
    visited_folders = set()

    print("[SURVEY] Bắt đầu duyệt thư mục để tìm file đích 'Kế hoạch mua hàng.xlsx'...")
    while queue and not target_item:
        curr_folder = queue.pop(0)
        if curr_folder in visited_folders:
            continue
        visited_folders.add(curr_folder)

        if curr_folder:
            encoded = quote(curr_folder, safe="/")
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded}:/children"
        else:
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"

        try:
            data = graph.get_json(url)
            children = data.get("value", [])
        except Exception as e:
            print(f"[SURVEY] Lỗi đọc thư mục '{curr_folder}': {e}")
            continue

        for ch in children:
            name = ch.get("name", "").strip()
            ch_path = f"{curr_folder}/{name}" if curr_folder else name
            if "folder" in ch:
                if ch_path.count("/") < 3:
                    queue.append(ch_path)
            else:
                lower_name = name.lower()
                if "mua h" in lower_name or "nvl" in lower_name or lower_name in target_names:
                    print(f"  - [CANDIDATE] {ch_path} (id={ch.get('id')}, size={ch.get('size')})")
                    candidate_items.append((ch_path, ch))
                if lower_name in target_names:
                    target_item = ch
                    exact_target_path = ch_path
                    print(f"[SURVEY] >>> TÌM THẤY CHÍNH XÁC FILE ĐÍCH: {exact_target_path} (ID: {ch.get('id')})")
                    break

    out_dir = Path("dry_run_artifacts")
    out_dir.mkdir(parents=True, exist_ok=True)

    if not target_item:
        print("[SURVEY] CẢNH BÁO: Chưa tìm thấy file chính xác 'Kế hoạch mua hàng.xlsx' trên drive chính!")
        return

    # 2. Tải file nguồn XNT_ketoan_Vikoda.xlsm
    source_path = "Tinh san xuat Mua hang 2027/Ton He thong/Ton Ke Toan/XNT_ketoan_Vikoda.xlsm"
    print(f"[SURVEY] Đang tải file nguồn từ SharePoint: {source_path}")
    source_item = graph.get_item_by_path(drive_id, source_path)
    source_bytes = graph.download_file(drive_id, source_item["id"])
    (out_dir / "real_source_XNT_ketoan_Vikoda.xlsm").write_bytes(source_bytes)
    print(f"[SURVEY] Đã tải nguồn: {len(source_bytes)} bytes, eTag={source_item.get('eTag')}")

    # 3. Tải file đích Kế hoạch mua hàng.xlsx
    print(f"[SURVEY] Đang tải file đích từ SharePoint: {exact_target_path}")
    target_bytes = graph.download_file(drive_id, target_item["id"])
    (out_dir / "real_target_Ke_hoach_mua_hang.xlsx").write_bytes(target_bytes)
    print(f"[SURVEY] Đã tải đích: {len(target_bytes)} bytes, eTag={target_item.get('eTag')}")

    # 4. Kiểm tra cấu trúc file đích
    wb_tgt = load_workbook(BytesIO(target_bytes), read_only=True)
    sheetnames = wb_tgt.sheetnames
    print(f"[SURVEY] Sheet names trong file đích: {sheetnames}")
    has_ton_nvl = "Ton_NVL" in sheetnames
    if not has_ton_nvl:
        print("[SURVEY] LỖI: Sheet 'Ton_NVL' không tồn tại trong file đích!")
        wb_tgt.close()
        return

    ws_tgt = wb_tgt["Ton_NVL"]
    header_vals = [ws_tgt.cell(1, c).value for c in range(1, 15)]
    sample_row2 = [ws_tgt.cell(2, c).value for c in range(1, 15)]
    print(f"[SURVEY] Header row 1 của Ton_NVL: {header_vals}")
    print(f"[SURVEY] Sample row 2 của Ton_NVL: {sample_row2}")
    wb_tgt.close()

    # 5. Cấu hình và chạy DRY-RUN
    cfg = load_nvl_config("nvl_stock_config.json")
    cfg.target_path = exact_target_path

    # Cập nhật lại config file với path chính xác vừa tìm được
    cfg_data = json.loads(Path("nvl_stock_config.json").read_text(encoding="utf-8"))
    cfg_data["target"]["sharepoint_path"] = exact_target_path
    Path("nvl_stock_config.json").write_text(json.dumps(cfg_data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"[SURVEY] Đã cập nhật nvl_stock_config.json với target.sharepoint_path='{exact_target_path}'")

    print("[SURVEY] Bắt đầu chạy DRY-RUN (publish=False)...")
    rep = run_nvl_sync(
        cfg,
        source_file=str(out_dir / "real_source_XNT_ketoan_Vikoda.xlsm"),
        target_file=str(out_dir / "real_target_Ke_hoach_mua_hang.xlsx"),
        out_dir=str(out_dir),
        publish=False,
    )
    print(f"[SURVEY] Kết quả dry-run: {rep.get('status')} - {rep.get('message')}")

    # Copy artifacts sang thư mục gốc để workflow lưu trữ
    for fname in ["nvl_stock_proposal.xlsx", "nvl_stock_report.json"]:
        p_src = out_dir / fname
        if p_src.exists():
            Path(fname).write_bytes(p_src.read_bytes())

    # 6. Tạo bảng đối chiếu ít nhất 10 mã
    source_stock, _ = read_nvl_source_stock(source_bytes, cfg)
    rec = reconcile_nvl_target(target_bytes, source_stock, cfg)

    comparison_rows = []
    # Ưu tiên các ca đổi giá trị
    for ch in rec.changes[:10]:
        comparison_rows.append({
            "code": ch["code"],
            "row": ch["row"],
            "current_target_D": ch["before"],
            "source_stock_M": ch["after"],
            "action": "UPDATE",
            "diff": ch["diff"],
        })
    # Lấy thêm các ca không đổi nếu có
    for un in rec.unchanged[:5]:
        comparison_rows.append({
            "code": un["code"],
            "row": un["row"],
            "current_target_D": un["value"],
            "source_stock_M": un["value"],
            "action": "UNCHANGED",
            "diff": 0.0,
        })
    # Lấy thêm các ca thiếu ở nguồn nếu có
    for mis in rec.missing_in_source[:5]:
        comparison_rows.append({
            "code": mis["code"],
            "row": mis["row"],
            "current_target_D": mis["current_value"],
            "source_stock_M": "NOT_FOUND_IN_SOURCE",
            "action": "PRESERVE",
            "diff": None,
        })

    audit_summary = {
        "sharepoint_target_path": exact_target_path,
        "sharepoint_source_path": source_path,
        "source_item_id": source_item.get("id"),
        "target_item_id": target_item.get("id"),
        "target_sheet": cfg.target_sheet,
        "headers": header_vals[:8],
        "total_source_codes": len(source_stock),
        "total_target_rows": rec.total_rows,
        "metrics": {
            "changed_count": len(rec.changes),
            "unchanged_count": len(rec.unchanged),
            "missing_in_source_count": len(rec.missing_in_source),
            "duplicate_codes_source": rec.duplicate_codes_source,
            "duplicate_codes_target": rec.duplicate_codes_target,
        },
        "sample_reconciliation_10_plus": comparison_rows,
    }

    (out_dir / "audit_summary.json").write_text(
        json.dumps(audit_summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"[SURVEY] Hoàn tất khảo sát và ghi nhận audit vào {out_dir}")


if __name__ == "__main__":
    main()
