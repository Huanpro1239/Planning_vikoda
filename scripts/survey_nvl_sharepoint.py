"""Script to survey SharePoint / local snapshots for NVL Stock,
verify SharePoint identity, examine reporting period and units,
execute dry-run reconciliation, and generate audit_summary.json.
"""

import argparse
import json
import os
import sys
from datetime import datetime, timezone

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import quote, unquote

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
from sync_stock import GraphClient, GraphRequestError, get_access_token


EQUIVALENT_UNIT_ALIASES: dict[str, str] = {
    "cai": "cái",
    "cái": "cái",
    "kg": "kg",
}


def normalize_unit_for_report(u: str) -> str:
    """Chuẩn hóa chuỗi đơn vị chỉ để phục vụ đối chiếu phân loại trong báo cáo kiểm toán (audit).
    TUYỆT ĐỐI KHÔNG sửa đổi cột C của workbook đích và KHÔNG tự ý quy đổi số lượng.
    """
    s = (u or "").strip().lower()
    return EQUIVALENT_UNIT_ALIASES.get(s, s)


def format_nvl_quantity(val: Any) -> str:
    """Định dạng số lượng tồn kho hiển thị trong audit (ví dụ 473992 -> '473,992', 934.388 -> '934.388')."""
    if val is None:
        return ""
    try:
        fval = float(val)
        if fval.is_integer():
            return f"{int(fval):,}"
        return f"{fval:,.4f}".rstrip("0").rstrip(".")
    except (ValueError, TypeError):
        return str(val)


# Xác nhận chính thức của người dùng được gắn cụ thể với kỳ báo cáo, mã và đơn vị:
# 1. Dùng số tồn theo báo cáo nguồn cho kỳ 'Từ ngày 01-08-2026 đến ngày 31-08-2026'.
# 2. Mã 430200173: chép trực tiếp theo ĐVT 'CAI' (Cái), không quy đổi sang Kg.
# 3. 7 mã thiếu ở nguồn: xác nhận theo file cũ, bảo toàn (PRESERVE), không gán 0.
# 4. 17 dòng thiếu ĐVT đích ở cột C: chừa trống, không tự ý điền.
APPROVED_USER_SCOPE: dict[str, Any] = {
    "reporting_period": "Từ ngày 01-08-2026 đến ngày 31-08-2026",
    "confirmed_units": {
        "430200173": "CAI",
    },
    "preserve_missing": True,
    "leave_empty_target_units": True,
}


def is_approved_reporting_period(period: str | None) -> bool:
    """Kiểm tra kỳ báo cáo nguồn hiện tại có khớp với kỳ người dùng đã phê duyệt không."""
    if not period:
        return False
    return period.strip() == APPROVED_USER_SCOPE["reporting_period"].strip()


def parse_args():
    parser = argparse.ArgumentParser(description="Khảo sát và đối soát dữ liệu tồn kho NVL.")
    parser.add_argument("--config", default="nvl_stock_config.json", help="Đường dẫn file cấu hình JSON.")
    parser.add_argument("--source-file", help="Đường dẫn file nguồn offline (tùy chọn).")
    parser.add_argument("--target-file", help="Đường dẫn file đích offline (tùy chọn).")
    parser.add_argument("--out-dir", default=".", help="Thư mục xuất kết quả.")
    return parser.parse_args()


def run_survey(
    config_path: str = "nvl_stock_config.json",
    source_file: str | None = None,
    target_file: str | None = None,
    out_dir_path: str = ".",
    graph: GraphClient | None = None,
) -> dict[str, Any]:
    current_phase = "init"
    out_dir = Path(out_dir_path)
    out_dir.mkdir(parents=True, exist_ok=True)
    cfg: NVLConfig | None = None
    client_graph: GraphClient | None = graph
    drive_id: str | None = None

    # Loại bỏ proposal cũ trong out_dir ngay từ đầu để không lưu sót kết quả trước
    (out_dir / "nvl_stock_proposal.xlsx").unlink(missing_ok=True)

    try:
        current_phase = "load_config"
        cfg = load_nvl_config(config_path)
        print(f"[SURVEY] Đã load cấu hình từ {config_path}")
        print(f"[SURVEY] Nguồn: {cfg.source_name} (path: '{cfg.source_path}', start_row={cfg.source_start_row})")
        print(f"[SURVEY] Đích: {cfg.target_name} (path: '{cfg.target_path}', start_row={cfg.target_start_row})")

        source_bytes: bytes
        target_bytes: bytes
        source_item: dict[str, Any] = {}
        target_item: dict[str, Any] = {}

        # 1. Thu thập dữ liệu: Offline hoặc Online
        if source_file and target_file:
            current_phase = "read_offline_files"
            p_src = Path(source_file)
            p_tgt = Path(target_file)
            if not p_src.exists():
                raise FileNotFoundError(f"Không tìm thấy file nguồn offline: {p_src}")
            if not p_tgt.exists():
                raise FileNotFoundError(f"Không tìm thấy file đích offline: {p_tgt}")
            source_bytes = p_src.read_bytes()
            target_bytes = p_tgt.read_bytes()
            source_local_path = out_dir / f"real_source_{p_src.name}"
            target_local_path = out_dir / f"real_target_{p_tgt.name}"
            source_local_path.write_bytes(source_bytes)
            target_local_path.write_bytes(target_bytes)
            print(f"[SURVEY] Đọc thành công file offline: nguồn={len(source_bytes)}B, đích={len(target_bytes)}B")
            source_item = {"id": "offline_local", "name": p_src.name, "eTag": "local_snapshot"}
            target_item = {"id": "offline_local", "name": p_tgt.name, "eTag": "local_snapshot"}
        else:
            current_phase = "connect_graph"
            if not os.environ.get("MS_CLIENT_SECRET") and client_graph is None:
                raise RuntimeError("Không có MS_CLIENT_SECRET trong môi trường và không truyền graph client.")

            if client_graph is None:
                token = get_access_token()
                client_graph = GraphClient(token)
            site_id = client_graph.get_site_id()
            drive_id = client_graph.get_default_drive_id(site_id)
            print(f"[SURVEY] Site ID: {site_id}, Drive ID: {drive_id}")

            # Xác minh file đích theo config path
            current_phase = "fetch_target"
            if not cfg.target_path:
                raise ValueError("Chưa cấu hình 'target.sharepoint_path' trong nvl_stock_config.json")

            print(f"[SURVEY] Kiểm tra file đích theo path đã cấu hình: '{cfg.target_path}'")
            target_item = client_graph.get_item_by_path(drive_id, cfg.target_path)
            # Xác minh identity
            target_etag = target_item.get("eTag", "")
            if cfg.target_sourcedoc and cfg.target_sourcedoc.lower() not in target_etag.lower():
                raise RuntimeError(
                    f"Target sourcedoc '{cfg.target_sourcedoc}' không khớp eTag '{target_etag}' của file đích ({target_item.get('id')})."
                )
            print(f"[SURVEY] Đã xác minh identity file đích: ID={target_item.get('id')}, eTag={target_etag}")

            target_bytes = client_graph.download_file(drive_id, target_item["id"])
            tgt_basename = Path(cfg.target_path).name or "target.xlsx"
            target_local_path = out_dir / f"real_target_{tgt_basename}"
            target_local_path.write_bytes(target_bytes)

            # Xác minh file nguồn
            current_phase = "fetch_source"
            if not cfg.source_path:
                raise ValueError("Chưa cấu hình 'source.sharepoint_path' trong nvl_stock_config.json")

            print(f"[SURVEY] Kiểm tra file nguồn theo path đã cấu hình: '{cfg.source_path}'")
            source_item = client_graph.get_item_by_path(drive_id, cfg.source_path)
            source_etag = source_item.get("eTag", "")
            if cfg.source_sourcedoc and cfg.source_sourcedoc.lower() not in source_etag.lower():
                raise RuntimeError(
                    f"Source sourcedoc '{cfg.source_sourcedoc}' không khớp eTag '{source_etag}' của file nguồn ({source_item.get('id')})."
                )
            print(f"[SURVEY] Đã xác minh identity file nguồn: ID={source_item.get('id')}, eTag={source_etag}")

            source_bytes = client_graph.download_file(drive_id, source_item["id"])
            src_basename = Path(cfg.source_path).name or "source.xlsm"
            source_local_path = out_dir / f"real_source_{src_basename}"
            source_local_path.write_bytes(source_bytes)

        # 2. Khảo sát cấu trúc & kỳ báo cáo ở file nguồn
        current_phase = "survey_source"
        wb_src = load_workbook(BytesIO(source_bytes), data_only=True)
        if cfg.source_sheet not in wb_src.sheetnames:
            _safe_close_workbook(wb_src)
            raise ValueError(f"Sheet nguồn '{cfg.source_sheet}' không tồn tại trong {cfg.source_name}")
        ws_src = wb_src[cfg.source_sheet]

        # Đọc kỳ báo cáo từ ô A7 (hoặc quét dòng 1 đến dòng 15)
        reporting_period = None
        for r in range(1, min(cfg.source_start_row, 15)):
            val = ws_src.cell(r, 1).value
            if isinstance(val, str) and any(k in val.lower() for k in ("từ ngày", "đến ngày", "kỳ", "tháng")):
                reporting_period = val.strip()
                break

        # Đọc danh mục ĐVT nguồn từ cột F (cột 6)
        source_units: dict[str, str] = {}
        for r in range(cfg.source_start_row, ws_src.max_row + 1):
            c_val = ws_src.cell(r, cfg.source_code_col).value
            u_val = ws_src.cell(r, 6).value  # Cột F
            if c_val is not None:
                code_str = str(c_val).strip()
                if code_str:
                    source_units[code_str] = str(u_val).strip() if u_val is not None else ""
        _safe_close_workbook(wb_src)
        print(f"[SURVEY] Kỳ báo cáo phát hiện trong nguồn: {reporting_period}")

        # 3. Khảo sát cấu trúc & ĐVT ở file đích
        current_phase = "survey_target"
        wb_tgt = load_workbook(BytesIO(target_bytes), data_only=True)
        if cfg.target_sheet not in wb_tgt.sheetnames:
            _safe_close_workbook(wb_tgt)
            raise ValueError(f"Sheet đích '{cfg.target_sheet}' không tồn tại trong {cfg.target_name}")
        ws_tgt = wb_tgt[cfg.target_sheet]

        header_vals = [ws_tgt.cell(1, c).value for c in range(1, 10)]

        # Đọc danh mục ĐVT & Tên ở đích
        target_units: dict[str, str] = {}
        target_names_dict: dict[str, str] = {}
        for r in range(cfg.target_start_row, ws_tgt.max_row + 1):
            c_val = ws_tgt.cell(r, cfg.target_code_col).value
            name_val = ws_tgt.cell(r, 2).value  # Cột B
            u_val = ws_tgt.cell(r, 3).value     # Cột C (ĐVT)
            if c_val is not None:
                code_str = str(c_val).strip()
                if code_str:
                    target_units[code_str] = str(u_val).strip() if u_val is not None else ""
                    target_names_dict[code_str] = str(name_val).strip() if name_val is not None else ""
        _safe_close_workbook(wb_tgt)

        # 4. Chạy Dry-run đồng bộ
        current_phase = "run_dry_run"
        print("[SURVEY] Bắt đầu chạy run_nvl_sync (publish=False)...")
        rep = run_nvl_sync(
            cfg,
            source_file=str(source_local_path),
            target_file=str(target_local_path),
            out_dir=str(out_dir),
            publish=False,
        )
        print(f"[SURVEY] Kết quả dry-run: status={rep.get('status')}")

        # 5. Đối chiếu số liệu và ĐVT
        current_phase = "build_audit"
        source_stock, source_meta = read_nvl_source_stock(source_bytes, cfg)
        rec = reconcile_nvl_target(target_bytes, source_stock, cfg)

        # Tra cứu giá trị trước thay đổi của cột D
        target_before_vals: dict[str, Any] = {}
        for ch in rec.changes:
            target_before_vals[ch["code"]] = ch["before"]
        for un in rec.unchanged:
            target_before_vals[un["code"]] = un["value"]

        # Đối chiếu ĐVT nguồn (F) vs đích (C) cho toàn bộ mã khớp
        is_approved_period = is_approved_reporting_period(reporting_period)
        exact_matches: list[dict[str, Any]] = []
        alias_matches: list[dict[str, Any]] = []
        missing_target_units: list[dict[str, Any]] = []
        divergent_units: list[dict[str, Any]] = []
        unit_comparisons: list[dict[str, Any]] = []

        for code, r in sorted(rec.target_codes.items(), key=lambda kv: kv[1]):
            if code in source_stock:
                src_u = source_units.get(code, "")
                tgt_u = target_units.get(code, "")
                qty = source_stock[code]
                qty_str = format_nvl_quantity(qty)
                src_clean = src_u.strip().lower()
                tgt_clean = tgt_u.strip().lower()

                item_comp = {
                    "row": r,
                    "code": code,
                    "name": target_names_dict.get(code, ""),
                    "source_unit_F": src_u,
                    "target_unit_C": tgt_u,
                    "source_stock_M": qty,
                    "target_stock_D_before": target_before_vals.get(code),
                }

                if not tgt_clean:
                    item_comp["classification"] = "MISSING_TARGET_UNIT"
                    item_comp["match"] = False
                    if is_approved_period:
                        item_comp["note"] = (
                            f"Đơn vị đích trống (nguồn: '{src_u}'). Người dùng đã CHỐT cho kỳ '{reporting_period}': Dòng trống không có thì chừa "
                            "(bảo toàn ô C đích trống, không tự ý điền)."
                        )
                    else:
                        item_comp["note"] = (
                            f"Đơn vị đích trống (nguồn: '{src_u}'). Chưa có xác nhận cho kỳ '{reporting_period}' "
                            "(áp dụng chính sách mặc định: bảo toàn ô C đích trống, không tự ý điền)."
                        )
                    missing_target_units.append(item_comp)
                elif src_clean == tgt_clean:
                    item_comp["classification"] = "EXACT_MATCH"
                    item_comp["match"] = True
                    item_comp["note"] = f"Khớp chính xác ('{src_u}' == '{tgt_u}')."
                    exact_matches.append(item_comp)
                elif normalize_unit_for_report(src_u) == normalize_unit_for_report(tgt_u):
                    item_comp["classification"] = "ALIAS_MATCH"
                    item_comp["match"] = True
                    item_comp["note"] = (
                        f"Khớp alias ('{src_u}' vs '{tgt_u}'): cùng đại lượng chuẩn hóa "
                        f"'{normalize_unit_for_report(src_u)}', khác cách viết/chữ hoa/dấu."
                    )
                    alias_matches.append(item_comp)
                else:
                    item_comp["classification"] = "DIVERGENT"
                    item_comp["match"] = False
                    if code == "430200173":
                        is_code_approved = (
                            is_approved_period
                            and src_u.strip().upper() == APPROVED_USER_SCOPE["confirmed_units"].get("430200173")
                        )
                        if is_code_approved:
                            note = (
                                f"ĐVT nguồn '{src_u}' ({qty_str} nhãn thân PET 1.5L) vs đích ghi '{tgt_u}'. "
                                f"Người dùng đã CHỐT cho kỳ '{reporting_period}': Ghi nhận theo ĐVT Cái, "
                                f"chép trực tiếp {qty_str} vào cột D, không quy đổi sang Kg."
                            )
                        else:
                            if not is_approved_period:
                                reason = f"Xác nhận trước đó gắn với kỳ '{APPROVED_USER_SCOPE['reporting_period']}' (ĐVT 'CAI')"
                            else:
                                reason = f"ĐVT nguồn hiện tại là '{src_u}' (thay đổi so với ĐVT 'CAI' đã chốt)"
                            note = (
                                f"ĐVT nguồn '{src_u}' ({qty_str} nhãn thân PET 1.5L) vs đích ghi '{tgt_u}'. "
                                f"{reason} và không tự áp dụng cho kỳ hiện tại '{reporting_period}'. Cần người dùng xác nhận cho kỳ mới."
                            )
                    else:
                        if is_approved_period:
                            note = (
                                f"Khác đơn vị đo lường: nguồn='{src_u}' vs đích='{tgt_u}'. "
                                f"Người dùng đã CHỐT cho kỳ '{reporting_period}': dùng theo báo cáo nguồn (chép trực tiếp số tồn {qty_str})."
                            )
                        else:
                            note = (
                                f"Khác đơn vị đo lường: nguồn='{src_u}' vs đích='{tgt_u}' (số tồn {qty_str}). "
                                f"Kỳ '{reporting_period}' chưa có xác nhận người dùng; áp dụng chính sách mặc định: chép trực tiếp số tồn."
                            )
                    item_comp["note"] = note
                    divergent_units.append(item_comp)

                unit_comparisons.append(item_comp)

        # Mẫu đối chiếu ít nhất 10 mã (bao gồm cập nhật, số 0, số lẻ, mã thiếu)
        comparison_rows = []
        for ch in rec.changes[:10]:
            before_val = ch["before"]
            try:
                diff_val = round(ch["after"] - float(before_val), 4) if before_val is not None else ch["after"]
            except Exception:
                diff_val = None
            comparison_rows.append({
                "code": ch["code"],
                "row": ch["row"],
                "name": target_names_dict.get(ch["code"], ""),
                "current_target_D": ch["before"],
                "source_stock_M": ch["after"],
                "action": "UPDATE",
                "diff": diff_val,
                "source_unit_F": source_units.get(ch["code"], ""),
                "target_unit_C": target_units.get(ch["code"], ""),
            })

        for un in rec.unchanged[:5]:
            comparison_rows.append({
                "code": un["code"],
                "row": un["row"],
                "name": target_names_dict.get(un["code"], ""),
                "current_target_D": un["value"],
                "source_stock_M": un["value"],
                "action": "UNCHANGED",
                "diff": 0.0,
                "source_unit_F": source_units.get(un["code"], ""),
                "target_unit_C": target_units.get(un["code"], ""),
            })

        for mis in rec.missing_in_source[:5]:
            comparison_rows.append({
                "code": mis["code"],
                "row": mis["row"],
                "name": target_names_dict.get(mis["code"], ""),
                "current_target_D": mis["current_value"],
                "source_stock_M": "NOT_FOUND_IN_SOURCE",
                "action": "PRESERVE",
                "diff": None,
                "source_unit_F": "NOT_FOUND",
                "target_unit_C": target_units.get(mis["code"], ""),
            })

        # Danh sách chi tiết các mã thiếu ở nguồn
        missing_items_detail = []
        for mis in rec.missing_in_source:
            if is_approved_period:
                note_mis = (
                    f"Mã đích không có trong báo cáo kế toán kỳ '{reporting_period}'. "
                    "Người dùng đã CHỐT cho kỳ này: Xác nhận theo file cũ, bảo toàn ô đích (PRESERVE), không gán 0."
                )
            else:
                note_mis = (
                    f"Mã đích không có trong báo cáo kế toán kỳ '{reporting_period}'. "
                    f"Xác nhận cũ cho kỳ '{APPROVED_USER_SCOPE['reporting_period']}' không tự áp dụng; "
                    "áp dụng chính sách mặc định: bảo toàn ô đích (PRESERVE), không gán 0."
                )
            missing_items_detail.append({
                "row": mis["row"],
                "code": mis["code"],
                "name": target_names_dict.get(mis["code"], ""),
                "target_unit_C": target_units.get(mis["code"], ""),
                "current_value": mis["current_value"],
                "action": "PRESERVE",
                "note": note_mis,
            })

        stock_430200173 = source_stock.get("430200173")
        qty_430200173_str = format_nvl_quantity(stock_430200173) if stock_430200173 is not None else "N/A"
        has_code_430200173 = "430200173" in source_stock
        unit_430200173 = source_units.get("430200173", "").strip().upper()
        is_unit_430200173_approved = (unit_430200173 == APPROVED_USER_SCOPE["confirmed_units"].get("430200173"))
        is_430200173_fully_approved = (is_approved_period and has_code_430200173 and is_unit_430200173_approved)

        if not is_approved_period:
            reporting_period_note = (
                f"Kỳ nguồn từ {cfg.source_name}: '{reporting_period}'. "
                f"CẢNH BÁO: Xác nhận của người dùng trước đây gắn với kỳ '{APPROVED_USER_SCOPE['reporting_period']}' "
                f"và không tự động áp dụng cho kỳ hiện tại '{reporting_period}'. "
                f"Cần người dùng rà soát và xác nhận lại các mã lệch ĐVT và mã thiếu cho kỳ mới."
            )
        elif not has_code_430200173:
            reporting_period_note = (
                f"Kỳ nguồn từ {cfg.source_name}: '{reporting_period}'. "
                f"CẢNH BÁO: Mã 430200173 không có trong báo cáo nguồn (không áp dụng xác nhận chép ĐVT Cái). "
                f"Bảo toàn {len(rec.missing_in_source)} mã thiếu theo file cũ và chừa trống cột C cho {len(missing_target_units)} dòng thiếu ĐVT đích. "
                f"Cần người dùng rà soát và xác nhận lại."
            )
        elif not is_unit_430200173_approved:
            actual_u = source_units.get("430200173", "").strip()
            reporting_period_note = (
                f"Kỳ nguồn từ {cfg.source_name}: '{reporting_period}'. "
                f"CẢNH BÁO: Mã 430200173 có ĐVT nguồn là '{actual_u}' (thay đổi so với ĐVT 'CAI' đã chốt). "
                f"Không tự áp dụng xác nhận 'chép trực tiếp Cái' (số tồn {qty_430200173_str}); cần người dùng xác nhận lại cho kỳ này."
            )
        else:
            reporting_period_note = (
                f"Kỳ nguồn từ {cfg.source_name}: '{reporting_period}'. "
                f"Đã được người dùng xác nhận và CHỐT chính thức cho kỳ này: sử dụng số tồn theo báo cáo nguồn, "
                f"ĐVT Cái (chép trực tiếp {qty_430200173_str} Cái cho mã 430200173), bảo toàn {len(rec.missing_in_source)} mã thiếu theo file cũ, "
                f"và chừa trống cột C cho {len(missing_target_units)} dòng thiếu ĐVT đích."
            )

        audit_summary = {
            "status": "success",
            "config_file": str(config_path),
            "reporting_period": reporting_period,
            "reporting_period_note": reporting_period_note,
            "sharepoint_target_path": cfg.target_path,
            "sharepoint_source_path": cfg.source_path,
            "source_item_id": source_item.get("id"),
            "target_item_id": target_item.get("id"),
            "source_etag": source_item.get("eTag"),
            "target_etag": target_item.get("eTag"),
            "target_sheet": cfg.target_sheet,
            "headers": header_vals[:8],
            "total_source_codes": len(source_stock),
            "total_target_rows": len(rec.target_codes),
            "metrics": {
                "matched_count": len(rec.changes) + len(rec.unchanged),
                "changed_count": len(rec.changes),
                "unchanged_count": len(rec.unchanged),
                "missing_in_source_count": len(rec.missing_in_source),
                "source_only_count": len(rec.source_only),
                "duplicate_codes_source": 0,
                "duplicate_codes_target": 0,
            },
            "unit_reconciliation": {
                "total_matched": len(unit_comparisons),
                "exact_matches_count": len(exact_matches),
                "alias_matches_count": len(alias_matches),
                "missing_target_units_count": len(missing_target_units),
                "divergent_units_count": len(divergent_units),
                "unit_mismatches_count": len(divergent_units) + len(missing_target_units),
                "policy_note": (
                    "Chép trực tiếp số tồn từ nguồn sang đích; không tự ý quy đổi đơn vị tính, "
                    "không sửa đổi cột C (ĐVT đích), và không biến mã thiếu thành 0."
                ),
                "alias_mapping_used": EQUIVALENT_UNIT_ALIASES,
                "divergent_units": divergent_units,
                "missing_target_units": missing_target_units,
                "alias_matches": alias_matches,
                "exact_matches": exact_matches,
                "unit_mismatches": divergent_units + missing_target_units,
            },
            "missing_in_source_items": missing_items_detail,
            "sample_reconciliation_10_plus": comparison_rows,
        }

        audit_json = json.dumps(audit_summary, ensure_ascii=False, indent=2) + "\n"
        (out_dir / "audit_summary.json").write_text(audit_json, encoding="utf-8")

        print(f"[SURVEY] Hoàn tất khảo sát thành công. Đã ghi audit_summary.json")
        return audit_summary

    except Exception as exc:
        print(f"[SURVEY] LỖI tại pha '{current_phase}': {exc}", file=sys.stderr)

        # 1. BẮT BUỘC loại bỏ proposal cũ để không lưu sót kết quả không hợp lệ
        try:
            (out_dir / "nvl_stock_proposal.xlsx").unlink(missing_ok=True)
        except OSError:
            pass

        # 2. Ghi report failed của lần chạy hiện tại
        error_report = {
            "schema_version": 1,
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "mode": "failed",
            "status": "failed",
            "phase": current_phase,
            "error_type": type(exc).__name__,
            "error_message": str(exc),
            "source": {
                "name": cfg.source_name if cfg else "",
                "sharepoint_path": cfg.source_path if cfg else "",
                "sourcedoc": cfg.source_sourcedoc if cfg else "",
            },
            "target": {
                "name": cfg.target_name if cfg else "",
                "sharepoint_path": cfg.target_path if cfg else "",
                "sourcedoc": cfg.target_sourcedoc if cfg else "",
            },
            "metrics": {
                "matched_count": 0,
                "changed_count": 0,
                "unchanged_count": 0,
                "missing_in_source_count": 0,
                "source_only_count": 0,
            },
        }
        try:
            (out_dir / "nvl_stock_report.json").write_text(
                json.dumps(error_report, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        except Exception:
            pass

        # 3. Ghi audit_summary.json failed
        error_summary: dict[str, Any] = {
            "status": "failed",
            "phase": current_phase,
            "error_type": type(exc).__name__,
            "error_message": str(exc),
            "config_file": str(config_path),
            "sharepoint_target_path": cfg.target_path if cfg else "",
            "sharepoint_source_path": cfg.source_path if cfg else "",
        }

        # Nếu lỗi tại pha fetch_target, kiểm tra và liệt kê danh sách file thực tế trong thư mục cha
        if current_phase == "fetch_target" and client_graph and drive_id and cfg and cfg.target_path:
            parent_dir = str(Path(cfg.target_path).parent).replace("\\", "/")
            if parent_dir and parent_dir != ".":
                try:
                    print(f"[SURVEY] Đang liệt kê danh sách file trong thư mục cha '{parent_dir}' trên SharePoint...")
                    children = client_graph.list_folder_children(drive_id, parent_dir)
                    files_in_parent = []
                    for item in children:
                        # Bỏ qua nếu là thư mục con (folder)
                        if "folder" in item:
                            continue
                        fname = item.get("name", "")
                        fid = item.get("id", "")
                        fetag = item.get("eTag", "")
                        files_in_parent.append({
                            "name": fname,
                            "id": fid,
                            "eTag": fetag,
                            "size": item.get("size"),
                            "lastModifiedDateTime": item.get("lastModifiedDateTime"),
                        })
                        print(f"  - [File] '{fname}' (ID: {fid}, eTag: {fetag})")
                    error_summary["parent_directory"] = parent_dir
                    error_summary["available_files_in_parent"] = files_in_parent
                except Exception as list_exc:
                    print(f"[SURVEY] Cảnh báo: Không thể liệt kê file trong '{parent_dir}': {list_exc}", file=sys.stderr)

        err_json = json.dumps(error_summary, ensure_ascii=False, indent=2) + "\n"
        try:
            (out_dir / "audit_summary.json").write_text(err_json, encoding="utf-8")
        except Exception:
            pass
        raise


def main():
    args = parse_args()
    try:
        run_survey(
            config_path=args.config,
            source_file=args.source_file,
            target_file=args.target_file,
            out_dir_path=args.out_dir,
        )
        sys.exit(0)
    except Exception as exc:
        print(f"[SURVEY] Thất bại: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
