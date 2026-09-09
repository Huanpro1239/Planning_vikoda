"""Module đồng bộ số lượng tồn kho nguyên vật liệu (NVL).

Nguồn: XNT_ketoan_Vikoda.xlsm, Sheet1
  - Cột B: Mã vật tư
  - Cột M: Số tồn kho

Đích: Kế hoạch mua hàng.xlsx, sheet Ton_NVL
  - Cột A: Mã vật tư
  - Cột D: Số tồn kho

Quy tắc:
  - Ghép mã vật tư Ton_NVL!A với Sheet1!B (chuẩn hóa trim khoảng trắng, tương đương float/int, giữ số 0 ở đầu).
  - Gán Ton_NVL!D = Sheet1!M: chép trực tiếp số tồn, không chia quy cách, không trừ tồn nhà máy, không đổi tiền tố mã 2 thành 1.
  - Patch trực tiếp XML của sheet Ton_NVL trong file ZIP, bảo toàn 100% macro, định dạng, merge, drawing và các sheet khác.
  - Tiến trình độc lập với pipeline sản xuất; chỉ tải lên SharePoint khi cờ --publish được bật.
"""

from __future__ import annotations

import argparse
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
import hashlib
from io import BytesIO
import json
import math
import os
from pathlib import Path
import posixpath
import re
import sys
import time
from typing import Any
import zipfile

from lxml import etree
from openpyxl import load_workbook

from graph_retry import install_retry_after_support, retry_delay_seconds
import sync_stock
from sync_stock import (
    GraphClient,
    GraphRequestError,
    clean_number,
    column_number,
    find_sheet_xml_path,
    get_access_token,
    is_retryable_graph_error,
    load_shared_strings,
    set_numeric_cell,
)

DEFAULT_CONFIG_FILE = "nvl_stock_config.json"
EPS = 1e-6


@dataclass
class NVLConfig:
    source_name: str
    source_path: str
    source_sheet: str
    source_code_col: int
    source_code_col_letter: str
    source_value_col: int
    source_value_col_letter: str
    source_start_row: int
    source_sourcedoc: str

    target_name: str
    target_path: str
    target_sheet: str
    target_code_col: int
    target_code_col_letter: str
    target_value_col: int
    target_value_col_letter: str
    target_start_row: int
    target_sourcedoc: str

    direct_copy: bool = True
    allow_zero: bool = True
    allow_negative: bool = True
    preserve_missing_in_source: bool = True
    reject_duplicates: bool = True
    reject_formula_in_target_cell: bool = True


@dataclass
class NVLReconcileResult:
    changes: list[dict[str, Any]] = field(default_factory=list)
    unchanged: list[dict[str, Any]] = field(default_factory=list)
    missing_in_source: list[dict[str, Any]] = field(default_factory=list)
    source_only: list[str] = field(default_factory=list)
    target_codes: dict[str, int] = field(default_factory=dict)
    source_codes: dict[str, int] = field(default_factory=dict)
    status: str = "success"
    message: str = ""


def load_nvl_config(config_path: str | Path | None = None) -> NVLConfig:
    """Đọc cấu hình đồng bộ NVL từ file JSON."""
    if config_path is None:
        config_path = Path(__file__).parent / DEFAULT_CONFIG_FILE
    else:
        config_path = Path(config_path)

    if not config_path.exists():
        raise FileNotFoundError(f"Không tìm thấy file cấu hình: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    src = data.get("source", {})
    tgt = data.get("target", {})
    pol = data.get("policies", {})

    return NVLConfig(
        source_name=src.get("name", "XNT_ketoan_Vikoda.xlsm"),
        source_path=src.get("sharepoint_path", ""),
        source_sheet=src.get("sheet_name", "Sheet1"),
        source_code_col=int(src.get("code_column", 2)),
        source_code_col_letter=src.get("code_column_letter", "B"),
        source_value_col=int(src.get("value_column", 13)),
        source_value_col_letter=src.get("value_column_letter", "M"),
        source_start_row=int(src.get("start_row", 2)),
        source_sourcedoc=src.get("sourcedoc", ""),
        target_name=tgt.get("name", "Kế hoạch mua hàng.xlsx"),
        target_path=tgt.get("sharepoint_path", ""),
        target_sheet=tgt.get("sheet_name", "Ton_NVL"),
        target_code_col=int(tgt.get("code_column", 1)),
        target_code_col_letter=tgt.get("code_column_letter", "A"),
        target_value_col=int(tgt.get("value_column", 4)),
        target_value_col_letter=tgt.get("value_column_letter", "D"),
        target_start_row=int(tgt.get("start_row", 2)),
        target_sourcedoc=tgt.get("sourcedoc", ""),
        direct_copy=bool(pol.get("direct_copy", True)),
        allow_zero=bool(pol.get("allow_zero", True)),
        allow_negative=bool(pol.get("allow_negative", True)),
        preserve_missing_in_source=bool(pol.get("preserve_missing_in_source", True)),
        reject_duplicates=bool(pol.get("reject_duplicates", True)),
        reject_formula_in_target_cell=bool(pol.get("reject_formula_in_target_cell", True)),
    )


def normalize_nvl_code(value: Any) -> str | None:
    """Chuẩn hóa mã vật tư NVL:

    - None, chuỗi rỗng hoặc boolean -> None
    - float với phần thập phân nguyên (ví dụ 12345.0) -> '12345'
    - int (ví dụ 12345) -> '12345'
    - str -> trim khoảng trắng đầu/cuối, giữ nguyên số 0 ở đầu (ví dụ '012345'),
      giữ nguyên chữ hoa/thường và khoảng trắng ở giữa.
    """
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        if value.is_integer():
            return str(int(value))
        return str(value).strip()

    if isinstance(value, int):
        return str(value)

    text = str(value).strip()
    if not text:
        return None

    # Nếu chuỗi biểu diễn số thực nguyên (ví dụ '12345.0')
    if re.fullmatch(r"-?\d+\.0+", text):
        return text.split(".")[0]

    return text


def parse_nvl_quantity(value: Any, *, cell_name: str = "") -> float:
    """Phân tích số lượng tồn kho NVL:

    - Cho phép số 0, số âm, số thập phân.
    - Không làm tròn thành số nguyên.
    - None, chuỗi rỗng, boolean, lỗi Excel, NaN, Infinity -> raise ValueError.
    """
    if value is None or (isinstance(value, str) and value.strip() == ""):
        raise ValueError(f"[{cell_name}] Ô số lượng tồn kho rỗng.")

    if isinstance(value, bool):
        raise ValueError(f"[{cell_name}] Chứa giá trị boolean ({value}), không phải số.")

    if isinstance(value, (int, float)):
        if math.isnan(value) or math.isinf(value):
            raise ValueError(f"[{cell_name}] Chứa giá trị NaN hoặc vô cực.")
        return float(value)

    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        try:
            val = float(cleaned)
            if math.isnan(val) or math.isinf(val):
                raise ValueError(f"[{cell_name}] Chứa giá trị NaN hoặc vô cực.")
            return val
        except ValueError as exc:
            raise ValueError(
                f"[{cell_name}] Giá trị {value!r} không thể chuyển đổi thành số tồn hợp lệ."
            ) from exc

    raise ValueError(f"[{cell_name}] Kiểu dữ liệu không hợp lệ: {type(value)}")


def read_nvl_source_stock(
    source_bytes: bytes,
    config: NVLConfig,
) -> tuple[dict[str, float], dict[str, Any]]:
    """Đọc dữ liệu số tồn kho từ file nguồn XNT_ketoan_Vikoda.xlsm!Sheet1."""
    wb_val = load_workbook(BytesIO(source_bytes), data_only=True, read_only=True, keep_vba=True)
    wb_raw = load_workbook(BytesIO(source_bytes), data_only=False, read_only=True, keep_vba=True)

    try:
        if config.source_sheet not in wb_val.sheetnames:
            raise RuntimeError(
                f"[Nguồn: {config.source_name}] Không tìm thấy sheet '{config.source_sheet}'."
            )

        ws_val = wb_val[config.source_sheet]
        ws_raw = wb_raw[config.source_sheet]

        stock: dict[str, float] = {}
        seen_rows: dict[str, int] = {}
        total_scanned = 0

        for r in range(config.source_start_row, ws_val.max_row + 1):
            total_scanned += 1
            raw_code = ws_val.cell(row=r, column=config.source_code_col).value
            if isinstance(raw_code, str):
                lower_code = raw_code.strip().lower()
                if any(k in lower_code for k in ("tổng cộng", "tong cong", "total")):
                    continue

            code = normalize_nvl_code(raw_code)
            if not code:
                continue

            if config.reject_duplicates and code in stock:
                raise RuntimeError(
                    f"[Nguồn: {config.source_name}] Mã vật tư '{code}' bị lặp tại dòng {r} "
                    f"(đã xuất hiện tại dòng {seen_rows[code]})."
                )

            seen_rows[code] = r
            cell_ref = f"{config.source_sheet}!{config.source_value_col_letter}{r}"

            c_val = ws_val.cell(row=r, column=config.source_value_col).value
            c_raw = ws_raw.cell(row=r, column=config.source_value_col).value

            if isinstance(c_raw, str) and c_raw.startswith("="):
                if c_val is None:
                    raise RuntimeError(
                        f"[Nguồn: {config.source_name}] Ô {cell_ref} chứa công thức '{c_raw}' "
                        f"nhưng không có giá trị cached (chưa được Excel tính toán và lưu)."
                    )

            qty = parse_nvl_quantity(c_val, cell_name=cell_ref)
            stock[code] = qty

        if not stock:
            raise RuntimeError(
                f"[Nguồn: {config.source_name}] Không đọc được bất kỳ mã vật tư hợp lệ nào "
                f"trong {config.source_sheet}!{config.source_code_col_letter}."
            )

        metadata = {
            "source_name": config.source_name,
            "sheet_name": config.source_sheet,
            "total_scanned_rows": total_scanned,
            "valid_codes_count": len(stock),
            "seen_rows": seen_rows,
        }
        return stock, metadata
    finally:
        wb_val.close()
        wb_raw.close()


def reconcile_nvl_target(
    target_bytes: bytes,
    source_stock: dict[str, float],
    config: NVLConfig,
) -> NVLReconcileResult:
    """Đối chiếu mã vật tư và lập danh sách thay đổi cho Kế hoạch mua hàng.xlsx!Ton_NVL."""
    wb_val = load_workbook(BytesIO(target_bytes), data_only=True, read_only=True)
    wb_raw = load_workbook(BytesIO(target_bytes), data_only=False, read_only=True)

    try:
        if config.target_sheet not in wb_val.sheetnames:
            raise RuntimeError(
                f"[Đích: {config.target_name}] Không tìm thấy sheet '{config.target_sheet}'."
            )

        ws_val = wb_val[config.target_sheet]
        ws_raw = wb_raw[config.target_sheet]

        seen_target_codes: dict[str, int] = {}
        changes: list[dict[str, Any]] = []
        unchanged: list[dict[str, Any]] = []
        missing_in_source: list[dict[str, Any]] = []

        for r in range(config.target_start_row, ws_val.max_row + 1):
            raw_code = ws_val.cell(row=r, column=config.target_code_col).value
            if isinstance(raw_code, str):
                lower_code = raw_code.strip().lower()
                if any(k in lower_code for k in ("tổng cộng", "tong cong", "total")):
                    continue

            code = normalize_nvl_code(raw_code)
            if not code:
                continue

            if config.reject_duplicates and code in seen_target_codes:
                raise RuntimeError(
                    f"[Đích: {config.target_name}] Mã vật tư '{code}' bị lặp trong "
                    f"{config.target_sheet}!{config.target_code_col_letter} tại dòng {r} "
                    f"(đã xuất hiện tại dòng {seen_target_codes[code]})."
                )

            seen_target_codes[code] = r
            cell_ref = f"{config.target_sheet}!{config.target_value_col_letter}{r}"

            target_raw = ws_raw.cell(row=r, column=config.target_value_col).value
            target_val = ws_val.cell(row=r, column=config.target_value_col).value

            if config.reject_formula_in_target_cell and isinstance(target_raw, str) and target_raw.startswith("="):
                raise RuntimeError(
                    f"[Đích: {config.target_name}] Ô {cell_ref} chứa công thức '{target_raw}'. "
                    f"Cột {config.target_value_col_letter} phải là cột nhập tồn, không được ghi đè lên công thức người dùng."
                )

            if code not in source_stock:
                missing_in_source.append({
                    "row": r,
                    "code": code,
                    "current_value": target_val,
                })
                continue

            new_qty = source_stock[code]
            needs_change = False

            if target_val is None or (isinstance(target_val, str) and target_val.strip() == ""):
                needs_change = True
            else:
                try:
                    curr_float = float(target_val)
                    if abs(curr_float - new_qty) > EPS:
                        needs_change = True
                except Exception:
                    needs_change = True

            if needs_change:
                changes.append({
                    "row": r,
                    "code": code,
                    "before": target_val,
                    "after": new_qty,
                })
            else:
                unchanged.append({
                    "row": r,
                    "code": code,
                    "value": target_val,
                })

        if not seen_target_codes:
            raise RuntimeError(
                f"[Đích: {config.target_name}] Không đọc được bất kỳ mã vật tư hợp lệ nào trong "
                f"{config.target_sheet}!{config.target_code_col_letter}."
            )

        matched_count = len(changes) + len(unchanged)
        if matched_count == 0:
            raise RuntimeError(
                f"[Đích: {config.target_name}] Không có mã vật tư nào trong "
                f"{config.target_sheet}!{config.target_code_col_letter} khớp với nguồn {config.source_name}. Bị chặn cập nhật."
            )

        source_only = [c for c in source_stock if c not in seen_target_codes]

        if missing_in_source:
            status = "completed_with_warnings"
            msg = f"Đồng bộ hoàn tất nhưng còn {len(missing_in_source)} mã đích không có trong nguồn."
        elif changes:
            status = "success"
            msg = f"Đồng bộ thành công: {len(changes)} ô cần cập nhật, {len(unchanged)} ô không đổi."
        else:
            status = "unchanged"
            msg = "Dữ liệu tồn kho khớp hoàn toàn, không có ô nào cần thay đổi."

        return NVLReconcileResult(
            changes=changes,
            unchanged=unchanged,
            missing_in_source=missing_in_source,
            source_only=source_only,
            target_codes=seen_target_codes,
            status=status,
            message=msg,
        )
    finally:
        wb_val.close()
        wb_raw.close()


def patch_nvl_destination_workbook(
    target_bytes: bytes,
    reconcile_result: NVLReconcileResult,
    config: NVLConfig,
) -> bytes:
    """Cập nhật các ô số tồn kho vào file Excel đích ở cấp độ ZIP XML."""
    if not reconcile_result.changes:
        return target_bytes

    source_buffer = BytesIO(target_bytes)
    output_buffer = BytesIO()

    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        sheet_path = find_sheet_xml_path(source_zip, config.target_sheet)
        sheet_root = etree.fromstring(source_zip.read(sheet_path))

        sheet_data_nodes = sheet_root.xpath('//*[local-name()="sheetData"]')
        if not sheet_data_nodes:
            raise RuntimeError(f"Không tìm thấy thẻ sheetData trong {sheet_path}.")
        sheet_data = sheet_data_nodes[0]

        row_map: dict[int, Any] = {}
        for r_elem in sheet_data.xpath('./*[local-name()="row"]'):
            r_num_text = r_elem.get("r")
            if r_num_text and r_num_text.isdigit():
                row_map[int(r_num_text)] = r_elem

        changes_by_row = {item["row"]: item["after"] for item in reconcile_result.changes}

        for row_num, new_val in changes_by_row.items():
            row_element = row_map.get(row_num)
            if row_element is None:
                namespace = etree.QName(sheet_data).namespace
                row_element = etree.Element(f"{{{namespace}}}row", r=str(row_num))
                inserted = False
                for existing_row in sheet_data.xpath('./*[local-name()="row"]'):
                    er = existing_row.get("r")
                    if er and int(er) > row_num:
                        existing_row.addprevious(row_element)
                        inserted = True
                        break
                if not inserted:
                    sheet_data.append(row_element)
                row_map[row_num] = row_element

            set_numeric_cell(
                row_element,
                row_num,
                config.target_value_col_letter,
                new_val,
            )

        new_sheet_xml = etree.tostring(
            sheet_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        with zipfile.ZipFile(output_buffer, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            for item in source_zip.infolist():
                data = new_sheet_xml if item.filename == sheet_path else source_zip.read(item.filename)
                output_zip.writestr(item, data)

    return output_buffer.getvalue()


def verify_nvl_patched_workbook(
    original_bytes: bytes,
    patched_bytes: bytes,
    reconcile_result: NVLReconcileResult,
    config: NVLConfig,
) -> dict[str, Any]:
    """Kiểm tra tính toàn vẹn của workbook sau khi patch XML:

    1. Tất cả các ô changed_cells đều nhận đúng giá trị mới.
    2. Các ô unchanged và missing_in_source giữ nguyên giá trị.
    3. Tất cả các phần tử trong file ZIP ngoài sheet đích đều giống nhau từng byte.
    """
    if not reconcile_result.changes:
        return {"ok": True, "message": "Không có thay đổi cần xác minh."}

    # 1. So sánh các ZIP parts không liên quan
    with zipfile.ZipFile(BytesIO(original_bytes), "r") as z_orig, zipfile.ZipFile(BytesIO(patched_bytes), "r") as z_patch:
        target_sheet_path = find_sheet_xml_path(z_orig, config.target_sheet)
        for name in z_orig.namelist():
            if name != target_sheet_path:
                orig_hash = hashlib.sha256(z_orig.read(name)).hexdigest()
                patch_hash = hashlib.sha256(z_patch.read(name)).hexdigest()
                if orig_hash != patch_hash:
                    raise RuntimeError(
                        f"Phần tử không liên quan {name} trong file ZIP bị thay đổi ngoài ý muốn!"
                    )

    # 2. Đọc lại workbook đích bằng openpyxl để xác minh dữ liệu
    wb = load_workbook(BytesIO(patched_bytes), data_only=True)
    try:
        ws = wb[config.target_sheet]

        for item in reconcile_result.changes:
            row = item["row"]
            expected = item["after"]
            actual = ws.cell(row=row, column=config.target_value_col).value
            if actual is None or abs(float(actual) - float(expected)) > EPS:
                raise RuntimeError(
                    f"Xác minh thất bại tại dòng {row}: kỳ vọng {expected}, thực tế {actual}."
                )

        for item in reconcile_result.missing_in_source:
            row = item["row"]
            expected = item["current_value"]
            actual = ws.cell(row=row, column=config.target_value_col).value
            if expected is None:
                if actual is not None and str(actual).strip() != "":
                    raise RuntimeError(
                        f"Ô giữ nguyên tại dòng {row} bị thay đổi từ rỗng thành {actual}."
                    )
            else:
                if abs(float(actual) - float(expected)) > EPS:
                    raise RuntimeError(
                        f"Ô giữ nguyên tại dòng {row} bị thay đổi: cũ {expected}, mới {actual}."
                    )
    finally:
        wb.close()

    return {"ok": True, "message": "Xác minh toàn vẹn thành công 100%."}


def generate_nvl_report(
    reconcile_result: NVLReconcileResult,
    config: NVLConfig,
    *,
    mode: str,
    source_revision: str | None = None,
    target_revision: str | None = None,
) -> dict[str, Any]:
    """Tạo báo cáo JSON đối soát chi tiết đồng bộ tồn NVL."""
    return {
        "schema_version": 1,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "mode": mode,
        "status": reconcile_result.status,
        "message": reconcile_result.message,
        "source": {
            "name": config.source_name,
            "sharepoint_path": config.source_path,
            "sourcedoc": config.source_sourcedoc,
            "sheet_name": config.source_sheet,
            "code_column": config.source_code_col_letter,
            "value_column": config.source_value_col_letter,
            "revision": source_revision,
        },
        "target": {
            "name": config.target_name,
            "sharepoint_path": config.target_path,
            "sourcedoc": config.target_sourcedoc,
            "sheet_name": config.target_sheet,
            "code_column": config.target_code_col_letter,
            "value_column": config.target_value_col_letter,
            "revision": target_revision,
        },
        "metrics": {
            "matched_count": len(reconcile_result.changes) + len(reconcile_result.unchanged),
            "changed_count": len(reconcile_result.changes),
            "unchanged_count": len(reconcile_result.unchanged),
            "missing_in_source_count": len(reconcile_result.missing_in_source),
            "source_only_count": len(reconcile_result.source_only),
        },
        "changes": reconcile_result.changes,
        "warnings": [
            {
                "type": "missing_in_source",
                "row": item["row"],
                "code": item["code"],
                "message": f"Mã vật tư '{item['code']}' ở đích không có trong file nguồn; giữ nguyên số tồn cũ ({item['current_value']}).",
            }
            for item in reconcile_result.missing_in_source
        ],
        "source_only_codes": reconcile_result.source_only,
    }


def run_nvl_sync(
    config: NVLConfig,
    *,
    source_file: str | Path | None = None,
    target_file: str | Path | None = None,
    out_dir: str | Path | None = None,
    publish: bool = False,
    graph: GraphClient | None = None,
    max_publish_attempts: int = 3,
) -> dict[str, Any]:
    """Chạy quy trình đồng bộ tồn NVL ở chế độ offline hoặc online."""
    install_retry_after_support()

    if (source_file or target_file) and publish:
        raise ValueError("Không được kết hợp cờ --publish khi chạy với file cục bộ.")

    if out_dir is None:
        out_dir = Path("offline_out/nvl") if (source_file or target_file) else Path(".")
    else:
        out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    proposal_path = out_dir / "nvl_stock_proposal.xlsx"
    report_path = out_dir / "nvl_stock_report.json"

    # 1. Chế độ OFFLINE (chạy từ file cục bộ)
    if source_file or target_file:
        if not (source_file and target_file):
            raise ValueError("Phải cung cấp đồng thời cả --source-file và --target-file.")
        src_path = Path(source_file)
        tgt_path = Path(target_file)
        if not src_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file nguồn: {src_path}")
        if not tgt_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file đích: {tgt_path}")

        print(f"[OFFLINE] Đọc nguồn: {src_path}")
        source_bytes = src_path.read_bytes()
        print(f"[OFFLINE] Đọc đích: {tgt_path}")
        target_bytes = tgt_path.read_bytes()

        source_stock, _ = read_nvl_source_stock(source_bytes, config)
        print(f"[OFFLINE] Đọc được {len(source_stock)} mã vật tư từ nguồn.")

        reconcile_res = reconcile_nvl_target(target_bytes, source_stock, config)
        print(
            f"[OFFLINE] Đối soát đích: {len(reconcile_res.changes)} ô đổi, "
            f"{len(reconcile_res.unchanged)} ô không đổi, "
            f"{len(reconcile_res.missing_in_source)} mã thiếu ở nguồn."
        )

        patched_bytes = patch_nvl_destination_workbook(target_bytes, reconcile_res, config)
        verify_nvl_patched_workbook(target_bytes, patched_bytes, reconcile_res, config)

        proposal_path.write_bytes(patched_bytes)
        print(f"[OFFLINE] Đã ghi proposal: {proposal_path}")

        report = generate_nvl_report(reconcile_res, config, mode="offline")
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"[OFFLINE] Đã ghi báo cáo: {report_path}")
        return report

    # 2. Chế độ ONLINE (kết nối Microsoft Graph)
    if graph is None:
        token = get_access_token()
        graph = GraphClient(token)

    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    if not config.source_path:
        raise ValueError("Chưa cấu hình 'source.sharepoint_path' trong file cấu hình.")
    if not config.target_path:
        raise ValueError(
            f"Chưa cấu hình đường dẫn 'target.sharepoint_path' cho '{config.target_name}' trên SharePoint.\n"
            f"(Sourcedoc ID đã biết: {config.target_sourcedoc}). Vui lòng điền đường dẫn thư mục chính xác."
        )

    for attempt in range(1, max_publish_attempts + 1):
        print(f"[ONLINE] Lượt {attempt}/{max_publish_attempts}: Đọc metadata từ SharePoint...")
        source_item = graph.get_item_by_path(drive_id, config.source_path)
        target_item = graph.get_item_by_path(drive_id, config.target_path)

        source_bytes = graph.download_file(drive_id, source_item["id"])
        target_bytes = graph.download_file(drive_id, target_item["id"])

        source_stock, _ = read_nvl_source_stock(source_bytes, config)
        reconcile_res = reconcile_nvl_target(target_bytes, source_stock, config)

        patched_bytes = patch_nvl_destination_workbook(target_bytes, reconcile_res, config)
        verify_nvl_patched_workbook(target_bytes, patched_bytes, reconcile_res, config)

        proposal_path.write_bytes(patched_bytes)
        report = generate_nvl_report(
            reconcile_res,
            config,
            mode="publish" if publish else "dry_run",
            source_revision=source_item.get("eTag"),
            target_revision=target_item.get("eTag"),
        )
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

        if not publish:
            print(f"[ONLINE] Dry-run hoàn tất. Proposal: {proposal_path}, Report: {report_path}")
            return report

        # Publish mode
        if not reconcile_res.changes:
            print("[ONLINE] Dữ liệu khớp 100%, không có ô nào cần upload.")
            return report

        print(f"[ONLINE] Đang upload file đích với ETag {target_item['eTag']}...")
        try:
            res = graph.upload_file(drive_id, target_item["id"], patched_bytes, expected_etag=target_item["eTag"])
            print(f"[ONLINE] Upload thành công: {res.get('name')} (Lượt {attempt})")
            report["upload_result"] = res
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            return report
        except GraphRequestError as exc:
            if exc.status_code == 412:
                print("[ONLINE] HTTP 412: File đích đã thay đổi đồng thời; đang lấy snapshot mới để tính lại...")
                time.sleep(1)
                continue
            if is_retryable_graph_error(exc) and attempt < max_publish_attempts:
                delay = retry_delay_seconds(exc, 3.0)
                print(f"[ONLINE] Lỗi tạm thời: {exc}. Chờ {delay}s rồi thử lại...")
                time.sleep(delay)
                continue
            raise

    raise RuntimeError("Vượt quá số lần thử tải lên SharePoint do xung đột liên tục.")


def main():
    parser = argparse.ArgumentParser(description="Đồng bộ tồn kho nguyên vật liệu (NVL) từ XNT_ketoan_Vikoda sang Kế hoạch mua hàng.")
    parser.add_argument("--config", default=DEFAULT_CONFIG_FILE, help="Đường dẫn file cấu hình JSON.")
    parser.add_argument("--source-file", help="File nguồn cục bộ (chế độ offline).")
    parser.add_argument("--target-file", help="File đích cục bộ (chế độ offline).")
    parser.add_argument("--target-path", help="Ghi đè đường dẫn file đích trên SharePoint.")
    parser.add_argument("--out", help="Thư mục xuất proposal và report.")
    parser.add_argument("--publish", action="store_true", help="Publish lên SharePoint (mặc định là dry-run).")

    args = parser.parse_args()

    cfg = load_nvl_config(args.config)
    if args.target_path:
        cfg.target_path = args.target_path

    try:
        rep = run_nvl_sync(
            cfg,
            source_file=args.source_file,
            target_file=args.target_file,
            out_dir=args.out,
            publish=args.publish,
        )
        print(f"Hoàn thành ({rep.get('status')}): {rep.get('message')}")
    except Exception as exc:
        print(f"LỖI: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
