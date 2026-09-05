import calendar
import json
import math
import re
import zipfile
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path

from lxml import etree
from openpyxl import load_workbook

from sync_planning_fc import (
    FC_SELECTOR_CELL,
    FC_SHEET,
    PLANNING_SHEET,
    _find_sheet_xml_path,
    _get_or_create_cell,
    _load_shared_strings,
    _read_cell_text,
)
from sync_stock import (
    DEST_PATH,
    SOURCE_ACTUAL_PATH,
    SOURCE_FACTORY_VIKODA_PATH,
    GraphClient,
    clean_number,
    get_access_token,
    normalize_code,
    read_conversion_factors,
    to_number,
)


RUNTIME_STATE_FILE = Path("planning_runtime.json")
RUNTIME_STATE_VERSION = 1

# Leadtime lấy từ logic DMSP của file kế hoạch cũ, nhưng được cấu hình
# trực tiếp trong code. Khi chạy production KHÔNG đọc sheet/file mẫu.
LEADTIME_BY_CODE = {
    "130100096": 3,
    "130100091": 3,
    "130100169": 3,
    "130200026": 1,
    "130100055": 4,
    "130100017": 3,
    "130100011": 4,
    "130100036": 4,
    "130100116": 3,
    "130100115": 3,
    "130100110": 3,
    "130100172": 3,
    "130100006": 3,
    "130100151": 3,
    "130100152": 3,
    "130100008": 3,
    "130100013": 4,
    "130100149": 3,
    "130300005": 4,
    "130300006": 4,
}

# Nguồn thô: Bao cao ton thuc te hien tai.xlsx
ACTUAL_CODE_COL = 3          # C - MÃ SP VIKODA
ACTUAL_RECEIPT_COL = 7       # G - TỔNG NHẬP THÁNG
ACTUAL_CONSIGNMENT_COL = 17  # Q - SL HÀNG ĐÃ BÁN (GỬI KHO)

# Nguồn thô: NXT_Vikoda.xlsm!Sheet1
SYSTEM_CODE_COL = 2          # B - Mã vật tư
SYSTEM_RECEIPT_COL = 9       # I - Nhập trong kỳ

# Ke_hoach_SX - dữ liệu đầu vào cùng dòng.
COL_BATCH = 4                # D
COL_PER_SHIFT = 5            # E
COL_CLASSIFICATION = 8       # H
COL_SHIFTS_PER_DAY = 9       # I
COL_ACTUAL_STOCK = 10        # J
COL_BOOK_STOCK = 11          # K
COL_FC = 12                  # L
COL_CURRENT_DEBT = 14        # N (chỉ dùng để bootstrap state lần đầu)

# Kết quả được tính và ghi TRỰC TIẾP tại Ke_hoach_SX!M:R.
OUTPUT_COLUMNS = {
    "M": "expected_end_stock",
    "N": "warehouse_debt",
    "O": "required_production",
    "P": "rounded_production",
    "Q": "production_days",
    "R": "production_start",
}


def load_runtime_state():
    if not RUNTIME_STATE_FILE.exists():
        return {
            "version": RUNTIME_STATE_VERSION,
            "opening_debt_by_month": {},
            "opening_consignment_by_month": {},
        }

    try:
        state = json.loads(RUNTIME_STATE_FILE.read_text(encoding="utf-8"))
    except Exception as exc:
        raise RuntimeError(
            f"Không đọc được {RUNTIME_STATE_FILE}: {exc}"
        ) from exc

    state.setdefault("version", RUNTIME_STATE_VERSION)
    state.setdefault("opening_debt_by_month", {})
    state.setdefault("opening_consignment_by_month", {})
    return state


def save_runtime_state(state):
    state["version"] = RUNTIME_STATE_VERSION
    RUNTIME_STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _month_key(year, month):
    return f"{year:04d}-{month:02d}"


def _next_month(year, month):
    if month == 12:
        return year + 1, 1
    return year, month + 1


def _is_month_end(value):
    return value.day == calendar.monthrange(value.year, value.month)[1]


def _find_report_date(worksheet):
    for row in range(1, min(worksheet.max_row, 6) + 1):
        for column in range(1, min(worksheet.max_column, 20) + 1):
            value = worksheet.cell(row=row, column=column).value
            if isinstance(value, datetime):
                return value.date()
            if isinstance(value, date):
                return value

    raise RuntimeError(
        "Không tìm thấy ngày báo cáo trong file tồn thực tế."
    )


def read_actual_inputs(source_bytes):
    workbook = load_workbook(
        BytesIO(source_bytes),
        data_only=True,
        read_only=True,
    )

    try:
        worksheet = workbook.worksheets[0]
        report_date = _find_report_date(worksheet)
        receipts = {}
        consignments = {}

        for row in range(1, worksheet.max_row + 1):
            code = normalize_code(
                worksheet.cell(row=row, column=ACTUAL_CODE_COL).value
            )
            if not code:
                continue

            if code in receipts:
                raise RuntimeError(
                    f"[Tồn thực tế] Mã {code} bị lặp trong cột C."
                )

            receipts[code] = clean_number(
                to_number(
                    worksheet.cell(row=row, column=ACTUAL_RECEIPT_COL).value,
                    f"G{row}",
                )
            )
            consignments[code] = clean_number(
                to_number(
                    worksheet.cell(row=row, column=ACTUAL_CONSIGNMENT_COL).value,
                    f"Q{row}",
                )
            )

        if not receipts:
            raise RuntimeError(
                "[Tồn thực tế] Không đọc được dữ liệu nhập tháng."
            )

        return report_date, receipts, consignments
    finally:
        workbook.close()


def read_system_receipts(source_bytes, conversion_factors, planning_codes):
    workbook = load_workbook(
        BytesIO(source_bytes),
        data_only=True,
        read_only=True,
        keep_vba=True,
    )

    try:
        if "Sheet1" not in workbook.sheetnames:
            raise RuntimeError("Không tìm thấy Sheet1 trong NXT_Vikoda.xlsm.")

        worksheet = workbook["Sheet1"]
        receipts = {}

        for row in range(1, worksheet.max_row + 1):
            code = normalize_code(
                worksheet.cell(row=row, column=SYSTEM_CODE_COL).value
            )
            if not code or code not in planning_codes:
                continue

            if code in receipts:
                raise RuntimeError(
                    f"[NXT_Vikoda] Mã {code} bị lặp trong cột B."
                )

            factor = conversion_factors.get(code)
            if factor is None or factor <= 0:
                raise RuntimeError(
                    f"Không có quy cách hợp lệ cho mã {code}."
                )

            raw_value = to_number(
                worksheet.cell(row=row, column=SYSTEM_RECEIPT_COL).value,
                f"I{row}",
            )
            receipts[code] = clean_number(raw_value / factor)

        # Mã không có trong NXT được xem là chưa nhập hệ thống.
        for code in planning_codes:
            receipts.setdefault(code, 0)

        return receipts
    finally:
        workbook.close()


def _parse_plan_month(selector):
    match = re.search(r"(\d{1,2})", str(selector or ""))
    if not match:
        raise RuntimeError(
            f"Không xác định được tháng từ {FC_SHEET}!{FC_SELECTOR_CELL}="
            f"{selector!r}."
        )

    month = int(match.group(1))
    if not 1 <= month <= 12:
        raise RuntimeError(f"Tháng kế hoạch không hợp lệ: {month}.")
    return month


def _resolve_plan_year(report_date, plan_month):
    year = report_date.year
    if report_date.month >= 10 and plan_month <= 3:
        year += 1
    elif report_date.month <= 3 and plan_month >= 10:
        year -= 1
    return year


def read_planning_rows(dest_bytes):
    workbook = load_workbook(
        BytesIO(dest_bytes),
        data_only=True,
        read_only=True,
    )

    try:
        if PLANNING_SHEET not in workbook.sheetnames:
            raise RuntimeError(
                f"Không tìm thấy sheet {PLANNING_SHEET!r}."
            )
        if FC_SHEET not in workbook.sheetnames:
            raise RuntimeError(f"Không tìm thấy sheet {FC_SHEET!r}.")

        worksheet = workbook[PLANNING_SHEET]
        selector = workbook[FC_SHEET][FC_SELECTOR_CELL].value
        plan_month = _parse_plan_month(selector)

        rows = {}
        for row in range(2, worksheet.max_row + 1):
            code = normalize_code(worksheet.cell(row=row, column=1).value)
            if not code:
                continue

            if code in rows:
                raise RuntimeError(
                    f"Mã {code} bị lặp trong {PLANNING_SHEET}!A."
                )

            rows[code] = {
                "row": row,
                "batch": to_number(
                    worksheet.cell(row=row, column=COL_BATCH).value,
                    f"{PLANNING_SHEET}!D{row}",
                ),
                "per_shift": to_number(
                    worksheet.cell(row=row, column=COL_PER_SHIFT).value,
                    f"{PLANNING_SHEET}!E{row}",
                ),
                "classification": str(
                    worksheet.cell(row=row, column=COL_CLASSIFICATION).value
                    or ""
                ).strip(),
                "shifts_per_day": to_number(
                    worksheet.cell(row=row, column=COL_SHIFTS_PER_DAY).value,
                    f"{PLANNING_SHEET}!I{row}",
                ),
                "actual_stock": to_number(
                    worksheet.cell(row=row, column=COL_ACTUAL_STOCK).value,
                    f"{PLANNING_SHEET}!J{row}",
                ),
                "book_stock": to_number(
                    worksheet.cell(row=row, column=COL_BOOK_STOCK).value,
                    f"{PLANNING_SHEET}!K{row}",
                ),
                "fc": to_number(
                    worksheet.cell(row=row, column=COL_FC).value,
                    f"{PLANNING_SHEET}!L{row}",
                ),
                "current_debt": to_number(
                    worksheet.cell(row=row, column=COL_CURRENT_DEBT).value,
                    f"{PLANNING_SHEET}!N{row}",
                ),
                "current_outputs": {
                    key: worksheet.cell(
                        row=row,
                        column=column,
                    ).value
                    for key, column in (
                        ("expected_end_stock", 13),
                        ("warehouse_debt", 14),
                        ("required_production", 15),
                        ("rounded_production", 16),
                        ("production_days", 17),
                        ("production_start", 18),
                    )
                },
            }

        if not rows:
            raise RuntimeError(
                f"Không đọc được dữ liệu trong {PLANNING_SHEET}."
            )

        return selector, plan_month, rows
    finally:
        workbook.close()


def excel_roundup_integer(value):
    """Mô phỏng Excel ROUNDUP(value, 0): làm tròn ra xa số 0."""
    if value > 0:
        return math.ceil(value)
    if value < 0:
        return math.floor(value)
    return 0


def calculate_row(
    *,
    fc,
    actual_stock,
    book_stock,
    opening_consignment,
    warehouse_debt,
    leadtime,
    batch,
    per_shift,
    shifts_per_day,
    classification,
    plan_year,
    plan_month,
):
    if batch <= 0:
        raise ValueError("Số lượng/mẻ phải > 0.")
    if per_shift <= 0:
        raise ValueError("Số lượng/ca phải > 0.")
    if shifts_per_day <= 0:
        raise ValueError("Số ca theo ngày phải > 0.")

    # M - Tồn cuối dự kiến.
    daily_fc = fc / 26 if fc else 0
    minimum_stock = daily_fc * (leadtime + 2)
    expected_end_stock = (
        0 if opening_consignment > minimum_stock else minimum_stock
    )

    # N - Nợ kho được truyền vào từ công thức nợ trực tiếp.

    # O - Số lượng cần sản xuất.
    if warehouse_debt > 0:
        required_production = fc + warehouse_debt - book_stock
    else:
        required_production = (
            fc + expected_end_stock - book_stock + warehouse_debt
        )

    # P - Làm tròn theo mẻ/ca giống Excel ROUNDUP.
    if required_production == 0:
        rounded_production = 0
    else:
        is_sugar = classification.casefold() == "có đường".casefold()
        base = batch if is_sugar else per_shift
        rounded_production = (
            excel_roundup_integer(required_production / base) * base
        )

    # Q - Số ngày cần sản xuất.
    production_days = rounded_production / per_shift / shifts_per_day

    # R - Ngày bắt đầu sản xuất.
    if daily_fc <= 0:
        production_start = None
    else:
        first_day = datetime(plan_year, plan_month, 1)
        calculated = first_day + timedelta(
            days=(actual_stock / daily_fc) - leadtime
        )
        production_start = max(first_day, calculated)

    return {
        "expected_end_stock": clean_number(expected_end_stock),
        "warehouse_debt": clean_number(warehouse_debt),
        "required_production": clean_number(required_production),
        "rounded_production": clean_number(rounded_production),
        "production_days": clean_number(production_days),
        "production_start": production_start,
    }


def _bootstrap_opening_debt(
    *, current_debt, actual_receipt, system_receipt
):
    # Migration lần đầu: suy ngược số đầu kỳ từ chính N hiện tại và
    # dữ liệu nguồn thật. Không lấy số từ sheet/file mẫu.
    return clean_number(current_debt + actual_receipt - system_receipt)


def calculate_metrics(
    *,
    report_date,
    plan_month,
    planning_rows,
    actual_receipts,
    system_receipts,
    current_consignments,
    state,
):
    source_key = _month_key(report_date.year, report_date.month)
    plan_year = _resolve_plan_year(report_date, plan_month)
    plan_key = _month_key(plan_year, plan_month)

    debt_state = state.setdefault("opening_debt_by_month", {})
    opening_debt = debt_state.get(source_key)
    state_changed = False

    if opening_debt is None:
        opening_debt = {}
        for code, row in planning_rows.items():
            opening_debt[code] = _bootstrap_opening_debt(
                current_debt=row["current_debt"],
                actual_receipt=actual_receipts.get(code, 0),
                system_receipt=system_receipts.get(code, 0),
            )
        debt_state[source_key] = opening_debt
        state_changed = True

    current_debt = {}
    for code in planning_rows:
        current_debt[code] = clean_number(
            max(
                float(opening_debt.get(code, 0) or 0)
                - float(actual_receipts.get(code, 0) or 0)
                + float(system_receipts.get(code, 0) or 0),
                0,
            )
        )

    consign_state = state.setdefault("opening_consignment_by_month", {})
    opening_consignment = consign_state.get(plan_key)
    if opening_consignment is None:
        # Khi lập tháng kế tiếp, Q của báo cáo cuối tháng chính là số gửi
        # kho đầu tháng kế hoạch. Nếu đang chạy giữa tháng, dùng snapshot
        # hiện tại làm baseline cho tới khi có chốt cuối tháng.
        opening_consignment = {
            code: clean_number(current_consignments.get(code, 0))
            for code in planning_rows
        }
        consign_state[plan_key] = opening_consignment
        state_changed = True

    if _is_month_end(report_date):
        next_year, next_month = _next_month(
            report_date.year, report_date.month
        )
        next_key = _month_key(next_year, next_month)

        next_debt = {
            code: clean_number(current_debt.get(code, 0))
            for code in planning_rows
        }
        next_consignment = {
            code: clean_number(current_consignments.get(code, 0))
            for code in planning_rows
        }

        if debt_state.get(next_key) != next_debt:
            debt_state[next_key] = next_debt
            state_changed = True
        if consign_state.get(next_key) != next_consignment:
            consign_state[next_key] = next_consignment
            state_changed = True

    metrics = {}
    for code, row in planning_rows.items():
        leadtime = LEADTIME_BY_CODE.get(code)
        if leadtime is None:
            raise RuntimeError(f"Chưa cấu hình Leadtime cho mã {code}.")

        metrics[code] = calculate_row(
            fc=row["fc"],
            actual_stock=row["actual_stock"],
            book_stock=row["book_stock"],
            opening_consignment=float(
                opening_consignment.get(code, 0) or 0
            ),
            warehouse_debt=float(current_debt.get(code, 0) or 0),
            leadtime=leadtime,
            batch=row["batch"],
            per_shift=row["per_shift"],
            shifts_per_day=row["shifts_per_day"],
            classification=row["classification"],
            plan_year=plan_year,
            plan_month=plan_month,
        )

    return metrics, state_changed, source_key, plan_key


def _datetime_to_excel_serial(value):
    if value is None:
        return None
    epoch = datetime(1899, 12, 30)
    return (value - epoch).total_seconds() / 86400


def _set_cell_value(row_element, row_number, column_letter, value):
    cell = _get_or_create_cell(row_element, row_number, column_letter)

    # Giữ nguyên style của cell; chỉ thay formula/value.
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)

    if value is None:
        return
    if isinstance(value, datetime):
        value = _datetime_to_excel_serial(value)

    namespace = etree.QName(cell).namespace
    value_node = etree.SubElement(cell, f"{{{namespace}}}v")
    value = clean_number(value)
    value_node.text = (
        str(value) if isinstance(value, int) else format(float(value), ".15g")
    )


def patch_workbook(workbook_bytes, metrics):
    source_buffer = BytesIO(workbook_bytes)
    output_buffer = BytesIO()

    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        sheet_path = _find_sheet_xml_path(source_zip, PLANNING_SHEET)
        shared_strings = _load_shared_strings(source_zip)
        sheet_root = etree.fromstring(source_zip.read(sheet_path))

        seen = set()
        patched = 0
        rows = sheet_root.xpath(
            '//*[local-name()="sheetData"]/*[local-name()="row"]'
        )

        for row_element in rows:
            row_number_text = row_element.get("r")
            if not row_number_text:
                continue
            row_number = int(row_number_text)

            code = None
            for cell in row_element.xpath('./*[local-name()="c"]'):
                if cell.get("r") == f"A{row_number}":
                    code = normalize_code(
                        _read_cell_text(cell, shared_strings)
                    )
                    break

            if not code or code not in metrics:
                continue
            if code in seen:
                raise RuntimeError(
                    f"Mã {code} bị lặp trong XML {PLANNING_SHEET}!A."
                )
            seen.add(code)

            for column, key in OUTPUT_COLUMNS.items():
                _set_cell_value(
                    row_element,
                    row_number,
                    column,
                    metrics[code][key],
                )
            patched += 1

        if patched != len(metrics):
            missing = sorted(set(metrics) - seen)
            raise RuntimeError(
                f"Chỉ cập nhật {patched}/{len(metrics)} mã vào "
                f"{PLANNING_SHEET}!M:R. Thiếu: "
                + ", ".join(missing[:10])
            )

        new_sheet_xml = etree.tostring(
            sheet_root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        with zipfile.ZipFile(
            output_buffer,
            "w",
            compression=zipfile.ZIP_DEFLATED,
        ) as output_zip:
            for item in source_zip.infolist():
                data = (
                    new_sheet_xml
                    if item.filename == sheet_path
                    else source_zip.read(item.filename)
                )
                output_zip.writestr(item, data)

    return output_buffer.getvalue(), patched


def _same_value(current, target):
    if target is None:
        return current in (None, "")

    if isinstance(target, datetime):
        if isinstance(current, datetime):
            return abs((current - target).total_seconds()) < 1
        if isinstance(current, date):
            return current == target.date()
        return False

    if isinstance(current, (int, float)) and isinstance(target, (int, float)):
        return math.isclose(
            float(current), float(target), rel_tol=1e-10, abs_tol=1e-7
        )
    return current == target


def count_changes(planning_rows, metrics):
    changed = 0
    for code, row in planning_rows.items():
        current = row["current_outputs"]
        target = metrics[code]
        if any(
            not _same_value(current[key], target[key])
            for key in OUTPUT_COLUMNS.values()
        ):
            changed += 1
    return changed


def main():
    token = get_access_token()
    graph = GraphClient(token)

    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
    actual_item = graph.get_item_by_path(drive_id, SOURCE_ACTUAL_PATH)
    system_item = graph.get_item_by_path(
        drive_id, SOURCE_FACTORY_VIKODA_PATH
    )

    dest_bytes = graph.download_file(drive_id, dest_item["id"])
    actual_bytes = graph.download_file(drive_id, actual_item["id"])
    system_bytes = graph.download_file(drive_id, system_item["id"])

    selector, plan_month, planning_rows = read_planning_rows(dest_bytes)
    conversion_factors, _ = read_conversion_factors(dest_bytes)
    report_date, actual_receipts, consignments = read_actual_inputs(
        actual_bytes
    )
    system_receipts = read_system_receipts(
        system_bytes,
        conversion_factors,
        set(planning_rows),
    )

    state = load_runtime_state()
    metrics, state_changed, source_key, plan_key = calculate_metrics(
        report_date=report_date,
        plan_month=plan_month,
        planning_rows=planning_rows,
        actual_receipts=actual_receipts,
        system_receipts=system_receipts,
        current_consignments=consignments,
        state=state,
    )

    changed_count = count_changes(planning_rows, metrics)

    if changed_count:
        updated_bytes, patched = patch_workbook(dest_bytes, metrics)
        result = graph.upload_file(
            drive_id,
            dest_item["id"],
            updated_bytes,
            expected_etag=dest_item["eTag"],
        )
        print(
            f"[{PLANNING_SHEET}] Đã tính trực tiếp M:R cho "
            f"{patched} mã; {changed_count} dòng thay đổi."
        )
        print("Upload thành công:", result.get("name", DEST_PATH))
    else:
        print(
            f"[{PLANNING_SHEET}] M:R đã đúng, không cần upload lại."
        )

    if state_changed:
        save_runtime_state(state)

    print(
        f"Nguồn tháng {source_key}; kế hoạch {plan_key}; "
        f"{FC_SHEET}!{FC_SELECTOR_CELL}={selector!r}."
    )
    print(
        "Runtime không đọc các sheet mẫu Ke hoach SX tuan / "
        "FC thang nay / Nokho."
    )


if __name__ == "__main__":
    main()
