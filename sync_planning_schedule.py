import math
import time
import zipfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import BytesIO
from zoneinfo import ZoneInfo

from lxml import etree
from openpyxl import load_workbook

from sync_planning_fc import (
    _find_sheet_xml_path,
    _get_or_create_cell,
    _load_shared_strings,
    _read_cell_text,
)
from sync_stock import DEST_PATH, GraphClient, get_access_token, normalize_code


PLANNING_SHEET = "Ke_hoach_SX"
START_COLUMN_NUMBER = 19  # S
MAX_DAYS = 31
TIMEZONE = ZoneInfo("Asia/Ho_Chi_Minh")
EPSILON = 1e-9


def _column_letter(number):
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _to_number(value, label):
    if value in (None, ""):
        return 0.0
    if isinstance(value, bool):
        raise RuntimeError(f"{label} chứa TRUE/FALSE, không phải số.")
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise RuntimeError(f"{label} không phải số: {value!r}") from exc


def _clean_number(value):
    value = float(value)
    if abs(value) < EPSILON:
        return 0
    if math.isclose(value, round(value), rel_tol=1e-12, abs_tol=1e-9):
        return int(round(value))
    return value


def _as_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raise RuntimeError(f"Ngày bắt đầu sản xuất không hợp lệ: {value!r}")


def _parse_header_date(value, plan_year):
    text = str(value or "").strip()
    if not text:
        return None
    first_line = text.splitlines()[0].strip()
    try:
        day_text, month_text = first_line.split("/", 1)
        return date(plan_year, int(month_text), int(day_text))
    except Exception as exc:
        raise RuntimeError(f"Tiêu đề ngày không hợp lệ: {value!r}") from exc


def _risk_date(first_day, actual_stock, target_stock, fc):
    if fc <= 0:
        return date.max
    daily_fc = fc / 26.0
    if daily_fc <= 0:
        return date.max
    days_to_target = max((actual_stock - target_stock) / daily_fc, 0.0)
    return first_day + timedelta(days=days_to_target)


def _is_sugar(classification):
    return str(classification or "").strip().casefold() == "có đường".casefold()


def _validate_quantum(product):
    if product["planned_qty"] <= 0:
        return
    if not product["line"]:
        raise RuntimeError(f"Mã {product['code']} chưa có Chuyền tại cột F.")
    if product["per_shift"] <= 0:
        raise RuntimeError(f"Mã {product['code']} có Số lượng/ca <= 0.")
    if product["max_shifts_per_day"] <= 0:
        raise RuntimeError(f"Mã {product['code']} có Số ca theo ngày <= 0.")

    quantum_qty = product["batch"] if product["is_sugar"] else product["per_shift"]
    if quantum_qty <= 0:
        source = "Số lượng/mẻ" if product["is_sugar"] else "Số lượng/ca"
        raise RuntimeError(f"Mã {product['code']} có {source} <= 0.")

    unit_count = product["planned_qty"] / quantum_qty
    rounded_units = round(unit_count)
    if not math.isclose(unit_count, rounded_units, rel_tol=1e-9, abs_tol=1e-7):
        raise RuntimeError(
            f"Mã {product['code']} có P={product['planned_qty']} không chia hết "
            f"cho quantum={quantum_qty}."
        )

    product["quantum_qty"] = quantum_qty
    product["quantum_shift"] = quantum_qty / product["per_shift"]
    product["remaining_units"] = int(rounded_units)
    product["required_units"] = int(rounded_units)

    if product["quantum_shift"] - product["max_shifts_per_day"] > EPSILON:
        raise RuntimeError(
            f"Mã {product['code']} cần {product['quantum_shift']:.3f} ca cho một "
            f"quantum nhưng giới hạn ngày chỉ {product['max_shifts_per_day']:.3f} ca."
        )


def read_schedule_inputs(workbook_bytes, *, plan_year=None):
    workbook = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        if PLANNING_SHEET not in workbook.sheetnames:
            raise RuntimeError(f"Không tìm thấy sheet {PLANNING_SHEET!r}.")
        worksheet = workbook[PLANNING_SHEET]

        if plan_year is None:
            plan_year = datetime.now(TIMEZONE).year

        headers = []
        for offset in range(MAX_DAYS):
            column = START_COLUMN_NUMBER + offset
            value = worksheet.cell(row=1, column=column).value
            parsed = _parse_header_date(value, plan_year)
            if parsed is None:
                if headers:
                    break
                continue
            headers.append(parsed)

        if not headers:
            raise RuntimeError(
                f"{PLANNING_SHEET}!S1 trở đi chưa có dải ngày. "
                "Hãy chạy sync_planning_calendar.py trước."
            )

        first_day = headers[0]
        expected = [first_day + timedelta(days=i) for i in range(len(headers))]
        if headers != expected:
            raise RuntimeError("Dải ngày S1:... không liên tục theo ngày.")

        products = []
        seen_codes = set()
        for row_number, values in enumerate(
            worksheet.iter_rows(min_row=2, max_col=START_COLUMN_NUMBER + MAX_DAYS - 1, values_only=True),
            start=2,
        ):
            code = normalize_code(values[0] if values else None)
            if not code:
                continue
            if code in seen_codes:
                raise RuntimeError(f"Mã {code} bị lặp trong {PLANNING_SHEET}!A.")
            seen_codes.add(code)

            classification = str(values[7] or "").strip()
            earliest = _as_date(values[17])
            planned_qty = _to_number(values[15], f"{PLANNING_SHEET}!P{row_number}")

            product = {
                "row": row_number,
                "code": code,
                "batch": _to_number(values[3], f"{PLANNING_SHEET}!D{row_number}"),
                "per_shift": _to_number(values[4], f"{PLANNING_SHEET}!E{row_number}"),
                "line": str(values[5] or "").strip(),
                "product_group": str(values[6] or "").strip(),
                "classification": classification,
                "is_sugar": _is_sugar(classification),
                "max_shifts_per_day": _to_number(values[8], f"{PLANNING_SHEET}!I{row_number}"),
                "actual_stock": _to_number(values[9], f"{PLANNING_SHEET}!J{row_number}"),
                "fc": _to_number(values[11], f"{PLANNING_SHEET}!L{row_number}"),
                "target_stock": _to_number(values[12], f"{PLANNING_SHEET}!M{row_number}"),
                "debt": _to_number(values[13], f"{PLANNING_SHEET}!N{row_number}"),
                "planned_qty": max(planned_qty, 0.0),
                "earliest_date": earliest,
                "existing_daily": list(values[18:18 + MAX_DAYS]),
            }
            product["risk_date"] = _risk_date(
                first_day,
                product["actual_stock"],
                product["target_stock"],
                product["fc"],
            )
            _validate_quantum(product)
            products.append(product)

        if not products:
            raise RuntimeError(f"Không đọc được SKU trong {PLANNING_SHEET}.")

        return headers, products
    finally:
        workbook.close()


def _priority_key(product):
    earliest = product["earliest_date"] or date.max
    remaining_shift = product["remaining_units"] * product.get("quantum_shift", 0)
    return (
        0 if product["debt"] > 0 else 1,
        product["risk_date"],
        earliest,
        -remaining_shift,
        product["code"],
    )


def build_schedule(headers, products):
    schedule = {
        product["code"]: {day: 0.0 for day in headers}
        for product in products
    }

    active = [product for product in products if product["planned_qty"] > 0]
    line_capacity = {}
    for product in active:
        current = line_capacity.get(product["line"], 0.0)
        line_capacity[product["line"]] = max(current, product["max_shifts_per_day"])

    line_usage = defaultdict(lambda: defaultdict(float))
    sku_usage = defaultdict(lambda: defaultdict(float))

    by_line = defaultdict(list)
    for product in active:
        by_line[product["line"]].append(product)

    for current_day in headers:
        for line, line_products in sorted(by_line.items()):
            available = line_capacity[line] - line_usage[line][current_day]
            while available > EPSILON:
                candidates = []
                for product in line_products:
                    if product["remaining_units"] <= 0:
                        continue
                    earliest = product["earliest_date"] or headers[0]
                    if current_day < earliest:
                        continue
                    sku_remaining_capacity = (
                        product["max_shifts_per_day"] - sku_usage[product["code"]][current_day]
                    )
                    if product["quantum_shift"] - available > EPSILON:
                        continue
                    if product["quantum_shift"] - sku_remaining_capacity > EPSILON:
                        continue
                    candidates.append(product)

                if not candidates:
                    break

                product = min(candidates, key=_priority_key)
                sku_remaining_capacity = (
                    product["max_shifts_per_day"] - sku_usage[product["code"]][current_day]
                )
                fit_line = int(math.floor((available + EPSILON) / product["quantum_shift"]))
                fit_sku = int(
                    math.floor((sku_remaining_capacity + EPSILON) / product["quantum_shift"])
                )
                units = min(product["remaining_units"], fit_line, fit_sku)
                if units <= 0:
                    break

                used_shift = units * product["quantum_shift"]
                produced = units * product["quantum_qty"]
                schedule[product["code"]][current_day] += produced
                product["remaining_units"] -= units
                line_usage[line][current_day] += used_shift
                sku_usage[product["code"]][current_day] += used_shift
                available -= used_shift

    carryover = {}
    for product in active:
        remaining_qty = product["remaining_units"] * product["quantum_qty"]
        if remaining_qty > EPSILON:
            carryover[product["code"]] = _clean_number(remaining_qty)

    validate_schedule(headers, products, schedule, line_capacity, line_usage)
    return schedule, line_capacity, line_usage, carryover


def validate_schedule(headers, products, schedule, line_capacity, line_usage):
    for line, days in line_usage.items():
        capacity = line_capacity[line]
        for current_day, used in days.items():
            if used - capacity > 1e-7:
                raise RuntimeError(
                    f"Chuyền {line} ngày {current_day:%d/%m} dùng {used:.3f} ca > "
                    f"capacity {capacity:.3f}."
                )

    for product in products:
        total = sum(schedule[product["code"]].values())
        if total - product["planned_qty"] > 1e-6:
            raise RuntimeError(
                f"Mã {product['code']} được xếp {total} > P={product['planned_qty']}."
            )
        earliest = product["earliest_date"]
        if earliest:
            for current_day, qty in schedule[product["code"]].items():
                if qty > EPSILON and current_day < earliest:
                    raise RuntimeError(
                        f"Mã {product['code']} bị xếp ngày {current_day:%d/%m} trước R={earliest:%d/%m}."
                    )


def _number_equal(current, target):
    current = 0.0 if current in (None, "") else float(current)
    target = 0.0 if target in (None, "") else float(target)
    return math.isclose(current, target, rel_tol=1e-12, abs_tol=1e-9)


def _set_numeric_or_blank(cell, value):
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)

    if value in (None, "") or abs(float(value)) < EPSILON:
        return

    namespace = etree.QName(cell).namespace
    node = etree.SubElement(cell, f"{{{namespace}}}v")
    cleaned = _clean_number(value)
    node.text = str(cleaned) if isinstance(cleaned, int) else format(float(cleaned), ".15g")


def patch_schedule_workbook(workbook_bytes, headers, products, schedule):
    target_values = {}
    day_count = len(headers)
    changed_count = 0

    for product in products:
        values = []
        for index in range(MAX_DAYS):
            target = schedule[product["code"]][headers[index]] if index < day_count else 0
            target = _clean_number(target)
            values.append(target)
            current = product["existing_daily"][index] if index < len(product["existing_daily"]) else None
            if not _number_equal(current, target):
                changed_count += 1
        target_values[product["code"]] = values

    if changed_count == 0:
        return workbook_bytes, 0

    source_buffer = BytesIO(workbook_bytes)
    output_buffer = BytesIO()
    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        sheet_path = _find_sheet_xml_path(source_zip, PLANNING_SHEET)
        shared_strings = _load_shared_strings(source_zip)
        root = etree.fromstring(source_zip.read(sheet_path))

        seen = set()
        rows = root.xpath('//*[local-name()="sheetData"]/*[local-name()="row"]')
        for row in rows:
            row_number_text = row.get("r")
            if not row_number_text or int(row_number_text) == 1:
                continue
            row_number = int(row_number_text)
            code = None
            for cell in row.xpath('./*[local-name()="c"]'):
                if cell.get("r") == f"A{row_number}":
                    code = normalize_code(_read_cell_text(cell, shared_strings))
                    break
            if not code or code not in target_values:
                continue
            seen.add(code)

            p_nodes = row.xpath(f'./*[local-name()="c"][@r="P{row_number}"]')
            body_style = p_nodes[0].get("s") if p_nodes else None

            for index, value in enumerate(target_values[code]):
                column_letter = _column_letter(START_COLUMN_NUMBER + index)
                cell = _get_or_create_cell(row, row_number, column_letter)
                if body_style is not None:
                    cell.set("s", body_style)
                _set_numeric_or_blank(cell, value)

        missing = sorted(set(target_values) - seen)
        if missing:
            raise RuntimeError(
                f"Không tìm thấy {len(missing)} mã trong XML {PLANNING_SHEET}: "
                + ", ".join(missing[:10])
            )

        new_xml = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        with zipfile.ZipFile(output_buffer, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            for item in source_zip.infolist():
                data = new_xml if item.filename == sheet_path else source_zip.read(item.filename)
                output_zip.writestr(item, data)

    return output_buffer.getvalue(), changed_count


def prepare_schedule_update(workbook_bytes, *, plan_year=None):
    headers, products = read_schedule_inputs(workbook_bytes, plan_year=plan_year)
    schedule, line_capacity, line_usage, carryover = build_schedule(headers, products)
    updated_bytes, changed_count = patch_schedule_workbook(
        workbook_bytes,
        headers,
        products,
        schedule,
    )

    utilization = {}
    for line, capacity in line_capacity.items():
        total_capacity = capacity * len(headers)
        total_used = sum(line_usage[line].values())
        utilization[line] = 0 if total_capacity <= 0 else total_used / total_capacity

    return updated_bytes, {
        "headers": headers,
        "products": products,
        "schedule": schedule,
        "line_capacity": line_capacity,
        "line_usage": line_usage,
        "utilization": utilization,
        "carryover": carryover,
        "changed_count": changed_count,
    }


def main_with_retry(*, sleep_func=time.sleep, max_attempts=6, retry_delay_seconds=10):
    token = get_access_token()
    graph = GraphClient(token)
    site_id = graph.get_site_id()
    drive_id = graph.get_default_drive_id(site_id)

    for attempt in range(1, max_attempts + 1):
        dest_item = graph.get_item_by_path(drive_id, DEST_PATH)
        dest_bytes = graph.download_file(drive_id, dest_item["id"])
        updated_bytes, info = prepare_schedule_update(dest_bytes)

        for line, utilization in sorted(info["utilization"].items()):
            capacity = info["line_capacity"][line]
            print(
                f"[Scheduler] {line}: {capacity:g} ca/ngày; "
                f"utilization tháng {utilization * 100:.1f}%."
            )

        if info["carryover"]:
            print("[Scheduler][CARRYOVER] Chưa xếp hết trong tháng:")
            for code, qty in sorted(info["carryover"].items()):
                print(f"  - {code}: {qty}")
        else:
            print("[Scheduler] Tất cả P > 0 đã được xếp hết trong tháng.")

        if info["changed_count"] == 0:
            print(f"[{PLANNING_SHEET}] Lịch S:... đã đúng; không cần upload lại.")
            return

        try:
            graph.upload_file(
                drive_id,
                dest_item["id"],
                updated_bytes,
                expected_etag=dest_item["eTag"],
            )
            print(
                f"[{PLANNING_SHEET}] Đã cập nhật {info['changed_count']} ô lịch sản xuất S:..."
            )
            return
        except RuntimeError as exc:
            message = str(exc)
            retryable = "423" in message or "resourceLocked" in message or "412" in message
            if not retryable or attempt == max_attempts:
                raise
            print(
                f"[{PLANNING_SHEET}] File đang khóa/thay đổi; thử scheduler lại "
                f"({attempt}/{max_attempts})."
            )
            sleep_func(retry_delay_seconds)


if __name__ == "__main__":
    main_with_retry()
