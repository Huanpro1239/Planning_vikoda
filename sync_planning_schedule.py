import math
import time
import zipfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from io import BytesIO
from zoneinfo import ZoneInfo

from lxml import etree
from openpyxl import load_workbook
from ortools.sat.python import cp_model

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
SHIFT_SCALE = 10000
QTY_SCALE = 1000
MAX_SOLVE_SECONDS = 8.0


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


def _scale(value, scale):
    return int(round(float(value) * scale))


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


def _is_sugar(classification):
    return str(classification or "").strip().casefold() == "có đường".casefold()


def _demand_profile(headers, fc):
    """Giữ convention FC/26: nhu cầu T2-T7, CN vẫn là ngày sản xuất bình thường."""
    demand_days = sum(1 for current_day in headers if current_day.weekday() != 6)
    if fc <= 0 or demand_days <= 0:
        return {current_day: 0.0 for current_day in headers}
    daily = fc / demand_days
    return {
        current_day: (0.0 if current_day.weekday() == 6 else daily)
        for current_day in headers
    }


def _validate_quantum(product):
    if product["planned_qty"] <= 0:
        product["quantum_qty"] = 0.0
        product["quantum_shift"] = 0.0
        product["required_units"] = 0
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

    quantum_shift = quantum_qty / product["per_shift"]
    if quantum_shift - product["max_shifts_per_day"] > EPSILON:
        raise RuntimeError(
            f"Mã {product['code']} cần {quantum_shift:.3f} ca cho một quantum "
            f"nhưng giới hạn ngày chỉ {product['max_shifts_per_day']:.3f} ca."
        )

    product["quantum_qty"] = quantum_qty
    product["quantum_shift"] = quantum_shift
    product["required_units"] = int(rounded_units)


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
            value = worksheet.cell(
                row=1,
                column=START_COLUMN_NUMBER + offset,
            ).value
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
            worksheet.iter_rows(
                min_row=2,
                max_col=START_COLUMN_NUMBER + MAX_DAYS - 1,
                values_only=True,
            ),
            start=2,
        ):
            code = normalize_code(values[0] if values else None)
            if not code:
                continue
            if code in seen_codes:
                raise RuntimeError(f"Mã {code} bị lặp trong {PLANNING_SHEET}!A.")
            seen_codes.add(code)

            classification = str(values[7] or "").strip()
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
                "debt": max(_to_number(values[13], f"{PLANNING_SHEET}!N{row_number}"), 0.0),
                "planned_qty": max(_to_number(values[15], f"{PLANNING_SHEET}!P{row_number}"), 0.0),
                "earliest_date": _as_date(values[17]),
                "existing_daily": list(values[18:18 + MAX_DAYS]),
            }
            product["demand_by_day"] = _demand_profile(headers, product["fc"])
            _validate_quantum(product)
            products.append(product)

        if not products:
            raise RuntimeError(f"Không đọc được SKU trong {PLANNING_SHEET}.")

        return headers, products
    finally:
        workbook.close()


def _new_solver():
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = MAX_SOLVE_SECONDS
    solver.parameters.num_search_workers = 8
    solver.parameters.random_seed = 17
    return solver


def _solve_phase(model, objective, label):
    model.Minimize(objective)
    solver = _new_solver()
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError(f"Scheduler CP-SAT không tìm được nghiệm ở phase {label}.")
    return solver, int(round(solver.ObjectiveValue()))


def _optimize_line(headers, line_products, line_capacity):
    model = cp_model.CpModel()
    day_index = {current_day: index for index, current_day in enumerate(headers)}
    capacity_scaled = _scale(line_capacity, SHIFT_SCALE)

    x = {}
    active = {}
    unscheduled = {}
    product_by_code = {product["code"]: product for product in line_products}

    for product in line_products:
        code = product["code"]
        required = product["required_units"]
        unscheduled[code] = model.NewIntVar(0, required, f"unscheduled_{code}")
        vars_for_product = []

        quantum_shift_scaled = _scale(product["quantum_shift"], SHIFT_SCALE)
        sku_daily_capacity_scaled = _scale(product["max_shifts_per_day"], SHIFT_SCALE)
        max_units_day = min(
            required,
            sku_daily_capacity_scaled // quantum_shift_scaled,
        )

        for current_day in headers:
            eligible = (
                product["earliest_date"] is None
                or current_day >= product["earliest_date"]
            )
            upper = max_units_day if eligible else 0
            variable = model.NewIntVar(0, upper, f"x_{code}_{day_index[current_day]}")
            x[(code, current_day)] = variable
            vars_for_product.append(variable)

            used = model.NewBoolVar(f"active_{code}_{day_index[current_day]}")
            active[(code, current_day)] = used
            if upper == 0:
                model.Add(used == 0)
            else:
                model.Add(variable <= upper * used)
                model.Add(variable >= used)

        model.Add(sum(vars_for_product) + unscheduled[code] == required)

    for current_day in headers:
        terms = []
        for product in line_products:
            shift_scaled = _scale(product["quantum_shift"], SHIFT_SCALE)
            terms.append(x[(product["code"], current_day)] * shift_scaled)
        model.Add(sum(terms) <= capacity_scaled)

    stockout_vars = []
    safety_vars = []
    for product in line_products:
        code = product["code"]
        quantum_qty_scaled = _scale(product["quantum_qty"], QTY_SCALE)
        opening_scaled = _scale(product["actual_stock"], QTY_SCALE)
        target_scaled = _scale(product["target_stock"], QTY_SCALE)
        cumulative_demand = 0
        cumulative_prod_terms = []

        for current_day in headers:
            cumulative_demand += _scale(product["demand_by_day"][current_day], QTY_SCALE)
            cumulative_prod_terms.append(x[(code, current_day)] * quantum_qty_scaled)
            stock_expr = opening_scaled + sum(cumulative_prod_terms) - cumulative_demand

            stockout = model.NewIntVar(0, 10**12, f"stockout_{code}_{day_index[current_day]}")
            model.Add(stockout >= -stock_expr)
            stockout_vars.append(stockout)

            safety = model.NewIntVar(0, 10**12, f"safety_{code}_{day_index[current_day]}")
            model.Add(safety >= target_scaled - stock_expr)
            safety_vars.append(safety)

    unscheduled_qty_expr = sum(
        unscheduled[product["code"]] * _scale(product["quantum_qty"], QTY_SCALE)
        for product in line_products
    )
    solver, best_unscheduled = _solve_phase(model, unscheduled_qty_expr, "P-completion")
    model.Add(unscheduled_qty_expr == best_unscheduled)

    stockout_expr = sum(stockout_vars)
    solver, best_stockout = _solve_phase(model, stockout_expr, "stockout")
    model.Add(stockout_expr == best_stockout)

    safety_expr = sum(safety_vars)
    solver, best_safety = _solve_phase(model, safety_expr, "safety-stock")
    model.Add(safety_expr == best_safety)

    active_days_expr = sum(active.values())
    timing_terms = []
    last_index = len(headers) - 1
    for product in line_products:
        code = product["code"]
        for current_day in headers:
            index = day_index[current_day]
            # Nợ kho ưu tiên về sớm; SKU khác sau khi đã bảo vệ inventory
            # thì xếp gần nhu cầu hơn để giảm tồn trung gian.
            timing_weight = index if product["debt"] > 0 else last_index - index
            timing_terms.append(x[(code, current_day)] * timing_weight)

    final_expr = active_days_expr * 10000 + sum(timing_terms)
    solver, _ = _solve_phase(model, final_expr, "compact-schedule")

    assignments = defaultdict(dict)
    carryover = {}
    usage = defaultdict(float)
    for product in line_products:
        code = product["code"]
        for current_day in headers:
            units = solver.Value(x[(code, current_day)])
            assignments[code][current_day] = units
            usage[current_day] += units * product["quantum_shift"]
        missing_units = solver.Value(unscheduled[code])
        if missing_units > 0:
            carryover[code] = _clean_number(missing_units * product["quantum_qty"])

    return assignments, usage, carryover, {
        "unscheduled_scaled": best_unscheduled,
        "stockout_scaled_days": best_stockout,
        "safety_scaled_days": best_safety,
    }


def build_schedule(headers, products):
    schedule = {
        product["code"]: {current_day: 0.0 for current_day in headers}
        for product in products
    }
    active_products = [product for product in products if product["planned_qty"] > 0]

    line_capacity = {}
    by_line = defaultdict(list)
    for product in active_products:
        by_line[product["line"]].append(product)
        line_capacity[product["line"]] = max(
            line_capacity.get(product["line"], 0.0),
            product["max_shifts_per_day"],
        )

    line_usage = defaultdict(lambda: defaultdict(float))
    carryover = {}
    optimizer_meta = {}

    for line, line_products in sorted(by_line.items()):
        assignments, usage, line_carryover, meta = _optimize_line(
            headers,
            line_products,
            line_capacity[line],
        )
        optimizer_meta[line] = meta
        carryover.update(line_carryover)

        for current_day, used_shift in usage.items():
            line_usage[line][current_day] = used_shift
        for product in line_products:
            code = product["code"]
            for current_day in headers:
                units = assignments[code][current_day]
                schedule[code][current_day] = units * product["quantum_qty"]

    validate_schedule(headers, products, schedule, line_capacity, line_usage)
    inventory = simulate_inventory(headers, products, schedule)
    return schedule, line_capacity, line_usage, carryover, inventory, optimizer_meta


def simulate_inventory(headers, products, schedule):
    result = {}
    for product in products:
        stock = product["actual_stock"]
        min_stock = stock
        first_stockout = None
        first_below_safety = None
        for current_day in headers:
            stock += (
                schedule[product["code"]][current_day]
                - product["demand_by_day"][current_day]
            )
            min_stock = min(min_stock, stock)
            if stock < -EPSILON and first_stockout is None:
                first_stockout = current_day
            if (
                stock < product["target_stock"] - EPSILON
                and first_below_safety is None
            ):
                first_below_safety = current_day

        result[product["code"]] = {
            "ending_stock": _clean_number(stock),
            "min_stock": _clean_number(min_stock),
            "first_stockout": first_stockout,
            "first_below_safety": first_below_safety,
        }
    return result


def validate_schedule(headers, products, schedule, line_capacity, line_usage):
    for line, days in line_usage.items():
        capacity = line_capacity[line]
        for current_day, used in days.items():
            if used - capacity > 1e-6:
                raise RuntimeError(
                    f"Chuyền {line} ngày {current_day:%d/%m} dùng {used:.3f} ca > "
                    f"capacity {capacity:.3f}."
                )

    for product in products:
        code = product["code"]
        total = sum(schedule[code].values())
        if total - product["planned_qty"] > 1e-5:
            raise RuntimeError(
                f"Mã {code} được xếp {total} > P={product['planned_qty']}."
            )
        earliest = product["earliest_date"]
        if earliest:
            for current_day, qty in schedule[code].items():
                if qty > EPSILON and current_day < earliest:
                    raise RuntimeError(
                        f"Mã {code} bị xếp ngày {current_day:%d/%m} trước R={earliest:%d/%m}."
                    )


def _number_equal(current, target):
    current = 0.0 if current in (None, "") else float(current)
    target = 0.0 if target in (None, "") else float(target)
    return math.isclose(current, target, rel_tol=1e-12, abs_tol=1e-8)


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
    (
        schedule,
        line_capacity,
        line_usage,
        carryover,
        inventory,
        optimizer_meta,
    ) = build_schedule(headers, products)
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
        "inventory": inventory,
        "optimizer_meta": optimizer_meta,
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

        stockout_codes = []
        safety_codes = []
        for product in info["products"]:
            result = info["inventory"][product["code"]]
            if result["first_stockout"]:
                stockout_codes.append((product["code"], result["first_stockout"]))
            elif result["first_below_safety"]:
                safety_codes.append((product["code"], result["first_below_safety"]))

        if stockout_codes:
            print("[Scheduler][INVENTORY_RISK] Các mã vẫn có nguy cơ tồn âm:")
            for code, risk_day in stockout_codes:
                print(f"  - {code}: từ {risk_day:%d/%m}")
        if safety_codes:
            print("[Scheduler][SAFETY_RISK] Các mã xuống dưới tồn mục tiêu:")
            for code, risk_day in safety_codes:
                print(f"  - {code}: từ {risk_day:%d/%m}")
        if not stockout_codes and not safety_codes:
            print("[Scheduler] Inventory simulation không phát hiện stockout/safety breach.")

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
