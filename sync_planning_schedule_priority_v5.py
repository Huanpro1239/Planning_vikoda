import math
import zipfile
from datetime import datetime
from io import BytesIO

from lxml import etree
from openpyxl import load_workbook

import sync_planning_metrics as metrics_core
import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v2 as v2
import sync_planning_schedule_priority_v3 as v3
from sync_planning_calendar_all_months import resolve_plan_year
from sync_planning_fc import _find_sheet_xml_path, _load_shared_strings, _read_cell_text
from sync_stock import normalize_code


_ORIGINAL_READ_INPUTS = priority.base.read_schedule_inputs
_ORIGINAL_PATCH_SCHEDULE = priority.base.patch_schedule_workbook
_ORIGINAL_CONTINUOUS_ALLOCATOR = priority.base._allocate_continuous_campaign_line
GROUP_SLACK_TOLERANCE_DAYS = 1


def _resolve_schedule_year(workbook_bytes):
    workbook = load_workbook(BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        worksheet = workbook[priority.base.PLANNING_SHEET]
        raw = str(
            worksheet.cell(row=1, column=priority.base.START_COLUMN_NUMBER).value or ""
        ).strip()
        if not raw:
            return datetime.now(priority.base.TIMEZONE).year
        first_line = raw.splitlines()[0].strip()
        _, month_text = first_line.split("/", 1)
        return resolve_plan_year(int(month_text))
    finally:
        workbook.close()


def _preferred_index(headers, preferred_date):
    if preferred_date is None or preferred_date <= headers[0]:
        return 0
    for index, current_day in enumerate(headers):
        if current_day >= preferred_date:
            return index
    return len(headers)


def _effective_release_index(headers, product, preferred_date):
    """R là ngày dự kiến; cho phép pull-forward khi workload/risk yêu cầu."""
    preferred = _preferred_index(headers, preferred_date)
    if product["planned_qty"] <= priority.base.EPSILON:
        return preferred

    empty_schedule = priority.base._empty_schedule(headers, [product])
    first_stockout, first_safety = v3._dynamic_risk(headers, product, empty_schedule)

    if float(product.get("debt", 0) or 0) > 0:
        return 0

    risk_index = first_stockout if first_stockout is not None else first_safety
    if risk_index is None:
        return preferred

    required_shifts = product["planned_qty"] / product["per_shift"]
    daily_capacity = max(
        float(product.get("max_shifts_per_day", 0) or 0),
        priority.base.EPSILON,
    )
    workload_days = required_shifts / daily_capacity
    setup_days = priority.base.SETUP_SHIFTS / daily_capacity
    release = int(math.floor(risk_index - workload_days - setup_days))
    return min(preferred, max(0, release))


def read_schedule_inputs_v5(workbook_bytes, *, plan_year=None):
    if plan_year is None:
        plan_year = _resolve_schedule_year(workbook_bytes)

    headers, products = _ORIGINAL_READ_INPUTS(workbook_bytes, plan_year=plan_year)
    for product in products:
        preferred = product["earliest_date"]
        product["preferred_date"] = preferred
        release_index = _effective_release_index(headers, product, preferred)
        product["release_index"] = release_index
        if release_index < len(headers):
            product["earliest_date"] = headers[release_index]
    return headers, products


def _risk_snapshot(headers, product):
    return priority._risk_snapshot(headers, product)


def _processing_shifts(product):
    per_shift = float(product.get("per_shift", 0) or 0)
    if per_shift <= 0:
        return math.inf
    return float(product.get("planned_qty", 0) or 0) / per_shift


def _campaign_key(headers, product, cursor_shift, capacity):
    risk = _risk_snapshot(headers, product)
    release = priority.base._earliest_index(headers, product) * capacity
    start = max(cursor_shift, release)
    processing = _processing_shifts(product)

    if risk["first_stockout"] is not None:
        tier = 0
        deadline_shift = (risk["first_stockout"] + 1) * capacity
        deadline_index = risk["first_stockout"]
    elif risk["first_safety"] is not None:
        tier = 1
        deadline_shift = (risk["first_safety"] + 1) * capacity
        deadline_index = risk["first_safety"]
    elif float(product.get("debt", 0) or 0) > 0:
        tier = 2
        deadline_shift = (len(headers) + 1) * capacity
        deadline_index = len(headers)
    else:
        tier = 3
        deadline_shift = math.inf
        deadline_index = len(headers) + 30

    latest_start = deadline_shift - processing if math.isfinite(deadline_shift) else math.inf
    slack = latest_start - start
    return (
        tier,
        slack,
        deadline_index,
        0 if float(product.get("debt", 0) or 0) > 0 else 1,
        priority.base._earliest_index(headers, product),
        product.get("row", 0),
    )


def _latest_start(headers, product, capacity):
    risk = _risk_snapshot(headers, product)
    release = priority.base._earliest_index(headers, product) * capacity
    deadline_index = (
        risk["first_stockout"]
        if risk["first_stockout"] is not None
        else risk["first_safety"]
    )
    if deadline_index is None:
        return math.inf
    return max(
        release,
        (deadline_index + 1) * capacity
        - _processing_shifts(product)
        - priority.base.SETUP_SHIFTS,
    )


def _finish_shift(headers, product, cursor_shift, capacity):
    release = priority.base._earliest_index(headers, product) * capacity
    start = max(cursor_shift + priority.base.SETUP_SHIFTS, release)
    return start + _processing_shifts(product)


def minimum_slack_campaign_candidate(remaining, headers, cursor_shift, capacity, last_group):
    """Một SKU = một campaign; shortage/minimum-slack trước, grouping sau."""
    if not remaining:
        return None, cursor_shift

    current_index = min(
        len(headers),
        int(math.floor(cursor_shift / capacity + priority.base.EPSILON)),
    )
    eligible = [
        product
        for product in remaining
        if priority.base._earliest_index(headers, product) <= current_index
    ]

    if not eligible:
        next_index = min(
            priority.base._earliest_index(headers, product)
            for product in remaining
        )
        cursor_shift = max(cursor_shift, next_index * capacity)
        current_index = next_index
        eligible = [
            product
            for product in remaining
            if priority.base._earliest_index(headers, product) <= current_index
        ]

    eligible.sort(key=lambda p: _campaign_key(headers, p, cursor_shift, capacity))
    primary = eligible[0]

    # Look-ahead SKU chưa release: chỉ dùng gap cho campaign khác nếu nó chạy xong
    # trước latest-start của SKU tương lai cấp bách.
    future = [
        product
        for product in remaining
        if priority.base._earliest_index(headers, product) > current_index
    ]
    if future:
        future.sort(key=lambda p: _campaign_key(headers, p, cursor_shift, capacity))
        urgent_future = future[0]
        if _campaign_key(headers, urgent_future, cursor_shift, capacity) < _campaign_key(
            headers, primary, cursor_shift, capacity
        ):
            latest = _latest_start(headers, urgent_future, capacity)
            safe_fillers = [
                p
                for p in eligible
                if _finish_shift(headers, p, cursor_shift, capacity) <= latest + priority.base.EPSILON
            ]
            if safe_fillers:
                safe_fillers.sort(key=lambda p: _campaign_key(headers, p, cursor_shift, capacity))
                primary = safe_fillers[0]
            else:
                release = priority.base._earliest_index(headers, urgent_future) * capacity
                return urgent_future, max(cursor_shift, release)

    # Giữ cùng nhóm nếu độ khẩn cấp gần tương đương và không làm primary trễ.
    if last_group:
        primary_key = _campaign_key(headers, primary, cursor_shift, capacity)
        tolerance = GROUP_SLACK_TOLERANCE_DAYS * capacity
        same_group = [
            p
            for p in eligible
            if p.get("product_group") == last_group
            and _campaign_key(headers, p, cursor_shift, capacity)[0] <= primary_key[0]
            and _campaign_key(headers, p, cursor_shift, capacity)[1] <= primary_key[1] + tolerance
        ]
        same_group.sort(key=lambda p: _campaign_key(headers, p, cursor_shift, capacity))
        for candidate in same_group:
            if candidate["code"] == primary["code"]:
                return primary, cursor_shift
            if _finish_shift(headers, candidate, cursor_shift, capacity) <= _latest_start(
                headers, primary, capacity
            ) + priority.base.EPSILON:
                return candidate, cursor_shift

    return primary, cursor_shift


def _actual_first_date(headers, product, schedule):
    for current_day in headers:
        if schedule[product["code"]][current_day] > priority.base.EPSILON:
            return current_day
    return None


def _same_day(left, right):
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    if isinstance(left, datetime):
        left = left.date()
    if isinstance(right, datetime):
        right = right.date()
    return left == right


def _patch_actual_start_dates(workbook_bytes, headers, products, schedule):
    targets = {}
    for product in products:
        if product["planned_qty"] <= priority.base.EPSILON:
            continue
        actual = _actual_first_date(headers, product, schedule)
        if actual is None:
            continue
        if not _same_day(product.get("preferred_date"), actual):
            targets[product["code"]] = datetime(actual.year, actual.month, actual.day)

    if not targets:
        return workbook_bytes, 0

    source_buffer = BytesIO(workbook_bytes)
    output_buffer = BytesIO()
    with zipfile.ZipFile(source_buffer, "r") as source_zip:
        sheet_path = _find_sheet_xml_path(source_zip, priority.base.PLANNING_SHEET)
        shared_strings = _load_shared_strings(source_zip)
        root = etree.fromstring(source_zip.read(sheet_path))
        seen = set()

        rows = root.xpath('//*[local-name()="sheetData"]/*[local-name()="row"]')
        for row in rows:
            row_number_text = row.get("r")
            if not row_number_text:
                continue
            row_number = int(row_number_text)
            code = None
            for cell in row.xpath('./*[local-name()="c"]'):
                if cell.get("r") == f"A{row_number}":
                    code = normalize_code(_read_cell_text(cell, shared_strings))
                    break
            if code not in targets:
                continue
            metrics_core._set_cell_value(row, row_number, "R", targets[code])
            seen.add(code)

        missing = sorted(set(targets) - seen)
        if missing:
            raise RuntimeError("Không ghi được R thực tế cho: " + ", ".join(missing[:10]))

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

    return output_buffer.getvalue(), len(targets)


def patch_schedule_workbook_v5(workbook_bytes, headers, products, schedule):
    updated, daily_changed = _ORIGINAL_PATCH_SCHEDULE(
        workbook_bytes,
        headers,
        products,
        schedule,
    )
    updated, r_changed = _patch_actual_start_dates(updated, headers, products, schedule)
    return updated, daily_changed + r_changed


def install_priority_scheduler_v5():
    # RGB/Galon: shortage-aware weekly/spread; KHS/PET: one-campaign MST.
    v2.install_priority_scheduler_v2()
    priority.shortage_first_campaign_candidate = minimum_slack_campaign_candidate
    priority.base._campaign_candidate = minimum_slack_campaign_candidate
    priority.base._allocate_continuous_campaign_line = _ORIGINAL_CONTINUOUS_ALLOCATOR
    priority.base.read_schedule_inputs = read_schedule_inputs_v5
    priority.base.patch_schedule_workbook = patch_schedule_workbook_v5


if __name__ == "__main__":
    install_priority_scheduler_v5()
    priority.base.main_with_retry()
