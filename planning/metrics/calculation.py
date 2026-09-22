"""Pure/static Planning metrics calculations.

Production behavior is wired directly here. No module mutates these functions at
import time.
"""

import calendar
import math
from datetime import datetime, timedelta

from sync_stock import clean_number

from .state import is_month_end, month_key, next_month


def excel_roundup_integer(value):
    """Mô phỏng Excel ROUNDUP(value, 0): làm tròn ra xa số 0."""
    if value > 0:
        return math.ceil(value)
    if value < 0:
        return math.floor(value)
    return 0


def demand_days(plan_year, plan_month):
    days_in_month = calendar.monthrange(plan_year, plan_month)[1]
    return sum(
        1
        for day in range(1, days_in_month + 1)
        if datetime(plan_year, plan_month, day).weekday() != 6
    )


def resolve_plan_year(report_date, plan_month):
    year = report_date.year
    delta = plan_month - report_date.month
    if delta <= -6:
        year += 1
    elif delta >= 6:
        year -= 1
    return year


def planning_stock(actual_stock, book_stock):
    return min(
        max(float(actual_stock or 0), 0.0),
        max(float(book_stock or 0), 0.0),
    )


def urgent_supply_need(fc, warehouse_debt, actual_stock):
    return max(
        float(fc or 0)
        + float(warehouse_debt or 0)
        - max(float(actual_stock or 0), 0.0),
        0.0,
    )


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
    """Production M:R base calculation used for every planning month."""
    if batch <= 0:
        raise ValueError("Số lượng/mẻ phải > 0.")
    if per_shift <= 0:
        raise ValueError("Số lượng/ca phải > 0.")
    if shifts_per_day <= 0:
        raise ValueError("Số ca theo ngày phải > 0.")

    days = demand_days(plan_year, plan_month)
    daily_fc = fc / days if fc and days else 0

    minimum_stock = daily_fc * (leadtime + 2)
    expected_end_stock = (
        0 if opening_consignment > minimum_stock else minimum_stock
    )

    stock_for_plan = planning_stock(actual_stock, book_stock)
    urgent_need = urgent_supply_need(fc, warehouse_debt, actual_stock)

    if urgent_need > 0:
        required_production = urgent_need
    else:
        required_production = (
            float(fc or 0)
            + float(warehouse_debt or 0)
            + float(expected_end_stock or 0)
            - stock_for_plan
        )
    required_production = max(required_production, 0)

    if required_production <= 0:
        rounded_production = 0
    else:
        is_sugar = classification.casefold() == "có đường".casefold()
        base_qty = batch if is_sugar else per_shift
        rounded_production = (
            excel_roundup_integer(required_production / base_qty) * base_qty
        )
        rounded_production = max(rounded_production, 0)

    production_days = rounded_production / per_shift / shifts_per_day

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


def calculate_row_default(
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
    """Legacy direct formula kept only as an explicit regression reference."""
    if batch <= 0:
        raise ValueError("Số lượng/mẻ phải > 0.")
    if per_shift <= 0:
        raise ValueError("Số lượng/ca phải > 0.")
    if shifts_per_day <= 0:
        raise ValueError("Số ca theo ngày phải > 0.")

    daily_fc = fc / 26 if fc else 0
    minimum_stock = daily_fc * (leadtime + 2)
    expected_end_stock = (
        0 if opening_consignment > minimum_stock else minimum_stock
    )

    if warehouse_debt > 0:
        required_production = fc + warehouse_debt - book_stock
    else:
        required_production = (
            fc + expected_end_stock - book_stock + warehouse_debt
        )

    if required_production == 0:
        rounded_production = 0
    else:
        is_sugar = classification.casefold() == "có đường".casefold()
        base_qty = batch if is_sugar else per_shift
        rounded_production = (
            excel_roundup_integer(required_production / base_qty) * base_qty
        )

    production_days = rounded_production / per_shift / shifts_per_day
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


def bootstrap_opening_debt(
    *, current_debt, actual_receipt, system_receipt
):
    return clean_number(current_debt + actual_receipt - system_receipt)


def clamp_nonnegative_production(output):
    output = dict(output)
    required = float(output.get("required_production", 0) or 0)
    rounded = float(output.get("rounded_production", 0) or 0)

    if required <= 0:
        output["required_production"] = 0
        output["rounded_production"] = 0
        output["production_days"] = 0
        return output

    output["required_production"] = clean_number(required)
    if rounded <= 0:
        output["rounded_production"] = 0
        output["production_days"] = 0
    else:
        output["rounded_production"] = clean_number(rounded)
    return output


def calculate_metrics(
    *,
    report_date,
    plan_month,
    planning_rows,
    actual_receipts,
    system_receipts,
    current_consignments,
    state,
    leadtimes,
):
    """Calculate production M:R with No kho debt and explicit Danh_muc leadtimes."""
    source_key = month_key(report_date.year, report_date.month)
    plan_year = resolve_plan_year(report_date, plan_month)
    plan_key = month_key(plan_year, plan_month)
    state_changed = False

    # These inputs remain in the contract because they are part of the same
    # production snapshot and are used elsewhere in pipeline provenance.
    _ = actual_receipts, system_receipts

    consign_state = state.setdefault("opening_consignment_by_month", {})
    opening_consignment = consign_state.get(plan_key)
    if opening_consignment is None:
        opening_consignment = {
            code: clean_number(current_consignments.get(code, 0))
            for code in planning_rows
        }
        consign_state[plan_key] = opening_consignment
        state_changed = True

    if is_month_end(report_date):
        next_year, next_month_value = next_month(
            report_date.year,
            report_date.month,
        )
        next_key = month_key(next_year, next_month_value)
        next_consignment = {
            code: clean_number(current_consignments.get(code, 0))
            for code in planning_rows
        }
        if consign_state.get(next_key) != next_consignment:
            consign_state[next_key] = next_consignment
            state_changed = True

    result = {}
    for code, row in planning_rows.items():
        leadtime = leadtimes.get(code)
        if leadtime is None:
            raise RuntimeError(
                f"Chưa có Leadtime Danh_muc!J cho mã {code}."
            )

        calculated = calculate_row(
            fc=row["fc"],
            actual_stock=row["actual_stock"],
            book_stock=row["book_stock"],
            opening_consignment=float(
                opening_consignment.get(code, 0) or 0
            ),
            warehouse_debt=float(row["current_debt"] or 0),
            leadtime=leadtime,
            batch=row["batch"],
            per_shift=row["per_shift"],
            shifts_per_day=row["shifts_per_day"],
            classification=row["classification"],
            plan_year=plan_year,
            plan_month=plan_month,
        )
        result[code] = clamp_nonnegative_production(calculated)

    return result, state_changed, source_key, plan_key
