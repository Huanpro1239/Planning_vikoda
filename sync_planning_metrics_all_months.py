import calendar
from datetime import datetime, timedelta

import sync_planning_metrics as metrics
import sync_planning_metrics_compat as compat
import sync_planning_metrics_direct  # noqa: F401 - installs No kho / Leadtime runtime hooks


def _demand_days(plan_year, plan_month):
    days_in_month = calendar.monthrange(plan_year, plan_month)[1]
    return sum(
        1
        for day in range(1, days_in_month + 1)
        if datetime(plan_year, plan_month, day).weekday() != 6
    )


def calculate_row_all_months(
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

    demand_days = _demand_days(plan_year, plan_month)
    daily_fc = fc / demand_days if fc and demand_days else 0

    # M - tồn mục tiêu theo đúng số ngày nhu cầu của tháng.
    minimum_stock = daily_fc * (leadtime + 2)
    expected_end_stock = (
        0 if opening_consignment > minimum_stock else minimum_stock
    )

    # O - nhu cầu sản xuất.
    if warehouse_debt > 0:
        required_production = fc + warehouse_debt - book_stock
    else:
        required_production = fc + expected_end_stock - book_stock + warehouse_debt

    # P - làm tròn mẻ/ca.
    if required_production == 0:
        rounded_production = 0
    else:
        is_sugar = classification.casefold() == "có đường".casefold()
        base_qty = batch if is_sugar else per_shift
        rounded_production = (
            metrics.excel_roundup_integer(required_production / base_qty) * base_qty
        )

    production_days = rounded_production / per_shift / shifts_per_day

    # R - ngày sớm nhất cần bắt đầu sản xuất.
    if daily_fc <= 0:
        production_start = None
    else:
        first_day = datetime(plan_year, plan_month, 1)
        calculated = first_day + timedelta(
            days=(actual_stock / daily_fc) - leadtime
        )
        production_start = max(first_day, calculated)

    return {
        "expected_end_stock": metrics.clean_number(expected_end_stock),
        "warehouse_debt": metrics.clean_number(warehouse_debt),
        "required_production": metrics.clean_number(required_production),
        "rounded_production": metrics.clean_number(rounded_production),
        "production_days": metrics.clean_number(production_days),
        "production_start": production_start,
    }


# sync_planning_metrics_direct.calculate_metrics_from_no_kho gọi metrics.calculate_row ở runtime.
metrics.calculate_row = calculate_row_all_months


if __name__ == "__main__":
    compat.main_with_retry()
