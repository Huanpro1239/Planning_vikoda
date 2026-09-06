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


def resolve_plan_year(report_date, plan_month):
    year = report_date.year
    delta = plan_month - report_date.month
    if delta <= -6:
        year += 1
    elif delta >= 6:
        year -= 1
    return year


def _planning_stock(actual_stock, book_stock):
    """Tồn dùng cho kế hoạch thường: lấy mức bảo thủ hơn giữa thực tế và sổ sách."""
    return min(max(float(actual_stock or 0), 0.0), max(float(book_stock or 0), 0.0))


def _urgent_supply_need(fc, warehouse_debt, actual_stock):
    """
    Nhu cầu bắt buộc để không thiếu cung trong tháng.

    Khi FC + nợ kho lớn hơn tồn thực tế J, SKU được xem là urgent. Phần urgent
    chỉ bù đủ cung ứng thực tế, không cộng tồn cuối mục tiêu M. Nhờ vậy capacity
    máy được ưu tiên cho stockout prevention trước safety-stock build.
    """
    return max(
        float(fc or 0) + float(warehouse_debt or 0) - max(float(actual_stock or 0), 0.0),
        0.0,
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

    planning_stock = _planning_stock(actual_stock, book_stock)
    urgent_supply_need = _urgent_supply_need(fc, warehouse_debt, actual_stock)
    is_urgent = urgent_supply_need > 0

    # O - SKU urgent ưu tiên đủ cung ứng trước, bỏ qua M.
    # SKU chưa urgent mới build tồn cuối mục tiêu theo Planning Stock bảo thủ.
    if is_urgent:
        required_production = urgent_supply_need
    else:
        required_production = (
            float(fc or 0)
            + float(warehouse_debt or 0)
            + float(expected_end_stock or 0)
            - planning_stock
        )

    required_production = max(required_production, 0)

    # P - làm tròn mẻ/ca.
    if required_production <= 0:
        rounded_production = 0
    else:
        is_sugar = classification.casefold() == "có đường".casefold()
        base_qty = batch if is_sugar else per_shift
        rounded_production = (
            metrics.excel_roundup_integer(required_production / base_qty) * base_qty
        )
        rounded_production = max(rounded_production, 0)

    production_days = rounded_production / per_shift / shifts_per_day

    # R - ngày nhu cầu danh nghĩa; scheduler có thể kéo sớm hơn nếu risk/capacity yêu cầu.
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


# sync_planning_metrics_direct.calculate_metrics_from_no_kho gọi các hàm metrics.* ở runtime.
metrics._resolve_plan_year = resolve_plan_year
metrics.calculate_row = calculate_row_all_months


if __name__ == "__main__":
    compat.main_with_retry()
