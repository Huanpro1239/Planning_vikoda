import math
from collections import defaultdict

import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v2 as v2


MAX_REPLENISH_DAYS = 7


def _dynamic_risk(headers, product, schedule):
    stock = float(product.get("actual_stock", 0) or 0)
    target = max(float(product.get("target_stock", 0) or 0), 0.0)
    first_safety = None
    first_stockout = None

    for index, current_day in enumerate(headers):
        stock += float(schedule[product["code"]].get(current_day, 0) or 0)
        stock -= float(product["demand_by_day"].get(current_day, 0) or 0)
        if first_safety is None and stock < target - priority.base.EPSILON:
            first_safety = index
        if first_stockout is None and stock < -priority.base.EPSILON:
            first_stockout = index

    return first_stockout, first_safety


def _dynamic_priority_key(headers, product, schedule, current_index):
    first_stockout, first_safety = _dynamic_risk(headers, product, schedule)
    if first_stockout is not None:
        tier = 0
        deadline = first_stockout
    elif first_safety is not None:
        tier = 1
        deadline = first_safety
    elif float(product.get("debt", 0) or 0) > 0:
        tier = 2
        deadline = current_index
    else:
        tier = 3
        deadline = len(headers) + 7

    return (
        tier,
        deadline,
        0 if float(product.get("debt", 0) or 0) > 0 else 1,
        priority.base._earliest_index(headers, product),
        product.get("row", 0),
    )


def _choose_dynamic_candidate(
    headers,
    products,
    schedule,
    current_index,
    last_code,
    last_group,
):
    primary = min(
        products,
        key=lambda product: _dynamic_priority_key(
            headers, product, schedule, current_index
        ),
    )
    primary_key = _dynamic_priority_key(headers, primary, schedule, current_index)

    # Stockout hôm nay/ngày kế tiếp: tuyệt đối không continuity override.
    if primary_key[0] == 0 and primary_key[1] <= current_index + 1:
        return primary

    # Nếu đang chạy một SKU và mức rủi ro của nó gần tương đương SKU đầu bảng,
    # tiếp tục để tránh đổi mã vô ích.
    if last_code:
        same_code = next((p for p in products if p["code"] == last_code), None)
        if same_code is not None:
            same_key = _dynamic_priority_key(
                headers, same_code, schedule, current_index
            )
            if same_key[0] == primary_key[0] and same_key[1] <= primary_key[1] + 1:
                return same_code

    # Cùng nhóm chỉ được ưu tiên khi SKU primary chưa stockout trong 2 ngày tới.
    if last_group and not (primary_key[0] == 0 and primary_key[1] <= current_index + 2):
        same_group = [
            p for p in products if p.get("product_group") == last_group
        ]
        if same_group:
            return min(
                same_group,
                key=lambda product: _dynamic_priority_key(
                    headers, product, schedule, current_index
                ),
            )

    return primary


def _required_units_through(
    headers,
    product,
    schedule,
    horizon_index,
    remaining_units,
):
    stock = float(product.get("actual_stock", 0) or 0)
    target = max(float(product.get("target_stock", 0) or 0), 0.0)

    for index in range(0, min(horizon_index, len(headers) - 1) + 1):
        current_day = headers[index]
        stock += float(schedule[product["code"]].get(current_day, 0) or 0)
        stock -= float(product["demand_by_day"].get(current_day, 0) or 0)

    shortage_qty = max(0.0, target - stock)
    quantum = float(product["quantum_qty"] or 0)
    if quantum <= 0:
        return 0

    units = int(math.ceil(shortage_qty / quantum - priority.base.EPSILON))
    if units <= 0:
        units = 1
    return min(remaining_units, units)


def allocate_continuous_stockout_first(headers, products, capacity):
    """
    KHS/PET: campaign liên tục nhưng rolling-replenishment.
    Chỉ ngắt campaign dài khi SKU khác có stockout gần hơn; sau mỗi block tính lại tồn.
    """
    schedule = priority.base._empty_schedule(headers, products)
    usage = defaultdict(float)
    carryover = {}
    campaigns = []
    remaining_units = {
        product["code"]: int(product["required_units"])
        for product in products
    }
    by_code = {product["code"]: product for product in products}

    cursor_shift = 0.0
    month_end_shift = len(headers) * capacity
    last_code = None
    last_group = None
    setup_total = 0.0

    while any(units > 0 for units in remaining_units.values()):
        current_index = min(
            len(headers),
            int(math.floor(cursor_shift / capacity + priority.base.EPSILON)),
        )
        active = [
            product
            for product in products
            if remaining_units[product["code"]] > 0
            and priority.base._earliest_index(headers, product) <= current_index
        ]

        if not active:
            future = [
                priority.base._earliest_index(headers, product)
                for product in products
                if remaining_units[product["code"]] > 0
            ]
            if not future:
                break
            next_index = min(future)
            cursor_shift = max(cursor_shift, next_index * capacity)
            if cursor_shift >= month_end_shift - priority.base.EPSILON:
                break
            continue

        product = _choose_dynamic_candidate(
            headers,
            active,
            schedule,
            current_index,
            last_code,
            last_group,
        )
        code = product["code"]

        if last_code is not None and code != last_code:
            if cursor_shift + priority.base.SETUP_SHIFTS >= month_end_shift + priority.base.EPSILON:
                break
            setup_duration = min(
                priority.base.SETUP_SHIFTS,
                max(0.0, month_end_shift - cursor_shift),
            )
            priority.base._add_interval_usage(
                usage,
                headers,
                cursor_shift,
                setup_duration,
                capacity,
            )
            setup_total += setup_duration
            cursor_shift += priority.base.SETUP_SHIFTS

        earliest_shift = priority.base._earliest_index(headers, product) * capacity
        cursor_shift = max(cursor_shift, earliest_shift)
        if cursor_shift >= month_end_shift - priority.base.EPSILON:
            break

        product_units_left = remaining_units[code]
        full_duration = product_units_left * product["quantum_shift"]
        chunk_units = product_units_left

        # Nếu chạy hết SKU này sẽ làm một SKU khác stockout trước khi campaign kết thúc,
        # chỉ sản xuất đủ để bảo vệ SKU hiện tại tới horizon kế tiếp rồi đánh giá lại.
        full_finish = cursor_shift + full_duration
        competitor_stockouts = []
        for competitor in active:
            if competitor["code"] == code:
                continue
            stockout, _ = _dynamic_risk(headers, competitor, schedule)
            if stockout is not None:
                competitor_stockouts.append(stockout)

        if competitor_stockouts:
            next_stockout = min(competitor_stockouts)
            competitor_deadline_shift = (next_stockout + 1) * capacity
            if full_finish > competitor_deadline_shift + priority.base.EPSILON:
                horizon = min(
                    len(headers) - 1,
                    current_index + MAX_REPLENISH_DAYS,
                    next_stockout,
                )
                chunk_units = _required_units_through(
                    headers,
                    product,
                    schedule,
                    horizon,
                    product_units_left,
                )

        chunk_units = max(1, min(product_units_left, int(chunk_units)))
        duration = chunk_units * product["quantum_shift"]
        if cursor_shift + duration > month_end_shift + priority.base.EPSILON:
            available = max(0.0, month_end_shift - cursor_shift)
            chunk_units = int(
                math.floor(available / product["quantum_shift"] + priority.base.EPSILON)
            )
            if chunk_units <= 0:
                break
            duration = chunk_units * product["quantum_shift"]

        produced = priority.base._add_production_interval(
            schedule,
            usage,
            headers,
            product,
            cursor_shift,
            duration,
            capacity,
        )
        expected_qty = chunk_units * product["quantum_qty"]
        if not math.isclose(produced, expected_qty, rel_tol=1e-9, abs_tol=1e-5):
            raise RuntimeError(
                f"Mã {code} block dự kiến {expected_qty} nhưng ghi {produced}."
            )

        campaigns.append(
            {
                "code": code,
                "group": product["product_group"],
                "start_shift": cursor_shift,
                "end_shift": cursor_shift + duration,
                "scheduled_qty": priority.base._clean_number(produced),
            }
        )
        remaining_units[code] -= chunk_units
        cursor_shift += duration
        last_code = code
        last_group = product["product_group"]

    for code, units in remaining_units.items():
        if units > 0:
            carryover[code] = priority.base._clean_number(
                units * by_code[code]["quantum_qty"]
            )

    return schedule, usage, carryover, {
        "mode": "continuous_stockout_first",
        "setup_shifts": setup_total,
        "campaigns": campaigns,
    }


def install_priority_scheduler_v3():
    v2.install_priority_scheduler_v2()
    priority.base._allocate_continuous_campaign_line = allocate_continuous_stockout_first


if __name__ == "__main__":
    install_priority_scheduler_v3()
    priority.base.main_with_retry()
