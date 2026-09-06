import math
from collections import defaultdict

import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v5 as v5


MAX_EXACT_SEQUENCE_SKUS = 12


def _preferred_index(headers, product):
    preferred = product.get("preferred_date") or product.get("earliest_date")
    if preferred is None or preferred <= headers[0]:
        return 0
    for index, current_day in enumerate(headers):
        if current_day >= preferred:
            return index
    return len(headers)


def _production_by_day(headers, product, start_shift, capacity):
    duration = product["planned_qty"] / product["per_shift"]
    month_end = len(headers) * capacity
    actual_duration = min(duration, max(0.0, month_end - start_shift))
    end_shift = start_shift + actual_duration
    production = {day: 0.0 for day in headers}

    for index, current_day in enumerate(headers):
        day_start = index * capacity
        day_end = day_start + capacity
        overlap = max(
            0.0,
            min(end_shift, day_end) - max(start_shift, day_start),
        )
        if overlap > priority.base.EPSILON:
            production[current_day] = overlap * product["per_shift"]

    produced = actual_duration * product["per_shift"]
    carryover = max(0.0, product["planned_qty"] - produced)
    return production, carryover


def _product_sequence_cost(headers, product, start_shift, capacity):
    """Lexicographic service-level cost of scheduling one full campaign at start_shift."""
    production, carryover = _production_by_day(
        headers,
        product,
        start_shift,
        capacity,
    )
    stock = float(product.get("actual_stock", 0) or 0)
    target = max(float(product.get("target_stock", 0) or 0), 0.0)

    stockout_days = 0
    stockout_qty = 0.0
    safety_days = 0
    safety_qty = 0.0

    for current_day in headers:
        stock += production[current_day]
        stock -= float(product["demand_by_day"].get(current_day, 0) or 0)
        if stock < -priority.base.EPSILON:
            stockout_days += 1
            stockout_qty += -stock
        if stock < target - priority.base.EPSILON:
            safety_days += 1
            safety_qty += target - stock

    start_index = min(
        len(headers),
        int(math.floor(start_shift / capacity + priority.base.EPSILON)),
    )
    preferred_index = _preferred_index(headers, product)
    lateness = max(0, start_index - preferred_index)
    earliness = max(0, preferred_index - start_index)
    debt_delay = start_index if float(product.get("debt", 0) or 0) > 0 else 0

    # Thứ tự tuyệt đối: Stockout > Safety > Carryover > trễ kế hoạch > nợ > chạy sớm.
    return (
        stockout_days,
        round(stockout_qty, 6),
        safety_days,
        round(safety_qty, 6),
        round(carryover, 6),
        lateness,
        debt_delay,
        earliness,
        0,  # group-change được cộng ở transition
    )


def _add_cost(left, right):
    return tuple(a + b for a, b in zip(left, right))


def _transition_cost(previous, current):
    if previous is None:
        return (0, 0.0, 0, 0.0, 0.0, 0, 0, 0, 0)
    group_change = (
        0
        if previous.get("product_group") == current.get("product_group")
        else 1
    )
    return (0, 0.0, 0, 0.0, 0.0, 0, 0, 0, group_change)


def _static_greedy_key(headers, product):
    stock = float(product.get("actual_stock", 0) or 0)
    target = max(float(product.get("target_stock", 0) or 0), 0.0)
    first_stockout = None
    first_safety = None
    for index, current_day in enumerate(headers):
        stock -= float(product["demand_by_day"].get(current_day, 0) or 0)
        if first_safety is None and stock < target - priority.base.EPSILON:
            first_safety = index
        if first_stockout is None and stock < -priority.base.EPSILON:
            first_stockout = index

    return (
        0 if first_stockout is not None else 1,
        first_stockout if first_stockout is not None else len(headers) + 30,
        0 if first_safety is not None else 1,
        first_safety if first_safety is not None else len(headers) + 30,
        0 if float(product.get("debt", 0) or 0) > 0 else 1,
        _preferred_index(headers, product),
        product.get("row", 0),
    )


def optimize_sequence(headers, products, capacity):
    """
    Exact DP for <=12 SKUs. Start time of the next campaign depends only on the
    subset already scheduled, so subset-DP finds the globally best sequence.
    """
    if not products:
        return []

    if len(products) > MAX_EXACT_SEQUENCE_SKUS:
        # Scale-safe fallback; current Vikoda lines are below this threshold.
        return sorted(products, key=lambda product: _static_greedy_key(headers, product))

    n = len(products)
    durations = [product["planned_qty"] / product["per_shift"] for product in products]
    subset_duration = [0.0] * (1 << n)
    for mask in range(1, 1 << n):
        bit = mask & -mask
        index = bit.bit_length() - 1
        subset_duration[mask] = subset_duration[mask ^ bit] + durations[index]

    zero_cost = (0, 0.0, 0, 0.0, 0.0, 0, 0, 0, 0)
    states = {(0, -1): (zero_cost, ())}

    for depth in range(n):
        next_states = {}
        for (mask, last), (cost, order) in states.items():
            if mask.bit_count() != depth:
                continue

            # Có k SKU trước đó => đã có k-1 setup; trước SKU kế tiếp thêm 1 setup.
            count = mask.bit_count()
            start_shift = subset_duration[mask]
            if count > 0:
                start_shift += priority.base.SETUP_SHIFTS * count

            previous = products[last] if last >= 0 else None
            for index, product in enumerate(products):
                if mask & (1 << index):
                    continue
                product_cost = _product_sequence_cost(
                    headers,
                    product,
                    start_shift,
                    capacity,
                )
                transition = _transition_cost(previous, product)
                new_cost = _add_cost(_add_cost(cost, product_cost), transition)
                new_mask = mask | (1 << index)
                key = (new_mask, index)
                candidate = (new_cost, order + (index,))
                existing = next_states.get(key)
                if existing is None or candidate[0] < existing[0]:
                    next_states[key] = candidate

        # Giữ các state mới và các state sâu hơn nếu có.
        states.update(next_states)

    full_mask = (1 << n) - 1
    finalists = [
        value
        for (mask, _), value in states.items()
        if mask == full_mask
    ]
    if not finalists:
        raise RuntimeError("Sequence optimizer không tìm được thứ tự campaign.")

    _, best_order = min(finalists, key=lambda item: item[0])
    return [products[index] for index in best_order]


def allocate_sequence_optimized_line(headers, products, capacity):
    """Một SKU = một campaign. Tổng số đổi mã là tối thiểu (n-1)."""
    schedule = priority.base._empty_schedule(headers, products)
    usage = defaultdict(float)
    carryover = {}
    campaigns = []
    order = optimize_sequence(headers, products, capacity)

    cursor_shift = 0.0
    month_end_shift = len(headers) * capacity
    setup_total = 0.0
    group_changes = 0
    previous = None

    for product in order:
        if previous is not None:
            setup_duration = min(
                priority.base.SETUP_SHIFTS,
                max(0.0, month_end_shift - cursor_shift),
            )
            if setup_duration > priority.base.EPSILON:
                priority.base._add_interval_usage(
                    usage,
                    headers,
                    cursor_shift,
                    setup_duration,
                    capacity,
                )
                setup_total += setup_duration
            cursor_shift += priority.base.SETUP_SHIFTS
            if previous.get("product_group") != product.get("product_group"):
                group_changes += 1

        start_shift = cursor_shift
        required_duration = product["planned_qty"] / product["per_shift"]
        if start_shift >= month_end_shift - priority.base.EPSILON:
            carryover[product["code"]] = priority.base._clean_number(
                product["planned_qty"]
            )
            campaigns.append(
                {
                    "code": product["code"],
                    "start_shift": start_shift,
                    "end_shift": start_shift,
                    "scheduled_qty": 0,
                }
            )
            previous = product
            continue

        actual_duration = min(required_duration, month_end_shift - start_shift)
        produced = priority.base._add_production_interval(
            schedule,
            usage,
            headers,
            product,
            start_shift,
            actual_duration,
            capacity,
        )
        cursor_shift = start_shift + actual_duration

        missing = max(0.0, product["planned_qty"] - produced)
        if missing > 1e-6:
            carryover[product["code"]] = priority.base._clean_number(missing)

        campaigns.append(
            {
                "code": product["code"],
                "group": product.get("product_group"),
                "start_shift": start_shift,
                "end_shift": cursor_shift,
                "scheduled_qty": priority.base._clean_number(produced),
            }
        )
        previous = product

    return schedule, usage, carryover, {
        "mode": "global_sequence_stockout_first",
        "setup_shifts": setup_total,
        "group_changes": group_changes,
        "campaigns": campaigns,
        "sequence": [product["code"] for product in order],
    }


def validate_schedule_soft_deadline(headers, products, schedule, line_capacity, line_usage):
    for line, days in line_usage.items():
        capacity = line_capacity[line]
        for current_day, used in days.items():
            if used - capacity > 1e-6:
                raise RuntimeError(
                    f"Chuyền {line} ngày {current_day:%d/%m} dùng {used:.3f} ca > "
                    f"capacity {capacity:.3f}."
                )

    for product in products:
        total = sum(schedule[product["code"]].values())
        if total - product["planned_qty"] > 1e-5:
            raise RuntimeError(
                f"Mã {product['code']} được xếp {total} > P={product['planned_qty']}."
            )


def install_priority_scheduler_v6():
    v5.install_priority_scheduler_v5()
    priority.base._allocate_continuous_campaign_line = allocate_sequence_optimized_line
    priority.base.validate_schedule = validate_schedule_soft_deadline


if __name__ == "__main__":
    install_priority_scheduler_v6()
    priority.base.main_with_retry()
