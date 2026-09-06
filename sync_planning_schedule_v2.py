import math
from collections import defaultdict
from datetime import date, timedelta

import sync_planning_schedule as base


RISK_GROUP_TOLERANCE_DAYS = 2
FAR_FUTURE = date(9999, 12, 31)


def _compute_risk_profile(headers, product):
    """Tính ngày chạm tồn mục tiêu và ngày tồn âm nếu chưa có SX bổ sung."""
    stock = float(product["actual_stock"])
    first_below_safety = None
    first_stockout = None

    for current_day in headers:
        stock -= float(product["demand_by_day"].get(current_day, 0) or 0)
        if (
            first_below_safety is None
            and stock < float(product["target_stock"]) - base.EPSILON
        ):
            first_below_safety = current_day
        if first_stockout is None and stock < -base.EPSILON:
            first_stockout = current_day

    critical_date = first_below_safety or first_stockout
    return {
        "critical_date": critical_date,
        "stockout_date": first_stockout,
        "projected_without_production": stock,
    }


def _enrich_risk(headers, products):
    for product in products:
        profile = _compute_risk_profile(headers, product)
        product.update(profile)


def _risk_date(product):
    return product.get("critical_date") or FAR_FUTURE


def _stockout_date(product):
    return product.get("stockout_date") or FAR_FUTURE


def _priority_key(product):
    """Thiếu hàng/safety trước; nợ kho và R chỉ là tie-break sau risk."""
    return (
        _risk_date(product),
        _stockout_date(product),
        0 if product.get("debt", 0) > 0 else 1,
        product.get("earliest_date") or FAR_FUTURE,
        product.get("row", 10**9),
    )


def _within_group_tolerance(product, best):
    best_risk = _risk_date(best)
    risk = _risk_date(product)
    if best_risk == FAR_FUTURE:
        return True
    if risk == FAR_FUTURE:
        return False
    return (risk - best_risk).days <= RISK_GROUP_TOLERANCE_DAYS


def _campaign_candidate(remaining, headers, cursor_shift, capacity, last_group):
    """EDD theo Risk Date, có look-ahead và chỉ gom nhóm khi risk gần tương đương."""
    if not remaining:
        return None, cursor_shift

    current_index = min(len(headers), int(math.floor(cursor_shift / capacity + base.EPSILON)))
    eligible = [
        product
        for product in remaining
        if base._earliest_index(headers, product) <= current_index
    ]

    if not eligible:
        next_index = min(base._earliest_index(headers, product) for product in remaining)
        cursor_shift = max(cursor_shift, next_index * capacity)
        current_index = next_index
        eligible = [
            product
            for product in remaining
            if base._earliest_index(headers, product) <= current_index
        ]

    eligible.sort(key=_priority_key)
    best = eligible[0]

    # Nếu một SKU rất cấp bách sẽ release trong vài ngày tới, không cho một
    # campaign dài ít cấp bách chiếm chuyền xuyên qua deadline của nó.
    future = [
        product
        for product in remaining
        if base._earliest_index(headers, product) > current_index
    ]
    if future:
        future.sort(key=_priority_key)
        urgent = future[0]
        if _priority_key(urgent) < _priority_key(best):
            urgent_release = base._earliest_index(headers, urgent) * capacity
            urgent_deadline_index = len(headers)
            if _risk_date(urgent) != FAR_FUTURE:
                urgent_deadline_index = max(
                    base._earliest_index(headers, urgent),
                    min(len(headers), ( _risk_date(urgent) - headers[0] ).days),
                )
            urgent_deadline = urgent_deadline_index * capacity
            candidate_duration = best["planned_qty"] / best["per_shift"]
            candidate_finish = cursor_shift + candidate_duration + base.SETUP_SHIFTS
            if candidate_finish > max(urgent_release, urgent_deadline):
                return urgent, max(cursor_shift, urgent_release)

    # Giảm thay khuôn/đổi nhóm chỉ khi không hy sinh SKU cấp bách hơn quá 2 ngày.
    if last_group:
        same_group = [
            product
            for product in eligible
            if product.get("product_group") == last_group
            and _within_group_tolerance(product, best)
        ]
        if same_group:
            same_group.sort(key=_priority_key)
            best = same_group[0]

    return best, cursor_shift


def _risk_aware_weekly_unit_targets(headers, product):
    """Chia tuần nhưng front-load đủ lượng để tránh chạm safety trước cuối tuần."""
    buckets = base._month_week_buckets(headers)
    active = [
        bucket
        for bucket in buckets
        if headers[bucket[1]] >= (product["earliest_date"] or headers[0])
    ]
    if not active:
        return {}

    total_units = product["required_units"]
    if total_units <= 0:
        return {}

    cumulative_demand = {}
    running = 0.0
    for current_day in headers:
        running += float(product["demand_by_day"].get(current_day, 0) or 0)
        cumulative_demand[current_day] = running

    targets = {}
    previous_cumulative = 0
    for ordinal, bucket in enumerate(active, start=1):
        _, end_index = bucket
        end_day = headers[end_index]

        even_cumulative = int(math.ceil(total_units * ordinal / len(active) - base.EPSILON))
        qty_needed_for_safety = max(
            0.0,
            cumulative_demand[end_day]
            + float(product["target_stock"])
            - float(product["actual_stock"]),
        )
        safety_cumulative = int(
            math.ceil(qty_needed_for_safety / product["quantum_qty"] - base.EPSILON)
        )
        desired_cumulative = min(total_units, max(even_cumulative, safety_cumulative))
        desired_cumulative = max(desired_cumulative, previous_cumulative)
        targets[bucket] = desired_cumulative - previous_cumulative
        previous_cumulative = desired_cumulative

    if previous_cumulative < total_units:
        last_bucket = active[-1]
        targets[last_bucket] = targets.get(last_bucket, 0) + total_units - previous_cumulative

    return targets


def _weekly_candidate(products, pending, last_group):
    candidates = [product for product in products if pending.get(product["code"], 0) > 0]
    if not candidates:
        return None
    candidates.sort(key=_priority_key)
    best = candidates[0]
    if last_group:
        same_group = [
            product
            for product in candidates
            if product.get("product_group") == last_group
            and _within_group_tolerance(product, best)
        ]
        if same_group:
            same_group.sort(key=_priority_key)
            best = same_group[0]
    return best


def _allocate_weekly_rgb_line(headers, products, capacity):
    schedule = base._empty_schedule(headers, products)
    usage = defaultdict(float)
    remaining_capacity = base._new_remaining(headers, capacity)
    weekly_targets = {
        product["code"]: _risk_aware_weekly_unit_targets(headers, product)
        for product in products
    }
    pending = defaultdict(int)
    last_code = None
    last_group = None
    setup_total = 0.0
    campaigns = []

    for bucket in base._month_week_buckets(headers):
        start_index, end_index = bucket
        for product in products:
            pending[product["code"]] += weekly_targets[product["code"]].get(bucket, 0)

        while True:
            product = _weekly_candidate(products, pending, last_group)
            if product is None:
                break

            campaign_start = max(start_index, base._earliest_index(headers, product))
            if campaign_start > end_index:
                break

            if last_code is not None and last_code != product["code"]:
                setup_index, ok = base._consume_setup(
                    remaining_capacity,
                    usage,
                    headers,
                    campaign_start,
                    allow_sunday=True,
                    max_index=end_index,
                )
                if not ok:
                    break
                setup_total += base.SETUP_SHIFTS
                campaign_start = setup_index

            requested = pending[product["code"]]
            placed, _ = base._place_units(
                schedule,
                usage,
                remaining_capacity,
                headers,
                product,
                requested,
                campaign_start,
                allow_sunday=True,
                max_index=end_index,
            )
            if placed <= 0:
                break

            pending[product["code"]] -= placed
            campaigns.append(
                {
                    "code": product["code"],
                    "week": f"{headers[start_index]:%d/%m}-{headers[end_index]:%d/%m}",
                    "units": placed,
                    "risk_date": None if _risk_date(product) == FAR_FUTURE else _risk_date(product),
                }
            )
            last_code = product["code"]
            last_group = product["product_group"]

    # Phần còn thiếu dùng toàn bộ capacity còn lại, CN là ngày SX bình thường.
    while True:
        product = _weekly_candidate(products, pending, last_group)
        if product is None:
            break
        start_index = base._earliest_index(headers, product)
        if start_index >= len(headers):
            break

        if last_code is not None and last_code != product["code"]:
            setup_index, ok = base._consume_setup(
                remaining_capacity,
                usage,
                headers,
                start_index,
                allow_sunday=True,
            )
            if not ok:
                break
            setup_total += base.SETUP_SHIFTS
            start_index = setup_index

        placed, _ = base._place_units(
            schedule,
            usage,
            remaining_capacity,
            headers,
            product,
            pending[product["code"]],
            start_index,
            allow_sunday=True,
        )
        if placed <= 0:
            break
        pending[product["code"]] -= placed
        last_code = product["code"]
        last_group = product["product_group"]

    carryover = {}
    for product in products:
        missing = pending[product["code"]]
        if missing > 0:
            carryover[product["code"]] = base._clean_number(
                missing * product["quantum_qty"]
            )

    return schedule, usage, carryover, {
        "mode": "weekly_rgb_risk_aware",
        "setup_shifts": setup_total,
        "campaigns": campaigns,
    }


def _galon_spread_daily_plan(headers, product):
    """Rải đều 19L trên mọi ngày SX, nhưng front-load nếu tồn sắp chạm safety."""
    eligible = [
        index
        for index in range(len(headers))
        if index >= base._earliest_index(headers, product)
    ]
    if not eligible or product["planned_qty"] <= 0:
        return {}

    total_units = int(round(product["planned_qty"] / product["per_shift"]))
    cumulative_demand = 0.0
    previous_target = 0
    plan = {index: 0 for index in eligible}

    for pos, index in enumerate(eligible, start=1):
        current_day = headers[index]
        cumulative_demand += float(product["demand_by_day"].get(current_day, 0) or 0)
        even_target = int(math.ceil(total_units * pos / len(eligible) - base.EPSILON))
        safety_qty = max(
            0.0,
            cumulative_demand + float(product["target_stock"]) - float(product["actual_stock"]),
        )
        safety_target = int(math.ceil(safety_qty / product["per_shift"] - base.EPSILON))
        cumulative_target = min(total_units, max(previous_target, even_target, safety_target))
        plan[index] = cumulative_target - previous_target
        previous_target = cumulative_target

    if previous_target < total_units:
        plan[eligible[-1]] += total_units - previous_target
    return plan


def _allocate_galon_line(headers, products, capacity):
    schedule = base._empty_schedule(headers, products)
    usage = defaultdict(float)
    remaining_capacity = base._new_remaining(headers, capacity)
    carryover = {}
    campaigns = []
    setup_total = 0.0
    last_code = None

    spread = next(
        (product for product in products if product["code"] == base.GALON_SPREAD_CODE),
        None,
    )
    others = [product for product in products if product["code"] != base.GALON_SPREAD_CODE]

    if spread is not None:
        daily_plan = _galon_spread_daily_plan(headers, spread)
        remaining_units = spread["required_units"]
        for index in sorted(daily_plan):
            requested = min(daily_plan[index], remaining_units)
            if requested <= 0:
                continue
            placed, _ = base._place_units(
                schedule,
                usage,
                remaining_capacity,
                headers,
                spread,
                requested,
                index,
                allow_sunday=True,
                max_index=index,
            )
            remaining_units -= placed

        if remaining_units > 0:
            placed, _ = base._place_units(
                schedule,
                usage,
                remaining_capacity,
                headers,
                spread,
                remaining_units,
                base._earliest_index(headers, spread),
                allow_sunday=True,
            )
            remaining_units -= placed

        if remaining_units > 0:
            carryover[spread["code"]] = base._clean_number(
                remaining_units * spread["quantum_qty"]
            )
        last_code = spread["code"]
        campaigns.append({"code": spread["code"], "mode": "daily_risk_spread"})

    others.sort(key=_priority_key)
    last_group = spread["product_group"] if spread else None
    while others:
        best = others[0]
        if last_group:
            same_group = [
                item
                for item in others
                if item["product_group"] == last_group
                and _within_group_tolerance(item, best)
            ]
            if same_group:
                same_group.sort(key=_priority_key)
                best = same_group[0]
        others.remove(best)

        start_index = base._earliest_index(headers, best)
        if last_code is not None and last_code != best["code"]:
            setup_index, ok = base._consume_setup(
                remaining_capacity,
                usage,
                headers,
                start_index,
                allow_sunday=True,
            )
            if ok:
                setup_total += base.SETUP_SHIFTS
                start_index = setup_index

        placed, _ = base._place_units(
            schedule,
            usage,
            remaining_capacity,
            headers,
            best,
            best["required_units"],
            start_index,
            allow_sunday=True,
        )
        missing = best["required_units"] - placed
        if missing > 0:
            carryover[best["code"]] = base._clean_number(
                missing * best["quantum_qty"]
            )
        if placed > 0:
            last_code = best["code"]
            last_group = best["product_group"]
            campaigns.append(
                {
                    "code": best["code"],
                    "mode": "risk_continuous_remaining_capacity",
                    "risk_date": None if _risk_date(best) == FAR_FUTURE else _risk_date(best),
                }
            )

    return schedule, usage, carryover, {
        "mode": "galon_risk_hybrid",
        "setup_shifts": setup_total,
        "campaigns": campaigns,
    }


def build_schedule(headers, products):
    _enrich_risk(headers, products)
    schedule = base._empty_schedule(headers, products)
    active_products = [product for product in products if product["planned_qty"] > 0]

    by_line = defaultdict(list)
    for product in active_products:
        by_line[product["line"]].append(product)

    line_capacity = {}
    line_usage = defaultdict(lambda: defaultdict(float))
    carryover = {}
    optimizer_meta = {}

    # Patch candidate để continuous allocator của base dùng Risk Date.
    original_candidate = base._campaign_candidate
    base._campaign_candidate = _campaign_candidate
    try:
        for line, line_products in sorted(by_line.items()):
            if line in base.CONTINUOUS_LINES:
                capacity = min(product["max_shifts_per_day"] for product in line_products)
                line_schedule, usage, line_carryover, meta = base._allocate_continuous_campaign_line(
                    headers, line_products, capacity
                )
                meta["mode"] = "continuous_campaign_risk_aware"
            elif line in base.WEEKLY_LINES:
                capacity = max(product["max_shifts_per_day"] for product in line_products)
                line_schedule, usage, line_carryover, meta = _allocate_weekly_rgb_line(
                    headers, line_products, capacity
                )
            elif line == base.GALON_LINE:
                capacity = max(product["max_shifts_per_day"] for product in line_products)
                line_schedule, usage, line_carryover, meta = _allocate_galon_line(
                    headers, line_products, capacity
                )
            else:
                capacity = min(product["max_shifts_per_day"] for product in line_products)
                line_schedule, usage, line_carryover, meta = base._allocate_continuous_campaign_line(
                    headers, line_products, capacity
                )
                meta["mode"] = "generic_continuous_risk_aware"

            line_capacity[line] = capacity
            optimizer_meta[line] = meta
            carryover.update(line_carryover)

            for current_day, used_shift in usage.items():
                line_usage[line][current_day] = used_shift
            for product in line_products:
                code = product["code"]
                for current_day in headers:
                    schedule[code][current_day] = line_schedule[code][current_day]
    finally:
        base._campaign_candidate = original_candidate

    base.validate_schedule(headers, products, schedule, line_capacity, line_usage)
    inventory = base.simulate_inventory(headers, products, schedule)
    return schedule, line_capacity, line_usage, carryover, inventory, optimizer_meta


# prepare_schedule_update/main_with_retry của module gốc tra global build_schedule runtime.
base.build_schedule = build_schedule


if __name__ == "__main__":
    base.main_with_retry()
