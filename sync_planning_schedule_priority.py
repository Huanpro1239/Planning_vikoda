import math
from collections import defaultdict

import sync_planning_schedule as base


# Giữ campaign liền mạch nếu SKU khẩn cấp vẫn kịp trước ngày rủi ro.
# Nếu không, ưu tiên SKU sắp thiếu hàng và chấp nhận đổi nhóm/mã.


def _risk_snapshot(headers, product):
    stock = float(product.get("actual_stock", 0) or 0)
    target = max(float(product.get("target_stock", 0) or 0), 0.0)
    first_safety = None
    first_stockout = None

    for index, current_day in enumerate(headers):
        stock -= float(product["demand_by_day"].get(current_day, 0) or 0)
        if first_safety is None and stock < target - base.EPSILON:
            first_safety = index
        if first_stockout is None and stock < -base.EPSILON:
            first_stockout = index

    candidates = [
        index for index in (first_safety, first_stockout) if index is not None
    ]
    deadline_index = min(candidates) if candidates else None
    severity = 2
    if deadline_index is not None:
        if first_stockout == deadline_index:
            severity = 0
        elif first_safety == deadline_index:
            severity = 1

    return {
        "first_safety": first_safety,
        "first_stockout": first_stockout,
        "deadline_index": deadline_index,
        "severity": severity,
    }


def _processing_shifts(product):
    per_shift = float(product.get("per_shift", 0) or 0)
    if per_shift <= 0:
        return math.inf
    return float(product.get("planned_qty", 0) or 0) / per_shift


def _priority_key(headers, product, cursor_shift, capacity, last_group=None):
    risk = _risk_snapshot(headers, product)
    earliest_index = base._earliest_index(headers, product)
    setup = base.SETUP_SHIFTS if last_group is not None else 0.0
    start = max(cursor_shift + setup, earliest_index * capacity)
    finish = start + _processing_shifts(product)

    if risk["deadline_index"] is not None:
        tier = 0
        deadline_shift = (risk["deadline_index"] + 1) * capacity
    elif float(product.get("debt", 0) or 0) > 0:
        tier = 1
        # Không có stock-risk nhưng có nợ kho: ưu tiên ngay sau nhóm thiếu hàng.
        deadline_shift = max(start, earliest_index * capacity) + capacity
    else:
        tier = 2
        deadline_shift = (len(headers) + 7) * capacity

    slack = deadline_shift - finish
    return (
        tier,
        slack,
        risk["deadline_index"] if risk["deadline_index"] is not None else len(headers) + 7,
        risk["severity"],
        0 if float(product.get("debt", 0) or 0) > 0 else 1,
        earliest_index,
        product.get("row", 0),
    )


def _can_keep_same_group(headers, candidate, primary, cursor_shift, capacity, last_group):
    if candidate["code"] == primary["code"]:
        return True

    primary_risk = _risk_snapshot(headers, primary)
    if primary_risk["deadline_index"] is None:
        return True

    setup_before_candidate = base.SETUP_SHIFTS if last_group is not None else 0.0
    candidate_start = max(
        cursor_shift + setup_before_candidate,
        base._earliest_index(headers, candidate) * capacity,
    )
    candidate_finish = candidate_start + _processing_shifts(candidate)

    primary_start = max(
        candidate_finish + base.SETUP_SHIFTS,
        base._earliest_index(headers, primary) * capacity,
    )
    primary_finish = primary_start + _processing_shifts(primary)
    primary_deadline = (primary_risk["deadline_index"] + 1) * capacity

    return primary_finish <= primary_deadline + base.EPSILON


def shortage_first_campaign_candidate(remaining, headers, cursor_shift, capacity, last_group):
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

    primary = min(
        eligible,
        key=lambda product: _priority_key(
            headers, product, cursor_shift, capacity, last_group
        ),
    )

    # Chỉ giữ cùng nhóm khi việc đó không làm SKU khẩn cấp trễ deadline tồn kho.
    if last_group:
        same_group = [
            product for product in eligible
            if product.get("product_group") == last_group
        ]
        if same_group and primary not in same_group:
            candidate = min(
                same_group,
                key=lambda product: _priority_key(
                    headers, product, cursor_shift, capacity, last_group
                ),
            )
            if _can_keep_same_group(
                headers,
                candidate,
                primary,
                cursor_shift,
                capacity,
                last_group,
            ):
                return candidate, cursor_shift

    return primary, cursor_shift


def shortage_aware_weekly_unit_targets(headers, product):
    buckets = base._month_week_buckets(headers)
    active = [
        bucket for bucket in buckets
        if headers[bucket[1]] >= (product["earliest_date"] or headers[0])
    ]
    if not active:
        return {}

    total_units = product["required_units"]
    quantum_qty = float(product["quantum_qty"] or 0)
    target_stock = max(float(product.get("target_stock", 0) or 0), 0.0)
    opening_stock = float(product.get("actual_stock", 0) or 0)

    targets = {}
    previous_cumulative = 0
    cumulative_demand = 0.0
    demand_by_index = []
    for current_day in headers:
        cumulative_demand += float(product["demand_by_day"].get(current_day, 0) or 0)
        demand_by_index.append(cumulative_demand)

    for ordinal, bucket in enumerate(active, start=1):
        _, end_index = bucket
        even_cumulative = int(math.ceil(total_units * ordinal / len(active) - base.EPSILON))

        required_qty = max(
            0.0,
            demand_by_index[end_index] + target_stock - opening_stock,
        )
        required_units = (
            int(math.ceil(required_qty / quantum_qty - base.EPSILON))
            if quantum_qty > 0
            else 0
        )

        cumulative_target = min(
            total_units,
            max(previous_cumulative, even_cumulative, required_units),
        )
        targets[bucket] = cumulative_target - previous_cumulative
        previous_cumulative = cumulative_target

    # Bảo đảm tổng target đúng P ngay cả khi active bucket bị thay đổi do R.
    if previous_cumulative < total_units:
        last_bucket = active[-1]
        targets[last_bucket] = targets.get(last_bucket, 0) + (total_units - previous_cumulative)

    return targets


def _weekly_candidate(headers, products, start_index, end_index, capacity, last_group):
    cursor_shift = start_index * capacity
    primary = min(
        products,
        key=lambda product: _priority_key(
            headers, product, cursor_shift, capacity, last_group
        ),
    )
    risk = _risk_snapshot(headers, primary)

    # Nếu SKU khẩn cấp sẽ chạm safety/stockout ngay trong tuần này thì không gom nhóm thay nó.
    if risk["deadline_index"] is not None and risk["deadline_index"] <= end_index:
        return primary

    if last_group:
        same_group = [
            product for product in products
            if product.get("product_group") == last_group
        ]
        if same_group:
            candidate = min(
                same_group,
                key=lambda product: _priority_key(
                    headers, product, cursor_shift, capacity, last_group
                ),
            )
            if _can_keep_same_group(
                headers,
                candidate,
                primary,
                cursor_shift,
                capacity,
                last_group,
            ):
                return candidate

    return primary


def allocate_weekly_rgb_priority(headers, products, capacity):
    """RGB: chia theo tuần nhưng front-load SKU sắp thiếu và Chủ nhật là ngày SX bình thường."""
    schedule = base._empty_schedule(headers, products)
    usage = defaultdict(float)
    remaining_capacity = base._new_remaining(headers, capacity)
    weekly_targets = {
        product["code"]: shortage_aware_weekly_unit_targets(headers, product)
        for product in products
    }
    pending = defaultdict(int)
    last_code = None
    last_group = None
    setup_total = 0.0
    campaigns = []

    for bucket in base._month_week_buckets(headers):
        start_index, end_index = bucket
        weekly_units = {}
        for product in products:
            pending[product["code"]] += weekly_targets[product["code"]].get(bucket, 0)
            if pending[product["code"]] > 0:
                weekly_units[product["code"]] = pending[product["code"]]

        remaining_products = [
            product for product in products
            if weekly_units.get(product["code"], 0) > 0
        ]

        while remaining_products:
            product = _weekly_candidate(
                headers,
                remaining_products,
                start_index,
                end_index,
                capacity,
                last_group,
            )
            remaining_products.remove(product)
            campaign_start = start_index

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
                    continue
                setup_total += base.SETUP_SHIFTS
                campaign_start = setup_index

            requested = weekly_units[product["code"]]
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
            pending[product["code"]] -= placed
            if placed > 0:
                campaigns.append(
                    {
                        "code": product["code"],
                        "week": f"{headers[start_index]:%d/%m}-{headers[end_index]:%d/%m}",
                        "units": placed,
                    }
                )
                last_code = product["code"]
                last_group = product["product_group"]

    # Phần còn lại: vẫn shortage-first, dùng toàn bộ ngày kể cả Chủ nhật.
    pending_products = [
        product for product in products if pending[product["code"]] > 0
    ]
    while pending_products:
        product, _ = shortage_first_campaign_candidate(
            pending_products,
            headers,
            0.0,
            capacity,
            last_group,
        )
        pending_products.remove(product)
        start_index = base._earliest_index(headers, product)

        if last_code is not None and last_code != product["code"]:
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
            product,
            pending[product["code"]],
            start_index,
            allow_sunday=True,
        )
        pending[product["code"]] -= placed
        if placed > 0:
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
        "mode": "weekly_rgb_shortage_first",
        "setup_shifts": setup_total,
        "campaigns": campaigns,
    }


def allocate_galon_priority(headers, products, capacity):
    """19L rải đều nhưng tự front-load nếu tồn sắp thiếu; các SKU khác shortage-first."""
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
    others = [
        product for product in products if product["code"] != base.GALON_SPREAD_CODE
    ]

    if spread is not None and spread["planned_qty"] > base.EPSILON:
        eligible = list(range(base._earliest_index(headers, spread), len(headers)))
        total_full_units = int(
            math.floor(spread["planned_qty"] / spread["per_shift"] + base.EPSILON)
        )
        partial_qty = spread["planned_qty"] - total_full_units * spread["per_shift"]
        produced_units = 0
        cumulative_demand = 0.0
        demand_by_index = []
        for current_day in headers:
            cumulative_demand += float(spread["demand_by_day"].get(current_day, 0) or 0)
            demand_by_index.append(cumulative_demand)

        for ordinal, index in enumerate(eligible, start=1):
            if total_full_units <= produced_units:
                break

            even_cumulative = int(
                math.ceil(total_full_units * ordinal / len(eligible) - base.EPSILON)
            )
            required_qty = max(
                0.0,
                demand_by_index[index]
                + max(float(spread.get("target_stock", 0) or 0), 0.0)
                - float(spread.get("actual_stock", 0) or 0),
            )
            required_cumulative = int(
                math.ceil(required_qty / spread["per_shift"] - base.EPSILON)
            )
            desired_cumulative = min(
                total_full_units,
                max(even_cumulative, required_cumulative),
            )
            units_today = max(0, desired_cumulative - produced_units)

            for _ in range(units_today):
                made = base._place_fractional_shift(
                    schedule,
                    usage,
                    remaining_capacity,
                    headers,
                    spread,
                    index,
                    spread["per_shift"],
                )
                if made <= base.EPSILON:
                    break
                produced_units += 1

        # Nếu một ngày bị giới hạn capacity, quét lại các ngày còn trống để đủ tổng P.
        missing_units = total_full_units - produced_units
        if missing_units > 0:
            for index in eligible:
                while missing_units > 0:
                    made = base._place_fractional_shift(
                        schedule,
                        usage,
                        remaining_capacity,
                        headers,
                        spread,
                        index,
                        spread["per_shift"],
                    )
                    if made <= base.EPSILON:
                        break
                    produced_units += 1
                    missing_units -= 1
                if missing_units <= 0:
                    break

        if partial_qty > base.EPSILON:
            placed_partial = False
            for index in reversed(eligible):
                made = base._place_fractional_shift(
                    schedule,
                    usage,
                    remaining_capacity,
                    headers,
                    spread,
                    index,
                    partial_qty,
                )
                if made >= partial_qty - 1e-6:
                    placed_partial = True
                    break
            if not placed_partial:
                carryover[spread["code"]] = base._clean_number(partial_qty)

        scheduled_spread = sum(schedule[spread["code"]].values())
        missing_qty = max(0.0, spread["planned_qty"] - scheduled_spread)
        if missing_qty > 1e-6:
            carryover[spread["code"]] = base._clean_number(missing_qty)

        last_code = spread["code"]
        campaigns.append({"code": spread["code"], "mode": "spread_shortage_aware"})

    remaining_others = list(others)
    cursor_shift = 0.0
    last_group = spread["product_group"] if spread is not None else None
    while remaining_others:
        product, cursor_shift = shortage_first_campaign_candidate(
            remaining_others,
            headers,
            cursor_shift,
            capacity,
            last_group,
        )
        remaining_others.remove(product)
        start_index = base._earliest_index(headers, product)

        if last_code is not None and last_code != product["code"]:
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

        placed, last_index = base._place_units(
            schedule,
            usage,
            remaining_capacity,
            headers,
            product,
            product["required_units"],
            start_index,
            allow_sunday=True,
        )
        missing_units = product["required_units"] - placed
        if missing_units > 0:
            carryover[product["code"]] = base._clean_number(
                missing_units * product["quantum_qty"]
            )
        if placed > 0:
            last_code = product["code"]
            last_group = product["product_group"]
            cursor_shift = max(cursor_shift, last_index * capacity)
            campaigns.append(
                {"code": product["code"], "mode": "shortage_first_remaining_capacity"}
            )

    return schedule, usage, carryover, {
        "mode": "galon_shortage_first",
        "setup_shifts": setup_total,
        "campaigns": campaigns,
    }


def install_priority_scheduler():
    base._campaign_candidate = shortage_first_campaign_candidate
    base._weekly_unit_targets = shortage_aware_weekly_unit_targets
    base._allocate_weekly_rgb_line = allocate_weekly_rgb_priority
    base._allocate_galon_line = allocate_galon_priority


if __name__ == "__main__":
    install_priority_scheduler()
    base.main_with_retry()
