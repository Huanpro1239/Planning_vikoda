import math

import sync_planning_schedule as base
import sync_planning_schedule_priority as priority


def debt_aware_risk_snapshot(headers, product):
    """Risk uses net opening availability: physical stock minus opening debt."""
    stock = (
        float(product.get("actual_stock", 0) or 0)
        - max(float(product.get("debt", 0) or 0), 0.0)
    )
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


def debt_aware_weekly_unit_targets(headers, product):
    """RGB weekly targets front-load debt/backlog before safety/even spreading."""
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
    opening_net = (
        float(product.get("actual_stock", 0) or 0)
        - max(float(product.get("debt", 0) or 0), 0.0)
    )

    targets = {}
    previous_cumulative = 0
    cumulative_demand = 0.0
    demand_by_index = []
    for current_day in headers:
        cumulative_demand += float(product["demand_by_day"].get(current_day, 0) or 0)
        demand_by_index.append(cumulative_demand)

    for ordinal, bucket in enumerate(active, start=1):
        _, end_index = bucket
        even_cumulative = int(
            math.ceil(total_units * ordinal / len(active) - base.EPSILON)
        )
        required_qty = max(
            0.0,
            demand_by_index[end_index] + target_stock - opening_net,
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

    if previous_cumulative < total_units:
        last_bucket = active[-1]
        targets[last_bucket] = targets.get(last_bucket, 0) + (
            total_units - previous_cumulative
        )
    return targets


def debt_aware_galon_priority(headers, products, capacity):
    """
    Reuse the established Galon spread/block allocator, but feed 19L its net
    opening availability after debt so the spread phase front-loads urgent
    replenishment instead of treating backlog as free stock.
    """
    adjusted = []
    for product in products:
        item = dict(product)
        if item.get("code") == base.GALON_SPREAD_CODE:
            item["actual_stock"] = (
                float(item.get("actual_stock", 0) or 0)
                - max(float(item.get("debt", 0) or 0), 0.0)
            )
        adjusted.append(item)
    return priority.allocate_galon_priority(headers, adjusted, capacity)


def install_debt_aware_priority():
    # Existing priority functions resolve these helpers from their module
    # globals at call time, so RGB/campaign sequencing immediately becomes
    # debt-aware without duplicating the whole scheduler.
    priority._risk_snapshot = debt_aware_risk_snapshot
    priority.shortage_aware_weekly_unit_targets = debt_aware_weekly_unit_targets

    # base._allocate_weekly_rgb_line already points at priority's allocator;
    # that function will call the patched helper above. Galon needs a wrapper
    # because its spread formula reads actual_stock directly.
    base._allocate_galon_line = debt_aware_galon_priority
