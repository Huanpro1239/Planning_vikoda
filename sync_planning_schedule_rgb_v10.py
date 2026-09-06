import math
from collections import defaultdict

import sync_planning_schedule as base
import sync_planning_schedule_rgb_v9 as v9


EPS = 1e-7
BEAM_WIDTH = 20000


def _fit_interval(cursor, duration, capacity):
    """Earliest non-splitting slot at/after cursor on a capacity-shift day."""
    if duration > capacity + EPS:
        return None
    cursor = max(float(cursor), 0.0)
    day = int(math.floor((cursor + EPS) / capacity))
    offset = cursor - day * capacity
    if offset < EPS:
        offset = 0.0
    if offset + duration <= capacity + EPS:
        return day * capacity + offset, day * capacity + offset + duration
    day += 1
    return day * capacity, day * capacity + duration


def _release_shift(headers, product, capacity):
    return base._earliest_index(headers, product) * float(capacity)


def _transition(state, product_index, products, headers, capacity):
    product = products[product_index]
    counts = state["counts"]
    if counts[product_index] >= int(product.get("required_units", 0) or 0):
        return None

    cursor = float(state["cursor"])
    last = state["last"]
    timeline = state["timeline"]
    setup_total = float(state["setup_total"])

    if last is not None and last != product_index:
        setup_slot = _fit_interval(cursor, base.SETUP_SHIFTS, float(capacity))
        if setup_slot is None:
            return None
        setup_start, setup_end = setup_slot
        if setup_end > len(headers) * float(capacity) + EPS:
            return None
        timeline = timeline + (
            {
                "kind": "setup",
                "from_code": products[last]["code"],
                "to_code": product["code"],
                "start_shift": setup_start,
                "end_shift": setup_end,
                "duration_shifts": base.SETUP_SHIFTS,
            },
        )
        cursor = setup_end
        setup_total += base.SETUP_SHIFTS

    cursor = max(cursor, _release_shift(headers, product, capacity))
    duration = float(product.get("quantum_shift", 0) or 0)
    qty = float(product.get("quantum_qty", 0) or 0)
    slot = _fit_interval(cursor, duration, float(capacity))
    if slot is None:
        return None
    start, end = slot
    if end > len(headers) * float(capacity) + EPS:
        return None

    new_counts = list(counts)
    new_counts[product_index] += 1
    production_event = {
        "kind": "production",
        "code": product["code"],
        "unit_no": new_counts[product_index],
        "qty": base._clean_number(qty),
        "start_shift": start,
        "end_shift": end,
        "duration_shifts": duration,
    }
    return {
        "cursor": end,
        "last": product_index,
        "counts": tuple(new_counts),
        "timeline": timeline + (production_event,),
        "setup_total": setup_total,
    }


def _schedule_from_timeline(headers, products, capacity, timeline):
    schedule = {
        product["code"]: {day: 0.0 for day in headers}
        for product in products
    }
    usage = defaultdict(float)
    for event in timeline:
        start = float(event["start_shift"])
        end = float(event["end_shift"])
        day_index = min(
            len(headers) - 1,
            max(0, int(math.floor((start + EPS) / float(capacity)))),
        )
        day = headers[day_index]
        usage[day] += end - start
        if event["kind"] == "production":
            schedule[event["code"]][day] += float(event.get("qty", 0) or 0)
    return schedule, usage


def _service_metrics(headers, products, schedule, *, through_days=None):
    limit = len(headers) if through_days is None else max(0, min(int(through_days), len(headers)))
    stockout_codes = set()
    stockout_days = 0
    total_deficit = 0.0
    max_deficit = 0.0
    latest_stockout = -1
    safety_days = 0
    safety_deficit = 0.0

    for product in products:
        code = product["code"]
        balance = float(product.get("actual_stock", 0) or 0)
        debt = max(float(product.get("debt", 0) or 0), 0.0)
        target = max(float(product.get("target_stock", 0) or 0), 0.0)
        for index, day in enumerate(headers[:limit]):
            balance += float(schedule[code].get(day, 0) or 0)
            balance -= float(product["demand_by_day"].get(day, 0) or 0)
            if index == 0:
                balance -= debt
            if balance < -EPS:
                deficit = -balance
                stockout_codes.add(code)
                stockout_days += 1
                total_deficit += deficit
                max_deficit = max(max_deficit, deficit)
                latest_stockout = max(latest_stockout, index)
            if balance < target - EPS:
                safety_days += 1
                safety_deficit += target - balance

    return (
        len(stockout_codes),
        stockout_days,
        round(total_deficit, 6),
        round(max_deficit, 6),
        latest_stockout,
        safety_days,
        round(safety_deficit, 6),
    )


def _partial_key(state, headers, products, capacity):
    schedule, _ = _schedule_from_timeline(
        headers,
        products,
        capacity,
        state["timeline"],
    )
    # Days before the current cursor day are irreversible. If cursor is exactly
    # at a day boundary, the preceding day is also closed.
    closed_days = min(
        len(headers),
        int(math.floor((float(state["cursor"]) + EPS) / float(capacity))),
    )
    service = _service_metrics(
        headers,
        products,
        schedule,
        through_days=closed_days,
    )
    return service + (
        round(float(state["setup_total"]), 6),
        round(float(state["cursor"]), 6),
    )


def _final_key(state, headers, products, capacity):
    schedule, _ = _schedule_from_timeline(headers, products, capacity, state["timeline"])
    service = _service_metrics(headers, products, schedule)
    return service + (
        round(float(state["setup_total"]), 6),
        round(float(state["cursor"]), 6),
    )


def allocate_rgb_quantum_lookahead(headers, products, capacity):
    """
    RGB V10: setup-aware quantum-level search over the whole physical machine.

    Unlike full-campaign V9, this may interrupt a SKU campaign when doing so
    reduces whole-resource stockout impact. Every transition still reserves
    0.5 shift, every production event is one complete product quantum, and no
    quantum crosses a day boundary.
    """
    active = [
        product for product in products
        if int(product.get("required_units", 0) or 0) > 0
    ]
    inactive = [
        product for product in products
        if int(product.get("required_units", 0) or 0) <= 0
    ]
    if not active:
        return v9.allocate_rgb_quantized_service(headers, products, capacity)

    total_units = sum(int(product.get("required_units", 0) or 0) for product in active)
    initial = {
        "cursor": 0.0,
        "last": None,
        "counts": tuple(0 for _ in active),
        "timeline": tuple(),
        "setup_total": 0.0,
    }
    states = [initial]
    expanded_total = 0

    for _step in range(total_units):
        dedup = {}
        for state in states:
            for product_index in range(len(active)):
                candidate = _transition(
                    state,
                    product_index,
                    active,
                    headers,
                    float(capacity),
                )
                if candidate is None:
                    continue
                expanded_total += 1
                key = (
                    candidate["counts"],
                    candidate["last"],
                    round(float(candidate["cursor"]), 6),
                )
                score = _partial_key(candidate, headers, active, capacity)
                existing = dedup.get(key)
                if existing is None or score < existing[0]:
                    dedup[key] = (score, candidate)

        if not dedup:
            # Preserve safe carryover behavior if the quantum-level search
            # cannot place all units inside the month.
            return v9.allocate_rgb_quantized_service(headers, products, capacity)

        ranked = sorted(dedup.values(), key=lambda item: item[0])
        states = [item[1] for item in ranked[:BEAM_WIDTH]]

    best = min(states, key=lambda state: _final_key(state, headers, active, capacity))
    active_schedule, usage = _schedule_from_timeline(
        headers,
        active,
        float(capacity),
        best["timeline"],
    )
    schedule = {
        product["code"]: {day: 0.0 for day in headers}
        for product in products
    }
    for product in active:
        schedule[product["code"]].update(active_schedule[product["code"]])
    for product in inactive:
        schedule[product["code"]] = {day: 0.0 for day in headers}

    sequence = [
        event["code"]
        for event in best["timeline"]
        if event.get("kind") == "production"
    ]
    return schedule, usage, {}, {
        "mode": "rgb_quantum_lookahead_v10",
        "setup_shifts": float(best["setup_total"]),
        "timeline": list(best["timeline"]),
        "quantum_sequence": sequence,
        "objective": list(_final_key(best, headers, active, capacity)),
        "expanded_states": expanded_total,
        "beam_width": BEAM_WIDTH,
        "quantum_policy": "whole_quantum_no_cross_day",
    }


def install_rgb_service_scheduler():
    base._allocate_weekly_rgb_line = allocate_rgb_quantum_lookahead
