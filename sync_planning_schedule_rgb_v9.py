import itertools
import math
from collections import defaultdict

import sync_planning_schedule as base


MAX_EXACT_RGB_SKUS = 8


def _empty_schedule(headers, products):
    return {
        product["code"]: {day: 0.0 for day in headers}
        for product in products
    }


def _release_index(headers, product):
    return base._earliest_index(headers, product)


def _place_setup(remaining, usage, headers, start_index, capacity, timeline, from_code, to_code):
    for index in range(max(0, start_index), len(headers)):
        if remaining[index] + base.EPSILON < base.SETUP_SHIFTS:
            continue
        used_before = float(capacity) - remaining[index]
        start_shift = index * float(capacity) + used_before
        end_shift = start_shift + base.SETUP_SHIFTS
        remaining[index] -= base.SETUP_SHIFTS
        usage[headers[index]] += base.SETUP_SHIFTS
        timeline.append(
            {
                "kind": "setup",
                "from_code": from_code,
                "to_code": to_code,
                "start_shift": start_shift,
                "end_shift": end_shift,
                "duration_shifts": base.SETUP_SHIFTS,
            }
        )
        return index
    return None


def _place_quantum(remaining, usage, schedule, headers, product, start_index, capacity, timeline, unit_no):
    duration = float(product.get("quantum_shift", 0) or 0)
    qty = float(product.get("quantum_qty", 0) or 0)
    if duration <= 0 or qty <= 0:
        return None

    for index in range(max(0, start_index), len(headers)):
        if remaining[index] + base.EPSILON < duration:
            continue
        used_before = float(capacity) - remaining[index]
        start_shift = index * float(capacity) + used_before
        end_shift = start_shift + duration
        remaining[index] -= duration
        usage[headers[index]] += duration
        schedule[product["code"]][headers[index]] += qty
        timeline.append(
            {
                "kind": "production",
                "code": product["code"],
                "unit_no": unit_no,
                "qty": base._clean_number(qty),
                "start_shift": start_shift,
                "end_shift": end_shift,
                "duration_shifts": duration,
            }
        )
        return index
    return None


def _simulate_sequence(headers, products, capacity, order):
    by_code = {product["code"]: product for product in products}
    schedule = _empty_schedule(headers, products)
    usage = defaultdict(float)
    timeline = []
    remaining = [float(capacity) for _ in headers]
    carryover = {}
    cursor = 0
    last_code = None
    setup_total = 0.0

    for code in order:
        product = by_code[code]
        required_units = int(product.get("required_units", 0) or 0)
        if required_units <= 0:
            continue
        cursor = max(cursor, _release_index(headers, product))

        if last_code is not None and last_code != code:
            setup_day = _place_setup(
                remaining,
                usage,
                headers,
                cursor,
                capacity,
                timeline,
                last_code,
                code,
            )
            if setup_day is None:
                carryover[code] = base._clean_number(
                    required_units * float(product.get("quantum_qty", 0) or 0)
                )
                continue
            setup_total += base.SETUP_SHIFTS
            cursor = setup_day

        placed = 0
        for unit_no in range(1, required_units + 1):
            day_index = _place_quantum(
                remaining,
                usage,
                schedule,
                headers,
                product,
                cursor,
                capacity,
                timeline,
                unit_no,
            )
            if day_index is None:
                break
            placed += 1
            cursor = day_index

        if placed < required_units:
            carryover[code] = base._clean_number(
                (required_units - placed) * float(product.get("quantum_qty", 0) or 0)
            )
        if placed > 0:
            last_code = code

    return schedule, usage, carryover, timeline, setup_total


def _service_cost(headers, products, schedule, carryover, setup_total, order):
    carryover_qty = sum(max(float(value or 0), 0.0) for value in carryover.values())
    carryover_skus = sum(1 for value in carryover.values() if float(value or 0) > base.EPSILON)
    stockout_skus = 0
    stockout_days = 0
    stockout_deficit = 0.0
    max_deficit = 0.0
    latest_stockout = -1
    safety_days = 0
    safety_deficit = 0.0

    for product in products:
        balance = float(product.get("actual_stock", 0) or 0)
        debt = max(float(product.get("debt", 0) or 0), 0.0)
        target = max(float(product.get("target_stock", 0) or 0), 0.0)
        code_has_stockout = False
        for index, day in enumerate(headers):
            balance += float(schedule[product["code"]].get(day, 0) or 0)
            balance -= float(product["demand_by_day"].get(day, 0) or 0)
            if index == 0:
                balance -= debt
            if balance < -base.EPSILON:
                code_has_stockout = True
                deficit = -balance
                stockout_days += 1
                stockout_deficit += deficit
                max_deficit = max(max_deficit, deficit)
                latest_stockout = max(latest_stockout, index)
            if balance < target - base.EPSILON:
                safety_days += 1
                safety_deficit += target - balance
        if code_has_stockout:
            stockout_skus += 1

    # Service is lexicographically dominant. Setup and earliness are only tie-breaks.
    release_penalty = 0
    by_code = {product["code"]: product for product in products}
    for ordinal, code in enumerate(order):
        release_penalty += (ordinal + 1) * _release_index(headers, by_code[code])

    return (
        carryover_skus,
        round(carryover_qty, 6),
        stockout_skus,
        stockout_days,
        round(stockout_deficit, 6),
        round(max_deficit, 6),
        latest_stockout,
        safety_days,
        round(safety_deficit, 6),
        round(setup_total, 6),
        release_penalty,
    )


def _fallback_order(headers, products):
    def key(product):
        balance = float(product.get("actual_stock", 0) or 0) - max(
            float(product.get("debt", 0) or 0), 0.0
        )
        first_stockout = len(headers) + 1
        for index, day in enumerate(headers):
            balance -= float(product["demand_by_day"].get(day, 0) or 0)
            if balance < -base.EPSILON:
                first_stockout = index
                break
        return (
            first_stockout,
            0 if float(product.get("debt", 0) or 0) > 0 else 1,
            _release_index(headers, product),
            product.get("row", 0),
        )

    return [product["code"] for product in sorted(products, key=key)]


def allocate_rgb_quantized_service(headers, products, capacity):
    """
    RGB V9: one physical serial machine, full production quanta only.

    Exact permutation search chooses the full-campaign order that minimizes
    whole-resource stockout impact. Every SKU switch consumes 0.5 shift and is
    represented in the returned timeline. No batch/shift is split across days.
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
        return _empty_schedule(headers, products), defaultdict(float), {}, {
            "mode": "rgb_quantized_service_v9",
            "setup_shifts": 0.0,
            "timeline": [],
            "sequence": [],
            "evaluated_sequences": 0,
        }

    codes = [product["code"] for product in active]
    if len(codes) <= MAX_EXACT_RGB_SKUS:
        orders = itertools.permutations(codes)
    else:
        orders = [_fallback_order(headers, active)]

    best = None
    evaluated = 0
    for order in orders:
        order = tuple(order)
        evaluated += 1
        schedule, usage, carryover, timeline, setup_total = _simulate_sequence(
            headers,
            active,
            float(capacity),
            order,
        )
        cost = _service_cost(
            headers,
            active,
            schedule,
            carryover,
            setup_total,
            order,
        )
        candidate = (cost, order, schedule, usage, carryover, timeline, setup_total)
        if best is None or candidate[0] < best[0]:
            best = candidate

    cost, order, active_schedule, usage, carryover, timeline, setup_total = best
    schedule = _empty_schedule(headers, products)
    for product in active:
        schedule[product["code"]].update(active_schedule[product["code"]])
    for product in inactive:
        schedule[product["code"]] = {day: 0.0 for day in headers}

    return schedule, usage, carryover, {
        "mode": "rgb_quantized_service_v9",
        "setup_shifts": setup_total,
        "timeline": timeline,
        "sequence": list(order),
        "objective": list(cost),
        "evaluated_sequences": evaluated,
        "quantum_policy": "whole_quantum_no_cross_day",
    }


def install_rgb_service_scheduler():
    base._allocate_weekly_rgb_line = allocate_rgb_quantized_service


if __name__ == "__main__":
    from sync_planning_schedule_production import install_production_output_cleanup

    install_production_output_cleanup()
    base.main_with_retry()
