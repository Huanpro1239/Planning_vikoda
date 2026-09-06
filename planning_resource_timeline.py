import itertools
import math


EPSILON = 1e-7
SETUP_SHIFTS = 0.5
SERIAL_RESOURCES = {
    "KHS + PET 9000": {"KHS", "PET 9000"},
    "RGB": {"RGB"},
    "Galon": {"Galon"},
}


def _active_products_for_resource(products, resource):
    lines = SERIAL_RESOURCES.get(resource)
    if not lines:
        return []
    return [product for product in products if product.get("line") in lines]


def _production_qty(schedule, code, day):
    return float((schedule.get(code) or {}).get(day, 0) or 0)


def _order_switches(previous_code, order):
    switches = 0
    previous = previous_code
    for code in order:
        if previous is not None and code != previous:
            switches += 1
        previous = code
    return switches


def build_daily_serial_timeline(headers, products, schedule, capacity, resource):
    """
    Build a physical within-day sequence for RGB/Galon from the exact daily plan.

    Production quantity is not changed. The DP only chooses the order of SKU
    blocks inside each day and inserts every required 0.5-shift changeover.
    A day is rejected if its production plus the required setup cannot fit the
    declared line capacity. This makes the timeline a feasibility proof, not a
    cosmetic reconstruction of aggregate setup totals.
    """
    if resource not in {"RGB", "Galon"}:
        raise RuntimeError(f"Resource {resource!r} không dùng daily serial builder.")

    by_code = {
        product["code"]: product
        for product in _active_products_for_resource(products, resource)
    }
    row_rank = {
        code: int(product.get("row", 0) or 0)
        for code, product in by_code.items()
    }

    # last_code -> (switch_count, deterministic_penalty, daily_orders)
    states = {None: (0, 0, [])}
    for day_index, current_day in enumerate(headers):
        active = []
        production_shifts = {}
        for code, product in by_code.items():
            qty = _production_qty(schedule, code, current_day)
            if qty <= EPSILON:
                continue
            per_shift = float(product.get("per_shift", 0) or 0)
            if per_shift <= 0:
                raise RuntimeError(
                    f"{resource} mã {code} có sản lượng nhưng E<=0."
                )
            shifts = qty / per_shift
            if not math.isfinite(shifts) or shifts <= 0:
                raise RuntimeError(
                    f"{resource} mã {code} ngày {current_day} có số ca không hợp lệ {shifts!r}."
                )
            active.append(code)
            production_shifts[code] = shifts

        if not active:
            states = {
                last: (cost, penalty, orders + [()])
                for last, (cost, penalty, orders) in states.items()
            }
            continue

        if len(active) > 8:
            raise RuntimeError(
                f"{resource} ngày {current_day} có {len(active)} SKU cùng chạy; "
                "không thể chứng minh thứ tự setup bằng exact day-order search."
            )

        active.sort(key=lambda code: (row_rank.get(code, 0), code))
        permutations = list(itertools.permutations(active))
        next_states = {}
        production_total = sum(production_shifts.values())

        for last_code, (switch_count, penalty, orders) in states.items():
            for order in permutations:
                switches = _order_switches(last_code, order)
                day_usage = production_total + switches * SETUP_SHIFTS
                if day_usage > float(capacity) + EPSILON:
                    continue

                end_code = order[-1]
                order_penalty = sum(
                    (position + 1) * (row_rank.get(code, 0) + 1)
                    for position, code in enumerate(order)
                )
                candidate = (
                    switch_count + switches,
                    penalty + order_penalty,
                    orders + [order],
                )
                existing = next_states.get(end_code)
                if existing is None or candidate[:2] < existing[:2]:
                    next_states[end_code] = candidate

        if not next_states:
            detail = ", ".join(
                f"{code}={production_shifts[code]:.3f}ca"
                for code in active
            )
            raise RuntimeError(
                f"{resource} ngày {current_day:%d/%m} không có thứ tự vận hành hợp lệ: "
                f"production={production_total:.3f}ca, capacity={float(capacity):.3f}ca, "
                f"cần setup khi đổi mã; {detail}."
            )
        states = next_states

    if not states:
        return [], 0.0, {day: 0.0 for day in headers}

    _, _, daily_orders = min(states.values(), key=lambda value: value[:2])
    timeline = []
    daily_usage = {day: 0.0 for day in headers}
    previous_code = None
    setup_total = 0.0

    for day_index, (current_day, order) in enumerate(zip(headers, daily_orders)):
        day_start = day_index * float(capacity)
        cursor = day_start
        for code in order:
            product = by_code[code]
            if previous_code is not None and previous_code != code:
                timeline.append(
                    {
                        "kind": "setup",
                        "from_code": previous_code,
                        "to_code": code,
                        "start_shift": cursor,
                        "end_shift": cursor + SETUP_SHIFTS,
                        "duration_shifts": SETUP_SHIFTS,
                    }
                )
                cursor += SETUP_SHIFTS
                daily_usage[current_day] += SETUP_SHIFTS
                setup_total += SETUP_SHIFTS

            qty = _production_qty(schedule, code, current_day)
            duration = qty / float(product["per_shift"])
            timeline.append(
                {
                    "kind": "production",
                    "code": code,
                    "qty": qty,
                    "start_shift": cursor,
                    "end_shift": cursor + duration,
                    "duration_shifts": duration,
                }
            )
            cursor += duration
            daily_usage[current_day] += duration
            previous_code = code

        if cursor > day_start + float(capacity) + EPSILON:
            raise RuntimeError(
                f"{resource} ngày {current_day:%d/%m} timeline vượt capacity."
            )

    return timeline, setup_total, daily_usage


def validate_resource_timeline(headers, products, schedule, capacity, resource, timeline):
    """Independently validate sequence, setup, capacity, batch and daily quantities."""
    if resource not in SERIAL_RESOURCES:
        return {
            "resource": resource,
            "validated": False,
            "reason": "resource_not_serially_validated",
        }

    by_code = {
        product["code"]: product
        for product in _active_products_for_resource(products, resource)
    }
    scheduled_codes = {
        code
        for code in by_code
        if any(_production_qty(schedule, code, day) > EPSILON for day in headers)
    }
    if not scheduled_codes:
        if timeline:
            raise RuntimeError(f"{resource} không có sản lượng nhưng timeline không rỗng.")
        return {
            "resource": resource,
            "validated": True,
            "setup_shifts": 0.0,
            "timeline_events": 0,
        }
    if not isinstance(timeline, list) or not timeline:
        raise RuntimeError(f"{resource} có sản lượng nhưng thiếu timeline production/setup.")

    horizon = len(headers) * float(capacity)
    reconstructed = {
        code: {day: 0.0 for day in headers}
        for code in scheduled_codes
    }
    daily_usage = {day: 0.0 for day in headers}
    events = []
    for index, raw in enumerate(timeline):
        if not isinstance(raw, dict):
            raise RuntimeError(f"{resource} timeline event #{index + 1} không phải object.")
        event = dict(raw)
        kind = str(event.get("kind") or "").strip()
        if kind not in {"production", "setup"}:
            raise RuntimeError(
                f"{resource} timeline event #{index + 1} kind={kind!r} không hợp lệ."
            )
        try:
            start = float(event.get("start_shift"))
            end = float(event.get("end_shift"))
        except (TypeError, ValueError) as exc:
            raise RuntimeError(
                f"{resource} timeline event #{index + 1} thiếu start/end hợp lệ."
            ) from exc
        if not math.isfinite(start) or not math.isfinite(end):
            raise RuntimeError(f"{resource} timeline event #{index + 1} có shift không hữu hạn.")
        if start < -EPSILON or end < start - EPSILON or end > horizon + EPSILON:
            raise RuntimeError(
                f"{resource} timeline event #{index + 1} ngoài horizon: {start}..{end}."
            )
        event["_start"] = start
        event["_end"] = end
        events.append(event)

    events.sort(key=lambda item: (item["_start"], item["_end"], item.get("kind", "")))
    previous_end = 0.0
    previous_production_code = None
    setup_since_production = 0.0
    setup_total = 0.0

    for event in events:
        if event["_start"] < previous_end - EPSILON:
            raise RuntimeError(
                f"{resource} timeline chồng lấn tại shift {event['_start']:.3f}."
            )
        previous_end = max(previous_end, event["_end"])
        duration = event["_end"] - event["_start"]

        for day_index, current_day in enumerate(headers):
            day_start = day_index * float(capacity)
            day_end = day_start + float(capacity)
            overlap = max(
                0.0,
                min(event["_end"], day_end) - max(event["_start"], day_start),
            )
            if overlap > EPSILON:
                daily_usage[current_day] += overlap

        if event["kind"] == "setup":
            setup_since_production += duration
            setup_total += duration
            continue

        code = str(event.get("code") or "").strip()
        if code not in scheduled_codes:
            raise RuntimeError(
                f"{resource} timeline production chứa mã {code!r} không thuộc resource."
            )
        if previous_production_code is not None and code != previous_production_code:
            if setup_since_production < SETUP_SHIFTS - EPSILON:
                raise RuntimeError(
                    f"{resource} đổi mã {previous_production_code}->{code} thiếu setup "
                    f"{SETUP_SHIFTS:g}ca; chỉ có {setup_since_production:g}ca."
                )
        setup_since_production = 0.0
        previous_production_code = code

        product = by_code[code]
        per_shift = float(product.get("per_shift", 0) or 0)
        qty = float(event.get("qty", 0) or 0)
        expected_qty = duration * per_shift
        if not math.isclose(qty, expected_qty, rel_tol=1e-9, abs_tol=1e-5):
            raise RuntimeError(
                f"{resource} timeline {code}: qty={qty} khác duration*E={expected_qty}."
            )

        if bool(product.get("is_sugar")):
            batch = float(product.get("batch", 0) or 0)
            if batch <= 0:
                raise RuntimeError(f"{resource} mã đường {code} có batch<=0.")
            units = qty / batch
            if not math.isclose(units, round(units), rel_tol=1e-9, abs_tol=1e-6):
                raise RuntimeError(
                    f"{resource} mã {code} timeline ghi {qty} không nguyên mẻ {batch}."
                )

        for day_index, current_day in enumerate(headers):
            day_start = day_index * float(capacity)
            day_end = day_start + float(capacity)
            overlap = max(
                0.0,
                min(event["_end"], day_end) - max(event["_start"], day_start),
            )
            if overlap > EPSILON:
                reconstructed[code][current_day] += overlap * per_shift

    for current_day, used in daily_usage.items():
        if used > float(capacity) + EPSILON:
            raise RuntimeError(
                f"{resource} ngày {current_day:%d/%m} dùng {used:.3f}ca > "
                f"capacity {float(capacity):.3f}ca gồm production+setup."
            )

    for code in scheduled_codes:
        for current_day in headers:
            expected = _production_qty(schedule, code, current_day)
            actual = reconstructed[code][current_day]
            if not math.isclose(expected, actual, rel_tol=1e-9, abs_tol=1e-5):
                raise RuntimeError(
                    f"{resource} timeline {code} ngày {current_day:%d/%m}={actual} "
                    f"khác lịch={expected}."
                )

    return {
        "resource": resource,
        "validated": True,
        "setup_shifts": setup_total,
        "timeline_events": len(events),
        "max_daily_usage": max(daily_usage.values()) if daily_usage else 0.0,
    }


def ensure_resource_timelines(info):
    """
    Return timeline/validation metadata for every serial resource in schedule info.

    KHS/PET uses the scheduler-native timeline. RGB/Galon receive a timeline
    derived from the exact daily schedule and are rejected if no physically
    valid within-day sequence with all setup can be found.
    """
    headers = list(info.get("headers") or [])
    products = list(info.get("products") or [])
    schedule = info.get("schedule") or {}
    capacities = dict(info.get("line_capacity") or {})
    optimizer_meta = info.get("optimizer_meta") or {}

    result = {}
    for resource, lines in SERIAL_RESOURCES.items():
        capacity = capacities.get(resource)
        active = [
            product
            for product in products
            if product.get("line") in lines
            and any(_production_qty(schedule, product["code"], day) > EPSILON for day in headers)
        ]
        if not active:
            continue
        if capacity is None:
            raise RuntimeError(f"Thiếu capacity cho serial resource {resource}.")

        raw_meta = dict(optimizer_meta.get(resource, {}) or {})
        scheduler_setup = float(raw_meta.get("setup_shifts", 0) or 0)
        timeline = raw_meta.get("timeline")
        if resource in {"RGB", "Galon"}:
            timeline, derived_setup, _ = build_daily_serial_timeline(
                headers,
                products,
                schedule,
                float(capacity),
                resource,
            )
        else:
            derived_setup = scheduler_setup

        validation = validate_resource_timeline(
            headers,
            products,
            schedule,
            float(capacity),
            resource,
            timeline,
        )
        actual_setup = float(validation.get("setup_shifts", derived_setup) or 0)
        result[resource] = {
            "timeline": timeline,
            "scheduler_setup_shifts": scheduler_setup,
            "setup_shifts": actual_setup,
            "setup_delta_shifts": actual_setup - scheduler_setup,
            "validation": validation,
        }
    return result
