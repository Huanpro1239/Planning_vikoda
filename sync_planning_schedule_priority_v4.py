import math
from collections import defaultdict

import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v2 as v2


REPLAN_HORIZON_DAYS = 7
HARD_STOCKOUT_DAYS = 2
CONTINUITY_SLACK_DAYS = 1


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


def _remaining_processing_shifts(product, remaining_units):
    return remaining_units * float(product.get("quantum_shift", 0) or 0)


def _deadline_shift(deadline_index, capacity):
    if deadline_index is None:
        return math.inf
    return (deadline_index + 1) * capacity


def _priority_key(headers, product, schedule, remaining_units, cursor_shift, capacity):
    stockout, safety = _dynamic_risk(headers, product, schedule)
    release = max(
        cursor_shift,
        priority.base._earliest_index(headers, product) * capacity,
    )
    processing = _remaining_processing_shifts(
        product,
        remaining_units[product["code"]],
    )
    finish_without_wait = release + processing

    stockout_slack = _deadline_shift(stockout, capacity) - finish_without_wait
    safety_slack = _deadline_shift(safety, capacity) - finish_without_wait

    current_index = min(
        len(headers),
        int(math.floor(cursor_shift / capacity + priority.base.EPSILON)),
    )

    if stockout is not None and stockout <= current_index + HARD_STOCKOUT_DAYS:
        tier = 0
        effective_slack = stockout_slack
        deadline = stockout
    elif stockout is not None:
        tier = 1
        effective_slack = stockout_slack
        deadline = stockout
    elif safety is not None:
        tier = 2
        effective_slack = safety_slack
        deadline = safety
    elif float(product.get("debt", 0) or 0) > 0:
        tier = 3
        effective_slack = capacity
        deadline = current_index
    else:
        tier = 4
        effective_slack = math.inf
        deadline = len(headers) + 30

    return (
        tier,
        effective_slack,
        deadline,
        safety if safety is not None else len(headers) + 30,
        0 if float(product.get("debt", 0) or 0) > 0 else 1,
        priority.base._earliest_index(headers, product),
        product.get("row", 0),
    )


def _can_finish_before_primary_deadline(
    headers,
    candidate,
    primary,
    schedule,
    remaining_units,
    cursor_shift,
    capacity,
):
    primary_stockout, primary_safety = _dynamic_risk(headers, primary, schedule)
    deadline_index = primary_stockout if primary_stockout is not None else primary_safety
    if deadline_index is None:
        return True

    candidate_release = max(
        cursor_shift,
        priority.base._earliest_index(headers, candidate) * capacity,
    )
    candidate_duration = _remaining_processing_shifts(
        candidate,
        remaining_units[candidate["code"]],
    )
    primary_release = priority.base._earliest_index(headers, primary) * capacity
    primary_start = max(
        candidate_release + candidate_duration + priority.base.SETUP_SHIFTS,
        primary_release,
    )
    return primary_start <= _deadline_shift(deadline_index, capacity) + priority.base.EPSILON


def _choose_candidate(
    headers,
    active,
    schedule,
    remaining_units,
    cursor_shift,
    capacity,
    last_code,
    last_group,
):
    primary = min(
        active,
        key=lambda product: _priority_key(
            headers,
            product,
            schedule,
            remaining_units,
            cursor_shift,
            capacity,
        ),
    )
    primary_key = _priority_key(
        headers,
        primary,
        schedule,
        remaining_units,
        cursor_shift,
        capacity,
    )

    # Stockout sát ngay: service level thắng tuyệt đối continuity.
    if primary_key[0] == 0:
        return primary

    # Giữ nguyên mã nếu độ khẩn cấp gần tương đương và không làm primary trễ.
    if last_code:
        same_code = next((p for p in active if p["code"] == last_code), None)
        if same_code is not None:
            same_key = _priority_key(
                headers,
                same_code,
                schedule,
                remaining_units,
                cursor_shift,
                capacity,
            )
            slack_margin = CONTINUITY_SLACK_DAYS * capacity
            if (
                same_key[0] <= primary_key[0] + 1
                and same_key[1] <= primary_key[1] + slack_margin
                and _can_finish_before_primary_deadline(
                    headers,
                    same_code,
                    primary,
                    schedule,
                    remaining_units,
                    cursor_shift,
                    capacity,
                )
            ):
                return same_code

    # Cùng nhóm/khuôn chỉ được chen trước nếu chạy trọn campaign vẫn không làm
    # SKU shortage-first lỡ deadline tồn kho.
    if last_group:
        same_group = [
            p
            for p in active
            if p.get("product_group") == last_group
            and p["code"] != primary["code"]
        ]
        same_group.sort(
            key=lambda product: _priority_key(
                headers,
                product,
                schedule,
                remaining_units,
                cursor_shift,
                capacity,
            )
        )
        for candidate in same_group:
            if _can_finish_before_primary_deadline(
                headers,
                candidate,
                primary,
                schedule,
                remaining_units,
                cursor_shift,
                capacity,
            ):
                return candidate

    return primary


def _units_needed_through(headers, product, schedule, horizon_index, remaining_units):
    stock = float(product.get("actual_stock", 0) or 0)
    target = max(float(product.get("target_stock", 0) or 0), 0.0)

    for index in range(0, min(horizon_index, len(headers) - 1) + 1):
        current_day = headers[index]
        stock += float(schedule[product["code"]].get(current_day, 0) or 0)
        stock -= float(product["demand_by_day"].get(current_day, 0) or 0)

    shortage_qty = max(0.0, target - stock)
    quantum = float(product.get("quantum_qty", 0) or 0)
    if quantum <= 0:
        return 0

    needed = int(math.ceil(shortage_qty / quantum - priority.base.EPSILON))
    return max(1, min(remaining_units, needed))


def allocate_continuous_min_slack(headers, products, capacity):
    """KHS/PET: shortage-first + minimum slack + campaign continuity có kiểm soát."""
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
            future_release = [
                priority.base._earliest_index(headers, product)
                for product in products
                if remaining_units[product["code"]] > 0
            ]
            if not future_release:
                break
            cursor_shift = max(cursor_shift, min(future_release) * capacity)
            if cursor_shift >= month_end_shift - priority.base.EPSILON:
                break
            continue

        product = _choose_candidate(
            headers,
            active,
            schedule,
            remaining_units,
            cursor_shift,
            capacity,
            last_code,
            last_group,
        )
        code = product["code"]

        if last_code is not None and code != last_code:
            if cursor_shift + priority.base.SETUP_SHIFTS > month_end_shift + priority.base.EPSILON:
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

        cursor_shift = max(
            cursor_shift,
            priority.base._earliest_index(headers, product) * capacity,
        )
        if cursor_shift >= month_end_shift - priority.base.EPSILON:
            break

        units_left = remaining_units[code]
        chunk_units = units_left
        full_finish = cursor_shift + units_left * product["quantum_shift"]

        # Tìm competitor có deadline tồn kho gần nhất sau lịch đã xếp hiện tại.
        competitors = []
        for competitor in products:
            if competitor["code"] == code or remaining_units[competitor["code"]] <= 0:
                continue
            stockout, safety = _dynamic_risk(headers, competitor, schedule)
            deadline = stockout if stockout is not None else safety
            if deadline is None:
                continue
            release = priority.base._earliest_index(headers, competitor) * capacity
            latest_start = max(release, _deadline_shift(deadline, capacity) - priority.base.SETUP_SHIFTS)
            competitors.append((latest_start, deadline, competitor))

        if competitors:
            competitors.sort(key=lambda item: (item[0], item[1], item[2].get("row", 0)))
            latest_start, deadline, _ = competitors[0]
            if full_finish > latest_start + priority.base.EPSILON:
                max_before_switch = int(
                    math.floor(
                        max(0.0, latest_start - cursor_shift)
                        / product["quantum_shift"]
                        + priority.base.EPSILON
                    )
                )
                horizon = min(
                    len(headers) - 1,
                    current_index + REPLAN_HORIZON_DAYS,
                    deadline,
                )
                protect_units = _units_needed_through(
                    headers,
                    product,
                    schedule,
                    horizon,
                    units_left,
                )
                if max_before_switch <= 0:
                    # Đánh giá lại ngay để competitor được chọn ở vòng kế tiếp.
                    if last_code == code:
                        last_code = None
                    cursor_shift += priority.base.EPSILON
                    continue
                chunk_units = min(units_left, max(1, min(protect_units, max_before_switch)))

        available = max(0.0, month_end_shift - cursor_shift)
        max_month_units = int(
            math.floor(available / product["quantum_shift"] + priority.base.EPSILON)
        )
        chunk_units = min(chunk_units, max_month_units)
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
                "group": product.get("product_group"),
                "start_shift": cursor_shift,
                "end_shift": cursor_shift + duration,
                "scheduled_qty": priority.base._clean_number(produced),
            }
        )
        remaining_units[code] -= chunk_units
        cursor_shift += duration
        last_code = code
        last_group = product.get("product_group")

    for code, units in remaining_units.items():
        if units > 0:
            carryover[code] = priority.base._clean_number(
                units * by_code[code]["quantum_qty"]
            )

    return schedule, usage, carryover, {
        "mode": "continuous_min_slack_shortage_first",
        "setup_shifts": setup_total,
        "campaigns": campaigns,
    }


def install_priority_scheduler_v4():
    # RGB + Galon dùng shortage-aware weekly/spread của v2; KHS/PET dùng MST v4.
    v2.install_priority_scheduler_v2()
    priority.base._allocate_continuous_campaign_line = allocate_continuous_min_slack


if __name__ == "__main__":
    install_priority_scheduler_v4()
    priority.base.main_with_retry()
