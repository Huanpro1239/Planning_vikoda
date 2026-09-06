import math
from collections import defaultdict

import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v7 as v7


SHARED_MACHINE_LINES = {"KHS", "PET 9000"}
SHARED_MACHINE_NAME = "KHS + PET 9000"
URGENT_BUFFER_DAYS = 1


_ORIGINAL_BUILD_SCHEDULE = priority.base.build_schedule


def _next_unit_info(headers, product, capacity, scheduled_units):
    """Return (deadline_shift, critical, release_shift) for the next quantum."""
    unit_no = scheduled_units + 1
    required_units = int(product.get("required_units", 0) or 0)
    if unit_no > required_units:
        return None

    quantum_qty = float(product.get("quantum_qty", 0) or 0)
    if quantum_qty <= 0:
        return None

    cumulative_demand = float(product.get("debt", 0) or 0)
    actual_stock = max(float(product.get("actual_stock", 0) or 0), 0.0)

    for index, current_day in enumerate(headers):
        cumulative_demand += float(product["demand_by_day"].get(current_day, 0) or 0)
        required_qty = max(cumulative_demand - actual_stock, 0.0)
        cumulative_units = int(
            math.ceil(required_qty / quantum_qty - priority.base.EPSILON)
        )
        if unit_no <= cumulative_units:
            # Urgent supply phải hoàn thành trước ngày dự kiến thiếu 1 ngày.
            # Riêng thiếu ngay ngày đầu tháng thì cho phép hoàn thành trong ngày 1.
            raw_deadline = (index + 1) * capacity
            buffered_deadline = max(
                capacity,
                raw_deadline - URGENT_BUFFER_DAYS * capacity,
            )
            return (buffered_deadline, True, 0.0)

    # Quantum dư do làm tròn hoặc safety build: không được chen trước urgent.
    release_index = priority.base._earliest_index(headers, product)
    return (len(headers) * capacity, False, release_index * capacity)


def _candidate_key(
    headers,
    product,
    capacity,
    scheduled_units,
    cursor_shift,
    last_code,
    last_group,
):
    info = _next_unit_info(
        headers,
        product,
        capacity,
        scheduled_units[product["code"]],
    )
    deadline, critical, release = info
    same_code = product["code"] == last_code
    setup_needed = 0.0 if last_code is None or same_code else priority.base.SETUP_SHIFTS
    duration = float(product.get("quantum_shift", 0) or 0)
    latest_start = deadline - duration - setup_needed
    slack = latest_start - cursor_shift

    if critical:
        return (
            0,
            slack,
            latest_start,
            deadline,
            0 if same_code else 1,
            0 if last_group and product.get("product_group") == last_group else 1,
            0 if float(product.get("debt", 0) or 0) > 0 else 1,
            product.get("row", 0),
        )

    return (
        1,
        release,
        0 if same_code else 1,
        0 if last_group and product.get("product_group") == last_group else 1,
        0 if float(product.get("debt", 0) or 0) > 0 else 1,
        deadline,
        product.get("row", 0),
        0,
    )


def allocate_shared_deadline_guarded(headers, products, capacity):
    """
    Shared KHS/PET machine scheduler.

    - Urgent supply (FC + debt vượt tồn thực tế) is deadline-protected.
    - Urgent deadlines have a one-day production buffer.
    - Safety/rounding quantity cannot jump ahead of an urgent quantum.
    - Urgent priority uses minimum slack / latest-start, including setup and
      quantum duration, so a SKU is started early enough to meet its deadline.
    - Continuity is kept only when one more quantum of the current SKU still
      leaves enough slack for the most urgent competitor.
    - A switch consumes 0.5 shift, so KHS and PET can never run in parallel.
    """
    schedule = priority.base._empty_schedule(headers, products)
    usage = defaultdict(float)
    carryover = {}
    campaigns = []
    deadline_misses = []

    scheduled_units = {product["code"]: 0 for product in products}
    by_code = {product["code"]: product for product in products}
    month_end_shift = len(headers) * capacity
    cursor_shift = 0.0
    last_code = None
    last_group = None
    setup_total = 0.0

    def units_left(product):
        return int(product.get("required_units", 0) or 0) - scheduled_units[product["code"]]

    while any(units_left(product) > 0 for product in products):
        pending = [product for product in products if units_left(product) > 0]
        if not pending:
            break

        infos = {
            product["code"]: _next_unit_info(
                headers,
                product,
                capacity,
                scheduled_units[product["code"]],
            )
            for product in pending
        }

        eligible = [
            product
            for product in pending
            if infos[product["code"]][1]
            or infos[product["code"]][2] <= cursor_shift + priority.base.EPSILON
        ]

        if not eligible:
            next_release = min(infos[product["code"]][2] for product in pending)
            cursor_shift = max(cursor_shift, next_release)
            if cursor_shift >= month_end_shift - priority.base.EPSILON:
                break
            continue

        # Chỉ xét ứng viên mà toàn bộ setup + quantum còn vừa horizon.
        # Nếu ứng viên ưu tiên nhất không vừa nhưng SKU khác vẫn vừa, phải thử
        # SKU đó thay vì break và bỏ phí capacity.
        feasible = []
        for product in eligible:
            same_code = product["code"] == last_code
            setup_needed = (
                0.0
                if last_code is None or same_code
                else priority.base.SETUP_SHIFTS
            )
            duration = float(product.get("quantum_shift", 0) or 0)
            if (
                cursor_shift + setup_needed + duration
                <= month_end_shift + priority.base.EPSILON
            ):
                feasible.append(product)

        if not feasible:
            break

        primary = min(
            feasible,
            key=lambda product: _candidate_key(
                headers,
                product,
                capacity,
                scheduled_units,
                cursor_shift,
                last_code,
                last_group,
            ),
        )
        chosen = primary

        if last_code and last_code in by_code:
            current = by_code[last_code]
            if units_left(current) > 0 and current in feasible and current["code"] != primary["code"]:
                current_deadline, current_critical, _ = infos[current["code"]]
                primary_deadline, primary_critical, _ = infos[primary["code"]]

                current_finish = cursor_shift + float(current["quantum_shift"])
                primary_finish_after_switch = (
                    current_finish
                    + priority.base.SETUP_SHIFTS
                    + float(primary["quantum_shift"])
                )

                if not (primary_critical and not current_critical):
                    if (
                        primary_finish_after_switch
                        <= primary_deadline + priority.base.EPSILON
                        and current_finish
                        <= current_deadline + priority.base.EPSILON
                    ):
                        chosen = current

        if last_code is not None and chosen["code"] != last_code:
            if (
                cursor_shift + priority.base.SETUP_SHIFTS
                > month_end_shift + priority.base.EPSILON
            ):
                break
            priority.base._add_interval_usage(
                usage,
                headers,
                cursor_shift,
                priority.base.SETUP_SHIFTS,
                capacity,
            )
            cursor_shift += priority.base.SETUP_SHIFTS
            setup_total += priority.base.SETUP_SHIFTS

        duration = float(chosen["quantum_shift"])
        if cursor_shift + duration > month_end_shift + priority.base.EPSILON:
            break

        deadline, critical, _ = _next_unit_info(
            headers,
            chosen,
            capacity,
            scheduled_units[chosen["code"]],
        )
        start_shift = cursor_shift
        produced = priority.base._add_production_interval(
            schedule,
            usage,
            headers,
            chosen,
            start_shift,
            duration,
            capacity,
        )
        expected = float(chosen["quantum_qty"])
        if not math.isclose(produced, expected, rel_tol=1e-9, abs_tol=1e-5):
            raise RuntimeError(
                f"Mã {chosen['code']} quantum dự kiến {expected} nhưng ghi {produced}."
            )

        cursor_shift += duration
        scheduled_units[chosen["code"]] += 1

        if critical and cursor_shift > deadline + priority.base.EPSILON:
            deadline_misses.append(
                {
                    "code": chosen["code"],
                    "qty": priority.base._clean_number(expected),
                    "deadline_shift": deadline,
                    "finish_shift": cursor_shift,
                }
            )

        if campaigns and campaigns[-1]["code"] == chosen["code"]:
            campaigns[-1]["end_shift"] = cursor_shift
            campaigns[-1]["scheduled_qty"] = priority.base._clean_number(
                campaigns[-1]["scheduled_qty"] + produced
            )
        else:
            campaigns.append(
                {
                    "code": chosen["code"],
                    "group": chosen.get("product_group"),
                    "start_shift": start_shift,
                    "end_shift": cursor_shift,
                    "scheduled_qty": priority.base._clean_number(produced),
                }
            )

        last_code = chosen["code"]
        last_group = chosen.get("product_group")

    unserved_due = []
    urgent_carryover = {}
    safety_carryover = {}
    shortage_by_day = defaultdict(float)

    for product in products:
        missing_units = units_left(product)
        if missing_units <= 0:
            continue

        code = product["code"]
        quantum_qty = float(product["quantum_qty"])
        carryover[code] = priority.base._clean_number(missing_units * quantum_qty)

        probe_units = scheduled_units[code]
        urgent_qty = 0.0
        safety_qty = 0.0
        earliest_deadline = None
        for _ in range(missing_units):
            info = _next_unit_info(
                headers,
                product,
                capacity,
                probe_units,
            )
            probe_units += 1
            if info is None:
                continue
            deadline, critical, _ = info
            if critical:
                urgent_qty += quantum_qty
                earliest_deadline = (
                    deadline
                    if earliest_deadline is None
                    else min(earliest_deadline, deadline)
                )
            else:
                safety_qty += quantum_qty

        if urgent_qty > priority.base.EPSILON:
            urgent_carryover[code] = priority.base._clean_number(urgent_qty)
            deadline_index = max(
                0,
                min(
                    len(headers) - 1,
                    int(math.ceil(earliest_deadline / capacity - priority.base.EPSILON)) - 1,
                ),
            )
            due_day = headers[deadline_index]
            shortage_by_day[due_day] += urgent_qty
            unserved_due.append(
                {
                    "code": code,
                    "qty": priority.base._clean_number(urgent_qty),
                    "deadline_shift": earliest_deadline,
                    "due_date": due_day,
                }
            )

        if safety_qty > priority.base.EPSILON:
            safety_carryover[code] = priority.base._clean_number(safety_qty)

    return schedule, usage, carryover, {
        "mode": "shared_machine_deadline_guarded",
        "setup_shifts": setup_total,
        "campaigns": campaigns,
        # Backward-compatible field: late-completed urgent quantum only.
        "deadline_misses": deadline_misses,
        "late_completed": list(deadline_misses),
        "unserved_due": unserved_due,
        "urgent_carryover": urgent_carryover,
        "safety_carryover": safety_carryover,
        "shortage_by_day": {
            day: priority.base._clean_number(qty)
            for day, qty in sorted(shortage_by_day.items())
        },
        "urgent_buffer_days": URGENT_BUFFER_DAYS,
        "physical_lines": sorted(SHARED_MACHINE_LINES),
    }


def build_schedule_shared_machine(headers, products):
    """Build schedule with one physical machine shared by KHS and PET 9000."""
    schedule = priority.base._empty_schedule(headers, products)
    active_products = [product for product in products if product["planned_qty"] > 0]

    shared_products = [
        product for product in active_products if product["line"] in SHARED_MACHINE_LINES
    ]
    other_products = [
        product for product in products if product["line"] not in SHARED_MACHINE_LINES
    ]

    line_capacity = {}
    line_usage = defaultdict(lambda: defaultdict(float))
    carryover = {}
    optimizer_meta = {}

    if other_products:
        (
            other_schedule,
            other_capacity,
            other_usage,
            other_carryover,
            _,
            other_meta,
        ) = _ORIGINAL_BUILD_SCHEDULE(headers, other_products)

        line_capacity.update(other_capacity)
        carryover.update(other_carryover)
        optimizer_meta.update(other_meta)
        for line, days in other_usage.items():
            for current_day, used in days.items():
                line_usage[line][current_day] = used
        for product in other_products:
            for current_day in headers:
                schedule[product["code"]][current_day] = other_schedule[product["code"]][current_day]

    if shared_products:
        capacity = min(product["max_shifts_per_day"] for product in shared_products)
        shared_schedule, usage, shared_carryover, meta = allocate_shared_deadline_guarded(
            headers,
            shared_products,
            capacity,
        )

        line_capacity[SHARED_MACHINE_NAME] = capacity
        carryover.update(shared_carryover)
        optimizer_meta[SHARED_MACHINE_NAME] = meta

        for current_day, used in usage.items():
            line_usage[SHARED_MACHINE_NAME][current_day] = used
        for product in shared_products:
            for current_day in headers:
                schedule[product["code"]][current_day] = shared_schedule[product["code"]][current_day]

    priority.base.validate_schedule(
        headers,
        products,
        schedule,
        line_capacity,
        line_usage,
    )
    inventory = priority.base.simulate_inventory(headers, products, schedule)
    return schedule, line_capacity, line_usage, carryover, inventory, optimizer_meta


def install_priority_scheduler_v8():
    v7.install_priority_scheduler_v7()
    priority.base.build_schedule = build_schedule_shared_machine


if __name__ == "__main__":
    install_priority_scheduler_v8()
    priority.base.main_with_retry()
