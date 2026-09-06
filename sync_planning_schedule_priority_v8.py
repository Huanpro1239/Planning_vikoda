import math
from collections import defaultdict

import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v7 as v7


SHARED_MACHINE_LINES = {"KHS", "PET 9000"}
SHARED_MACHINE_NAME = "KHS + PET 9000"
URGENT_BUFFER_DAYS = 1


_ORIGINAL_BUILD_SCHEDULE = priority.base.build_schedule


def _next_unit_detail(headers, product, capacity, scheduled_units):
    """Return structured deadline/release data for the next production quantum."""
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
            # Business due date: the day demand makes this quantum necessary.
            demand_deadline_shift = (index + 1) * capacity
            # Production deadline: one production day earlier when possible.
            production_deadline_shift = max(
                capacity,
                demand_deadline_shift - URGENT_BUFFER_DAYS * capacity,
            )
            production_index = max(
                0,
                min(
                    len(headers) - 1,
                    int(
                        math.ceil(
                            production_deadline_shift / capacity
                            - priority.base.EPSILON
                        )
                    )
                    - 1,
                ),
            )
            return {
                "unit_no": unit_no,
                "critical": True,
                "release_shift": 0.0,
                "demand_deadline_shift": demand_deadline_shift,
                "production_deadline_shift": production_deadline_shift,
                "demand_due_date": current_day,
                "production_deadline_date": headers[production_index],
            }

    # Quantum dư do làm tròn hoặc safety build: không được chen trước urgent.
    release_index = priority.base._earliest_index(headers, product)
    return {
        "unit_no": unit_no,
        "critical": False,
        "release_shift": release_index * capacity,
        "demand_deadline_shift": None,
        "production_deadline_shift": len(headers) * capacity,
        "demand_due_date": None,
        "production_deadline_date": headers[-1] if headers else None,
    }


def _next_unit_info(headers, product, capacity, scheduled_units):
    """Backward-compatible tuple used by the priority functions."""
    detail = _next_unit_detail(
        headers,
        product,
        capacity,
        scheduled_units,
    )
    if detail is None:
        return None
    return (
        detail["production_deadline_shift"],
        detail["critical"],
        detail["release_shift"],
    )


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
    timeline = []
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
            setup_start = cursor_shift
            setup_end = cursor_shift + priority.base.SETUP_SHIFTS
            priority.base._add_interval_usage(
                usage,
                headers,
                setup_start,
                priority.base.SETUP_SHIFTS,
                capacity,
            )
            timeline.append(
                {
                    "kind": "setup",
                    "from_code": last_code,
                    "to_code": chosen["code"],
                    "start_shift": setup_start,
                    "end_shift": setup_end,
                    "duration_shifts": priority.base.SETUP_SHIFTS,
                }
            )
            cursor_shift = setup_end
            setup_total += priority.base.SETUP_SHIFTS

        duration = float(chosen["quantum_shift"])
        if cursor_shift + duration > month_end_shift + priority.base.EPSILON:
            break

        detail = _next_unit_detail(
            headers,
            chosen,
            capacity,
            scheduled_units[chosen["code"]],
        )
        deadline = detail["production_deadline_shift"]
        critical = detail["critical"]
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
        timeline.append(
            {
                "kind": "production",
                "code": chosen["code"],
                "uom": str(chosen.get("uom") or ""),
                "unit_no": scheduled_units[chosen["code"]],
                "qty": priority.base._clean_number(produced),
                "start_shift": start_shift,
                "end_shift": cursor_shift,
                "duration_shifts": duration,
                "critical": bool(critical),
                "demand_due_date": detail["demand_due_date"],
                "production_deadline_date": detail["production_deadline_date"],
                "demand_deadline_shift": detail["demand_deadline_shift"],
                "production_deadline_shift": detail["production_deadline_shift"],
            }
        )

        if critical and cursor_shift > deadline + priority.base.EPSILON:
            deadline_misses.append(
                {
                    "code": chosen["code"],
                    "uom": str(chosen.get("uom") or ""),
                    "unit_no": scheduled_units[chosen["code"]],
                    "qty": priority.base._clean_number(expected),
                    "deadline_shift": deadline,
                    "finish_shift": cursor_shift,
                    "demand_due_date": detail["demand_due_date"],
                    "production_deadline_date": detail["production_deadline_date"],
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
    shortage_groups = defaultdict(float)

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
        for _ in range(missing_units):
            detail = _next_unit_detail(
                headers,
                product,
                capacity,
                probe_units,
            )
            probe_units += 1
            if detail is None:
                continue
            if detail["critical"]:
                urgent_qty += quantum_qty
                event = {
                    "code": code,
                    "uom": str(product.get("uom") or ""),
                    "unit_no": detail["unit_no"],
                    "qty": priority.base._clean_number(quantum_qty),
                    # Legacy aliases retained for older consumers.
                    "deadline_shift": detail["production_deadline_shift"],
                    "due_date": detail["production_deadline_date"],
                    # Explicit business/prod deadlines for operations.
                    "demand_deadline_shift": detail["demand_deadline_shift"],
                    "production_deadline_shift": detail["production_deadline_shift"],
                    "demand_due_date": detail["demand_due_date"],
                    "production_deadline_date": detail["production_deadline_date"],
                }
                unserved_due.append(event)
                shortage_groups[
                    (
                        code,
                        str(product.get("uom") or ""),
                        detail["demand_due_date"],
                    )
                ] += quantum_qty
            else:
                safety_qty += quantum_qty

        if urgent_qty > priority.base.EPSILON:
            urgent_carryover[code] = priority.base._clean_number(urgent_qty)
        if safety_qty > priority.base.EPSILON:
            safety_carryover[code] = priority.base._clean_number(safety_qty)

    shortage_by_day = [
        {
            "code": code,
            "uom": uom,
            "demand_due_date": due_date,
            "qty": priority.base._clean_number(qty),
        }
        for (code, uom, due_date), qty in sorted(
            shortage_groups.items(),
            key=lambda item: (item[0][2], item[0][0]),
        )
    ]

    return schedule, usage, carryover, {
        "mode": "shared_machine_deadline_guarded",
        "setup_shifts": setup_total,
        "campaigns": campaigns,
        "timeline": timeline,
        # Backward-compatible field: late-completed urgent quantum only.
        "deadline_misses": deadline_misses,
        "late_completed": list(deadline_misses),
        "unserved_due": unserved_due,
        "urgent_carryover": urgent_carryover,
        "safety_carryover": safety_carryover,
        # List by SKU/UOM/date; never sums unlike units into one scalar.
        "shortage_by_day": shortage_by_day,
        "urgent_buffer_days": URGENT_BUFFER_DAYS,
        "physical_lines": sorted(SHARED_MACHINE_LINES),
        "capacity_shifts_per_day": capacity,
        "horizon_shifts": month_end_shift,
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
