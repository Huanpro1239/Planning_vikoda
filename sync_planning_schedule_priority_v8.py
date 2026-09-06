from collections import defaultdict

import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v7 as v7


SHARED_MACHINE_LINES = {"KHS", "PET 9000"}
SHARED_MACHINE_NAME = "KHS + PET 9000"


_ORIGINAL_BUILD_SCHEDULE = priority.base.build_schedule


def build_schedule_shared_machine(headers, products):
    """
    KHS và PET 9000 dùng chung một máy nên phải chia sẻ cùng một timeline ca.

    Hai line không còn được schedule độc lập. Toàn bộ SKU KHS/PET được đưa vào
    cùng sequence optimizer stockout-first; một SKU = một campaign liên tục.
    Các line khác (RGB/Galon/...) giữ allocator hiện hành.
    """
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

    # Các resource khác dùng logic hiện hành của base scheduler.
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
        allocator = priority.base._allocate_continuous_campaign_line
        shared_schedule, usage, shared_carryover, meta = allocator(
            headers,
            shared_products,
            capacity,
        )
        meta = dict(meta)
        meta["mode"] = "shared_machine_stockout_first"
        meta["physical_lines"] = sorted(SHARED_MACHINE_LINES)

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
