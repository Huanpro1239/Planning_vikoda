import sync_planning_schedule_rgb_v9 as serial_v9


def allocate_galon_quantized_service(headers, products, capacity):
    """
    Galon V9 uses the same exact serial-machine model as RGB V9.

    The previous spread allocator could put 19L and another Galon SKU into a
    full 2-shift day and still report a changeover separately. This wrapper
    searches whole-campaign SKU order while reserving every 0.5-shift setup in
    the physical capacity and respecting the product quantum/batch rules.
    """
    schedule, usage, carryover, meta = serial_v9.allocate_rgb_quantized_service(
        headers,
        products,
        capacity,
    )
    meta = dict(meta)
    meta["mode"] = "galon_quantized_service_v9"
    meta["resource"] = "Galon"
    return schedule, usage, carryover, meta


def install_galon_service_scheduler(base_module):
    base_module._allocate_galon_line = allocate_galon_quantized_service
