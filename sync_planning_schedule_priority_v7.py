import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v6 as v6


def product_sequence_cost_early_shortage(headers, product, start_shift, capacity):
    """
    Một SKU chỉ có một campaign. Objective ưu tiên:
    1) giảm số SKU từng stockout;
    2) đẩy stockout còn bất khả kháng càng muộn càng tốt;
    3) giảm độ sâu thiếu hàng;
    4) đẩy safety breach càng muộn càng tốt;
    5) carry-over, trễ R, nợ kho, chạy sớm, đổi nhóm.
    """
    production, carryover = v6._production_by_day(
        headers,
        product,
        start_shift,
        capacity,
    )
    stock = float(product.get("actual_stock", 0) or 0)
    target = max(float(product.get("target_stock", 0) or 0), 0.0)

    has_stockout = 0
    stockout_urgency = 0
    stockout_depth = 0.0
    safety_urgency = 0
    day_count = len(headers)

    for index, current_day in enumerate(headers):
        stock += production[current_day]
        stock -= float(product["demand_by_day"].get(current_day, 0) or 0)

        # Ngày đầu tháng có trọng số lớn nhất. Lũy thừa 3 đủ mạnh để
        # không "hy sinh" một SKU từ đầu tháng chỉ để giảm deficit ở cuối tháng.
        urgency_weight = (day_count - index) ** 3
        if stock < -priority.base.EPSILON:
            has_stockout = 1
            stockout_urgency += urgency_weight
            stockout_depth += (-stock) * (day_count - index)
        if stock < target - priority.base.EPSILON:
            safety_urgency += urgency_weight

    start_index = min(
        len(headers),
        int(start_shift / capacity + priority.base.EPSILON),
    )
    preferred_index = v6._preferred_index(headers, product)
    lateness = max(0, start_index - preferred_index)
    earliness = max(0, preferred_index - start_index)
    debt_delay = start_index if float(product.get("debt", 0) or 0) > 0 else 0

    return (
        has_stockout,
        stockout_urgency,
        round(stockout_depth, 6),
        safety_urgency,
        round(carryover, 6),
        lateness,
        debt_delay,
        earliness,
        0,
    )


def install_priority_scheduler_v7():
    v6.install_priority_scheduler_v6()
    # optimize_sequence của v6 lookup hàm này trong namespace module v6 ở runtime.
    v6._product_sequence_cost = product_sequence_cost_early_shortage


if __name__ == "__main__":
    install_priority_scheduler_v7()
    priority.base.main_with_retry()
