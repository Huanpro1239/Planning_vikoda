import math

import sync_planning_schedule_priority as priority


def shortage_first_campaign_candidate_v2(
    remaining,
    headers,
    cursor_shift,
    capacity,
    last_group,
):
    if not remaining:
        return None, cursor_shift

    current_index = min(
        len(headers),
        int(math.floor(cursor_shift / capacity + priority.base.EPSILON)),
    )
    eligible = [
        product
        for product in remaining
        if priority.base._earliest_index(headers, product) <= current_index
    ]

    if not eligible:
        next_index = min(
            priority.base._earliest_index(headers, product)
            for product in remaining
        )
        cursor_shift = max(cursor_shift, next_index * capacity)
        current_index = next_index
        eligible = [
            product
            for product in remaining
            if priority.base._earliest_index(headers, product) <= current_index
        ]

    primary = min(
        eligible,
        key=lambda product: priority._priority_key(
            headers,
            product,
            cursor_shift,
            capacity,
            last_group,
        ),
    )

    primary_risk = priority._risk_snapshot(headers, primary)

    # Rule cứng: nếu SKU sẽ chạm safety/stockout trong hôm nay hoặc ngày kế tiếp,
    # chạy SKU đó trước. Không cho tối ưu continuity/đổi khuôn lấn át service level.
    if (
        primary_risk["deadline_index"] is not None
        and primary_risk["deadline_index"] <= current_index + 1
    ):
        return primary, cursor_shift

    if last_group:
        same_group = [
            product
            for product in eligible
            if product.get("product_group") == last_group
        ]
        if same_group and primary not in same_group:
            candidate = min(
                same_group,
                key=lambda product: priority._priority_key(
                    headers,
                    product,
                    cursor_shift,
                    capacity,
                    last_group,
                ),
            )
            if priority._can_keep_same_group(
                headers,
                candidate,
                primary,
                cursor_shift,
                capacity,
                last_group,
            ):
                return candidate, cursor_shift

    return primary, cursor_shift


def install_priority_scheduler_v2():
    # Cập nhật global trong module priority để cả Galon fallback cũng dùng rule mới.
    priority.shortage_first_campaign_candidate = shortage_first_campaign_candidate_v2
    priority.install_priority_scheduler()
    priority.base._campaign_candidate = shortage_first_campaign_candidate_v2


if __name__ == "__main__":
    install_priority_scheduler_v2()
    priority.base.main_with_retry()
