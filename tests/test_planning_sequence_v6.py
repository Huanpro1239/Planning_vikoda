import unittest
from datetime import date

import sync_planning_schedule as base
from sync_planning_schedule_priority_v6 import (
    allocate_sequence_optimized_line,
    optimize_sequence,
)


class PlanningSequenceV6Tests(unittest.TestCase):
    def product(
        self,
        *,
        code,
        stock,
        demand,
        planned_qty=300,
        group="A",
        preferred=date(2026, 9, 1),
        row=2,
    ):
        headers = [date(2026, 9, day) for day in range(1, 11)]
        item = {
            "row": row,
            "code": code,
            "line": "KHS",
            "product_group": group,
            "classification": "Không đường",
            "is_sugar": False,
            "planned_qty": float(planned_qty),
            "per_shift": 100.0,
            "batch": 100.0,
            "max_shifts_per_day": 3.0,
            "earliest_date": headers[0],
            "preferred_date": preferred,
            "actual_stock": float(stock),
            "target_stock": 0.0,
            "debt": 0.0,
            "fc": float(demand * len(headers)),
            "demand_by_day": {day: float(demand) for day in headers},
        }
        base._validate_quantum(item)
        return headers, item

    def test_stockout_urgent_sku_is_placed_before_safe_sku(self):
        headers, urgent = self.product(
            code="URGENT",
            stock=50,
            demand=100,
            row=3,
            group="B",
        )
        _, safe = self.product(
            code="SAFE",
            stock=5000,
            demand=0,
            row=2,
            group="A",
        )
        order = optimize_sequence(headers, [safe, urgent], 3.0)
        self.assertEqual(order[0]["code"], "URGENT")

    def test_each_sku_has_one_campaign_and_minimum_setup(self):
        headers, a = self.product(code="A", stock=1000, demand=10, row=2)
        _, b = self.product(code="B", stock=1000, demand=10, row=3)
        _, c = self.product(code="C", stock=1000, demand=10, row=4)
        schedule, usage, carryover, meta = allocate_sequence_optimized_line(
            headers,
            [a, b, c],
            3.0,
        )
        self.assertEqual(len(meta["campaigns"]), 3)
        self.assertAlmostEqual(meta["setup_shifts"], 1.0)
        self.assertEqual(carryover, {})
        for product in (a, b, c):
            self.assertEqual(sum(schedule[product["code"]].values()), 300)
        for used in usage.values():
            self.assertLessEqual(used, 3.0 + 1e-9)

    def test_equal_service_prefers_fewer_group_changes(self):
        headers, a1 = self.product(
            code="A1", stock=5000, demand=0, group="A", row=2
        )
        _, b = self.product(
            code="B", stock=5000, demand=0, group="B", row=3
        )
        _, a2 = self.product(
            code="A2", stock=5000, demand=0, group="A", row=4
        )
        order = optimize_sequence(headers, [a1, b, a2], 3.0)
        groups = [item["product_group"] for item in order]
        transitions = sum(1 for x, y in zip(groups, groups[1:]) if x != y)
        self.assertEqual(transitions, 1)


if __name__ == "__main__":
    unittest.main()
