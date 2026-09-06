import unittest
from datetime import date

import sync_planning_schedule as base
import sync_planning_schedule_priority_v6 as v6
from sync_planning_schedule_priority_v7 import product_sequence_cost_early_shortage


class PlanningSequenceV7Tests(unittest.TestCase):
    def product(self, *, code, stock, demand, planned_qty, row):
        headers = [date(2026, 9, day) for day in range(1, 11)]
        item = {
            "row": row,
            "code": code,
            "line": "KHS",
            "product_group": code,
            "classification": "Không đường",
            "is_sugar": False,
            "planned_qty": float(planned_qty),
            "per_shift": 100.0,
            "batch": 100.0,
            "max_shifts_per_day": 1.0,
            "earliest_date": headers[0],
            "preferred_date": headers[0],
            "actual_stock": float(stock),
            "target_stock": 0.0,
            "debt": 0.0,
            "fc": float(demand * len(headers)),
            "demand_by_day": {day: float(demand) for day in headers},
        }
        base._validate_quantum(item)
        return headers, item

    def test_early_stockout_cost_is_worse_than_late_stockout(self):
        headers, product = self.product(
            code="A",
            stock=150,
            demand=100,
            planned_qty=300,
            row=2,
        )
        early = product_sequence_cost_early_shortage(headers, product, 5.0, 1.0)
        late = product_sequence_cost_early_shortage(headers, product, 1.0, 1.0)
        self.assertGreater(early[:4], late[:4])

    def test_optimizer_uses_v7_cost_after_patch(self):
        headers, urgent = self.product(
            code="URGENT",
            stock=150,
            demand=100,
            planned_qty=300,
            row=2,
        )
        _, safe = self.product(
            code="SAFE",
            stock=5000,
            demand=0,
            planned_qty=300,
            row=3,
        )
        original = v6._product_sequence_cost
        try:
            v6._product_sequence_cost = product_sequence_cost_early_shortage
            order = v6.optimize_sequence(headers, [safe, urgent], 1.0)
        finally:
            v6._product_sequence_cost = original
        self.assertEqual(order[0]["code"], "URGENT")


if __name__ == "__main__":
    unittest.main()
