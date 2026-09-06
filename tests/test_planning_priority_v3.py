import unittest
from datetime import date

import sync_planning_schedule as base
from sync_planning_schedule_priority_v3 import (
    _dynamic_priority_key,
    allocate_continuous_stockout_first,
)


class PlanningPriorityV3Tests(unittest.TestCase):
    def product(
        self,
        *,
        code,
        planned_qty,
        actual_stock,
        daily_demand,
        target_stock=0,
        earliest=date(2026, 9, 1),
        row=2,
        group="A",
    ):
        headers = [date(2026, 9, day) for day in range(1, 16)]
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
            "earliest_date": earliest,
            "actual_stock": float(actual_stock),
            "target_stock": float(target_stock),
            "debt": 0.0,
            "fc": float(daily_demand * len(headers)),
            "demand_by_day": {day: float(daily_demand) for day in headers},
        }
        base._validate_quantum(item)
        return headers, item

    def test_stockout_risk_outranks_safety_only_risk(self):
        headers, stockout = self.product(
            code="STOCKOUT",
            planned_qty=300,
            actual_stock=150,
            daily_demand=100,
            target_stock=0,
            row=3,
        )
        _, safety = self.product(
            code="SAFETY",
            planned_qty=300,
            actual_stock=500,
            daily_demand=20,
            target_stock=450,
            row=2,
        )
        schedule = base._empty_schedule(headers, [stockout, safety])
        self.assertLess(
            _dynamic_priority_key(headers, stockout, schedule, 0),
            _dynamic_priority_key(headers, safety, schedule, 0),
        )

    def test_long_campaign_can_be_split_to_protect_other_stockout(self):
        headers, a = self.product(
            code="A",
            planned_qty=900,
            actual_stock=150,
            daily_demand=100,
            row=2,
            group="A",
        )
        _, b = self.product(
            code="B",
            planned_qty=300,
            actual_stock=250,
            daily_demand=100,
            row=3,
            group="B",
        )
        schedule, usage, carryover, meta = allocate_continuous_stockout_first(
            headers,
            [a, b],
            3.0,
        )
        self.assertEqual(carryover, {})
        self.assertEqual(sum(schedule["A"].values()), 900)
        self.assertEqual(sum(schedule["B"].values()), 300)
        self.assertGreaterEqual(len(meta["campaigns"]), 2)
        for used in usage.values():
            self.assertLessEqual(used, 3.0 + 1e-9)


if __name__ == "__main__":
    unittest.main()
