import unittest
from datetime import date

import sync_planning_schedule as base
from sync_planning_schedule_priority_v4 import (
    _priority_key,
    allocate_continuous_min_slack,
)


class PriorityV4Tests(unittest.TestCase):
    def product(
        self,
        *,
        code,
        headers,
        actual_stock,
        fc,
        planned_qty,
        per_shift=100,
        earliest=None,
        row=2,
        group="A",
    ):
        demand_days = sum(1 for day in headers if day.weekday() != 6)
        daily = fc / demand_days if fc and demand_days else 0
        product = {
            "row": row,
            "code": code,
            "line": "KHS",
            "product_group": group,
            "planned_qty": float(planned_qty),
            "per_shift": float(per_shift),
            "batch": float(per_shift),
            "max_shifts_per_day": 1.0,
            "classification": "Không đường",
            "is_sugar": False,
            "earliest_date": earliest or headers[0],
            "debt": 0.0,
            "actual_stock": float(actual_stock),
            "fc": float(fc),
            "target_stock": 0.0,
            "demand_by_day": {
                day: (0.0 if day.weekday() == 6 else daily)
                for day in headers
            },
        }
        base._validate_quantum(product)
        return product

    def test_imminent_stockout_has_priority(self):
        headers = [date(2026, 9, day) for day in range(1, 8)]
        urgent = self.product(
            code="URGENT",
            headers=headers,
            actual_stock=0,
            fc=500,
            planned_qty=200,
            row=3,
        )
        safe = self.product(
            code="SAFE",
            headers=headers,
            actual_stock=10000,
            fc=0,
            planned_qty=200,
            row=2,
        )
        schedule = base._empty_schedule(headers, [urgent, safe])
        remaining = {"URGENT": 2, "SAFE": 2}
        self.assertLess(
            _priority_key(headers, urgent, schedule, remaining, 0, 1),
            _priority_key(headers, safe, schedule, remaining, 0, 1),
        )

        result, _, carryover, meta = allocate_continuous_min_slack(
            headers,
            [safe, urgent],
            1,
        )
        first_urgent = min(day for day, qty in result["URGENT"].items() if qty > 0)
        first_safe = min(day for day, qty in result["SAFE"].items() if qty > 0)
        self.assertLessEqual(first_urgent, first_safe)
        self.assertEqual(carryover, {})
        self.assertEqual(meta["mode"], "continuous_min_slack_shortage_first")

    def test_long_job_with_low_slack_beats_short_safe_job(self):
        headers = [date(2026, 10, day) for day in range(1, 11)]
        long_urgent = self.product(
            code="LONG",
            headers=headers,
            actual_stock=100,
            fc=900,
            planned_qty=500,
            row=4,
        )
        short_safe = self.product(
            code="SHORT",
            headers=headers,
            actual_stock=10000,
            fc=0,
            planned_qty=100,
            row=2,
        )
        schedule = base._empty_schedule(headers, [long_urgent, short_safe])
        remaining = {"LONG": 5, "SHORT": 1}
        self.assertLess(
            _priority_key(headers, long_urgent, schedule, remaining, 0, 1),
            _priority_key(headers, short_safe, schedule, remaining, 0, 1),
        )

    def test_february_leap_year_capacity(self):
        headers = [date(2028, 2, day) for day in range(1, 30)]
        item = self.product(
            code="LEAP",
            headers=headers,
            actual_stock=1000,
            fc=0,
            planned_qty=200,
            earliest=date(2028, 2, 28),
        )
        schedule, usage, carryover, _ = allocate_continuous_min_slack(
            headers,
            [item],
            1,
        )
        self.assertEqual(sum(schedule["LEAP"].values()), 200)
        self.assertEqual(carryover, {})
        self.assertLessEqual(max(usage.values()), 1.0)


if __name__ == "__main__":
    unittest.main()
