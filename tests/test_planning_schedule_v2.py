import unittest
from datetime import date

import sync_planning_schedule as base
from sync_planning_schedule_v2 import (
    _compute_risk_profile,
    _enrich_risk,
    _risk_aware_weekly_unit_targets,
    build_schedule,
)


class PlanningScheduleV2Tests(unittest.TestCase):
    def product(
        self,
        *,
        code,
        line="KHS",
        group="Pet 430",
        planned_qty=100,
        per_shift=100,
        batch=100,
        shifts_per_day=1,
        classification="Không đường",
        earliest=date(2026, 9, 1),
        debt=0,
        actual_stock=1000,
        fc=0,
        target_stock=0,
        row=2,
        headers=None,
    ):
        headers = headers or [date(2026, 9, day) for day in range(1, 31)]
        demand_days = sum(1 for day in headers if day.weekday() != 6)
        daily = fc / demand_days if demand_days and fc else 0
        item = {
            "row": row,
            "code": code,
            "line": line,
            "product_group": group,
            "planned_qty": float(planned_qty),
            "per_shift": float(per_shift),
            "batch": float(batch),
            "max_shifts_per_day": float(shifts_per_day),
            "classification": classification,
            "is_sugar": classification.casefold() == "có đường".casefold(),
            "earliest_date": earliest,
            "debt": float(debt),
            "actual_stock": float(actual_stock),
            "fc": float(fc),
            "target_stock": float(target_stock),
            "demand_by_day": {
                day: (0.0 if day.weekday() == 6 else daily)
                for day in headers
            },
        }
        base._validate_quantum(item)
        return item

    def test_imminent_shortage_runs_before_safe_sku(self):
        headers = [date(2026, 9, day) for day in range(1, 6)]
        urgent = self.product(
            code="URGENT",
            actual_stock=0,
            fc=400,
            target_stock=0,
            headers=headers,
            row=3,
        )
        safe = self.product(
            code="SAFE",
            actual_stock=10000,
            fc=0,
            headers=headers,
            row=2,
        )
        schedule, _, _, _, _, meta = build_schedule(headers, [safe, urgent])
        first_urgent = min(day for day, qty in schedule["URGENT"].items() if qty > 0)
        first_safe = min(day for day, qty in schedule["SAFE"].items() if qty > 0)
        self.assertLessEqual(first_urgent, first_safe)
        self.assertEqual(meta["KHS"]["mode"], "continuous_campaign_risk_aware")

    def test_risk_profile_uses_target_stock(self):
        headers = [date(2026, 9, day) for day in range(1, 4)]
        product = self.product(
            code="A",
            actual_stock=250,
            fc=300,
            target_stock=100,
            headers=headers,
        )
        product["demand_by_day"] = {headers[0]: 100, headers[1]: 100, headers[2]: 100}
        profile = _compute_risk_profile(headers, product)
        self.assertEqual(profile["critical_date"], headers[1])
        self.assertEqual(profile["stockout_date"], headers[2])

    def test_weekly_rgb_frontloads_when_stock_is_at_risk(self):
        headers = [date(2026, 9, day) for day in range(1, 15)]
        product = self.product(
            code="RGB",
            line="RGB",
            group="RGB có gas",
            planned_qty=400,
            per_shift=100,
            actual_stock=0,
            fc=400,
            target_stock=0,
            headers=headers,
        )
        _enrich_risk(headers, [product])
        targets = _risk_aware_weekly_unit_targets(headers, product)
        first_bucket = base._month_week_buckets(headers)[0]
        self.assertGreaterEqual(targets[first_bucket], 2)
        self.assertEqual(sum(targets.values()), product["required_units"])

    def test_sunday_remains_available_for_production(self):
        sunday = date(2026, 9, 6)
        headers = [sunday, date(2026, 9, 7)]
        product = self.product(
            code="A",
            earliest=sunday,
            actual_stock=0,
            fc=100,
            headers=headers,
        )
        product["demand_by_day"] = {day: 0 for day in headers}
        schedule, _, _, carryover, _, _ = build_schedule(headers, [product])
        self.assertEqual(schedule["A"][sunday], 100)
        self.assertEqual(carryover, {})

    def test_february_28_day_horizon(self):
        headers = [date(2027, 2, day) for day in range(1, 29)]
        product = self.product(
            code="FEB",
            planned_qty=200,
            per_shift=100,
            shifts_per_day=2,
            earliest=date(2027, 2, 27),
            headers=headers,
        )
        product["demand_by_day"] = {day: 0 for day in headers}
        schedule, _, _, carryover, _, _ = build_schedule(headers, [product])
        self.assertEqual(sum(schedule["FEB"].values()), 200)
        self.assertEqual(carryover, {})
        self.assertTrue(all(day.month == 2 for day in schedule["FEB"]))

    def test_february_leap_year_29_day_horizon(self):
        headers = [date(2028, 2, day) for day in range(1, 30)]
        product = self.product(
            code="LEAP",
            planned_qty=100,
            earliest=date(2028, 2, 29),
            headers=headers,
        )
        product["demand_by_day"] = {day: 0 for day in headers}
        schedule, _, _, carryover, _, _ = build_schedule(headers, [product])
        self.assertEqual(schedule["LEAP"][date(2028, 2, 29)], 100)
        self.assertEqual(carryover, {})


if __name__ == "__main__":
    unittest.main()
