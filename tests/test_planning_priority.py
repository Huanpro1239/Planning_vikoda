import unittest
from datetime import date, datetime

import sync_planning_schedule as base
from sync_planning_calendar import build_date_headers
from sync_planning_calendar_all_months import resolve_plan_year as resolve_calendar_year
from sync_planning_metrics_all_months import _demand_days, resolve_plan_year as resolve_metrics_year
from sync_planning_schedule_priority import (
    _risk_snapshot,
    shortage_aware_weekly_unit_targets,
)
from sync_planning_schedule_priority_v2 import shortage_first_campaign_candidate_v2


class PlanningPriorityTests(unittest.TestCase):
    def product(
        self,
        *,
        code,
        group="A",
        planned_qty=100,
        per_shift=100,
        batch=100,
        earliest=date(2026, 9, 1),
        actual_stock=1000,
        target_stock=0,
        debt=0,
        fc=0,
        row=2,
    ):
        headers = [date(2026, 9, day) for day in range(1, 31)]
        demand_days = sum(1 for day in headers if day.weekday() != 6)
        daily = fc / demand_days if fc and demand_days else 0
        item = {
            "row": row,
            "code": code,
            "line": "KHS",
            "product_group": group,
            "classification": "Không đường",
            "is_sugar": False,
            "planned_qty": float(planned_qty),
            "per_shift": float(per_shift),
            "batch": float(batch),
            "max_shifts_per_day": 3.0,
            "earliest_date": earliest,
            "actual_stock": float(actual_stock),
            "target_stock": float(target_stock),
            "debt": float(debt),
            "fc": float(fc),
            "demand_by_day": {
                day: (0.0 if day.weekday() == 6 else daily)
                for day in headers
            },
        }
        base._validate_quantum(item)
        return headers, item

    def test_shortage_risk_beats_same_group_continuity(self):
        headers, urgent = self.product(
            code="URGENT",
            group="B",
            actual_stock=0,
            fc=2600,
            row=3,
        )
        _, same_group = self.product(
            code="SAME",
            group="A",
            actual_stock=10000,
            fc=0,
            row=2,
        )
        candidate, _ = shortage_first_campaign_candidate_v2(
            [same_group, urgent], headers, 0.0, 3.0, "A"
        )
        self.assertEqual(candidate["code"], "URGENT")

    def test_keeps_same_group_when_urgent_sku_still_safe(self):
        headers, urgent = self.product(
            code="URGENT",
            group="B",
            planned_qty=100,
            actual_stock=900,
            fc=2600,
            row=3,
        )
        _, same_group = self.product(
            code="SAME",
            group="A",
            planned_qty=100,
            actual_stock=10000,
            fc=0,
            row=2,
        )
        candidate, _ = shortage_first_campaign_candidate_v2(
            [same_group, urgent], headers, 0.0, 3.0, "A"
        )
        self.assertEqual(candidate["code"], "SAME")

    def test_weekly_target_frontloads_when_inventory_is_low(self):
        headers, rgb = self.product(
            code="RGB",
            group="RGB có gas",
            planned_qty=400,
            per_shift=100,
            batch=100,
            actual_stock=0,
            target_stock=0,
            fc=400,
        )
        targets = shortage_aware_weekly_unit_targets(headers, rgb)
        first_bucket = list(targets)[0]
        self.assertGreaterEqual(targets[first_bucket], 1)
        self.assertEqual(sum(targets.values()), rgb["required_units"])

    def test_risk_snapshot_detects_stockout(self):
        headers, product = self.product(
            code="A",
            actual_stock=50,
            fc=2600,
        )
        risk = _risk_snapshot(headers, product)
        self.assertIsNotNone(risk["first_stockout"])

    def test_all_month_day_counts_and_leap_year(self):
        self.assertEqual(len(build_date_headers(2028, 2)), 29)
        self.assertEqual(len(build_date_headers(2027, 2)), 28)
        self.assertEqual(len(build_date_headers(2026, 4)), 30)
        self.assertEqual(len(build_date_headers(2026, 1)), 31)
        self.assertGreaterEqual(_demand_days(2026, 1), 24)
        self.assertLessEqual(_demand_days(2026, 1), 27)

    def test_year_rollover_is_consistent(self):
        self.assertEqual(resolve_calendar_year(1, datetime(2026, 12, 15)), 2027)
        self.assertEqual(resolve_metrics_year(date(2026, 12, 31), 1), 2027)
        self.assertEqual(resolve_calendar_year(12, datetime(2027, 1, 10)), 2026)
        self.assertEqual(resolve_metrics_year(date(2027, 1, 1), 12), 2026)


if __name__ == "__main__":
    unittest.main()
