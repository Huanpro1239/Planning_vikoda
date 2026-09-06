import unittest
from datetime import date

import sync_planning_schedule as base
from sync_planning_metrics_all_months import (
    _planning_stock,
    calculate_row_all_months,
)
from sync_planning_schedule_priority_v5 import _effective_release_index


class PlanningStockAwareV5Tests(unittest.TestCase):
    def test_planning_stock_uses_lower_of_actual_and_book(self):
        self.assertEqual(_planning_stock(336, 790), 336)
        self.assertEqual(_planning_stock(1000, 700), 700)

    def test_actual_shortage_increases_P_even_when_book_stock_looks_enough(self):
        result = calculate_row_all_months(
            fc=500,
            actual_stock=336,
            book_stock=790,
            opening_consignment=0,
            warehouse_debt=0,
            leadtime=3,
            batch=1400,
            per_shift=1400,
            shifts_per_day=1,
            classification="Không đường",
            plan_year=2026,
            plan_month=9,
        )
        self.assertGreater(result["required_production"], 0)
        self.assertEqual(result["rounded_production"], 1400)

    def test_debt_case_also_uses_conservative_stock(self):
        result = calculate_row_all_months(
            fc=17146,
            actual_stock=459,
            book_stock=2603.2,
            opening_consignment=0,
            warehouse_debt=879,
            leadtime=3,
            batch=1600,
            per_shift=1600,
            shifts_per_day=2,
            classification="Không đường",
            plan_year=2026,
            plan_month=9,
        )
        self.assertEqual(result["rounded_production"], 17600)

    def test_risk_can_pull_release_before_nominal_R(self):
        headers = [date(2026, 9, day) for day in range(1, 31)]
        product = {
            "code": "A",
            "planned_qty": 88000.0,
            "per_shift": 4000.0,
            "max_shifts_per_day": 3.0,
            "actual_stock": 23441.0,
            "target_stock": 0.0,
            "debt": 0.0,
            "demand_by_day": {
                day: (0.0 if day.weekday() == 6 else 91975 / 26)
                for day in headers
            },
        }
        preferred = date(2026, 9, 4)
        release = _effective_release_index(headers, product, preferred)
        self.assertLess(release, 3)


if __name__ == "__main__":
    unittest.main()
