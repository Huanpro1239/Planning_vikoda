import unittest
from datetime import datetime

from sync_planning_metrics import (
    _bootstrap_opening_debt,
    calculate_row,
    excel_roundup_integer,
)


class DirectPlanningMetricsTests(unittest.TestCase):
    def test_roundup_matches_excel(self):
        self.assertEqual(excel_roundup_integer(1.2), 2)
        self.assertEqual(excel_roundup_integer(-1.2), -2)
        self.assertEqual(excel_roundup_integer(0), 0)

    def test_bootstrap_debt_uses_current_sheet_and_raw_sources(self):
        # 130100006 tại file hiện hành:
        # N hiện tại 15,587; nhập thực tế 178,525; nhập hệ thống 170,000.
        self.assertEqual(
            _bootstrap_opening_debt(
                current_debt=15587,
                actual_receipt=178525,
                system_receipt=170000,
            ),
            24112,
        )

    def test_calculates_M_to_R_directly(self):
        result = calculate_row(
            fc=9370,
            actual_stock=6534,
            book_stock=7376.83333333334,
            opening_consignment=0,
            warehouse_debt=13,
            leadtime=3,
            batch=4500,
            per_shift=4500,
            shifts_per_day=3,
            classification="Không đường",
            plan_year=2026,
            plan_month=9,
        )

        self.assertAlmostEqual(
            result["expected_end_stock"],
            9370 / 26 * 5,
        )
        self.assertEqual(result["warehouse_debt"], 13)
        self.assertAlmostEqual(
            result["required_production"],
            9370 + 13 - 7376.83333333334,
        )
        self.assertEqual(result["rounded_production"], 4500)
        self.assertAlmostEqual(result["production_days"], 1 / 3)
        self.assertIsInstance(result["production_start"], datetime)

    def test_sugar_product_rounds_by_batch(self):
        result = calculate_row(
            fc=28389,
            actual_stock=18829,
            book_stock=18828.2083333333,
            opening_consignment=0,
            warehouse_debt=0,
            leadtime=1,
            batch=5200,
            per_shift=4000,
            shifts_per_day=3,
            classification="Có đường",
            plan_year=2026,
            plan_month=9,
        )

        self.assertAlmostEqual(
            result["expected_end_stock"],
            28389 / 26 * 3,
        )
        self.assertEqual(result["rounded_production"], 15600)
        self.assertAlmostEqual(result["production_days"], 1.3)

    def test_zero_fc_returns_blank_start_date(self):
        result = calculate_row(
            fc=0,
            actual_stock=895,
            book_stock=913,
            opening_consignment=0,
            warehouse_debt=0,
            leadtime=3,
            batch=500,
            per_shift=500,
            shifts_per_day=3,
            classification="Không đường",
            plan_year=2026,
            plan_month=9,
        )
        self.assertIsNone(result["production_start"])


if __name__ == "__main__":
    unittest.main()
