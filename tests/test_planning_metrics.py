import unittest
from datetime import datetime
from io import BytesIO

from openpyxl import Workbook

from sync_planning_metrics import (
    calculate_product_metrics,
    excel_roundup_integer,
    read_system_receipts,
)


class PlanningMetricsTests(unittest.TestCase):
    def test_excel_roundup_away_from_zero(self):
        self.assertEqual(excel_roundup_integer(1.2), 2)
        self.assertEqual(excel_roundup_integer(-1.2), -2)
        self.assertEqual(excel_roundup_integer(0), 0)

    def test_matches_reference_row_without_debt(self):
        result = calculate_product_metrics(
            fc=5343.6,
            actual_stock=8444,
            book_stock=13161.833333333334,
            opening_consignment=2978,
            warehouse_debt=0,
            leadtime=3,
            batch=4500,
            per_shift=4500,
            shifts_per_day=3,
            classification="Không đường",
            plan_year=2026,
            plan_month=8,
        )

        self.assertAlmostEqual(result["expected_end_stock"], 0)
        self.assertAlmostEqual(
            result["required_production"],
            -7818.233333333334,
        )
        self.assertAlmostEqual(result["rounded_production"], -9000)
        self.assertAlmostEqual(
            result["production_days"],
            -0.6666666666666666,
        )
        self.assertEqual(
            result["production_start"].date(),
            datetime(2026, 9, 8).date(),
        )

    def test_matches_reference_row_with_debt_and_sugar_rounding(self):
        result = calculate_product_metrics(
            fc=79815.47174026055,
            actual_stock=43002,
            book_stock=1665.9166666666667,
            opening_consignment=10098.875000000002,
            warehouse_debt=2650.25,
            leadtime=4,
            batch=2700,
            per_shift=4000,
            shifts_per_day=3,
            classification="Có đường",
            plan_year=2026,
            plan_month=8,
        )

        self.assertAlmostEqual(
            result["expected_end_stock"],
            18418.955016983204,
        )
        self.assertAlmostEqual(
            result["required_production"],
            80799.80507359387,
        )
        self.assertAlmostEqual(result["rounded_production"], 81000)
        self.assertAlmostEqual(result["production_days"], 6.75)

    def test_fc_zero_returns_blank_start_date(self):
        result = calculate_product_metrics(
            fc=0,
            actual_stock=0,
            book_stock=18,
            opening_consignment=0,
            warehouse_debt=0,
            leadtime=3,
            batch=500,
            per_shift=500,
            shifts_per_day=3,
            classification="Không đường",
            plan_year=2026,
            plan_month=8,
        )
        self.assertIsNone(result["production_start"])

    def test_system_receipts_ignore_codes_outside_planning_catalog(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Sheet1"
        worksheet.append([None, "Mã vật tư", None, None, None, None, None, None, "Nhập trong kỳ"])
        worksheet.append([None, 130100096, None, None, None, None, None, None, 9000])
        worksheet.append([None, 130100014, None, None, None, None, None, None, 5000])

        buffer = BytesIO()
        workbook.save(buffer)
        workbook.close()

        result = read_system_receipts(
            buffer.getvalue(),
            {"130100096": 24},
        )

        self.assertEqual(result, {"130100096": 375})
        self.assertNotIn("130100014", result)


if __name__ == "__main__":
    unittest.main()
