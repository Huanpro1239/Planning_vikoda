import unittest
from datetime import date

from planning import metrics
from sync_planning_metrics_direct import calculate_metrics_from_no_kho


class NonnegativeProductionTests(unittest.TestCase):
    def test_negative_o_and_p_are_clamped_to_zero(self):
        planning_rows = {
                "130100091": {
                    "fc": 1421,
                    "actual_stock": 1605,
                    "book_stock": 2051.20833333333,
                    "current_debt": 0,
                    "batch": 500,
                    "per_shift": 2500,
                    "shifts_per_day": 3,
                    "classification": "Không đường",
                }
            }
            result, _, _, _ = calculate_metrics(
                report_date=date(2026, 8, 31),
                plan_month=9,
                planning_rows=planning_rows,
                actual_receipts={},
                system_receipts={},
                current_consignments={"130100091": 10000},
            state={},
            leadtimes={"130100091": 3},
        )
        row = result["130100091"]
        self.assertEqual(row["required_production"], 0)
        self.assertEqual(row["rounded_production"], 0)
        self.assertEqual(row["production_days"], 0)


if __name__ == "__main__":
    unittest.main()
