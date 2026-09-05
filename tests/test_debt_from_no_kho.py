import unittest
from datetime import date
from io import BytesIO

from openpyxl import Workbook

import sync_planning_metrics as metrics
from sync_planning_metrics_direct import (
    calculate_metrics_from_no_kho,
    read_debt_from_no_kho,
)


class DebtFromNoKhoTests(unittest.TestCase):
    def make_workbook(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "No kho"
        worksheet.append(["Mã Sản Phẩm", "Tên Sản Phẩm", "Đơn vị tính", "Số lượng nợ"])
        worksheet.append([130100011, "SP A", "Thùng", 2650])
        worksheet.append([130100096, "SP B", "Thùng", None])
        output = BytesIO()
        workbook.save(output)
        return output.getvalue()

    def test_reads_column_d_by_product_code(self):
        debts = read_debt_from_no_kho(self.make_workbook())
        self.assertEqual(debts["130100011"], 2650)
        self.assertEqual(debts["130100096"], 0)

    def test_n_and_o_use_no_kho_value_directly(self):
        old_mapping = metrics.LEADTIME_BY_CODE
        metrics.LEADTIME_BY_CODE = {"130100011": 4}
        try:
            planning_rows = {
                "130100011": {
                    "fc": 57019,
                    "actual_stock": 7176,
                    "book_stock": 1708,
                    "current_debt": 2650,
                    "batch": 2700,
                    "per_shift": 4000,
                    "shifts_per_day": 3,
                    "classification": "Có đường",
                }
            }
            result, _, _, _ = calculate_metrics_from_no_kho(
                report_date=date(2026, 8, 31),
                plan_month=9,
                planning_rows=planning_rows,
                actual_receipts={},
                system_receipts={},
                current_consignments={"130100011": 0},
                state={},
            )
            self.assertEqual(result["130100011"]["warehouse_debt"], 2650)
            self.assertAlmostEqual(
                result["130100011"]["required_production"],
                57019 + 2650 - 1708,
            )
        finally:
            metrics.LEADTIME_BY_CODE = old_mapping


if __name__ == "__main__":
    unittest.main()
