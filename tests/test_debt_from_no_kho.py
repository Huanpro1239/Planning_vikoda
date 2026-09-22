import unittest
from datetime import date
from io import BytesIO

from openpyxl import Workbook

from planning.metrics import calculate_metrics, read_debt_from_no_kho


class DebtFromNoKhoTests(unittest.TestCase):
    def make_workbook(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "No kho"
        worksheet.append(
            ["Mã Sản Phẩm", "Tên Sản Phẩm", "Đơn vị tính", "Số lượng nợ"]
        )
        worksheet.append([130100011, "SP A", "Thùng", 2650])
        worksheet.append([130100096, "SP B", "Thùng", None])
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        return output.getvalue()

    def test_reads_column_d_by_product_code(self):
        debts = read_debt_from_no_kho(self.make_workbook())
        self.assertEqual(debts["130100011"], 2650)
        self.assertEqual(debts["130100096"], 0)

    def test_n_uses_no_kho_and_urgent_o_uses_actual_stock(self):
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
        result, _, _, _ = calculate_metrics(
            report_date=date(2026, 8, 31),
            plan_month=9,
            planning_rows=planning_rows,
            actual_receipts={},
            system_receipts={},
            current_consignments={"130100011": 0},
            state={},
            leadtimes={"130100011": 4},
        )
        self.assertEqual(result["130100011"]["warehouse_debt"], 2650)
        self.assertAlmostEqual(
            result["130100011"]["required_production"],
            57019 + 2650 - 7176,
        )


if __name__ == "__main__":
    unittest.main()
