import unittest
from io import BytesIO

from openpyxl import Workbook, load_workbook

import sync_stock


class PlanningFormulaRemovalTests(unittest.TestCase):
    def make_workbook(self):
        workbook = Workbook()
        planning = workbook.active
        planning.title = "Ke hoach SX"
        planning["A1"] = "Nhap tay"
        planning["B1"] = "=1+1"

        stock = workbook.create_sheet("Ton_kho")
        stock["A1"] = "=2+2"

        buffer = BytesIO()
        workbook.save(buffer)
        return buffer.getvalue()

    def test_removes_formula_only_from_ke_hoach_sx(self):
        self.assertTrue(
            hasattr(sync_stock, "remove_sheet_formulas"),
            "sync_stock.remove_sheet_formulas chưa tồn tại",
        )

        updated = sync_stock.remove_sheet_formulas(
            self.make_workbook(),
            "Ke hoach SX",
        )

        workbook = load_workbook(BytesIO(updated), data_only=False)
        self.assertEqual(workbook["Ke hoach SX"]["A1"].value, "Nhap tay")
        self.assertIsNone(workbook["Ke hoach SX"]["B1"].value)
        self.assertEqual(workbook["Ton_kho"]["A1"].value, "=2+2")


if __name__ == "__main__":
    unittest.main()
