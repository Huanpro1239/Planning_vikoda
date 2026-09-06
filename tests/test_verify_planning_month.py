import unittest
from io import BytesIO

from openpyxl import Workbook

from verify_planning_month import verify_workbook
from sync_planning_calendar import build_date_headers


class VerifyPlanningMonthTests(unittest.TestCase):
    def make_workbook(self, *, planning_fc=9370, stale_after_end=False):
        workbook = Workbook()
        fc = workbook.active
        fc.title = "FC"
        fc.append([
            "STT", "Mã SP", "Tên SP", "Đơn vị tính",
            "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4",
            "Tháng 5", "Tháng 6", "Tháng 7", "Tháng 8",
            "Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12", None, "Tháng 9",
        ])
        fc.append([1, 130100096, "A", "Thùng", 0, 0, 0, 0, 0, 0, 0, 5343.6, 9370, 0, 0, 0])

        planning = workbook.create_sheet("Ke_hoach_SX")
        headers = [
            "Mã Sản Phẩm", "Tên Sản Phẩm", "Đơn vị tính", "Số lượng /mẻ",
            "Số lượng/ ca", "Chuyền", "Nhóm sản phẩm", "Phân loại SP",
            "Số ca theo ngày", "Tồn đầu thực tế", "Tồn đầu sổ sách", "FC",
            "Tồn cuối dự kiến", "Nợ kho", "Số lượng cần sản xuất",
            "Số lượng sản xuất theo mẻ/ca", "Số ngày cần sản xuất", "Ngày bắt đầu sản xuất",
        ] + build_date_headers(2026, 9)
        planning.append(headers)
        planning.append([
            130100096, "A", "Thùng", 4500, 4500, "KHS", "Lon", "Không đường",
            3, 100, 100, planning_fc, 0, 0, 0, 0, 0, None,
        ] + [None] * 30)
        if stale_after_end:
            planning["AW2"] = 123

        output = BytesIO()
        workbook.save(output)
        return output.getvalue()

    def test_passes_when_l_and_calendar_match_september(self):
        info = verify_workbook(self.make_workbook())
        self.assertEqual(info["plan_month"], 9)
        self.assertEqual(info["source_column"], 13)
        self.assertEqual(info["checked"], 1)

    def test_fails_when_l_still_contains_august_fc(self):
        with self.assertRaisesRegex(RuntimeError, "L chưa theo"):
            verify_workbook(self.make_workbook(planning_fc=5343.6))

    def test_fails_when_data_survives_after_last_day(self):
        with self.assertRaisesRegex(RuntimeError, "Còn dữ liệu sau ngày cuối tháng"):
            verify_workbook(self.make_workbook(stale_after_end=True))


if __name__ == "__main__":
    unittest.main()
