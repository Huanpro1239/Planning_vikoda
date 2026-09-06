import unittest
from io import BytesIO

from openpyxl import Workbook, load_workbook

from verify_planning_month import verify_workbook
from sync_planning_calendar import build_date_headers


class VerifyPlanningMonthTests(unittest.TestCase):
    def make_workbook(
        self,
        *,
        planning_fc=9370,
        planning_actual=6534,
        planning_book=7376.833333333334,
        stale_after_end=False,
    ):
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

        stock = workbook.create_sheet("Ton_kho")
        stock.append([
            "Mã Sản Phẩm", "Tên Sản Phẩm", "Đơn vị tính", "Tồn thực tế",
            "Tồn Vikoda (Nhà Máy)", "Tồn VKD (Nhà Máy)",
            "Tồn Vikoda (Kho khác)", "Tồn VKD (Kho khác)",
        ])
        stock.append([130100096, "A", "Thùng", 6534, 5679.16666666667, 100.666666666667, 784.5, 812.5])

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
            3, planning_actual, planning_book, planning_fc, 1801.92307692308, 0,
            2836, 4500, 1 / 3, None,
        ] + [4500] + [None] * 29)
        if stale_after_end:
            planning["AW2"] = 123

        output = BytesIO()
        workbook.save(output)
        return output.getvalue()

    def test_passes_when_stock_fc_and_calendar_match(self):
        info = verify_workbook(self.make_workbook())
        self.assertEqual(info["plan_month"], 9)
        self.assertEqual(info["source_column"], 13)
        self.assertEqual(info["stock_checked"], 1)
        self.assertEqual(info["fc_checked"], 1)

    def test_fails_when_j_k_are_stale(self):
        with self.assertRaisesRegex(RuntimeError, "J:K chưa theo"):
            verify_workbook(self.make_workbook(planning_actual=8444, planning_book=13161.8333333333))

    def test_fails_when_l_still_contains_august_fc(self):
        with self.assertRaisesRegex(RuntimeError, "L chưa theo"):
            verify_workbook(self.make_workbook(planning_fc=5343.6))

    def test_fails_when_data_survives_after_last_day(self):
        with self.assertRaisesRegex(RuntimeError, "Còn dữ liệu sau ngày cuối tháng"):
            verify_workbook(self.make_workbook(stale_after_end=True))


    def test_fails_when_middle_calendar_header_is_corrupt(self):
        workbook = load_workbook(BytesIO(self.make_workbook()))
        workbook["Ke_hoach_SX"]["T1"] = "BAD MIDDLE HEADER"
        output = BytesIO()
        workbook.save(output)
        with self.assertRaisesRegex(RuntimeError, "Tiêu đề ngày sai"):
            verify_workbook(output.getvalue())

    def test_fails_when_daily_production_exceeds_P(self):
        workbook = load_workbook(BytesIO(self.make_workbook()))
        workbook["Ke_hoach_SX"]["S2"] = 999999
        output = BytesIO()
        workbook.save(output)
        with self.assertRaisesRegex(RuntimeError, "Tổng SX ngày"):
            verify_workbook(output.getvalue())


if __name__ == "__main__":
    unittest.main()
