import unittest
from io import BytesIO

from openpyxl import Workbook, load_workbook

from sync_planning_calendar import build_date_headers
from sync_planning_weekly_model import prepare_weekly_schedule_update
from verify_planning_month import verify_workbook


CODE_KHS = 130100011
CODE_PET = 130100022


def build_shared_machine_workbook(*, fc_khs=300, fc_pet=300, quy_cach_pet=2.0):
    """Workbook 2 SKU (KHS + PET 9000) dùng chung một máy vật lý."""
    wb = Workbook()
    fc = wb.active
    fc.title = "FC"
    fc.append([
        "STT", "Mã SP", "Tên", "ĐVT",
        "Tháng 1", "Tháng 2", "Tháng 3", "Tháng 4", "Tháng 5", "Tháng 6",
        "Tháng 7", "Tháng 8", "Tháng 9", "Tháng 10", "Tháng 11", "Tháng 12",
        None, "Tháng 9",
    ])
    fc.append([1, CODE_KHS, "A", "T", 0, 0, 0, 0, 0, 0, 0, 0, fc_khs, 0, 0, 0])
    fc.append([2, CODE_PET, "B", "T", 0, 0, 0, 0, 0, 0, 0, 0, fc_pet, 0, 0, 0])

    stock = wb.create_sheet("Ton_kho")
    stock.append(["Mã Sản Phẩm", "Tên", "ĐVT", "Tồn thực tế", "c5", "c6", "c7", "c8"])
    stock.append([CODE_KHS, "A", "T", 0, 0, 0, 0, 0])
    stock.append([CODE_PET, "B", "T", 0, 0, 0, 0, 0])

    master = wb.create_sheet("Danh_muc")
    master.append([
        "Mã", "Tên", "ĐVT", "SL/mẻ", "SL/ca", "Chuyền", "Nhóm",
        "Phân loại", "Quy cách", "Leadtime", "Debt mode",
    ])
    master.append([CODE_KHS, "A", "T", 100, 100, "KHS", "KHS", "Không đường", 1, 0, "SUBTRACT_BOOK_ON_DEBT"])
    master.append([CODE_PET, "B", "T", 100, 100, "PET 9000", "PET 9000", "Không đường", quy_cach_pet, 0, "SUBTRACT_BOOK_ON_DEBT"])

    planning = wb.create_sheet("Ke_hoach_SX")
    planning.append([
        "Mã Sản Phẩm", "Tên", "ĐVT", "Số lượng /mẻ", "Số lượng/ ca", "Chuyền",
        "Nhóm", "Phân loại SP", "Số ca theo ngày", "Tồn đầu thực tế",
        "Tồn đầu sổ sách", "FC", "Tồn cuối dự kiến", "Nợ kho",
        "O", "P", "Q", "R",
    ] + build_date_headers(2026, 9))
    planning.append([CODE_KHS, "A", "T", 100, 100, "KHS", "KHS", "Không đường", 2, 0, 0, fc_khs, 0, 0, 0, 0, 0, None] + [None] * 30)
    planning.append([CODE_PET, "B", "T", 100, 100, "PET 9000", "PET 9000", "Không đường", 2, 0, 0, fc_pet, 0, 0, 0, 0, 0, None] + [None] * 30)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


class VerifySharedMachineTests(unittest.TestCase):
    def _plan(self):
        wb = build_shared_machine_workbook()
        return prepare_weekly_schedule_update(
            wb, plan_year=2026, plan_month=9, input_revision={"target": {"etag": "x"}}
        )

    def test_valid_two_sku_shared_machine_passes_with_report(self):
        updated, report, _ = self._plan()
        self.assertEqual(report["publish_status"], "ready_for_publish")
        info = verify_workbook(updated, schedule_report=report)
        self.assertEqual(info["schedule_checked"], 2)
        self.assertEqual(info["plan_year"], 2026)

    def test_plan_year_is_derived_from_report(self):
        updated, report, _ = self._plan()
        # Không truyền plan_year: verifier phải lấy 2026 từ report.plan_month.
        info = verify_workbook(updated, schedule_report=report)
        self.assertEqual(info["plan_year"], 2026)

    def test_valid_two_sku_shared_machine_passes_without_report(self):
        updated, _, _ = self._plan()
        info = verify_workbook(updated, plan_year=2026)
        self.assertEqual(info["schedule_checked"], 2)

    def test_shared_machine_over_capacity_is_rejected(self):
        updated, _, _ = self._plan()
        wb = load_workbook(BytesIO(updated))
        ws = wb["Ke_hoach_SX"]
        # Bơm sản lượng máy chung vượt xa năng lực nhưng vẫn <= capacity/ngày cho từng ô.
        for col in range(19, 49):
            ws.cell(row=2, column=col).value = 100000
        out = BytesIO()
        wb.save(out)
        with self.assertRaises(RuntimeError):
            verify_workbook(out.getvalue(), plan_year=2026)


if __name__ == "__main__":
    unittest.main()
