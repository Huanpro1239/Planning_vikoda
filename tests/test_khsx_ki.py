"""Tests for KHSX_ki weekly summary aggregation, formatting, and verification."""
import calendar
import datetime
from io import BytesIO
from pathlib import Path
import unittest

from openpyxl import Workbook, load_workbook

import sync_planning_khsx_ki as ki
from sync_planning_weekly_model import compute_planning_inputs_hash

AUDIT_SNAPSHOT = Path("temp_audit/survey_run_34190029619/planning_proposal.xlsx")


def make_mock_khsx_ki_workbook(*, days=30, skus=None, month=9, year=2026):
    """Tạo workbook mẫu có cả Ke_hoach_SX và KHSX_ki để kiểm thử độc lập."""
    wb = Workbook()
    ws_kh = wb.active
    ws_kh.title = "Ke_hoach_SX"

    headers_kh = [
        "Mã Sản Phẩm", "Tên Sản Phẩm", "Đơn vị tính", "Số lượng /mẻ",
        "Số lượng/ ca", "Chuyền", "Nhóm sản phẩm", "Phân loại SP",
        "Số ca theo ngày", "Tồn đầu thực tế", "Tồn đầu sổ sách", "FC",
        "Tồn cuối dự kiến", "Nợ kho", "Số lượng cần sản xuất",
        "Số lượng sản xuất theo mẻ/ca", "Số ngày cần sản xuất",
        "Ngày bắt đầu sản xuất",
    ] + [f"{d:02d}/{month:02d}" for d in range(1, days + 1)]
    ws_kh.append(headers_kh)

    sku_list = skus or [
        {"code": 130100011, "name": "Đảnh Thạnh chanh", "uom": "Thùng", "p": 86400.0, "daily": [12000.0] * 6 + [12000.0, 2400.0] + [0.0] * (days - 8)},
        {"code": 130100096, "name": "Vikoda Alkaline", "uom": "Thùng", "p": 0.0, "daily": [0.0] * days},
    ]

    for s in sku_list:
        row = [
            s["code"], s["name"], s["uom"], 100, 500, "KHS", "Nhom", "Không đường",
            1, 0, 0, s["p"], 0, 0, s["p"], s["p"], 1, None,
        ] + [s["daily"][d - 1] if d <= len(s["daily"]) and s["daily"][d - 1] > 0 else None for d in range(1, days + 1)]
        ws_kh.append(row)

    ws_ki = wb.create_sheet("KHSX_ki")
    ws_ki.append(["CÔNG TY CỔ PHẦN NƯỚC KHOÁNG KHÁNH HÒA"])
    ws_ki.append(["KẾ HOẠCH SẢN XUẤT THÁNG"])
    ws_ki.append([f"Kỳ kế hoạch: Tháng {month}"])
    ws_ki.append([f"Ngày lập: 01/{month:02d}/{year}"])
    ws_ki.append([])

    # Row 6: Tiêu đề
    std_wks = ki.compute_month_weeks(year, month, num_week_cols=5)
    header_ki = ["STT", "Mã SP", "Tên Sản Phẩm", "ĐVT"] + [w["label"] for w in std_wks] + ["Tổng cộng"]
    ws_ki.append(header_ki)

    for idx, s in enumerate(sku_list, start=1):
        ws_ki.append([idx, s["code"], s["name"], s["uom"], None, None, None, None, None, None])

    ws_ki.append(["TỔNG CỘNG", None, None, None, None, None, None, None, None, None])

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


class KHSXKiTests(unittest.TestCase):

    def test_has_khsx_ki_sheet(self):
        wb = Workbook()
        buf = BytesIO()
        wb.save(buf)
        self.assertFalse(ki.has_khsx_ki_sheet(buf.getvalue()))

        wb.create_sheet("KHSX_ki")
        buf = BytesIO()
        wb.save(buf)
        self.assertTrue(ki.has_khsx_ki_sheet(buf.getvalue()))

    def test_compute_month_weeks_coverage_all_month_lengths(self):
        """Kiểm tra ranh giới tuần cho các độ dài tháng 28, 29, 30, 31 ngày."""
        test_cases = [
            (2026, 2, 28),   # Tháng 2 năm thường (28 ngày, CN đầu tiên = ngày 1)
            (2028, 2, 29),   # Tháng 2 năm nhuận (29 ngày, CN đầu tiên = ngày 6)
            (2027, 2, 28),   # Tháng 2 bắt đầu vào Thứ Hai (đủ 4 tuần x 7 ngày)
            (2026, 9, 30),   # Tháng 9 (30 ngày)
            (2026, 11, 30),  # Tháng 11 (30 ngày, CN đầu tiên = ngày 1)
            (2026, 1, 31),   # Tháng 1 (31 ngày)
            (2026, 10, 31),  # Tháng 10 (31 ngày)
        ]
        for year, month, expected_days in test_cases:
            weeks = ki.compute_month_weeks(year, month, num_week_cols=5)
            covered = []
            for w in weeks:
                covered.extend(w["days"])
                if w["days"]:
                    self.assertEqual(w["days"], list(range(w["start_day"], w["end_day"] + 1)))
            self.assertEqual(covered, list(range(1, expected_days + 1)), f"Mismatch for {year}-{month:02d}")

    def test_mock_workbook_patch_and_verify(self):
        """Kiểm thử tính toán tuần trên mock workbook."""
        wb_bytes = make_mock_khsx_ki_workbook(days=30, month=9, year=2026)
        patched_bytes, report = ki.patch_khsx_ki_workbook(wb_bytes, plan_year=2026, plan_month=9)
        self.assertTrue(report["ok"])
        self.assertEqual(report["checked_skus"], 2)
        self.assertAlmostEqual(report["grand_total"], 86400.0, places=4)

        verify = ki.verify_khsx_ki(patched_bytes, plan_year=2026, plan_month=9)
        self.assertTrue(verify["ok"])
        self.assertEqual(verify["checked_skus"], 2)

    def test_sku_order_independence(self):
        """Kiểm tra thứ tự SKU khác nhau giữa Ke_hoach_SX và KHSX_ki vẫn khớp chính xác."""
        skus = [
            {"code": 130100001, "name": "SP 1", "uom": "Thùng", "p": 5000.0, "daily": [5000.0] + [0.0] * 29},
            {"code": 130100002, "name": "SP 2", "uom": "Thùng", "p": 10000.0, "daily": [0.0] * 7 + [10000.0] + [0.0] * 22},
        ]
        # Tạo workbook với KHSX_ki có thứ tự đảo ngược
        wb = Workbook()
        ws_kh = wb.active
        ws_kh.title = "Ke_hoach_SX"
        headers_kh = ["Mã SP", "Tên", "ĐVT", "", "", "", "", "", "", "", "", "", "", "", "Nhu cầu", "P", "Q", "Start"] + [f"{d:02d}/09" for d in range(1, 31)]
        ws_kh.append(headers_kh)
        # Ke_hoach_SX: SKU 1 trước, SKU 2 sau
        ws_kh.append([130100001, "SP 1", "Thùng", 0,0,0,0,0,0,0,0,0,0,0,5000, 5000.0, 1, None] + [5000.0] + [None] * 29)
        ws_kh.append([130100002, "SP 2", "Thùng", 0,0,0,0,0,0,0,0,0,0,0,10000, 10000.0, 1, None] + [None] * 7 + [10000.0] + [None] * 22)

        ws_ki = wb.create_sheet("KHSX_ki")
        ws_ki.append(["Title"])
        ws_ki.append(["Subtitle"])
        ws_ki.append(["Kỳ kế hoạch: Tháng 9"])
        ws_ki.append(["Ngày lập: 01/09/2026"])
        ws_ki.append([])
        std_wks = ki.compute_month_weeks(2026, 9)
        ws_ki.append(["STT", "Mã SP", "Tên", "ĐVT"] + [w["label"] for w in std_wks] + ["Tổng cộng"])
        # KHSX_ki: SKU 2 trước, SKU 1 sau
        ws_ki.append([1, 130100002, "SP 2", "Thùng", None, None, None, None, None, None])
        ws_ki.append([2, 130100001, "SP 1", "Thùng", None, None, None, None, None, None])
        ws_ki.append(["TỔNG CỘNG", None, None, None, None, None, None, None, None, None])

        buf = BytesIO()
        wb.save(buf)

        patched_bytes, report = ki.patch_khsx_ki_workbook(buf.getvalue(), plan_year=2026, plan_month=9)
        verify = ki.verify_khsx_ki(patched_bytes, plan_year=2026, plan_month=9)
        self.assertTrue(verify["ok"])

        # Kiểm tra nội dung trực tiếp
        wb_check = load_workbook(BytesIO(patched_bytes), data_only=True)
        ws_check = wb_check["KHSX_ki"]
        # Row 7 is SKU 2 (10000 in Week 2, total 10000)
        self.assertEqual(ws_check.cell(7, 2).value, 130100002)
        self.assertAlmostEqual(float(ws_check.cell(7, 6).value), 10000.0)  # Week 2 is col 6
        self.assertAlmostEqual(float(ws_check.cell(7, 10).value), 10000.0) # Total is col 10

        # Row 8 is SKU 1 (5000 in Week 1, total 5000)
        self.assertEqual(ws_check.cell(8, 2).value, 130100001)
        self.assertAlmostEqual(float(ws_check.cell(8, 5).value), 5000.0)   # Week 1 is col 5
        self.assertAlmostEqual(float(ws_check.cell(8, 10).value), 5000.0)  # Total is col 10

        # Row 9 is TỔNG CỘNG
        self.assertAlmostEqual(float(ws_check.cell(9, 5).value), 5000.0)
        self.assertAlmostEqual(float(ws_check.cell(9, 6).value), 10000.0)
        self.assertAlmostEqual(float(ws_check.cell(9, 10).value), 15000.0)

    def test_real_sharepoint_proposal_snapshot_parity(self):
        """Kiểm tra đối chiếu toàn diện 100% SKU trên file proposal SharePoint thật."""
        if not AUDIT_SNAPSHOT.exists():
            self.skipTest(f"Không tìm thấy file snapshot {AUDIT_SNAPSHOT}")

        wb_bytes = AUDIT_SNAPSHOT.read_bytes()
        patched_bytes, report = ki.patch_khsx_ki_workbook(wb_bytes, plan_year=2026, plan_month=9)
        self.assertTrue(report["ok"])
        self.assertEqual(report["checked_skus"], 20)
        self.assertAlmostEqual(report["grand_total"], 673483.3333333333, places=4)

        verify = ki.verify_khsx_ki(patched_bytes, plan_year=2026, plan_month=9)
        self.assertTrue(verify["ok"])
        self.assertEqual(verify["checked_skus"], 20)
        self.assertAlmostEqual(verify["grand_total"], 673483.3333333333, places=4)

        # Kiểm chứng riêng SKU 130100011
        wb_check = load_workbook(BytesIO(patched_bytes), data_only=True)
        ws_ki = wb_check["KHSX_ki"]
        row_11 = None
        for r in range(7, 27):
            if ws_ki.cell(r, 2).value == 130100011:
                row_11 = r
                break
        self.assertIsNotNone(row_11)
        # Tuần 1: 72000, Tuần 2: 14400, Tuần 3-5: 0, Tổng: 86400
        self.assertAlmostEqual(float(ws_ki.cell(row_11, 5).value), 72000.0)
        self.assertAlmostEqual(float(ws_ki.cell(row_11, 6).value), 14400.0)
        self.assertAlmostEqual(float(ws_ki.cell(row_11, 7).value), 0.0)
        self.assertAlmostEqual(float(ws_ki.cell(row_11, 8).value), 0.0)
        self.assertAlmostEqual(float(ws_ki.cell(row_11, 9).value), 0.0)
        self.assertAlmostEqual(float(ws_ki.cell(row_11, 10).value), 86400.0)

    def test_input_hash_excludes_khsx_ki(self):
        """Khẳng định sửa đổi trên KHSX_ki không làm thay đổi compute_planning_inputs_hash."""
        wb_bytes = make_mock_khsx_ki_workbook(days=30, month=9, year=2026)
        h1 = compute_planning_inputs_hash(wb_bytes)

        patched_bytes, _ = ki.patch_khsx_ki_workbook(wb_bytes, plan_year=2026, plan_month=9)
        h2 = compute_planning_inputs_hash(patched_bytes)

        self.assertEqual(h1, h2, "compute_planning_inputs_hash bị thay đổi bởi KHSX_ki!")

    def test_two_round_pipeline_idempotence(self):
        """Kiểm tra tính lũy đẳng 2 vòng (idempotence)."""
        wb_bytes = make_mock_khsx_ki_workbook(days=30, month=9, year=2026)
        p1, r1 = ki.patch_khsx_ki_workbook(wb_bytes, plan_year=2026, plan_month=9)
        p2, r2 = ki.patch_khsx_ki_workbook(p1, plan_year=2026, plan_month=9)

        self.assertEqual(r1["grand_total"], r2["grand_total"])
        self.assertEqual(r1["checked_skus"], r2["checked_skus"])

        v1 = ki.verify_khsx_ki(p1, plan_year=2026, plan_month=9)
        v2 = ki.verify_khsx_ki(p2, plan_year=2026, plan_month=9)
        self.assertTrue(v1["ok"])
        self.assertTrue(v2["ok"])

    def test_pipeline_integration_with_khsx_ki(self):
        """Kiểm tra toàn bộ luồng prepare_pipeline_output khi workbook có sheet KHSX_ki."""
        from planning_pipeline import prepare_pipeline_output
        from tests.test_run_offline import (
            _actual_source,
            _single_value_source,
            _target_workbook,
            CODE,
            CODE_VKD,
        )

        source_bytes = {
            "actual_stock": _actual_source(),
            "factory_vikoda": _single_value_source(CODE, value_col_index=11, value=0, receipt=50),
            "factory_vkd": _single_value_source(CODE_VKD, value_col_index=11, value=0),
            "accounting_vikoda": _single_value_source(CODE, value_col_index=12, value=0),
            "accounting_vkd": _single_value_source(CODE_VKD, value_col_index=12, value=0),
        }

        wb = load_workbook(BytesIO(_target_workbook()))
        wb["FC"]["M2"] = 2600

        # Thêm sheet KHSX_ki vào target workbook
        ws_ki = wb.create_sheet("KHSX_ki")
        ws_ki.append(["CÔNG TY CỔ PHẦN NƯỚC KHOÁNG KHÁNH HÒA"])
        ws_ki.append(["KẾ HOẠCH SẢN XUẤT THÁNG"])
        ws_ki.append(["Kỳ kế hoạch: Tháng 9"])
        ws_ki.append(["Ngày lập: 01/09/2026"])
        ws_ki.append([])
        std_wks = ki.compute_month_weeks(2026, 9)
        ws_ki.append(["STT", "Mã SP", "Tên", "ĐVT"] + [w["label"] for w in std_wks] + ["Tổng cộng"])
        ws_ki.append([1, CODE, "SP", "Thùng", None, None, None, None, None, None])
        ws_ki.append(["TỔNG CỘNG", None, None, None, None, None, None, None, None, None])

        buf = BytesIO()
        wb.save(buf)
        target_bytes = buf.getvalue()

        # Round 1
        out1, report1, state1 = prepare_pipeline_output(
            target_bytes,
            source_bytes,
            runtime_state={},
            input_revision={"target": {"etag": "test-etag"}},
        )

        pipeline_info1 = report1["pipeline"]
        self.assertIn("khsx_ki", pipeline_info1["steps"])
        self.assertEqual(pipeline_info1["engine_version"], "ke_hoach_sx_tuan_v3_khsx_ki_20260908")
        self.assertTrue(pipeline_info1["verify"]["khsx_ki"]["ok"])

        # Round 2: Idempotence test
        out2, report2, state2 = prepare_pipeline_output(
            out1,
            source_bytes,
            runtime_state=state1,
            input_revision={"target": {"etag": "test-etag"}},
        )
        self.assertTrue(report2["pipeline"]["verify"]["khsx_ki"]["ok"])


if __name__ == "__main__":
    unittest.main()

