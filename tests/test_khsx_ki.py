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


def make_mock_khsx_ki_workbook(*, days=None, skus=None, month=9, year=2026, with_merges=True):
    """Tạo workbook mẫu có cả Ke_hoach_SX và KHSX_ki để kiểm thử độc lập."""
    days_in_month = calendar.monthrange(year, month)[1]
    total_days = days if days is not None else days_in_month
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
    ] + [f"{d:02d}/{month:02d}" for d in range(1, total_days + 1)]
    ws_kh.append(headers_kh)

    sku_list = skus or [
        {"code": 130100011, "name": "Đảnh Thạnh chanh", "uom": "Thùng", "p": 86400.0, "daily": [12000.0] * 6 + [12000.0, 2400.0] + [0.0] * max(0, total_days - 8)},
        {"code": 130100096, "name": "Vikoda Alkaline", "uom": "Thùng", "p": 0.0, "daily": [0.0] * total_days},
    ]

    for s in sku_list:
        row = [
            s["code"], s["name"], s["uom"], 100, 500, "KHS", "Nhom", "Không đường",
            1, 0, 0, s["p"], 0, 0, s["p"], s["p"], 1, None,
        ] + [s["daily"][d - 1] if d <= len(s["daily"]) and s["daily"][d - 1] > 0 else None for d in range(1, total_days + 1)]
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

    empty_cols = [None] * (len(std_wks) + 1)
    for idx, s in enumerate(sku_list, start=1):
        ws_ki.append([idx, s["code"], s["name"], s["uom"]] + empty_cols)

    ws_ki.append(["TỔNG CỘNG", None, None, None] + empty_cols)

    if with_merges:
        ws_ki.merge_cells("A1:J1")
        ws_ki.merge_cells("A2:J2")
        ws_ki.merge_cells("A3:J3")
        ws_ki.merge_cells("A4:J4")
        ws_ki.cell(31, 9).value = "TP.KẾ HOẠCH"
        ws_ki.merge_cells("I31:J31")

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

    def test_missing_active_sku_blocked(self):
        """[P1 Defect 1] Chặn khi KHSX_ki bị xóa/thiếu mã SKU so với Ke_hoach_SX."""
        base_bytes = make_mock_khsx_ki_workbook()
        wb = load_workbook(BytesIO(base_bytes))
        # Row 7 is SKU 130100011
        wb["KHSX_ki"].delete_rows(7)
        buf = BytesIO()
        wb.save(buf)
        tampered_bytes = buf.getvalue()

        # Writer must block
        with self.assertRaises(RuntimeError) as ctx_patch:
            ki.patch_khsx_ki_workbook(tampered_bytes, plan_year=2026, plan_month=9)
        self.assertIn("thiếu 1 mã SKU", str(ctx_patch.exception))
        self.assertIn("130100011", str(ctx_patch.exception))

        # Verifier must also block
        with self.assertRaises(RuntimeError) as ctx_verify:
            ki.verify_khsx_ki(tampered_bytes, plan_year=2026, plan_month=9)
        self.assertIn("thiếu 1 mã SKU", str(ctx_verify.exception))
        self.assertIn("130100011", str(ctx_verify.exception))

    def test_duplicate_sku_blocked_in_both_sheets(self):
        """[P1 Defect 1] Chặn khi có mã SKU trùng lặp trong KHSX_ki hoặc Ke_hoach_SX."""
        base_bytes = make_mock_khsx_ki_workbook()

        # Ca 1: Trùng trong KHSX_ki
        wb1 = load_workbook(BytesIO(base_bytes))
        ws_ki = wb1["KHSX_ki"]
        ws_ki.cell(8, 2).value = ws_ki.cell(7, 2).value  # Đặt dòng 8 cùng mã dòng 7 (130100011)
        buf1 = BytesIO()
        wb1.save(buf1)
        tampered_ki = buf1.getvalue()

        with self.assertRaises(RuntimeError) as ctx1_patch:
            ki.patch_khsx_ki_workbook(tampered_ki, plan_year=2026, plan_month=9)
        self.assertIn("Trùng mã SKU 130100011 trong sheet KHSX_ki", str(ctx1_patch.exception))

        with self.assertRaises(RuntimeError) as ctx1_ver:
            ki.verify_khsx_ki(tampered_ki, plan_year=2026, plan_month=9)
        self.assertIn("Trùng mã SKU 130100011 trong sheet KHSX_ki", str(ctx1_ver.exception))

        # Ca 2: Trùng trong Ke_hoach_SX
        wb2 = load_workbook(BytesIO(base_bytes))
        ws_kh = wb2["Ke_hoach_SX"]
        ws_kh.cell(3, 1).value = ws_kh.cell(2, 1).value  # Dòng 3 trùng mã dòng 2
        buf2 = BytesIO()
        wb2.save(buf2)
        tampered_kh = buf2.getvalue()

        with self.assertRaises(RuntimeError) as ctx2_patch:
            ki.patch_khsx_ki_workbook(tampered_kh, plan_year=2026, plan_month=9)
        self.assertIn("Trùng mã SKU 130100011 trong sheet Ke_hoach_SX", str(ctx2_patch.exception))

        with self.assertRaises(RuntimeError) as ctx2_ver:
            ki.verify_khsx_ki(tampered_kh, plan_year=2026, plan_month=9)
        self.assertIn("Trùng mã SKU 130100011 trong sheet Ke_hoach_SX", str(ctx2_ver.exception))

    def test_unknown_sku_blocked(self):
        """[P1 Defect 1] Chặn khi KHSX_ki chứa mã SKU lạ không có trong Ke_hoach_SX."""
        base_bytes = make_mock_khsx_ki_workbook()
        wb = load_workbook(BytesIO(base_bytes))
        ws_ki = wb["KHSX_ki"]
        # Chèn thêm 1 dòng SKU lạ trước dòng TỔNG CỘNG
        ws_ki.insert_rows(9)
        ws_ki.cell(9, 1).value = 3
        ws_ki.cell(9, 2).value = 130199999
        ws_ki.cell(9, 3).value = "Mã SKU lạ không có trong Ke_hoach_SX"
        ws_ki.cell(9, 4).value = "Thùng"
        buf = BytesIO()
        wb.save(buf)
        tampered_bytes = buf.getvalue()

        with self.assertRaises(RuntimeError) as ctx_patch:
            ki.patch_khsx_ki_workbook(tampered_bytes, plan_year=2026, plan_month=9)
        self.assertIn("chứa 1 mã SKU không có trong Ke_hoach_SX", str(ctx_patch.exception))
        self.assertIn("130199999", str(ctx_patch.exception))

        with self.assertRaises(RuntimeError) as ctx_ver:
            ki.verify_khsx_ki(tampered_bytes, plan_year=2026, plan_month=9)
        self.assertIn("chứa 1 mã SKU không có trong Ke_hoach_SX", str(ctx_ver.exception))
        self.assertIn("130199999", str(ctx_ver.exception))

    def test_four_week_month_readonly_verification(self):
        """[P1 Defect 2] Tháng có 4 tuần (Tháng 2/2027) xác minh thành công trên read-only workbook."""
        wb_bytes = make_mock_khsx_ki_workbook(month=2, year=2027)
        patched, report = ki.patch_khsx_ki_workbook(wb_bytes, plan_year=2027, plan_month=2)
        self.assertTrue(report["ok"])
        self.assertEqual(report["num_weeks"], 4)

        # Verifier thuần đọc mở read_only=True không được ném AttributeError: Cell is read only
        verify = ki.verify_khsx_ki(patched, plan_year=2027, plan_month=2)
        self.assertTrue(verify["ok"])
        self.assertEqual(verify["num_weeks"], 4)

        # Kiểm tra nội dung chi tiết trên workbook đã patch
        wb_check = load_workbook(BytesIO(patched), data_only=True)
        ws_ki = wb_check["KHSX_ki"]
        # Cột 9 (Tuần 5) là tuần rỗng: giá trị 0.0
        self.assertEqual(ws_ki.cell(6, 9).value, "Tuần 5\n-")
        self.assertAlmostEqual(float(ws_ki.cell(7, 9).value or 0.0), 0.0)
        self.assertAlmostEqual(float(ws_ki.cell(8, 9).value or 0.0), 0.0)
        # Cột 10 (Tổng cộng) có giá trị 86400
        self.assertAlmostEqual(float(ws_ki.cell(7, 10).value), 86400.0)
        # Cột 11 (K) hoàn toàn rỗng
        self.assertIsNone(ws_ki.cell(6, 11).value)
        self.assertIsNone(ws_ki.cell(7, 11).value)
        wb_check.close()

    def test_six_week_month_layout_and_verification(self):
        """[P2 Defect 3] Tháng có 6 tuần (Tháng 11/2026) mở rộng sang Cột K và phân hoạch đúng ranh giới."""
        wb_bytes = make_mock_khsx_ki_workbook(month=11, year=2026)
        patched, report = ki.patch_khsx_ki_workbook(wb_bytes, plan_year=2026, plan_month=11)
        self.assertTrue(report["ok"])
        self.assertEqual(report["num_weeks"], 6)

        verify = ki.verify_khsx_ki(patched, plan_year=2026, plan_month=11)
        self.assertTrue(verify["ok"])
        self.assertEqual(verify["num_weeks"], 6)
        self.assertAlmostEqual(verify["grand_total"], 86400.0)

        wb_check = load_workbook(BytesIO(patched), data_only=True)
        ws_ki = wb_check["KHSX_ki"]

        # Cột 10 (J) là Tuần 6 (30/11-30/11)
        self.assertEqual(ws_ki.cell(6, 10).value, "Tuần 6\n30/11-30/11")
        # Cột 11 (K) là Cột Tổng cộng
        self.assertEqual(ws_ki.cell(6, 11).value, "Tổng cộng")
        self.assertAlmostEqual(float(ws_ki.cell(7, 11).value), 86400.0)

        # Kiểm tra tiêu đề A1..A4 và dòng chữ ký 31 đã mở rộng sang cột K
        merged_coords = [rng.coord for rng in ws_ki.merged_cells.ranges]
        self.assertIn("A1:K1", merged_coords)
        self.assertIn("A2:K2", merged_coords)
        self.assertIn("A3:K3", merged_coords)
        self.assertIn("A4:K4", merged_coords)
        self.assertIn("I31:K31", merged_coords)
        wb_check.close()

        # Kiểm tra thêm Tháng 3/2026 (31 ngày, 6 tuần)
        mar_bytes = make_mock_khsx_ki_workbook(month=3, year=2026)
        mar_patched, mar_rep = ki.patch_khsx_ki_workbook(mar_bytes, plan_year=2026, plan_month=3)
        self.assertEqual(mar_rep["num_weeks"], 6)
        mar_ver = ki.verify_khsx_ki(mar_patched, plan_year=2026, plan_month=3)
        self.assertTrue(mar_ver["ok"])
        self.assertEqual(mar_ver["num_weeks"], 6)

    def test_switching_between_6_5_4_weeks_cleans_managed_cells(self):
        """[P2 Defect 3] Chuyển đổi giữa tháng 6 tuần, 5 tuần và 4 tuần dọn sạch Cột K và phục hồi merge."""
        # 1. Bắt đầu với tháng 6 tuần (Tháng 11/2026) -> Cột 11 có dữ liệu
        wb_bytes = make_mock_khsx_ki_workbook(month=11, year=2026)
        p_6wk, _ = ki.patch_khsx_ki_workbook(wb_bytes, plan_year=2026, plan_month=11)
        self.assertTrue(ki.verify_khsx_ki(p_6wk, plan_year=2026, plan_month=11)["ok"])

        # 2. Lấy output 6 tuần chạy sang Tháng 9/2026 (5 tuần)
        p_5wk, r_5wk = ki.patch_khsx_ki_workbook(p_6wk, plan_year=2026, plan_month=9)
        self.assertTrue(r_5wk["ok"])
        self.assertEqual(r_5wk["num_weeks"], 5)
        v_5wk = ki.verify_khsx_ki(p_5wk, plan_year=2026, plan_month=9)
        self.assertTrue(v_5wk["ok"])

        wb_5 = load_workbook(BytesIO(p_5wk), data_only=True)
        ws_5 = wb_5["KHSX_ki"]
        self.assertEqual(ws_5.cell(6, 10).value, "Tổng cộng")
        # Cột K phải được xóa sạch
        self.assertIsNone(ws_5.cell(6, 11).value)
        self.assertIsNone(ws_5.cell(7, 11).value)
        # Hợp nhất phục hồi về cột J
        merged_5 = [rng.coord for rng in ws_5.merged_cells.ranges]
        self.assertIn("A1:J1", merged_5)
        self.assertIn("I31:J31", merged_5)
        self.assertNotIn("A1:K1", merged_5)
        self.assertNotIn("I31:K31", merged_5)
        wb_5.close()

        # 3. Lấy output 5 tuần chạy sang Tháng 2/2027 (4 tuần)
        p_4wk, r_4wk = ki.patch_khsx_ki_workbook(p_5wk, plan_year=2027, plan_month=2)
        self.assertTrue(r_4wk["ok"])
        self.assertEqual(r_4wk["num_weeks"], 4)
        v_4wk = ki.verify_khsx_ki(p_4wk, plan_year=2027, plan_month=2)
        self.assertTrue(v_4wk["ok"])

        wb_4 = load_workbook(BytesIO(p_4wk), data_only=True)
        ws_4 = wb_4["KHSX_ki"]
        self.assertEqual(ws_4.cell(6, 9).value, "Tuần 5\n-")
        self.assertEqual(ws_4.cell(6, 10).value, "Tổng cộng")
        self.assertIsNone(ws_4.cell(6, 11).value)
        wb_4.close()

    def test_overlapping_or_corrupted_week_header_rejected(self):
        """[P2 Defect 3 & 4] Tiêu đề tuần bị chồng lấn/sai lệch bị verifier từ chối và được patch sửa chuẩn."""
        base_bytes = make_mock_khsx_ki_workbook()
        wb = load_workbook(BytesIO(base_bytes))
        ws = wb["KHSX_ki"]
        # Giả lập người dùng sửa tay tiêu đề tuần 3 và tuần 4 chồng lấn ngày 21/09
        ws["G6"] = "Tuần 3\n14/09-21/09"
        ws["H6"] = "Tuần 4\n21/09-27/09"
        buf = BytesIO()
        wb.save(buf)
        tampered_bytes = buf.getvalue()

        # Verifier chạy trực tiếp trên file bị sửa phải chặn lại
        with self.assertRaises(RuntimeError) as ctx:
            ki.verify_khsx_ki(tampered_bytes, plan_year=2026, plan_month=9)
        self.assertIn("không khớp dải ngày chuẩn", str(ctx.exception))

        # Writer patch lại phải ghi đè tiêu đề chuẩn và kết quả pass verify
        repaired_bytes, report = ki.patch_khsx_ki_workbook(tampered_bytes, plan_year=2026, plan_month=9)
        self.assertTrue(report["ok"])
        verify = ki.verify_khsx_ki(repaired_bytes, plan_year=2026, plan_month=9)
        self.assertTrue(verify["ok"])

    def test_user_edited_headers_do_not_alter_schedule_rules(self):
        """[P2 Defect 4] Code làm chủ tiêu đề tuần; chỉnh sửa ô header không làm thay đổi kết quả kế hoạch."""
        base_bytes = make_mock_khsx_ki_workbook()
        wb = load_workbook(BytesIO(base_bytes))
        ws = wb["KHSX_ki"]
        # Người dùng sửa tiêu đề E6 và F6
        ws["E6"] = "Tuần 1\n01/09-07/09"
        ws["F6"] = "Tuần 2\n08/09-13/09"
        buf = BytesIO()
        wb.save(buf)
        altered_bytes = buf.getvalue()

        orig_out, _ = ki.patch_khsx_ki_workbook(base_bytes, plan_year=2026, plan_month=9)
        altered_out, _ = ki.patch_khsx_ki_workbook(altered_bytes, plan_year=2026, plan_month=9)

        b1 = load_workbook(BytesIO(orig_out), data_only=True)
        b2 = load_workbook(BytesIO(altered_out), data_only=True)
        # Sản lượng tuần 1 phải bằng nhau (72000), không bị ăn lẹm ngày 7
        self.assertEqual(b1["KHSX_ki"]["E7"].value, b2["KHSX_ki"]["E7"].value)
        self.assertEqual(b1["KHSX_ki"]["F7"].value, b2["KHSX_ki"]["F7"].value)
        self.assertEqual(b2["KHSX_ki"]["E6"].value, "Tuần 1\n01/09-06/09")
        b1.close()
        b2.close()

    def test_calendar_partition_invariant_all_24_months(self):
        """[P2 Defect 3] Xác minh phân hoạch tuần Monday-Sunday trên toàn bộ 24 tháng 2026-2027."""
        for year in (2026, 2027):
            for month in range(1, 13):
                days_in_month = calendar.monthrange(year, month)[1]
                weeks = ki.compute_standard_calendar_weeks(year, month)
                self.assertIn(len(weeks), (4, 5, 6), f"Số tuần ngoài dải 4-6: {year}-{month:02d}")

                all_days = []
                for idx, w in enumerate(weeks, start=1):
                    self.assertEqual(w["week_num"], idx)
                    self.assertGreaterEqual(len(w["days"]), 1)
                    self.assertLessEqual(len(w["days"]), 7, f"Tuần dài hơn 7 ngày tại {year}-{month:02d} tuần {idx}")
                    self.assertEqual(w["start_day"], w["days"][0])
                    self.assertEqual(w["end_day"], w["days"][-1])
                    self.assertEqual(w["days"], list(range(w["start_day"], w["end_day"] + 1)))
                    all_days.extend(w["days"])

                self.assertEqual(all_days, list(range(1, days_in_month + 1)))
                self.assertEqual(len(all_days), len(set(all_days)), "Trùng ngày trong phân hoạch tuần!")


if __name__ == "__main__":
    unittest.main()

