"""Tests for scripts/survey_nvl_sharepoint.py."""

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock

import openpyxl

from scripts.survey_nvl_sharepoint import run_survey


class SurveyNVLTests(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.workdir = Path(self.tmp_dir.name)

        # 1. Tạo file nguồn mock
        self.src_path = self.workdir / "mock_source.xlsm"
        wb_src = openpyxl.Workbook()
        ws_src = wb_src.active
        ws_src.title = "Sheet1"
        ws_src["A5"] = "BÁO CÁO NHẬP XUẤT TỒN KẾ TOÁN"
        ws_src["A7"] = "Từ ngày 01-08-2026 đến ngày 31-08-2026"

        # Headers tại dòng 9, 10
        ws_src.cell(9, 2, "Mã vật tư")
        ws_src.cell(9, 6, "Đvt")
        ws_src.cell(9, 13, "Tồn cuối kỳ")
        ws_src.cell(10, 13, "Số lượng")

        # Dữ liệu từ dòng 11
        # Row 11: 330100005, KG, 159569.279
        ws_src.cell(11, 2, "330100005")
        ws_src.cell(11, 6, "KG")
        ws_src.cell(11, 13, 159569.279)

        # Row 12: 330100010, KG, 1619
        ws_src.cell(12, 2, "330100010")
        ws_src.cell(12, 6, "KG")
        ws_src.cell(12, 13, 1619.0)

        # Row 13: 330200010, KG, 0
        ws_src.cell(13, 2, "330200010")
        ws_src.cell(13, 6, "KG")
        ws_src.cell(13, 13, 0.0)

        # Row 14: 430100002, CAI, 500
        ws_src.cell(14, 2, "430100002")
        ws_src.cell(14, 6, "CAI")
        ws_src.cell(14, 13, 500.0)

        # Row 15: 430100118, BIH, 743 (Divergent unit)
        ws_src.cell(15, 2, "430100118")
        ws_src.cell(15, 6, "BIH")
        ws_src.cell(15, 13, 743.0)

        # Row 16: 430200173, CAI, 473992 (Divergent unit nhưng có xác nhận người dùng)
        ws_src.cell(16, 2, "430200173")
        ws_src.cell(16, 6, "CAI")
        ws_src.cell(16, 13, 473992.0)

        wb_src.save(self.src_path)
        wb_src.close()

        # 2. Tạo file đích mock
        self.tgt_path = self.workdir / "mock_target.xlsx"
        wb_tgt = openpyxl.Workbook()
        ws_tgt = wb_tgt.active
        ws_tgt.title = "Ton_NVL"

        # Headers tại dòng 1
        ws_tgt.cell(1, 1, "Mã NVL")
        ws_tgt.cell(1, 2, "Tên NVL")
        ws_tgt.cell(1, 3, "ĐVT")
        ws_tgt.cell(1, 4, "Tồn Cuối")

        # Row 2: 330100005, Đường RE, Kg (khác case với KG), None
        ws_tgt.cell(2, 1, "330100005")
        ws_tgt.cell(2, 2, "Đường tinh luyện RE")
        ws_tgt.cell(2, 3, "Kg")
        ws_tgt.cell(2, 4, None)

        # Row 3: 330100010, Nitơ lỏng, None (thiếu ĐVT ở đích), None
        ws_tgt.cell(3, 1, "330100010")
        ws_tgt.cell(3, 2, "Nitơ lỏng")
        ws_tgt.cell(3, 3, None)
        ws_tgt.cell(3, 4, None)

        # Row 4: 330200010, Taurine, KG (khớp), None
        ws_tgt.cell(4, 1, "330200010")
        ws_tgt.cell(4, 2, "Taurine")
        ws_tgt.cell(4, 3, "KG")
        ws_tgt.cell(4, 4, None)

        # Row 5: 430100002, Phôi PET, Cái (khác dấu với CAI), None
        ws_tgt.cell(5, 1, "430100002")
        ws_tgt.cell(5, 2, "Phôi PET")
        ws_tgt.cell(5, 3, "Cái")
        ws_tgt.cell(5, 4, None)

        # Row 6: 330500109, Premix, Cái, None (Thiếu ở nguồn)
        ws_tgt.cell(6, 1, "330500109")
        ws_tgt.cell(6, 2, "Premix PR0136")
        ws_tgt.cell(6, 3, "Cái")
        ws_tgt.cell(6, 4, None)

        # Row 7: 430100118, Bình 5 gallon, Cái (khác đơn vị BIH vs Cái), None
        ws_tgt.cell(7, 1, "430100118")
        ws_tgt.cell(7, 2, "Bình 5 gallon")
        ws_tgt.cell(7, 3, "Cái")
        ws_tgt.cell(7, 4, None)

        # Row 8: 430200173, Nhãn thân PVC, Kg (khác đơn vị CAI vs Kg), None
        ws_tgt.cell(8, 1, "430200173")
        ws_tgt.cell(8, 2, "Nhãn thân PVC")
        ws_tgt.cell(8, 3, "Kg")
        ws_tgt.cell(8, 4, None)

        wb_tgt.save(self.tgt_path)
        wb_tgt.close()

        # 3. Tạo file config mock
        self.cfg_path = self.workdir / "test_config.json"
        cfg_data = {
            "source": {
                "name": "XNT_ketoan_Vikoda.xlsm",
                "sharepoint_path": "Tinh san xuat Mua hang 2027/Ton He thong/Ton Ke Toan/XNT_ketoan_Vikoda.xlsm",
                "sourcedoc": "C0372C81-A402-4A11-B9E0-847768CC9CFB",
                "sheet_name": "Sheet1",
                "code_column": 2,
                "code_column_letter": "B",
                "value_column": 13,
                "value_column_letter": "M",
                "start_row": 11,
            },
            "target": {
                "name": "Kế hoạch mua hàng.xlsx",
                "sharepoint_path": "Tinh san xuat Mua hang 2027/Kế hoạch mua hàng.xlsx",
                "sourcedoc": "89D1BA7B-006E-4527-B879-ABF120309214",
                "sheet_name": "Ton_NVL",
                "code_column": 1,
                "code_column_letter": "A",
                "value_column": 4,
                "value_column_letter": "D",
                "start_row": 2,
            },
            "policies": {
                "reject_duplicates": True,
                "forbid_formulas": True,
                "forbid_merged_cells": True,
                "forbid_sheet_protection": True,
                "preserve_unmatched": True,
            },
        }
        self.cfg_path.write_text(json.dumps(cfg_data, indent=2), encoding="utf-8")
        self.out_dir = self.workdir / "survey_output"

    def tearDown(self):
        self.tmp_dir.cleanup()

    def test_survey_offline_snapshots_success(self):
        """Khảo sát offline bằng snapshot thành công, tạo đầy đủ audit_summary với kỳ và ĐVT."""
        res = run_survey(
            config_path=str(self.cfg_path),
            source_file=str(self.src_path),
            target_file=str(self.tgt_path),
            out_dir_path=str(self.out_dir),
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["reporting_period"], "Từ ngày 01-08-2026 đến ngày 31-08-2026")
        self.assertIn("Từ ngày 01-08-2026 đến ngày 31-08-2026", res["reporting_period_note"])
        self.assertIn("xác nhận và CHỐT chính thức cho kỳ này", res["reporting_period_note"])

        # Kiểm tra metrics
        metrics = res["metrics"]
        self.assertEqual(metrics["matched_count"], 6)
        self.assertEqual(metrics["changed_count"], 6)
        self.assertEqual(metrics["missing_in_source_count"], 1)
        self.assertEqual(metrics["duplicate_codes_source"], 0)
        self.assertEqual(metrics["duplicate_codes_target"], 0)

        # Kiểm tra đơn vị tính (unit reconciliation)
        unit_rec = res["unit_reconciliation"]
        self.assertEqual(unit_rec["total_matched"], 6)
        self.assertEqual(unit_rec["exact_matches_count"], 2)       # 330100005 (KG vs Kg), 330200010 (KG vs KG)
        self.assertEqual(unit_rec["alias_matches_count"], 1)       # 430100002 (CAI vs Cái)
        self.assertEqual(unit_rec["missing_target_units_count"], 1) # 330100010 (KG vs '')
        self.assertEqual(unit_rec["divergent_units_count"], 2)     # 430100118 (BIH vs Cái), 430200173 (CAI vs Kg)
        self.assertEqual(unit_rec["unit_mismatches_count"], 3)     # missing + divergent
        self.assertIn("không tự ý quy đổi", unit_rec["policy_note"])
        self.assertIn("473,992 Cái cho mã 430200173", res["reporting_period_note"])

        # Kiểm tra danh sách thiếu ở nguồn
        missing_items = res["missing_in_source_items"]
        self.assertEqual(len(missing_items), 1)
        self.assertEqual(missing_items[0]["code"], "330500109")
        self.assertEqual(missing_items[0]["action"], "PRESERVE")

        # Kiểm tra file sinh ra trên đĩa
        audit_file = self.out_dir / "audit_summary.json"
        self.assertTrue(audit_file.exists())
        audit_disk = json.loads(audit_file.read_text(encoding="utf-8"))
        self.assertEqual(audit_disk["status"], "success")
        self.assertTrue((self.out_dir / "nvl_stock_proposal.xlsx").exists())
        self.assertTrue((self.out_dir / "nvl_stock_report.json").exists())

    def test_survey_mock_graph_success(self):
        """Khảo sát trực tuyến qua mock GraphClient: xác minh identity, kỳ báo cáo và xuất audit."""
        mock_graph = MagicMock()
        mock_graph.get_site_id.return_value = "site-test"
        mock_graph.get_default_drive_id.return_value = "drive-test"

        mock_target_item = {
            "id": "item-target-123",
            "name": "Kế hoạch mua hàng.xlsx",
            "eTag": '"{89D1BA7B-006E-4527-B879-ABF120309214},18"',
        }
        mock_source_item = {
            "id": "item-source-456",
            "name": "XNT_ketoan_Vikoda.xlsm",
            "eTag": '"{C0372C81-A402-4A11-B9E0-847768CC9CFB},5"',
        }

        def fake_get_item_by_path(drive_id, path):
            if "XNT" in path:
                return mock_source_item
            return mock_target_item

        mock_graph.get_item_by_path.side_effect = fake_get_item_by_path

        def fake_download(drive_id, item_id):
            if item_id == "item-source-456":
                return self.src_path.read_bytes()
            return self.tgt_path.read_bytes()

        mock_graph.download_file.side_effect = fake_download

        res = run_survey(
            config_path=str(self.cfg_path),
            out_dir_path=str(self.out_dir),
            graph=mock_graph,
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["source_item_id"], "item-source-456")
        self.assertEqual(res["target_item_id"], "item-target-123")
        self.assertEqual(res["reporting_period"], "Từ ngày 01-08-2026 đến ngày 31-08-2026")
        self.assertTrue((self.out_dir / "audit_summary.json").exists())

    def test_survey_target_identity_mismatch_raises_error_and_emits_failed_audit(self):
        """Khi target identity (sourcedoc trong eTag) không khớp, survey báo lỗi RuntimeError và ghi status=failed."""
        mock_graph = MagicMock()
        mock_graph.get_site_id.return_value = "site-test"
        mock_graph.get_default_drive_id.return_value = "drive-test"

        # eTag đích không chứa GUID '89D1BA7B-006E-4527-B879-ABF120309214'
        mock_target_item = {
            "id": "item-target-wrong",
            "name": "Kế hoạch mua hàng.xlsx",
            "eTag": '"{11111111-2222-3333-4444-555555555555},1"',
        }
        mock_source_item = {
            "id": "item-source-456",
            "name": "XNT_ketoan_Vikoda.xlsm",
            "eTag": '"{C0372C81-A402-4A11-B9E0-847768CC9CFB},5"',
        }

        mock_graph.get_item_by_path.side_effect = lambda drive_id, path: mock_source_item if "XNT" in path else mock_target_item

        with self.assertRaises(RuntimeError) as ctx:
            run_survey(
                config_path=str(self.cfg_path),
                out_dir_path=str(self.out_dir),
                graph=mock_graph,
            )

        self.assertIn("Target sourcedoc", str(ctx.exception))
        self.assertIn("không khớp eTag", str(ctx.exception))

        err_audit = self.out_dir / "audit_summary.json"
        self.assertTrue(err_audit.exists())
        data = json.loads(err_audit.read_text(encoding="utf-8"))
        self.assertEqual(data["status"], "failed")
        self.assertEqual(data["phase"], "fetch_target")
        self.assertIn("Target sourcedoc", data["error_message"])

    def test_survey_source_identity_mismatch_raises_error_and_emits_failed_audit(self):
        """Khi source identity (sourcedoc trong eTag) không khớp, survey báo lỗi RuntimeError và ghi status=failed."""
        mock_graph = MagicMock()
        mock_graph.get_site_id.return_value = "site-test"
        mock_graph.get_default_drive_id.return_value = "drive-test"

        mock_target_item = {
            "id": "item-target-123",
            "name": "Kế hoạch mua hàng.xlsx",
            "eTag": '"{89D1BA7B-006E-4527-B879-ABF120309214},18"',
        }
        # eTag nguồn không chứa GUID 'C0372C81-A402-4A11-B9E0-847768CC9CFB'
        mock_source_item = {
            "id": "item-source-wrong",
            "name": "XNT_ketoan_Vikoda.xlsm",
            "eTag": '"{99999999-8888-7777-6666-555555555555},1"',
        }

        mock_graph.get_item_by_path.side_effect = lambda drive_id, path: mock_source_item if "XNT" in path else mock_target_item
        mock_graph.download_file.return_value = self.tgt_path.read_bytes()

        with self.assertRaises(RuntimeError) as ctx:
            run_survey(
                config_path=str(self.cfg_path),
                out_dir_path=str(self.out_dir),
                graph=mock_graph,
            )

        self.assertIn("Source sourcedoc", str(ctx.exception))
        self.assertIn("không khớp eTag", str(ctx.exception))

        err_audit = self.out_dir / "audit_summary.json"
        self.assertTrue(err_audit.exists())
        data = json.loads(err_audit.read_text(encoding="utf-8"))
        self.assertEqual(data["status"], "failed")
        self.assertEqual(data["phase"], "fetch_source")
        self.assertIn("Source sourcedoc", data["error_message"])

    def test_survey_error_handling_emits_error_report_and_raises(self):
        """Khi gặp lỗi (ví dụ file đích không tồn tại), survey ghi error report và ném ngoại lệ."""
        bad_cfg = self.workdir / "bad_config.json"
        cfg_data = json.loads(self.cfg_path.read_text(encoding="utf-8"))
        cfg_data["target"]["sheet_name"] = "NonExistentSheet"
        bad_cfg.write_text(json.dumps(cfg_data), encoding="utf-8")

        with self.assertRaises(ValueError):
            run_survey(
                config_path=str(bad_cfg),
                source_file=str(self.src_path),
                target_file=str(self.tgt_path),
                out_dir_path=str(self.out_dir),
            )

        err_audit = self.out_dir / "audit_summary.json"
        self.assertTrue(err_audit.exists())
        data = json.loads(err_audit.read_text(encoding="utf-8"))
        self.assertEqual(data["status"], "failed")
        self.assertEqual(data["phase"], "survey_target")
        self.assertIn("NonExistentSheet", data["error_message"])

    def test_survey_cli_exit_codes(self):
        """Kiểm tra exit code qua CLI: 0 khi thành công, 1 khi thất bại."""
        # Ca thành công
        cmd_ok = [
            sys.executable,
            "-X",
            "utf8",
            "scripts/survey_nvl_sharepoint.py",
            "--config",
            str(self.cfg_path),
            "--source-file",
            str(self.src_path),
            "--target-file",
            str(self.tgt_path),
            "--out-dir",
            str(self.out_dir),
        ]
        p_ok = subprocess.run(cmd_ok, capture_output=True, text=True)
        self.assertEqual(p_ok.returncode, 0)

        # Ca thất bại (file nguồn không tồn tại)
        cmd_fail = [
            sys.executable,
            "-X",
            "utf8",
            "scripts/survey_nvl_sharepoint.py",
            "--config",
            str(self.cfg_path),
            "--source-file",
            "nonexistent_source.xlsm",
            "--target-file",
            str(self.tgt_path),
            "--out-dir",
            str(self.out_dir),
        ]
        p_fail = subprocess.run(cmd_fail, capture_output=True, text=True)
        self.assertEqual(p_fail.returncode, 1)

    def test_survey_dynamic_snapshot_values_and_approved_period(self):
        """Khi số tồn snapshot thay đổi (ví dụ mã 430200173 có 555,666), audit hiển thị số động,
        không rơi về số hardcode 473,992, và áp dụng xác nhận khi đúng kỳ được duyệt.
        """
        # Thêm mã 430200173 vào source với tồn 555666.0
        wb_src = openpyxl.load_workbook(self.src_path)
        ws_src = wb_src["Sheet1"]
        ws_src.cell(16, 2, "430200173")
        ws_src.cell(16, 6, "CAI")
        ws_src.cell(16, 13, 555666.0)
        wb_src.save(self.src_path)
        wb_src.close()

        # Thêm mã 430200173 vào target với ĐVT Kg
        wb_tgt = openpyxl.load_workbook(self.tgt_path)
        ws_tgt = wb_tgt["Ton_NVL"]
        ws_tgt.cell(8, 1, "430200173")
        ws_tgt.cell(8, 2, "Nhãn thân PVC")
        ws_tgt.cell(8, 3, "Kg")
        ws_tgt.cell(8, 4, None)
        wb_tgt.save(self.tgt_path)
        wb_tgt.close()

        res = run_survey(
            config_path=str(self.cfg_path),
            source_file=str(self.src_path),
            target_file=str(self.tgt_path),
            out_dir_path=str(self.out_dir),
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["reporting_period"], "Từ ngày 01-08-2026 đến ngày 31-08-2026")
        # Kiểm tra ghi chú tổng quan kỳ báo cáo hiển thị số tồn động 555,666
        self.assertIn("555,666 Cái cho mã 430200173", res["reporting_period_note"])
        self.assertNotIn("473,992", res["reporting_period_note"])

        # Kiểm tra ghi chú cụ thể của mã 430200173 trong divergent_units
        divergent = res["unit_reconciliation"]["divergent_units"]
        item_430 = next((item for item in divergent if item["code"] == "430200173"), None)
        self.assertIsNotNone(item_430)
        self.assertIn("555,666", item_430["note"])
        self.assertIn("Người dùng đã CHỐT cho kỳ 'Từ ngày 01-08-2026 đến ngày 31-08-2026'", item_430["note"])
        self.assertNotIn("473,992", item_430["note"])

    def test_survey_period_changed_does_not_apply_stale_approvals(self):
        """Khi kỳ báo cáo nguồn thay đổi (ví dụ sang kỳ tháng 9/2026), audit ghi rõ kỳ mới,
        hiển thị số tồn động mới, và cảnh báo rằng xác nhận cũ không tự động áp dụng.
        """
        # Thay đổi kỳ báo cáo trong file nguồn sang tháng 9/2026 và thêm mã 430200173 với tồn 888999.0
        wb_src = openpyxl.load_workbook(self.src_path)
        ws_src = wb_src["Sheet1"]
        ws_src["A7"] = "Từ ngày 01-09-2026 đến ngày 30-09-2026"
        ws_src.cell(16, 2, "430200173")
        ws_src.cell(16, 6, "CAI")
        ws_src.cell(16, 13, 888999.0)
        wb_src.save(self.src_path)
        wb_src.close()

        # Thêm mã 430200173 vào target với ĐVT Kg
        wb_tgt = openpyxl.load_workbook(self.tgt_path)
        ws_tgt = wb_tgt["Ton_NVL"]
        ws_tgt.cell(8, 1, "430200173")
        ws_tgt.cell(8, 2, "Nhãn thân PVC")
        ws_tgt.cell(8, 3, "Kg")
        ws_tgt.cell(8, 4, None)
        wb_tgt.save(self.tgt_path)
        wb_tgt.close()

        res = run_survey(
            config_path=str(self.cfg_path),
            source_file=str(self.src_path),
            target_file=str(self.tgt_path),
            out_dir_path=str(self.out_dir),
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["reporting_period"], "Từ ngày 01-09-2026 đến ngày 30-09-2026")

        # Ghi chú tổng quan phải cảnh báo xác nhận cũ không áp dụng
        self.assertIn("CẢNH BÁO: Xác nhận của người dùng trước đây gắn với kỳ 'Từ ngày 01-08-2026 đến ngày 31-08-2026'", res["reporting_period_note"])
        self.assertIn("không tự động áp dụng cho kỳ hiện tại 'Từ ngày 01-09-2026 đến ngày 30-09-2026'", res["reporting_period_note"])
        self.assertNotIn("473,992", res["reporting_period_note"])

        # Kiểm tra ghi chú cụ thể của mã 430200173: không được ghi 'Người dùng đã CHỐT cho kỳ này'
        divergent = res["unit_reconciliation"]["divergent_units"]
        item_430 = next((item for item in divergent if item["code"] == "430200173"), None)
        self.assertIsNotNone(item_430)
        self.assertIn("888,999", item_430["note"])
        self.assertIn("Xác nhận trước đó gắn với kỳ 'Từ ngày 01-08-2026 đến ngày 31-08-2026'", item_430["note"])
        self.assertIn("không tự áp dụng cho kỳ hiện tại 'Từ ngày 01-09-2026 đến ngày 30-09-2026'", item_430["note"])
        self.assertNotIn("473,992", item_430["note"])

        # Kiểm tra mã thiếu ở nguồn: ghi chú bảo toàn nhưng nêu rõ xác nhận cũ không tự áp dụng
        missing_items = res["missing_in_source_items"]
        self.assertEqual(len(missing_items), 1)
        self.assertIn("Xác nhận cũ cho kỳ 'Từ ngày 01-08-2026 đến ngày 31-08-2026' không tự áp dụng", missing_items[0]["note"])

    def test_survey_approved_period_but_unit_changed_does_not_claim_approved(self):
        """Khi đúng kỳ báo cáo nhưng ĐVT nguồn của mã 430200173 thay đổi (ví dụ CAI -> KG),
        reporting_period_note và ghi chú chi tiết không được ghi 'đã chốt ... Cái',
        mà phải cảnh báo thay đổi ĐVT và yêu cầu xác nhận lại.
        """
        # Thay đổi ĐVT của mã 430200173 trong source thành 'KG'
        wb_src = openpyxl.load_workbook(self.src_path)
        ws_src = wb_src["Sheet1"]
        ws_src.cell(16, 2, "430200173")
        ws_src.cell(16, 6, "KG")
        ws_src.cell(16, 13, 12345.0)
        wb_src.save(self.src_path)
        wb_src.close()

        res = run_survey(
            config_path=str(self.cfg_path),
            source_file=str(self.src_path),
            target_file=str(self.tgt_path),
            out_dir_path=str(self.out_dir),
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["reporting_period"], "Từ ngày 01-08-2026 đến ngày 31-08-2026")

        # Ghi chú tổng quan KHÔNG được ghi 'ĐVT Cái (chép trực tiếp'
        self.assertNotIn("ĐVT Cái (chép trực tiếp", res["reporting_period_note"])
        self.assertIn("Mã 430200173 có ĐVT nguồn là 'KG' (thay đổi so với ĐVT 'CAI' đã chốt)", res["reporting_period_note"])
        self.assertIn("Không tự áp dụng xác nhận 'chép trực tiếp Cái'", res["reporting_period_note"])

        # Ghi chú chi tiết trong divergent_units (hoặc exact_matches vì KG == Kg)
        # Ở đây nguồn=KG vs đích=Kg -> exact match về mặt đơn vị tính, nhưng vẫn kiểm tra note của mã
        self.assertNotIn("Người dùng đã CHỐT", res["reporting_period_note"])

    def test_survey_approved_period_but_code_missing_does_not_claim_approved(self):
        """Khi đúng kỳ báo cáo nhưng mã 430200173 không có trong nguồn,
        reporting_period_note không được ghi 'đã chốt ... Cái cho mã 430200173',
        mà phải cảnh báo mã không có trong nguồn.
        """
        # Xóa mã 430200173 khỏi source (xóa dòng 16)
        wb_src = openpyxl.load_workbook(self.src_path)
        ws_src = wb_src["Sheet1"]
        ws_src.delete_rows(16, 1)
        wb_src.save(self.src_path)
        wb_src.close()

        res = run_survey(
            config_path=str(self.cfg_path),
            source_file=str(self.src_path),
            target_file=str(self.tgt_path),
            out_dir_path=str(self.out_dir),
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["reporting_period"], "Từ ngày 01-08-2026 đến ngày 31-08-2026")

        # Ghi chú tổng quan KHÔNG được ghi 'ĐVT Cái (chép trực tiếp'
        self.assertNotIn("ĐVT Cái (chép trực tiếp", res["reporting_period_note"])
        self.assertIn("Mã 430200173 không có trong báo cáo nguồn (không áp dụng xác nhận chép ĐVT Cái)", res["reporting_period_note"])

    def test_survey_fetch_target_failure_removes_stale_proposal_and_emits_failed_report(self):
        """Khi pha fetch_target thất bại (ví dụ Graph 404),
        survey phải:
        1. Xóa bỏ proposal cũ (không lưu sót artifact không hợp lệ).
        2. Ghi report failed của lần chạy hiện tại kèm đầy đủ thông tin source/target.
        3. Ghi audit_summary.json status=failed và danh sách file trong thư mục cha (nếu có).
        """
        # Tạo sẵn file proposal và report cũ trong out_dir
        self.out_dir.mkdir(parents=True, exist_ok=True)
        stale_prop = self.out_dir / "nvl_stock_proposal.xlsx"
        stale_prop.write_bytes(b"stale proposal bytes")
        stale_rep = self.out_dir / "nvl_stock_report.json"
        stale_rep.write_text(json.dumps({"mode": "dry_run", "status": "stale"}), encoding="utf-8")

        mock_graph = MagicMock()
        mock_graph.get_site_id.return_value = "site-123"
        mock_graph.get_default_drive_id.return_value = "drive-123"
        # Giả lập Graph 404 khi tìm file đích
        mock_graph.get_item_by_path.side_effect = RuntimeError(
            "Microsoft Graph lỗi 404: {'error': {'code': 'itemNotFound', 'message': 'The resource could not be found.'}}"
        )
        mock_graph.list_folder_children.return_value = [
            {"name": "Existing_File_1.xlsx", "id": "id-1", "eTag": "etag-1", "size": 1024, "lastModifiedDateTime": "2026-09-09T10:00:00Z"},
            {"name": "Another_File.xlsx", "id": "id-2", "eTag": "etag-2", "size": 2048, "lastModifiedDateTime": "2026-09-09T11:00:00Z"},
        ]

        with self.assertRaises(RuntimeError) as ctx:
            run_survey(
                config_path=str(self.cfg_path),
                out_dir_path=str(self.out_dir),
                graph=mock_graph,
            )
        self.assertIn("itemNotFound", str(ctx.exception))

        # 1. Proposal cũ PHẢI bị xóa
        self.assertFalse(stale_prop.exists(), "Proposal cũ phải bị xóa khi gặp lỗi fetch_target!")

        # 2. Report failed của lần chạy này PHẢI được ghi
        self.assertTrue(stale_rep.exists(), "Report failed phải được ghi!")
        rep_data = json.loads(stale_rep.read_text(encoding="utf-8"))
        self.assertEqual(rep_data["status"], "failed")
        self.assertEqual(rep_data["mode"], "failed")
        self.assertEqual(rep_data["phase"], "fetch_target")
        self.assertIn("itemNotFound", rep_data["error_message"])
        self.assertEqual(rep_data["source"]["name"], "XNT_ketoan_Vikoda.xlsm")
        self.assertEqual(rep_data["target"]["name"], "Kế hoạch mua hàng.xlsx")
        self.assertEqual(rep_data["target"]["sharepoint_path"], "Tinh san xuat Mua hang 2027/Kế hoạch mua hàng.xlsx")

        # 3. Audit summary failed PHẢI có danh sách file cha
        audit_file = self.out_dir / "audit_summary.json"
        self.assertTrue(audit_file.exists())
        audit_data = json.loads(audit_file.read_text(encoding="utf-8"))
        self.assertEqual(audit_data["status"], "failed")
        self.assertEqual(audit_data["phase"], "fetch_target")
        self.assertEqual(len(audit_data["available_files_in_parent"]), 2)
        self.assertEqual(audit_data["available_files_in_parent"][0]["name"], "Existing_File_1.xlsx")


if __name__ == "__main__":
    unittest.main()
