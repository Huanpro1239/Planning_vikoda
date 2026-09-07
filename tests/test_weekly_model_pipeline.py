import unittest
from io import BytesIO

from openpyxl import Workbook, load_workbook

import sync_planning_pipeline as publish_runner
from sync_planning_weekly_model import (
    prepare_weekly_schedule_update,
    verify_weekly_workbook,
)


def workbook_bytes(*, debt_mode=True, galon_profile=True, stale_daily=False):
    wb = Workbook()
    planning = wb.active
    planning.title = "Ke_hoach_SX"
    planning.append([
        "Mã Sản Phẩm", "Tên Sản Phẩm", "Đơn vị tính", "Số lượng /mẻ",
        "Số lượng/ ca", "Chuyền", "Nhóm sản phẩm", "Phân loại SP",
        "Số ca theo ngày", "Tồn đầu thực tế", "Tồn đầu sổ sách", "FC",
        "Tồn cuối dự kiến", "Nợ kho", "Số lượng cần sản xuất",
        "Số lượng sản xuất theo mẻ/ca", "Số ngày cần sản xuất",
        "Ngày bắt đầu sản xuất",
    ] + [f"{day:02d}/04" for day in range(1, 31)])
    planning.append([
        9001, "A", "Thùng", 100, 500, "KHS", "Nhom", "Không đường",
        2, 100, 50, 1000, 200, 100, 0, 0, 0, None,
    ] + ([777] + [None] * 29 if stale_daily else [None] * 30))
    planning.append([
        9002, "B", "Bình", 100, 100, "Galon", "Galon", "Không đường",
        1, 0, 0, 500, 0, 0, 0, 0, 0, None,
    ] + [None] * 30)

    master = wb.create_sheet("Danh_muc")
    headers = [
        "Mã Sản Phẩm", "Tên Sản Phẩm", "Đơn vị tính", "Số lượng /mẻ",
        "Số lượng/ ca", "Chuyền", "Nhóm sản phẩm", "Phân loại SP",
        "Quy cách", "Leadtime",
    ]
    if debt_mode:
        headers.append("Debt mode")
    if galon_profile:
        headers.append("Schedule profile")
    master.append(headers)

    row1 = [9001, "A", "Thùng", 100, 500, "KHS", "Nhom", "Không đường", 10, 2]
    row2 = [9002, "B", "Bình", 100, 100, "Galon", "Galon", "Không đường", 20, 1]
    if debt_mode:
        row1.append("IGNORE_BOOK_ON_DEBT")
        row2.append("SUBTRACT_BOOK_ON_DEBT")
    if galon_profile:
        row1.append("")
        row2.append("SPREAD_NON_SUNDAY")
    master.append(row1)
    master.append(row2)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


class WeeklyModelPipelineTests(unittest.TestCase):
    def test_explicit_policy_can_be_ready_for_publish(self):
        raw = workbook_bytes()
        updated, report, _ = prepare_weekly_schedule_update(
            raw,
            plan_year=2025,
            plan_month=4,
            input_revision={"target": {"etag": "snapshot-a"}},
        )
        self.assertEqual(report["algorithm"], "ke_hoach_sx_tuan_v1")
        self.assertEqual(report["publish_status"], "ready_for_publish")
        self.assertEqual(report["policy_warnings"], [])
        self.assertTrue(
            verify_weekly_workbook(
                updated,
                schedule_report=report,
                plan_year=2025,
                plan_month=4,
            )["validated"]
        )

        wb = load_workbook(BytesIO(updated), data_only=True, read_only=True)
        try:
            ws = wb["Ke_hoach_SX"]
            # Debt mode IGNORE_BOOK: P = FC + debt = 1100.
            self.assertEqual(ws["O2"].value, 1100)
            # Không đường: Q làm tròn theo 500/ca => 1500.
            self.assertEqual(ws["P2"].value, 1500)
        finally:
            wb.close()

    def test_missing_policy_metadata_forces_review_required(self):
        raw = workbook_bytes(debt_mode=False, galon_profile=False)
        _, report, _ = prepare_weekly_schedule_update(
            raw,
            plan_year=2025,
            plan_month=4,
            input_revision={"target": {"etag": "snapshot-b"}},
        )
        self.assertEqual(report["publish_status"], "review_required")
        self.assertEqual(report["status"]["service"]["state"], "policy_metadata_missing")
        warning_types = {item["type"] for item in report["policy_warnings"]}
        self.assertIn("MISSING_DEBT_MODE", warning_types)
        self.assertIn("MISSING_SCHEDULE_PROFILE", warning_types)

    def test_rebuild_clears_stale_daily_schedule_cells(self):
        raw = workbook_bytes(stale_daily=True)
        updated, report, _ = prepare_weekly_schedule_update(
            raw,
            plan_year=2025,
            plan_month=4,
            input_revision={"target": {"etag": "snapshot-c"}},
        )
        verify_weekly_workbook(
            updated,
            schedule_report=report,
            plan_year=2025,
            plan_month=4,
        )
        wb = load_workbook(BytesIO(updated), data_only=True, read_only=True)
        try:
            # T2 = 19th column = day 01/04. The stale 777 must not survive.
            self.assertNotEqual(wb["Ke_hoach_SX"]["S2"].value, 777)
        finally:
            wb.close()

    def test_policy_metadata_missing_cannot_be_review_override(self):
        report = {
            "publish_status": "review_required",
            "proposal_id": "proposal-policy",
            "status": {
                "monthly_quantity": {"ok": True, "state": "complete"},
                "resource_validation": {"ok": True, "state": "model_rules_passed"},
                "service": {"ok": False, "state": "policy_metadata_missing"},
            },
        }
        authorized, decision = publish_runner._publish_decision(
            report,
            {
                "proposal_id": "proposal-policy",
                "reason": "try to override",
                "approved_by": "tester",
            },
            "tester",
        )
        self.assertFalse(authorized)
        self.assertEqual(decision["basis"], "unsupported_review_reason")


if __name__ == "__main__":
    unittest.main()
