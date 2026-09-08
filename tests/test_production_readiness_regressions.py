import json
import unittest
from datetime import date, datetime
from io import BytesIO

from openpyxl import Workbook

import sync_planning_metrics_direct as direct_metrics
import sync_planning_pipeline as pipeline_runner
import sync_stock
import sync_stock_compat
from sync_planning_weekly_model import (
    compute_planning_inputs_hash,
    prepare_weekly_schedule_update,
    verify_weekly_workbook,
)
from weekly_planning_engine import (
    PlannerPolicy,
    WeeklyInputRow,
    build_daily_plan,
    calculate_rows,
)


def make_workbook_bytes(
    *,
    skus=None,
    debt_mode=True,
    galon_profile=True,
    debts=None,
):
    """Create a minimal in-memory Excel workbook for pipeline testing."""
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
    ] + [f"{day:02d}/09" for day in range(1, 31)])

    if skus is None:
        skus = [
            {
                "code": 130100011,
                "name": "SP A",
                "uom": "Thùng",
                "batch": 100,
                "per_shift": 200,
                "line": "KHS",
                "group": "Nhom",
                "classification": "Không đường",
                "shifts_per_day": 1,
                "ton_dau_thuc_te": 0,
                "ton_dau_so_sach": 0,
                "fc": 8000,
                "ton_cuoi_du_kien": 0,
                "no_kho": 0,
                "mold": 10,
                "leadtime": 0,
                "debt_mode": "SUBTRACT_BOOK_ON_DEBT",
                "profile": "",
            }
        ]

    for item in skus:
        planning.append([
            item["code"],
            item.get("name", "SP"),
            item.get("uom", "Thùng"),
            item.get("batch", 100),
            item.get("per_shift", 200),
            item.get("line", "KHS"),
            item.get("group", "Nhom"),
            item.get("classification", "Không đường"),
            item.get("shifts_per_day", 1),
            item.get("ton_dau_thuc_te", 0),
            item.get("ton_dau_so_sach", 0),
            item.get("fc", 0),
            item.get("ton_cuoi_du_kien", 0),
            item.get("no_kho", 0),
            0, 0, 0, None,
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

    for item in skus:
        row = [
            item["code"],
            item.get("name", "SP"),
            item.get("uom", "Thùng"),
            item.get("batch", 100),
            item.get("per_shift", 200),
            item.get("line", "KHS"),
            item.get("group", "Nhom"),
            item.get("classification", "Không đường"),
            item.get("mold", 10),
            item.get("leadtime", 0),
        ]
        if debt_mode:
            row.append(item.get("debt_mode", "SUBTRACT_BOOK_ON_DEBT"))
        if galon_profile:
            row.append(item.get("profile", ""))
        master.append(row)

    if debts is not None:
        debt_sheet = wb.create_sheet("No kho")
        debt_sheet.append(["Mã Sản Phẩm", "Tên", "ĐVT", "Số lượng nợ"])
        for code, val in debts.items():
            debt_sheet.append([code, "SP", "Thùng", val])

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


class ProductionReadinessRegressionTests(unittest.TestCase):
    def test_capacity_shortfall_blocks_autopublish_requires_approval(self):
        """Defect 1: Capacity shortfall must mark service.ok=False and require approval."""
        # 8000 target demand, but monthly capacity is only 6000 (30 days * 200/day).
        wb_bytes = make_workbook_bytes(
            skus=[{
                "code": 130100011,
                "name": "SP A",
                "uom": "Thùng",
                "batch": 100,
                "per_shift": 200,
                "line": "KHS",
                "group": "Nhom",
                "classification": "Không đường",
                "shifts_per_day": 1,
                "ton_dau_thuc_te": 0,
                "ton_dau_so_sach": 0,
                "fc": 8000,
                "ton_cuoi_du_kien": 0,
                "no_kho": 0,
                "mold": 10,
                "leadtime": 0,
                "debt_mode": "SUBTRACT_BOOK_ON_DEBT",
                "profile": "",
            }],
            debt_mode=True,
        )

        final_bytes, report, analysis = prepare_weekly_schedule_update(
            wb_bytes,
            plan_year=2026,
            plan_month=9,
            input_revision={"target": {"etag": "offline-review"}},
        )

        self.assertEqual(report["publish_status"], "review_required")
        self.assertFalse(report["status"]["service"]["ok"])
        self.assertEqual(report["status"]["service"]["stockout_skus"], ["130100011"])
        self.assertEqual(report["status"]["service"]["capacity_balanced_skus"], ["130100011"])

        # Auto-publish without approval MUST be blocked
        authorized, decision = pipeline_runner._publish_decision(report, None, "scheduler")
        self.assertFalse(authorized)
        self.assertEqual(decision["state"], "blocked")

        # Explicit approval with reason for the exact proposal MUST be accepted
        proposal_id = report.get("proposal_id", "test_prop")
        report["proposal_id"] = proposal_id
        authorized_approved, decision_approved = pipeline_runner._publish_decision(
            report,
            {"proposal_id": proposal_id, "reason": "Chấp nhận thiếu 2000 thùng vì hết công suất"},
            "planner_lead",
        )
        self.assertTrue(authorized_approved)
        self.assertEqual(decision_approved["state"], "authorized")
        self.assertEqual(decision_approved["basis"], "review_approval")

        # verify_weekly_workbook must pass with this report
        verification = verify_weekly_workbook(
            final_bytes,
            schedule_report=report,
            plan_year=2026,
            plan_month=9,
        )
        self.assertTrue(verification["validated"])

    def test_all_input_changes_detected_and_outputs_ignored(self):
        """Defect 2: All master policy, No kho, and Ke_hoach_SX inputs trigger recalculation."""
        base_skus = [{
            "code": 130100011,
            "name": "SP A",
            "uom": "Thùng",
            "batch": 100,
            "per_shift": 200,
            "line": "KHS",
            "group": "Nhom",
            "classification": "Không đường",
            "shifts_per_day": 1,
            "ton_dau_thuc_te": 0,
            "ton_dau_so_sach": 0,
            "fc": 1000,
            "ton_cuoi_du_kien": 0,
            "no_kho": 0,
            "mold": 10,
            "leadtime": 0,
            "debt_mode": "SUBTRACT_BOOK_ON_DEBT",
            "profile": "",
        }]

        base_bytes = make_workbook_bytes(skus=base_skus, debts={130100011: 0})
        _, base_conv_hash = sync_stock_compat.read_conversion_factors_robust(base_bytes)
        _, base_no_kho_hash = direct_metrics.hash_debt_sheet(base_bytes)
        base_planning_hash = compute_planning_inputs_hash(base_bytes)

        old_state = {
            "sources": {k: "etag_1" for k in pipeline_runner.SOURCES},
            "conversion_hash": base_conv_hash,
            "fc_hash": "fc_hash_1",
            "no_kho_hash": base_no_kho_hash,
            "planning_inputs_hash": base_planning_hash,
            "engine_version": "ke_hoach_sx_tuan_v2_service_first_20260908",
        }
        source_items = {k: {"eTag": "etag_1"} for k in pipeline_runner.SOURCES}

        # 1. Same inputs -> No changes
        info_same = {
            "conversion_hash": base_conv_hash,
            "fc_hash": "fc_hash_1",
            "no_kho_hash": base_no_kho_hash,
            "planning_inputs_hash": base_planning_hash,
            "engine_version": "ke_hoach_sx_tuan_v2_service_first_20260908",
        }
        self.assertEqual(
            pipeline_runner.detect_input_changes(old_state, source_items, info_same),
            [],
        )

        # 2. Leadtime changed in Danh_muc -> sheet:Danh_muc detected
        skus_leadtime = [dict(base_skus[0], leadtime=7)]
        bytes_leadtime = make_workbook_bytes(skus=skus_leadtime, debts={130100011: 0})
        _, conv_leadtime_hash = sync_stock_compat.read_conversion_factors_robust(bytes_leadtime)
        self.assertNotEqual(base_conv_hash, conv_leadtime_hash)
        info_leadtime = dict(info_same, conversion_hash=conv_leadtime_hash)
        self.assertEqual(
            pipeline_runner.detect_input_changes(old_state, source_items, info_leadtime),
            ["sheet:Danh_muc"],
        )

        # 3. Debt changed in No kho -> sheet:No_kho detected
        bytes_debt = make_workbook_bytes(skus=base_skus, debts={130100011: 500})
        _, debt_hash = direct_metrics.hash_debt_sheet(bytes_debt)
        self.assertNotEqual(base_no_kho_hash, debt_hash)
        info_debt = dict(info_same, no_kho_hash=debt_hash)
        self.assertEqual(
            pipeline_runner.detect_input_changes(old_state, source_items, info_debt),
            ["sheet:No_kho"],
        )

        # 4. shifts_per_day changed in Ke_hoach_SX -> sheet:Ke_hoach_SX detected
        skus_shifts = [dict(base_skus[0], shifts_per_day=2)]
        bytes_shifts = make_workbook_bytes(skus=skus_shifts, debts={130100011: 0})
        planning_shifts_hash = compute_planning_inputs_hash(bytes_shifts)
        self.assertNotEqual(base_planning_hash, planning_shifts_hash)
        info_shifts = dict(info_same, planning_inputs_hash=planning_shifts_hash)
        self.assertEqual(
            pipeline_runner.detect_input_changes(old_state, source_items, info_shifts),
            ["sheet:Ke_hoach_SX"],
        )

        # 5. Output cells in Ke_hoach_SX do NOT change planning_inputs_hash
        final_bytes, _, _ = prepare_weekly_schedule_update(
            base_bytes,
            plan_year=2026,
            plan_month=9,
            input_revision={"target": {"etag": "offline-review"}},
        )
        output_planning_hash = compute_planning_inputs_hash(final_bytes)
        self.assertEqual(base_planning_hash, output_planning_hash)

        # 6. Legacy old_state without new hashes triggers recalculation safely
        legacy_state = {
            "sources": {k: "etag_1" for k in pipeline_runner.SOURCES},
            "conversion_hash": base_conv_hash,
            "fc_hash": "fc_hash_1",
        }
        detected_legacy = pipeline_runner.detect_input_changes(legacy_state, source_items, info_same)
        self.assertIn("sheet:No_kho", detected_legacy)
        self.assertIn("sheet:Ke_hoach_SX", detected_legacy)

    def test_buffer_does_not_displace_service_timeline(self):
        """Defect 3: Timeline capacity constraints must prioritize service over earlier SKU buffer."""
        from dataclasses import replace
        from tests.test_service_first_safety_stock import row

        inputs = [
            replace(row(1001, 'KHS', 2, fc=100, target=900), ton_dau_thuc_te=2000),
            replace(row(1002, 'PET 9000', 3, fc=900, target=0), ton_dau_thuc_te=2100),
        ]
        calculated = calculate_rows(inputs, period_year=2026, period_month=9)
        plan = build_daily_plan(calculated, policy=PlannerPolicy(setup_shifts=0, allow_capacity_trim=True))

        scheduled_a = sum(p.qty for p in plan if p.ma_sp == 1001)
        scheduled_b = sum(p.qty for p in plan if p.ma_sp == 1002)

        # SKU A gets exactly 100 (service); buffer 900 is dropped because capacity on timeline cannot fit it
        self.assertEqual(scheduled_a, 100.0)
        # SKU B gets full 900 (service)
        self.assertEqual(scheduled_b, 900.0)

    def test_zero_forecast_with_debt_schedules_production(self):
        """Defect 4: When FC=0 and Nợ kho > 0, start_datetime must be day 1 and production must be scheduled."""
        for mode in ("SUBTRACT_BOOK_ON_DEBT", "IGNORE_BOOK_ON_DEBT"):
            wb_bytes = make_workbook_bytes(
                skus=[{
                    "code": 130100011,
                    "name": "SP A",
                    "uom": "Thùng",
                    "batch": 100,
                    "per_shift": 500,
                    "line": "KHS",
                    "group": "Nhom",
                    "classification": "Không đường",
                    "shifts_per_day": 1,
                    "ton_dau_thuc_te": 0,
                    "ton_dau_so_sach": 0,
                    "fc": 0,
                    "ton_cuoi_du_kien": 0,
                    "no_kho": 500,
                    "mold": 10,
                    "leadtime": 0,
                    "debt_mode": mode,
                    "profile": "",
                }],
                debt_mode=True,
            )

            final_bytes, report, analysis = prepare_weekly_schedule_update(
                wb_bytes,
                plan_year=2026,
                plan_month=9,
                input_revision={"target": {"etag": "offline-review"}},
            )

            self.assertEqual(report["publish_status"], "ready_for_publish")
            calc_row = analysis.calculated[0]
            self.assertEqual(calc_row.q_rounded, 500.0)
            self.assertEqual(calc_row.start_datetime, datetime(2026, 9, 1, 0, 0, 0))

            mb = report["mass_balance"]["130100011"]
            self.assertEqual(mb["planned_qty"], 500.0)
            self.assertEqual(mb["scheduled_qty"], 500.0)
            self.assertEqual(mb["carryover_qty"], 0.0)


if __name__ == "__main__":
    unittest.main()
