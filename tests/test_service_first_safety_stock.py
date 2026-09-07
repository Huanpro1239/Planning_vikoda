import unittest
from collections import defaultdict
from datetime import date
from io import BytesIO

from openpyxl import Workbook, load_workbook

from sync_planning_weekly_model import (
    analyze_weekly_workbook,
    build_weekly_schedule_report,
    patch_weekly_workbook,
)
from weekly_planning_engine import (
    PlannerPolicy,
    WeeklyInputRow,
    build_daily_plan,
    calculate_rows,
)


def row(
    code,
    line,
    source_row,
    *,
    fc,
    target,
    per_shift=100.0,
    shifts=1.0,
    mold=1.0,
):
    return WeeklyInputRow(
        source_row=source_row,
        ma_sp=code,
        ten_sp=str(code),
        don_vi_tinh="DV",
        sl_me=100.0,
        sl_ca=per_shift,
        chuyen=line,
        nhom_sp=line,
        phan_loai="Không đường",
        quy_cach=mold,
        shifts_per_day=shifts,
        ton_dau_thuc_te=0.0,
        ton_dau_so_sach=0.0,
        fc=fc,
        ton_cuoi_du_kien=target,
        no_kho=0.0,
        avg_daily_sales=100.0,
        leadtime=0.0,
        debt_formula_mode="SUBTRACT_BOOK_ON_DEBT",
    )


class ServiceFirstSafetyStockTests(unittest.TestCase):
    def _analysis(self, *, service_a=1000.0, service_b=1000.0):
        calculated = calculate_rows(
            [
                row(1001, "KHS", 2, fc=service_a, target=5000.0),
                row(1002, "PET 9000", 3, fc=service_b, target=0.0),
            ],
            period_year=2026,
            period_month=9,
        )
        daily = build_daily_plan(calculated, policy=PlannerPolicy(setup_shifts=0.0))
        from sync_planning_weekly_model import WeeklyAnalysis

        return WeeklyAnalysis(
            calculated=calculated,
            daily_plan=daily,
            policy_warnings=[],
            changed_cells=0,
            period_year=2026,
            period_month=9,
        )

    def test_sales_are_complete_when_only_safety_stock_exceeds_capacity(self):
        analysis = self._analysis()
        report = build_weekly_schedule_report(
            b"synthetic",
            analysis,
            input_revision={"target": {"etag": "synthetic"}},
        )

        self.assertEqual(report["publish_status"], "ready_for_publish")
        self.assertTrue(report["status"]["monthly_quantity"]["ok"])
        self.assertEqual(
            report["status"]["monthly_quantity"]["state"],
            "service_complete_buffer_shortfall",
        )
        self.assertTrue(report["status"]["service"]["ok"])
        self.assertEqual(report["status"]["service"]["stockout_skus"], [])
        self.assertFalse(report["status"]["safety_stock"]["ok"])
        self.assertEqual(
            report["status"]["safety_stock"]["state"],
            "partially_achieved",
        )

        a = report["mass_balance"]["1001"]
        b = report["mass_balance"]["1002"]
        self.assertEqual(a["service_target_qty"], 1000.0)
        self.assertEqual(a["service_carryover_qty"], 0.0)
        self.assertGreater(a["buffer_carryover_qty"], 0.0)
        self.assertEqual(b["service_carryover_qty"], 0.0)

    def test_shared_machine_schedules_all_service_before_buffer(self):
        analysis = self._analysis()
        phases_by_day = defaultdict(set)
        first_buffer_day = None
        last_service_day = None

        for item in analysis.daily_plan:
            if item.chuyen not in {"KHS", "PET 9000"}:
                continue
            phases_by_day[item.date].add(item.phase)
            if item.phase == "buffer":
                first_buffer_day = (
                    item.date
                    if first_buffer_day is None
                    else min(first_buffer_day, item.date)
                )
            if item.phase == "service":
                last_service_day = (
                    item.date
                    if last_service_day is None
                    else max(last_service_day, item.date)
                )

        self.assertIsNotNone(first_buffer_day)
        self.assertIsNotNone(last_service_day)
        self.assertGreaterEqual(first_buffer_day, last_service_day)
        # No day may exceed the one-shift shared machine.
        per_shift = {c.input.ma_sp: c.input.sl_ca for c in analysis.calculated}
        usage = defaultdict(float)
        for item in analysis.daily_plan:
            if item.chuyen in {"KHS", "PET 9000"}:
                usage[item.date] += item.qty / per_shift[item.ma_sp]
        for used in usage.values():
            self.assertLessEqual(used, 1.0 + 1e-6)

    def test_service_shortfall_remains_review_required(self):
        analysis = self._analysis(service_a=2000.0, service_b=2000.0)
        report = build_weekly_schedule_report(
            b"synthetic",
            analysis,
            input_revision={"target": {"etag": "synthetic"}},
        )
        self.assertEqual(report["publish_status"], "review_required")
        self.assertFalse(report["status"]["monthly_quantity"]["ok"])
        self.assertEqual(
            report["status"]["monthly_quantity"]["state"],
            "service_carryover",
        )
        self.assertFalse(report["status"]["service"]["ok"])
        self.assertEqual(report["status"]["service"]["state"], "stockout_risk")

    def test_workbook_commits_capacity_feasible_qty_not_unreachable_buffer(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Ke_hoach_SX"
        ws.append([
            "Mã Sản Phẩm", "Tên Sản Phẩm", "Đơn vị tính", "Số lượng /mẻ",
            "Số lượng/ ca", "Chuyền", "Nhóm sản phẩm", "Phân loại SP",
            "Số ca theo ngày", "Tồn đầu thực tế", "Tồn đầu sổ sách", "FC",
            "Tồn cuối dự kiến", "Nợ kho", "Số lượng cần sản xuất",
            "Số lượng sản xuất theo mẻ/ca", "Số ngày cần sản xuất",
            "Ngày bắt đầu sản xuất",
        ] + [f"{day:02d}/09" for day in range(1, 31)])
        ws.append([1001, "A", "DV", 100, 100, "KHS", "KHS", "Không đường",
                   1, 0, 0, 1000, 5000, 0, 0, 0, 0, None] + [None] * 30)
        ws.append([1002, "B", "DV", 100, 100, "PET 9000", "PET 9000", "Không đường",
                   1, 0, 0, 1000, 0, 0, 0, 0, 0, None] + [None] * 30)
        master = wb.create_sheet("Danh_muc")
        master.append([
            "Mã Sản Phẩm", "Tên Sản Phẩm", "Đơn vị tính", "Số lượng /mẻ",
            "Số lượng/ ca", "Chuyền", "Nhóm sản phẩm", "Phân loại SP",
            "Quy cách", "Leadtime", "Debt mode",
        ])
        master.append([1001, "A", "DV", 100, 100, "KHS", "KHS", "Không đường",
                       1, 0, "SUBTRACT_BOOK_ON_DEBT"])
        master.append([1002, "B", "DV", 100, 100, "PET 9000", "PET 9000", "Không đường",
                       1, 0, "SUBTRACT_BOOK_ON_DEBT"])

        raw = BytesIO()
        wb.save(raw)
        analysis = analyze_weekly_workbook(
            raw.getvalue(),
            plan_year=2026,
            plan_month=9,
        )
        updated = patch_weekly_workbook(raw.getvalue(), analysis)

        result = load_workbook(BytesIO(updated), data_only=True, read_only=True)
        try:
            plan = result["Ke_hoach_SX"]
            # O keeps desired need (1000 + 5000); P/Q reflect what fits.
            self.assertEqual(plan["O2"].value, 6000)
            self.assertEqual(plan["P2"].value, 2000)
            self.assertEqual(plan["Q2"].value, 20)
            self.assertEqual(plan["P3"].value, 1000)
            self.assertEqual(plan["Q3"].value, 10)
            self.assertEqual(
                sum(float(plan.cell(2, 19 + d).value or 0) for d in range(30)),
                2000.0,
            )
        finally:
            result.close()


if __name__ == "__main__":
    unittest.main()
