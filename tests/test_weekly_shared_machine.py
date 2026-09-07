import unittest
from collections import defaultdict
from datetime import date

from sync_planning_weekly_model import (
    WeeklyAnalysis,
    build_weekly_schedule_report,
)
from weekly_planning_engine import (
    DailyPlanRow,
    PlannerPolicy,
    WeeklyInputRow,
    build_daily_plan,
    calculate_rows,
)


def make_row(code, line, source_row, *, qty=200.0, quy_cach=1.0, shifts=2.0):
    return WeeklyInputRow(
        source_row=source_row,
        ma_sp=code,
        ten_sp=str(code),
        don_vi_tinh="DV",
        sl_me=100.0,
        sl_ca=100.0,
        chuyen=line,
        nhom_sp=line,
        phan_loai="Không đường",
        quy_cach=quy_cach,
        shifts_per_day=shifts,
        ton_dau_thuc_te=0.0,
        ton_dau_so_sach=0.0,
        fc=qty,
        ton_cuoi_du_kien=0.0,
        no_kho=0.0,
        avg_daily_sales=100.0,
        leadtime=0.0,
        debt_formula_mode="SUBTRACT_BOOK_ON_DEBT",
    )


class WeeklySharedMachineTests(unittest.TestCase):
    def test_khs_and_pet9000_cannot_run_in_parallel(self):
        rows = calculate_rows(
            [
                make_row(9001, "KHS", 2, qty=200.0, quy_cach=1.0),
                make_row(9002, "PET 9000", 3, qty=200.0, quy_cach=1.0),
            ],
            period_year=2026,
            period_month=9,
        )
        plan = build_daily_plan(rows, policy=PlannerPolicy())

        qty_by_code = defaultdict(float)
        shift_usage = defaultdict(float)
        per_shift = {row.input.ma_sp: row.input.sl_ca for row in rows}
        for item in plan:
            qty_by_code[item.ma_sp] += item.qty
            shift_usage[item.date] += item.qty / per_shift[item.ma_sp]

        self.assertEqual(qty_by_code[9001], 200.0)
        self.assertEqual(qty_by_code[9002], 200.0)
        self.assertTrue(shift_usage)
        for usage in shift_usage.values():
            self.assertLessEqual(usage, 2.0 + 1e-6)

        # KHS consumes both shifts on day 1, so PET cannot also run on day 1.
        self.assertEqual(
            sum(item.qty for item in plan if item.ma_sp == 9002 and item.date == date(2026, 9, 1)),
            0.0,
        )

    def test_changeover_applies_when_switching_between_shared_lines(self):
        rows = calculate_rows(
            [
                make_row(9101, "KHS", 2, qty=100.0, quy_cach=1.0),
                make_row(9102, "PET 9000", 3, qty=100.0, quy_cach=2.0),
            ],
            period_year=2026,
            period_month=9,
        )
        plan = build_daily_plan(rows, policy=PlannerPolicy(setup_shifts=0.5))

        pet_day1 = sum(
            item.qty
            for item in plan
            if item.ma_sp == 9102 and item.date == date(2026, 9, 1)
        )
        pet_day2 = sum(
            item.qty
            for item in plan
            if item.ma_sp == 9102 and item.date == date(2026, 9, 2)
        )

        # After KHS uses shift 1, 0.5 shift is setup, leaving only 0.5 PET shift.
        self.assertAlmostEqual(pet_day1, 50.0)
        self.assertAlmostEqual(pet_day2, 50.0)

    def test_shared_machine_requires_one_common_shift_calendar(self):
        rows = calculate_rows(
            [
                make_row(9201, "KHS", 2, shifts=2.0),
                make_row(9202, "PET 9000", 3, shifts=3.0),
            ],
            period_year=2026,
            period_month=9,
        )
        with self.assertRaisesRegex(ValueError, "chung một máy"):
            build_daily_plan(rows, policy=PlannerPolicy())


    def test_report_cannot_mark_parallel_shared_machine_as_passed(self):
        calculated = calculate_rows(
            [
                make_row(9301, "KHS", 2, qty=200.0, quy_cach=1.0),
                make_row(9302, "PET 9000", 3, qty=200.0, quy_cach=1.0),
            ],
            period_year=2026,
            period_month=9,
        )
        # Deliberately construct an impossible plan: both products consume the
        # full two-shift machine on the same day.
        impossible = [
            DailyPlanRow(9301, 2, "KHS", date(2026, 9, 1), 200.0),
            DailyPlanRow(9302, 3, "PET 9000", date(2026, 9, 1), 200.0),
        ]
        analysis = WeeklyAnalysis(
            calculated=calculated,
            daily_plan=impossible,
            policy_warnings=[],
            changed_cells=0,
            period_year=2026,
            period_month=9,
        )
        report = build_weekly_schedule_report(
            b"synthetic",
            analysis,
            input_revision={"target": {"etag": "synthetic"}},
        )
        resource = report["status"]["resource_validation"]
        self.assertFalse(resource["ok"])
        self.assertEqual(resource["state"], "shared_machine_over_capacity")
        self.assertEqual(resource["shared_machine"]["over_capacity_dates"], ["2026-09-01"])
        self.assertEqual(report["publish_status"], "review_required")

    def test_schedule_dates_never_drift_outside_planning_month(self):
        # ton_dau_thuc_te = 4000 with avg_daily_sales = 100 means stockout is 40 days away (in October)
        rows = calculate_rows(
            [
                WeeklyInputRow(
                    source_row=2, ma_sp=9401, ten_sp="A", don_vi_tinh="DV",
                    sl_me=100.0, sl_ca=100.0, chuyen="KHS", nhom_sp="KHS", phan_loai="Không đường",
                    quy_cach=1.0, shifts_per_day=1.0,
                    ton_dau_thuc_te=4000.0, ton_dau_so_sach=0.0, fc=2600.0, ton_cuoi_du_kien=1000.0,
                    no_kho=0.0, avg_daily_sales=100.0, leadtime=0.0, debt_formula_mode="SUBTRACT_BOOK_ON_DEBT",
                )
            ],
            period_year=2026,
            period_month=9,
        )
        # Column R (start_datetime) reflects the raw stockout date (in October)
        self.assertEqual(rows[0].start_datetime.month, 10)

        # But daily plan must remain strictly inside September 2026!
        plan = build_daily_plan(rows, policy=PlannerPolicy())
        self.assertTrue(len(plan) > 0)
        for item in plan:
            self.assertEqual(item.date.year, 2026)
            self.assertEqual(item.date.month, 9)

    def test_schedule_works_across_all_month_lengths(self):
        # 28 days (Feb 2025), 29 days (Feb 2028), 30 days (Apr 2026), 31 days (May 2026)
        cases = [(2025, 2, 28), (2028, 2, 29), (2026, 4, 30), (2026, 5, 31)]
        for year, month, expected_days in cases:
            rows = calculate_rows(
                [
                    make_row(9501, "KHS", 2, qty=100.0, quy_cach=1.0, shifts=1.0),
                    make_row(9502, "PET 9000", 3, qty=100.0, quy_cach=1.0, shifts=1.0),
                ],
                period_year=year,
                period_month=month,
            )
            plan = build_daily_plan(rows, policy=PlannerPolicy())
            dates = {item.date for item in plan}
            for d in dates:
                self.assertEqual(d.year, year)
                self.assertEqual(d.month, month)
                self.assertLessEqual(d.day, expected_days)

    def test_priority_and_capacity_trim_fits_within_monthly_limits(self):
        # 3 SKUs on shared machine KHS + PET 9000, month = Sep (30 days * 1 shift = 30 shifts)
        # SKU 9601: has debt, must run first
        # SKU 9602: huge FC = 2500 (25 shifts)
        # SKU 9603: normal FC = 1500 (15 shifts)
        # Total raw service shifts = 10 + 25 + 15 = 50 shifts > 30 shifts capacity!
        rows = calculate_rows(
            [
                WeeklyInputRow(
                    source_row=2, ma_sp=9601, ten_sp="Debt SKU", don_vi_tinh="DV",
                    sl_me=100.0, sl_ca=100.0, chuyen="KHS", nhom_sp="KHS", phan_loai="Không đường",
                    quy_cach=1.0, shifts_per_day=1.0,
                    ton_dau_thuc_te=0.0, ton_dau_so_sach=0.0, fc=500.0, ton_cuoi_du_kien=0.0,
                    no_kho=500.0, avg_daily_sales=50.0, leadtime=0.0, debt_formula_mode="SUBTRACT_BOOK_ON_DEBT",
                ),
                WeeklyInputRow(
                    source_row=3, ma_sp=9602, ten_sp="Huge FC SKU", don_vi_tinh="DV",
                    sl_me=100.0, sl_ca=100.0, chuyen="KHS", nhom_sp="KHS", phan_loai="Không đường",
                    quy_cach=1.0, shifts_per_day=1.0,
                    ton_dau_thuc_te=100.0, ton_dau_so_sach=0.0, fc=2500.0, ton_cuoi_du_kien=0.0,
                    no_kho=0.0, avg_daily_sales=100.0, leadtime=0.0, debt_formula_mode="SUBTRACT_BOOK_ON_DEBT",
                ),
                WeeklyInputRow(
                    source_row=4, ma_sp=9603, ten_sp="Normal FC SKU", don_vi_tinh="DV",
                    sl_me=100.0, sl_ca=100.0, chuyen="PET 9000", nhom_sp="PET 9000", phan_loai="Không đường",
                    quy_cach=1.0, shifts_per_day=1.0,
                    ton_dau_thuc_te=500.0, ton_dau_so_sach=0.0, fc=1500.0, ton_cuoi_du_kien=0.0,
                    no_kho=0.0, avg_daily_sales=50.0, leadtime=0.0, debt_formula_mode="SUBTRACT_BOOK_ON_DEBT",
                ),
            ],
            period_year=2026,
            period_month=9,
        )
        plan = build_daily_plan(rows, policy=PlannerPolicy(allow_capacity_trim=True))
        self.assertTrue(len(plan) > 0)

        # 1. Total scheduled shifts must be <= 30
        total_shifts = sum(item.qty / 100.0 for item in plan)
        self.assertLessEqual(total_shifts, 30.0 + 1e-6)

        # 2. Every single SKU must have at least 1 batch scheduled (no SKU left at 0)
        scheduled_by_code = {}
        for item in plan:
            scheduled_by_code[item.ma_sp] = scheduled_by_code.get(item.ma_sp, 0.0) + item.qty
        self.assertEqual(len(scheduled_by_code), 3)
        for code in [9601, 9602, 9603]:
            self.assertGreaterEqual(scheduled_by_code[code], 100.0)

        # 3. Priority: Debt SKU 9601 must start on day 1
        sku_9601_dates = [item.date for item in plan if item.ma_sp == 9601]
        self.assertEqual(min(sku_9601_dates), date(2026, 9, 1))

        # 4. Each SKU must run contiguously (no gap)
        for code in [9601, 9602, 9603]:
            dates = sorted(item.date for item in plan if item.ma_sp == code)
            for i in range(len(dates) - 1):
                self.assertEqual((dates[i + 1] - dates[i]).days, 1)


if __name__ == "__main__":
    unittest.main()
