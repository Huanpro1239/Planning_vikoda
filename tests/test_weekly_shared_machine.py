import unittest
from collections import defaultdict
from datetime import date

from weekly_planning_engine import (
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


if __name__ == "__main__":
    unittest.main()
