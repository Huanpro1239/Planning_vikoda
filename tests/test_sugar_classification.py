import unittest

from weekly_planning_engine import (
    WeeklyInputRow,
    calculate_row,
    is_sugar_classification,
)


def make_row(classification):
    # sl_me != sl_ca để phân biệt rõ cơ sở làm tròn nào được dùng.
    return WeeklyInputRow(
        source_row=2,
        ma_sp=130100011,
        ten_sp="A",
        don_vi_tinh="Thùng",
        sl_me=700.0,
        sl_ca=500.0,
        chuyen="KHS",
        nhom_sp="KHS",
        phan_loai=classification,
        quy_cach=1.0,
        shifts_per_day=2.0,
        ton_dau_thuc_te=0.0,
        ton_dau_so_sach=0.0,
        fc=650.0,
        ton_cuoi_du_kien=0.0,
        no_kho=0.0,
        avg_daily_sales=25.0,
        leadtime=0.0,
        debt_formula_mode="SUBTRACT_BOOK_ON_DEBT",
    )


class SugarClassificationTests(unittest.TestCase):
    def test_helper_is_case_and_space_insensitive(self):
        for value in ("Có đường", "có đường", "  CÓ ĐƯỜNG ", "có  đường".replace("  ", " ")):
            self.assertTrue(is_sugar_classification(value), value)
        for value in ("Không đường", "", None, "co duong"):
            self.assertFalse(is_sugar_classification(value), value)

    def test_lowercase_sugar_rounds_on_batch_like_canonical(self):
        # p_need = fc = 650 -> làm tròn lên bội số của SL/mẻ (700) cho hàng "có đường".
        canonical = calculate_row(make_row("Có đường"), period_year=2026, period_month=9)
        lowercase = calculate_row(make_row("có đường"), period_year=2026, period_month=9)
        self.assertEqual(canonical.q_rounded, 700.0)
        self.assertEqual(lowercase.q_rounded, 700.0)

    def test_non_sugar_rounds_on_shift(self):
        row = calculate_row(make_row("Không đường"), period_year=2026, period_month=9)
        # 650 -> bội số của SL/ca (500) = 1000.
        self.assertEqual(row.q_rounded, 1000.0)


if __name__ == "__main__":
    unittest.main()
