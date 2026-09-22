import unittest

from planning import khsx_ki


class KHSXKiModuleBoundaryTests(unittest.TestCase):
    def test_public_api_resolves_to_split_modules(self):
        self.assertEqual(
            khsx_ki.compute_standard_calendar_weeks.__module__,
            "planning.khsx_ki.calendar",
        )
        self.assertEqual(
            khsx_ki._detect_current_layout.__module__,
            "planning.khsx_ki.layout",
        )
        self.assertEqual(
            khsx_ki.patch_khsx_ki_workbook.__module__,
            "planning.khsx_ki.workbook",
        )
        self.assertEqual(
            khsx_ki.verify_khsx_ki.__module__,
            "planning.khsx_ki.verification",
        )

    def test_stable_facade_exports_core_api(self):
        expected = {
            "compute_month_weeks",
            "compute_standard_calendar_weeks",
            "get_layout_spec",
            "has_khsx_ki_sheet",
            "parse_week_columns",
            "patch_khsx_ki_workbook",
            "verify_khsx_ki",
        }
        self.assertTrue(expected <= set(khsx_ki.__all__))


if __name__ == "__main__":
    unittest.main()
