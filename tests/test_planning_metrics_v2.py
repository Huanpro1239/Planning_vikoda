import unittest

from sync_planning_metrics_v2 import demand_days_in_month


class PlanningMetricsV2Tests(unittest.TestCase):
    def test_september_2026_has_26_demand_days(self):
        self.assertEqual(demand_days_in_month(2026, 9), 26)

    def test_february_2027_is_dynamic_not_hardcoded_26(self):
        self.assertEqual(demand_days_in_month(2027, 2), 24)

    def test_february_2028_leap_year(self):
        self.assertEqual(demand_days_in_month(2028, 2), 25)

    def test_month_with_27_non_sundays(self):
        self.assertEqual(demand_days_in_month(2026, 8), 26)


if __name__ == "__main__":
    unittest.main()
