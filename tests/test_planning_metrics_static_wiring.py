import inspect
import unittest

from planning import metrics


class PlanningMetricsStaticWiringTests(unittest.TestCase):
    def test_calculate_metrics_requires_explicit_leadtimes(self):
        parameters = inspect.signature(metrics.calculate_metrics).parameters
        self.assertIn("leadtimes", parameters)

    def test_no_mutable_global_leadtime_mapping(self):
        self.assertFalse(hasattr(metrics, "LEADTIME_BY_CODE"))

    def test_production_calculation_is_canonical_module_function(self):
        self.assertEqual(
            metrics.calculate_row.__module__,
            "planning.metrics.calculation",
        )
        self.assertEqual(
            metrics.calculate_metrics.__module__,
            "planning.metrics.calculation",
        )

    def test_cross_year_resolution_is_static(self):
        from datetime import date

        self.assertEqual(metrics.resolve_plan_year(date(2026, 12, 1), 1), 2027)
        self.assertEqual(metrics.resolve_plan_year(date(2026, 1, 1), 12), 2025)


if __name__ == "__main__":
    unittest.main()
