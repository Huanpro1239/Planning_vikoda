import unittest

from planning import weekly_model
from planning.weekly_model import (
    inputs,
    policy,
    report,
    schedule,
    service,
    verification,
    workbook,
)


class WeeklyModelModuleBoundaryTests(unittest.TestCase):
    def test_public_api_resolves_to_split_modules(self):
        self.assertEqual(
            weekly_model.compute_planning_inputs_hash.__module__,
            "planning.weekly_model.inputs",
        )
        self.assertEqual(
            weekly_model.normalize_debt_mode.__module__,
            "planning.weekly_model.policy",
        )
        self.assertEqual(
            weekly_model.analyze_weekly_workbook.__module__,
            "planning.weekly_model.schedule",
        )
        self.assertEqual(
            weekly_model.patch_weekly_workbook.__module__,
            "planning.weekly_model.workbook",
        )
        self.assertEqual(
            weekly_model.build_weekly_schedule_report.__module__,
            "planning.weekly_model.report",
        )
        self.assertEqual(
            weekly_model.verify_weekly_workbook.__module__,
            "planning.weekly_model.verification",
        )
        self.assertEqual(
            weekly_model.prepare_weekly_schedule_update.__module__,
            "planning.weekly_model.service",
        )

    def test_weekly_engine_remains_separate(self):
        self.assertFalse(hasattr(schedule, "WeeklyInputRow"))
        self.assertTrue(callable(schedule.analyze_weekly_workbook))
        self.assertTrue(callable(service.prepare_weekly_schedule_update))
        self.assertTrue(callable(verification.verify_weekly_workbook))
        self.assertTrue(callable(report.build_weekly_schedule_report))
        self.assertTrue(callable(workbook.patch_weekly_workbook))
        self.assertTrue(callable(inputs.read_inputs))
        self.assertTrue(callable(policy.normalize_profile))


if __name__ == "__main__":
    unittest.main()
