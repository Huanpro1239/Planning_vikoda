import json
import unittest
from datetime import date, datetime
from io import BytesIO

from openpyxl import Workbook

from planning_schedule_report import build_schedule_report, json_safe


class ReportJsonSafeTests(unittest.TestCase):
    def test_converts_nested_date_and_datetime_after_pipeline_metadata_is_appended(self):
        report = {
            "pipeline": {
                "planning_changes": {
                    "130100006": {
                        "R": {
                            "before": datetime(2026, 9, 1, 8, 30),
                            "after": date(2026, 9, 2),
                        }
                    }
                }
            }
        }

        safe = json_safe(report)
        payload = json.dumps(safe, ensure_ascii=False)

        self.assertIn("2026-09-01T08:30:00", payload)
        self.assertIn("2026-09-02", payload)

    def test_late_completed_inherits_demand_deadline_from_matching_timeline_quantum(self):
        workbook = Workbook()
        workbook.active.title = "Ke_hoach_SX"
        output = BytesIO()
        workbook.save(output)

        due = date(2026, 9, 10)
        production_due = date(2026, 9, 9)
        info = {
            "headers": [date(2026, 9, 1)],
            "products": [],
            "schedule": {},
            "carryover": {},
            "line_capacity": {"KHS + PET 9000": 3.0},
            "utilization": {"KHS + PET 9000": 0.5},
            "inventory": {},
            "optimizer_meta": {
                "KHS + PET 9000": {
                    "timeline": [
                        {
                            "kind": "production",
                            "code": "130300005",
                            "unit_no": 1,
                            "demand_deadline_shift": 30.0,
                            "production_deadline_shift": 27.0,
                            "demand_due_date": due,
                            "production_deadline_date": production_due,
                        }
                    ],
                    "late_completed": [
                        {
                            "code": "130300005",
                            "unit_no": 1,
                            "qty": 3230,
                            "finish_shift": 28.125,
                            "demand_due_date": due,
                            "production_deadline_date": production_due,
                        }
                    ],
                }
            },
        }

        report = build_schedule_report(output.getvalue(), info)
        event = report["resources"]["KHS + PET 9000"]["meta"]["late_completed"][0]
        self.assertEqual(event["demand_deadline_shift"], 30.0)
        self.assertEqual(event["production_deadline_shift"], 27.0)


if __name__ == "__main__":
    unittest.main()
