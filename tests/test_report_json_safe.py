import json
import unittest
from datetime import date, datetime

from planning_schedule_report import json_safe


class ReportJsonSafeTests(unittest.TestCase):
    def test_converts_nested_date_and_datetime(self):
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


if __name__ == "__main__":
    unittest.main()
