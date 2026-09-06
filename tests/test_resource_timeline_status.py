import unittest
from datetime import date
from io import BytesIO

from openpyxl import Workbook

from planning_resource_timeline import (
    build_daily_serial_timeline,
    validate_resource_timeline,
)
from planning_schedule_report import build_schedule_report


class ResourceTimelineTests(unittest.TestCase):
    def _products(self):
        return [
            {
                "row": 2,
                "code": "A",
                "line": "RGB",
                "per_shift": 100.0,
                "batch": 100.0,
                "is_sugar": False,
                "planned_qty": 100.0,
            },
            {
                "row": 3,
                "code": "B",
                "line": "RGB",
                "per_shift": 100.0,
                "batch": 100.0,
                "is_sugar": False,
                "planned_qty": 50.0,
            },
        ]

    def test_rgb_timeline_inserts_every_switch_and_matches_daily_schedule(self):
        headers = [date(2026, 9, 1)]
        products = self._products()
        schedule = {
            "A": {headers[0]: 100.0},
            "B": {headers[0]: 50.0},
        }
        timeline, setup, usage = build_daily_serial_timeline(
            headers,
            products,
            schedule,
            2.0,
            "RGB",
        )
        self.assertEqual(setup, 0.5)
        self.assertAlmostEqual(usage[headers[0]], 2.0)
        self.assertEqual([event["kind"] for event in timeline], ["production", "setup", "production"])

        result = validate_resource_timeline(
            headers,
            products,
            schedule,
            2.0,
            "RGB",
            timeline,
        )
        self.assertTrue(result["validated"])
        self.assertAlmostEqual(result["setup_shifts"], 0.5)

    def test_rgb_timeline_rejects_daily_plan_that_cannot_fit_required_setup(self):
        headers = [date(2026, 9, 1)]
        products = self._products()
        products[1]["planned_qty"] = 100.0
        schedule = {
            "A": {headers[0]: 100.0},
            "B": {headers[0]: 100.0},
        }
        with self.assertRaisesRegex(RuntimeError, "không có thứ tự vận hành hợp lệ"):
            build_daily_serial_timeline(
                headers,
                products,
                schedule,
                2.0,
                "RGB",
            )


class SplitStatusTests(unittest.TestCase):
    @staticmethod
    def _workbook_bytes():
        wb = Workbook()
        ws = wb.active
        ws.title = "Ke_hoach_SX"
        ws.append(["Mã Sản Phẩm", "Tên", "ĐVT"])
        ws.append(["A", "A", "Thùng"])
        ws.append(["B", "B", "Thùng"])
        out = BytesIO()
        wb.save(out)
        return out.getvalue()

    def test_report_does_not_call_monthly_quantity_complete_plan_feasible_when_stockout_remains(self):
        day = date(2026, 9, 1)
        products = [
            {
                "row": 2,
                "code": "A",
                "line": "RGB",
                "per_shift": 100.0,
                "batch": 100.0,
                "is_sugar": False,
                "planned_qty": 100.0,
            },
            {
                "row": 3,
                "code": "B",
                "line": "RGB",
                "per_shift": 100.0,
                "batch": 100.0,
                "is_sugar": False,
                "planned_qty": 50.0,
            },
        ]
        info = {
            "headers": [day],
            "products": products,
            "schedule": {
                "A": {day: 100.0},
                "B": {day: 50.0},
            },
            "carryover": {},
            "line_capacity": {"RGB": 2.0},
            "optimizer_meta": {"RGB": {"setup_shifts": 0.5}},
            "utilization": {"RGB": 0.75},
            "inventory": {
                "A": {"first_stockout": day},
                "B": {"first_stockout": None},
            },
        }

        report = build_schedule_report(
            self._workbook_bytes(),
            info,
            input_revision={"etag": "x"},
        )
        self.assertEqual(report["status"]["monthly_quantity"]["state"], "complete")
        self.assertEqual(report["status"]["resource_validation"]["state"], "passed")
        self.assertEqual(report["status"]["service"]["state"], "stockout_risk")
        self.assertEqual(report["publish_status"], "review_required")
        self.assertNotEqual(report["publish_status"], "feasible")
        self.assertEqual(report["resources"]["RGB"]["meta"]["setup_shifts"], 0.5)
        self.assertTrue(report["resources"]["RGB"]["meta"]["resource_validation"]["validated"])


if __name__ == "__main__":
    unittest.main()
