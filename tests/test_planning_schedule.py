import unittest
from datetime import date

from sync_planning_schedule import _validate_quantum, build_schedule


class PlanningScheduleTests(unittest.TestCase):
    def product(
        self,
        *,
        code,
        line="KHS",
        planned_qty=100,
        per_shift=100,
        batch=100,
        shifts_per_day=1,
        classification="Không đường",
        earliest=date(2026, 9, 1),
        debt=0,
        risk=date(2026, 9, 10),
    ):
        item = {
            "code": code,
            "line": line,
            "planned_qty": float(planned_qty),
            "per_shift": float(per_shift),
            "batch": float(batch),
            "max_shifts_per_day": float(shifts_per_day),
            "classification": classification,
            "is_sugar": classification.casefold() == "có đường".casefold(),
            "earliest_date": earliest,
            "debt": float(debt),
            "risk_date": risk,
        }
        _validate_quantum(item)
        return item

    def test_sunday_is_a_normal_production_day(self):
        sunday = date(2026, 9, 6)
        headers = [sunday, date(2026, 9, 7)]
        product = self.product(code="A", earliest=sunday)

        schedule, _, _, carryover = build_schedule(headers, [product])

        self.assertEqual(schedule["A"][sunday], 100)
        self.assertEqual(carryover, {})

    def test_never_schedules_before_R(self):
        headers = [date(2026, 9, 1), date(2026, 9, 2), date(2026, 9, 3)]
        product = self.product(code="A", earliest=date(2026, 9, 3))

        schedule, _, _, _ = build_schedule(headers, [product])

        self.assertEqual(schedule["A"][date(2026, 9, 1)], 0)
        self.assertEqual(schedule["A"][date(2026, 9, 2)], 0)
        self.assertEqual(schedule["A"][date(2026, 9, 3)], 100)

    def test_line_capacity_prevents_overlap(self):
        headers = [date(2026, 9, day) for day in range(1, 5)]
        a = self.product(code="A", planned_qty=200, shifts_per_day=1)
        b = self.product(code="B", planned_qty=200, shifts_per_day=1)

        schedule, line_capacity, line_usage, carryover = build_schedule(headers, [a, b])

        self.assertEqual(line_capacity["KHS"], 1)
        for day in headers:
            self.assertLessEqual(line_usage["KHS"][day], 1)
        self.assertEqual(sum(schedule["A"].values()), 200)
        self.assertEqual(sum(schedule["B"].values()), 200)
        self.assertEqual(carryover, {})

    def test_sugar_product_is_scheduled_by_batch_quantum(self):
        headers = [date(2026, 9, 14), date(2026, 9, 15)]
        sumo = self.product(
            code="130200026",
            line="PET 9000",
            planned_qty=20800,
            per_shift=4000,
            batch=5200,
            shifts_per_day=3,
            classification="Có đường",
            earliest=date(2026, 9, 14),
        )

        schedule, _, line_usage, carryover = build_schedule(headers, [sumo])

        self.assertEqual(schedule["130200026"][date(2026, 9, 14)], 10400)
        self.assertEqual(schedule["130200026"][date(2026, 9, 15)], 10400)
        self.assertAlmostEqual(line_usage["PET 9000"][date(2026, 9, 14)], 2.6)
        self.assertEqual(carryover, {})

    def test_debt_has_priority_when_capacity_is_tight(self):
        headers = [date(2026, 9, 1)]
        normal = self.product(code="A", debt=0, risk=date(2026, 9, 1))
        debt = self.product(code="B", debt=10, risk=date(2026, 9, 20))

        schedule, _, _, carryover = build_schedule(headers, [normal, debt])

        self.assertEqual(schedule["B"][headers[0]], 100)
        self.assertEqual(schedule["A"][headers[0]], 0)
        self.assertEqual(carryover["A"], 100)

    def test_reports_carryover_instead_of_overbooking(self):
        headers = [date(2026, 9, 1)]
        product = self.product(code="A", planned_qty=200, shifts_per_day=1)

        schedule, _, line_usage, carryover = build_schedule(headers, [product])

        self.assertEqual(schedule["A"][headers[0]], 100)
        self.assertEqual(line_usage["KHS"][headers[0]], 1)
        self.assertEqual(carryover["A"], 100)


if __name__ == "__main__":
    unittest.main()
