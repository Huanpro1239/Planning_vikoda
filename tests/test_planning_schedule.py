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
        actual_stock=1000,
        fc=0,
        target_stock=0,
    ):
        headers = [date(2026, 9, day) for day in range(1, 31)]
        demand_days = sum(1 for day in headers if day.weekday() != 6)
        daily = fc / demand_days if demand_days and fc else 0
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
            "actual_stock": float(actual_stock),
            "fc": float(fc),
            "target_stock": float(target_stock),
            "demand_by_day": {
                day: (0.0 if day.weekday() == 6 else daily)
                for day in headers
            },
        }
        _validate_quantum(item)
        return item

    def test_sunday_is_a_normal_production_day(self):
        sunday = date(2026, 9, 6)
        headers = [sunday, date(2026, 9, 7)]
        product = self.product(code="A", earliest=sunday, debt=10)
        product["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, _, carryover, _, _ = build_schedule(headers, [product])

        self.assertEqual(schedule["A"][sunday], 100)
        self.assertEqual(carryover, {})

    def test_never_schedules_before_R(self):
        headers = [date(2026, 9, 1), date(2026, 9, 2), date(2026, 9, 3)]
        product = self.product(code="A", earliest=date(2026, 9, 3))
        product["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, _, _, _, _ = build_schedule(headers, [product])

        self.assertEqual(schedule["A"][date(2026, 9, 1)], 0)
        self.assertEqual(schedule["A"][date(2026, 9, 2)], 0)
        self.assertEqual(schedule["A"][date(2026, 9, 3)], 100)

    def test_line_capacity_prevents_overlap(self):
        headers = [date(2026, 9, day) for day in range(1, 5)]
        a = self.product(code="A", planned_qty=200, shifts_per_day=1)
        b = self.product(code="B", planned_qty=200, shifts_per_day=1)
        for product in (a, b):
            product["demand_by_day"] = {day: 0 for day in headers}

        schedule, line_capacity, line_usage, carryover, _, _ = build_schedule(headers, [a, b])

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
        sumo["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, line_usage, carryover, _, _ = build_schedule(headers, [sumo])

        self.assertEqual(sum(schedule["130200026"].values()), 20800)
        for qty in schedule["130200026"].values():
            self.assertEqual(qty % 5200, 0)
        for used in line_usage["PET 9000"].values():
            self.assertLessEqual(used, 3.0 + 1e-9)
        self.assertEqual(carryover, {})

    def test_stockout_risk_has_priority_over_debt(self):
        headers = [date(2026, 9, 1)]
        urgent = self.product(code="A", actual_stock=0, fc=100, debt=0)
        debt = self.product(code="B", actual_stock=1000, fc=0, debt=10)
        urgent["demand_by_day"] = {headers[0]: 100}
        debt["demand_by_day"] = {headers[0]: 0}

        schedule, _, _, carryover, inventory, _ = build_schedule(headers, [urgent, debt])

        self.assertEqual(schedule["A"][headers[0]], 100)
        self.assertEqual(schedule["B"][headers[0]], 0)
        self.assertEqual(carryover["B"], 100)
        self.assertIsNone(inventory["A"]["first_stockout"])

    def test_debt_breaks_tie_when_inventory_risk_is_equal(self):
        headers = [date(2026, 9, 1)]
        normal = self.product(code="A", debt=0)
        debt = self.product(code="B", debt=10)
        for product in (normal, debt):
            product["demand_by_day"] = {headers[0]: 0}

        schedule, _, _, carryover, _, _ = build_schedule(headers, [normal, debt])

        self.assertEqual(schedule["B"][headers[0]], 100)
        self.assertEqual(schedule["A"][headers[0]], 0)
        self.assertEqual(carryover["A"], 100)

    def test_reports_carryover_instead_of_overbooking(self):
        headers = [date(2026, 9, 1)]
        product = self.product(code="A", planned_qty=200, shifts_per_day=1)
        product["demand_by_day"] = {headers[0]: 0}

        schedule, _, line_usage, carryover, _, _ = build_schedule(headers, [product])

        self.assertEqual(schedule["A"][headers[0]], 100)
        self.assertEqual(line_usage["KHS"][headers[0]], 1)
        self.assertEqual(carryover["A"], 100)

    def test_optimizer_repacks_fractional_batches_when_total_capacity_fits(self):
        headers = [date(2026, 9, 1), date(2026, 9, 2)]
        a = self.product(
            code="A",
            planned_qty=280,
            per_shift=400,
            batch=280,
            shifts_per_day=1,
            classification="Có đường",
        )
        b = self.product(
            code="B",
            planned_qty=520,
            per_shift=400,
            batch=520,
            shifts_per_day=2,
            classification="Có đường",
        )
        for product in (a, b):
            product["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, line_usage, carryover, _, _ = build_schedule(headers, [a, b])

        self.assertEqual(sum(schedule["A"].values()), 280)
        self.assertEqual(sum(schedule["B"].values()), 520)
        self.assertEqual(carryover, {})
        for used in line_usage["KHS"].values():
            self.assertLessEqual(used, 2.0 + 1e-9)

    def test_inventory_simulation_reports_plan_quantity_shortfall(self):
        headers = [date(2026, 9, 1), date(2026, 9, 2)]
        product = self.product(code="A", planned_qty=0, actual_stock=50, fc=100)
        product["demand_by_day"] = {headers[0]: 50, headers[1]: 50}

        _, _, _, _, inventory, _ = build_schedule(headers, [product])

        self.assertEqual(inventory["A"]["ending_stock"], -50)
        self.assertEqual(inventory["A"]["first_stockout"], date(2026, 9, 2))


if __name__ == "__main__":
    unittest.main()
