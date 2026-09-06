import unittest
from datetime import date, timedelta

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
        shifts_per_day=2,
        classification="Không đường",
        earliest=date(2026, 9, 1),
        debt=0,
        actual_stock=1000,
        fc=0,
        target_stock=0,
        group="G",
        row=1,
    ):
        headers = [date(2026, 9, 1) + timedelta(days=i) for i in range(30)]
        demand_days = sum(1 for day in headers if day.weekday() != 6)
        daily = fc / demand_days if demand_days and fc else 0
        item = {
            "row": row,
            "code": code,
            "line": line,
            "product_group": group,
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

    def test_khs_campaign_is_continuous_and_setup_consumes_half_shift(self):
        headers = [date(2026, 9, 1) + timedelta(days=i) for i in range(4)]
        a = self.product(code="A", planned_qty=200, shifts_per_day=2, row=1)
        b = self.product(code="B", planned_qty=100, shifts_per_day=2, row=2)
        for product in (a, b):
            product["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, line_usage, carryover, _, meta = build_schedule(headers, [a, b])

        self.assertEqual(sum(schedule["A"].values()), 200)
        self.assertEqual(sum(schedule["B"].values()), 100)
        self.assertEqual(carryover, {})
        self.assertEqual(meta["KHS"]["mode"], "continuous_campaign")
        self.assertAlmostEqual(meta["KHS"]["setup_shifts"], 0.5)
        self.assertEqual(schedule["A"][headers[0]], 200)
        self.assertEqual(schedule["B"][headers[1]], 100)
        self.assertAlmostEqual(line_usage["KHS"][headers[1]], 1.5)

    def test_never_schedules_before_R(self):
        headers = [date(2026, 9, 1), date(2026, 9, 2), date(2026, 9, 3)]
        product = self.product(code="A", earliest=date(2026, 9, 3))
        product["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, _, _, _, _ = build_schedule(headers, [product])

        self.assertEqual(schedule["A"][date(2026, 9, 1)], 0)
        self.assertEqual(schedule["A"][date(2026, 9, 2)], 0)
        self.assertEqual(schedule["A"][date(2026, 9, 3)], 100)

    def test_khs_uses_one_block_per_sku_instead_of_fragmenting(self):
        headers = [date(2026, 9, 1) + timedelta(days=i) for i in range(8)]
        product = self.product(
            code="A",
            planned_qty=700,
            per_shift=100,
            batch=100,
            shifts_per_day=2,
        )
        product["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, _, carryover, _, _ = build_schedule(headers, [product])

        positive_indexes = [
            i for i, day in enumerate(headers) if schedule["A"][day] > 0
        ]
        self.assertEqual(positive_indexes, list(range(min(positive_indexes), max(positive_indexes) + 1)))
        self.assertEqual(sum(schedule["A"].values()), 700)
        self.assertEqual(carryover, {})

    def test_sunday_is_normal_for_continuous_campaign(self):
        sunday = date(2026, 9, 6)
        headers = [sunday, date(2026, 9, 7)]
        product = self.product(code="A", earliest=sunday)
        product["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, _, carryover, _, _ = build_schedule(headers, [product])

        self.assertEqual(schedule["A"][sunday], 100)
        self.assertEqual(carryover, {})

    def test_rgb_is_split_by_month_weeks(self):
        headers = [date(2026, 9, 1) + timedelta(days=i) for i in range(28)]
        product = self.product(
            code="RGB1",
            line="RGB",
            planned_qty=400,
            per_shift=100,
            batch=100,
            shifts_per_day=2,
        )
        product["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, _, carryover, _, meta = build_schedule(headers, [product])

        weekly_totals = [
            sum(schedule["RGB1"][day] for day in headers[start:start + 7])
            for start in range(0, 28, 7)
        ]
        self.assertEqual(weekly_totals, [100, 100, 100, 100])
        self.assertEqual(carryover, {})
        self.assertEqual(meta["RGB"]["mode"], "weekly_rgb")

    def test_galon_19l_is_spread_across_workdays(self):
        headers = [date(2026, 9, 1) + timedelta(days=i) for i in range(7)]
        product = self.product(
            code="130100006",
            line="Galon",
            planned_qty=600,
            per_shift=100,
            batch=100,
            shifts_per_day=2,
        )
        product["demand_by_day"] = {day: 0 for day in headers}

        schedule, _, _, carryover, _, meta = build_schedule(headers, [product])

        self.assertEqual(schedule["130100006"][date(2026, 9, 6)], 0)
        for day in headers:
            if day.weekday() != 6:
                self.assertEqual(schedule["130100006"][day], 100)
        self.assertEqual(carryover, {})
        self.assertEqual(meta["Galon"]["mode"], "galon_hybrid")

    def test_sugar_campaign_keeps_total_and_capacity(self):
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

        self.assertAlmostEqual(sum(schedule["130200026"].values()), 20800)
        for used in line_usage["PET 9000"].values():
            self.assertLessEqual(used, 3.0 + 1e-9)
        self.assertEqual(carryover, {})

    def test_reports_carryover_instead_of_overbooking(self):
        headers = [date(2026, 9, 1)]
        product = self.product(
            code="A",
            planned_qty=200,
            per_shift=100,
            shifts_per_day=1,
        )
        product["demand_by_day"] = {headers[0]: 0}

        schedule, _, line_usage, carryover, _, _ = build_schedule(headers, [product])

        self.assertEqual(schedule["A"][headers[0]], 100)
        self.assertEqual(line_usage["KHS"][headers[0]], 1)
        self.assertEqual(carryover["A"], 100)

    def test_inventory_simulation_reports_plan_quantity_shortfall(self):
        headers = [date(2026, 9, 1), date(2026, 9, 2)]
        product = self.product(code="A", planned_qty=0, actual_stock=50, fc=100)
        product["demand_by_day"] = {headers[0]: 50, headers[1]: 50}

        _, _, _, _, inventory, _ = build_schedule(headers, [product])

        self.assertEqual(inventory["A"]["ending_stock"], -50)
        self.assertEqual(inventory["A"]["first_stockout"], date(2026, 9, 2))


    def test_inventory_simulation_charges_debt_once_on_first_day(self):
        headers = [date(2026, 9, 1), date(2026, 9, 2)]
        product = self.product(
            code="DEBT",
            planned_qty=0,
            actual_stock=0,
            fc=0,
            debt=100,
        )
        product["demand_by_day"] = {day: 0 for day in headers}

        _, _, _, _, inventory, _ = build_schedule(headers, [product])
        result = inventory["DEBT"]

        self.assertEqual(result["ending_stock"], -100)
        self.assertEqual(result["ending_net_available"], -100)
        self.assertEqual(result["ending_physical_stock"], 0)
        self.assertEqual(result["ending_backlog"], 100)
        self.assertEqual(result["first_stockout"], headers[0])
        self.assertEqual(result["debt_due_date"], headers[0])


if __name__ == "__main__":
    unittest.main()
