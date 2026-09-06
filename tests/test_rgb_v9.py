import unittest
from datetime import date, timedelta

import sync_planning_schedule as base
from planning_resource_timeline import validate_resource_timeline
from sync_planning_schedule_production import install_production_output_cleanup
import sync_planning_schedule_rgb_v9 as rgb_v9
import sync_planning_schedule_rgb_v10 as rgb_v10


class RGBV9Tests(unittest.TestCase):
    @staticmethod
    def _product(code, row, *, planned=100.0, per_shift=100.0, batch=100.0, sugar=False, debt=0.0, stock=1000.0):
        product = {
            "row": row,
            "code": code,
            "line": "RGB",
            "product_group": code,
            "classification": "Có đường" if sugar else "Không đường",
            "is_sugar": sugar,
            "batch": batch,
            "per_shift": per_shift,
            "max_shifts_per_day": 2.0,
            "actual_stock": stock,
            "fc": 0.0,
            "target_stock": 0.0,
            "debt": debt,
            "planned_qty": planned,
            "earliest_date": None,
        }
        product["quantum_qty"] = batch if sugar else per_shift
        product["quantum_shift"] = product["quantum_qty"] / per_shift
        product["required_units"] = int(round(planned / product["quantum_qty"]))
        return product

    def test_production_installer_keeps_rgb_v10_last(self):
        install_production_output_cleanup()
        install_production_output_cleanup()
        self.assertIs(base._allocate_weekly_rgb_line, rgb_v10.allocate_rgb_quantum_lookahead)

    def test_rgb_v9_reserves_setup_in_capacity_and_emits_valid_timeline(self):
        headers = [date(2026, 9, 1) + timedelta(days=i) for i in range(2)]
        products = [
            self._product("A", 2),
            self._product("B", 3),
        ]
        for product in products:
            product["demand_by_day"] = {day: 0.0 for day in headers}

        schedule, usage, carryover, meta = rgb_v9.allocate_rgb_quantized_service(
            headers,
            products,
            2.0,
        )
        self.assertEqual(carryover, {})
        self.assertTrue(meta["timeline"])
        self.assertAlmostEqual(meta["setup_shifts"], 0.5)
        self.assertLessEqual(max(usage.values()), 2.0 + 1e-9)

        validation = validate_resource_timeline(
            headers,
            products,
            schedule,
            2.0,
            "RGB",
            meta["timeline"],
        )
        self.assertTrue(validation["validated"])
        self.assertAlmostEqual(validation["setup_shifts"], 0.5)

    def test_rgb_v9_never_splits_sugar_batch_across_days(self):
        headers = [date(2026, 9, 1) + timedelta(days=i) for i in range(3)]
        sugar = self._product(
            "S",
            2,
            planned=6500.0,
            per_shift=4500.0,
            batch=3250.0,
            sugar=True,
            stock=0.0,
        )
        sugar["demand_by_day"] = {day: 1000.0 for day in headers}

        schedule, usage, carryover, meta = rgb_v9.allocate_rgb_quantized_service(
            headers,
            [sugar],
            2.0,
        )
        self.assertEqual(carryover, {})
        for qty in schedule["S"].values():
            if qty > 1e-9:
                self.assertAlmostEqual(qty / 3250.0, round(qty / 3250.0))
        self.assertTrue(all(value <= 2.0 + 1e-9 for value in usage.values()))

    def test_rgb_v9_service_objective_puts_opening_debt_before_safe_sku(self):
        headers = [date(2026, 9, 1) + timedelta(days=i) for i in range(2)]
        urgent = self._product("URG", 2, debt=150.0, stock=0.0)
        safe = self._product("SAFE", 3, debt=0.0, stock=1000.0)
        for product in (urgent, safe):
            product["demand_by_day"] = {day: 0.0 for day in headers}

        _, _, _, meta = rgb_v9.allocate_rgb_quantized_service(
            headers,
            [safe, urgent],
            2.0,
        )
        self.assertEqual(meta["sequence"][0], "URG")


if __name__ == "__main__":
    unittest.main()
