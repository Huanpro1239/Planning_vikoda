import unittest
from datetime import date

from sync_planning_metrics_all_months import calculate_row_all_months
import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v8 as v8


class UrgentSupplyAndSharedMachineV8Tests(unittest.TestCase):
    def test_urgent_supply_ignores_target_stock_M(self):
        result = calculate_row_all_months(
            fc=1000,
            actual_stock=200,
            book_stock=200,
            opening_consignment=0,
            warehouse_debt=100,
            leadtime=3,
            batch=100,
            per_shift=100,
            shifts_per_day=3,
            classification="Không đường",
            plan_year=2026,
            plan_month=9,
        )
        # Urgent supply = FC + debt - J = 900. M must not be added.
        self.assertEqual(result["required_production"], 900)
        self.assertEqual(result["rounded_production"], 900)
        self.assertGreater(result["expected_end_stock"], 0)

    def _product(self, code, line, row):
        headers = [date(2026, 9, 1), date(2026, 9, 2)]
        return {
            "row": row,
            "code": code,
            "batch": 300,
            "per_shift": 100,
            "line": line,
            "product_group": line,
            "classification": "Không đường",
            "is_sugar": False,
            "max_shifts_per_day": 3.0,
            "actual_stock": 0.0,
            "fc": 300.0,
            "target_stock": 0.0,
            "debt": 0.0,
            "planned_qty": 300.0,
            "earliest_date": headers[0],
            "preferred_date": headers[0],
            "demand_by_day": {headers[0]: 150.0, headers[1]: 150.0},
            "quantum_qty": 100.0,
            "quantum_shift": 1.0,
            "required_units": 3,
            "existing_daily": [None] * 31,
        }

    def test_khs_and_pet9000_share_one_machine_capacity(self):
        headers = [date(2026, 9, 1), date(2026, 9, 2)]
        products = [
            self._product("KHS-A", "KHS", 2),
            self._product("PET-A", "PET 9000", 3),
        ]

        v8.install_priority_scheduler_v8()
        schedule, capacities, usage, carryover, _, meta = v8.build_schedule_shared_machine(
            headers,
            products,
        )

        self.assertEqual(capacities[v8.SHARED_MACHINE_NAME], 3.0)
        for current_day in headers:
            self.assertLessEqual(
                usage[v8.SHARED_MACHINE_NAME][current_day],
                3.0 + priority.base.EPSILON,
            )
        self.assertEqual(sum(schedule["KHS-A"].values()), 300.0)
        self.assertTrue(carryover or sum(schedule["PET-A"].values()) <= 300.0)
        self.assertEqual(meta[v8.SHARED_MACHINE_NAME]["mode"], "shared_machine_stockout_first")


if __name__ == "__main__":
    unittest.main()
