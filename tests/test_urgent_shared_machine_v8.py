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
        self.assertEqual(result["required_production"], 900)
        self.assertEqual(result["rounded_production"], 900)
        self.assertGreater(result["expected_end_stock"], 0)

    def _product(
        self,
        code,
        line,
        row,
        *,
        actual_stock=0.0,
        daily=(150.0, 150.0),
        planned_qty=300.0,
    ):
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
            "actual_stock": actual_stock,
            "fc": sum(daily),
            "target_stock": 0.0,
            "debt": 0.0,
            "planned_qty": planned_qty,
            "earliest_date": headers[0],
            "preferred_date": headers[0],
            "demand_by_day": {headers[0]: daily[0], headers[1]: daily[1]},
            "quantum_qty": 100.0,
            "quantum_shift": 1.0,
            "required_units": int(planned_qty / 100),
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

        # 6 production shifts + 0.5 setup cannot fit into a 6-shift horizon.
        # The scheduler must report physical carryover instead of overbooking.
        for product in products:
            produced = sum(schedule[product["code"]].values())
            missing = float(carryover.get(product["code"], 0) or 0)
            self.assertAlmostEqual(produced + missing, product["planned_qty"])
        self.assertTrue(carryover)
        self.assertEqual(
            meta[v8.SHARED_MACHINE_NAME]["mode"],
            "shared_machine_deadline_guarded",
        )

    def test_urgent_deadline_can_interrupt_longer_campaign(self):
        headers = [date(2026, 9, 1), date(2026, 9, 2)]
        # KHS-A can wait until day 2. PET-A has zero opening stock and needs
        # one quantum on day 1, so PET-A must not be sacrificed behind KHS-A.
        products = [
            self._product(
                "KHS-A",
                "KHS",
                2,
                actual_stock=300.0,
                daily=(0.0, 300.0),
                planned_qty=300.0,
            ),
            self._product(
                "PET-A",
                "PET 9000",
                3,
                actual_stock=0.0,
                daily=(100.0, 0.0),
                planned_qty=100.0,
            ),
        ]

        v8.install_priority_scheduler_v8()
        schedule, _, usage, _, inventory, meta = v8.build_schedule_shared_machine(
            headers,
            products,
        )

        self.assertGreater(schedule["PET-A"][headers[0]], 0)
        self.assertIsNone(inventory["PET-A"]["first_stockout"])
        self.assertFalse(meta[v8.SHARED_MACHINE_NAME]["deadline_misses"])
        for current_day in headers:
            self.assertLessEqual(
                usage[v8.SHARED_MACHINE_NAME][current_day],
                3.0 + priority.base.EPSILON,
            )


    def test_tries_smaller_feasible_candidate_before_stopping(self):
        headers = [date(2026, 9, 1)]
        a = self._product(
            "A",
            "KHS",
            2,
            actual_stock=0.0,
            daily=(400.0, 0.0),
            planned_qty=400.0,
        )
        b = self._product(
            "B",
            "PET 9000",
            3,
            actual_stock=0.0,
            daily=(100.0, 0.0),
            planned_qty=100.0,
        )
        # One A batch takes 2 shifts; B takes 0.5 shift. After A=200,
        # setup 0.5 + B 0.5 still fits the 3-shift horizon.
        a.update(
            batch=200.0,
            per_shift=100.0,
            is_sugar=True,
            quantum_qty=200.0,
            quantum_shift=2.0,
            required_units=2,
            demand_by_day={headers[0]: 400.0},
        )
        b.update(
            batch=100.0,
            per_shift=200.0,
            is_sugar=True,
            quantum_qty=100.0,
            quantum_shift=0.5,
            required_units=1,
            demand_by_day={headers[0]: 100.0},
        )

        schedule, usage, carryover, meta = v8.allocate_shared_deadline_guarded(
            headers, [a, b], 3.0
        )

        self.assertEqual(sum(schedule["A"].values()), 200)
        self.assertEqual(sum(schedule["B"].values()), 100)
        self.assertAlmostEqual(sum(usage.values()), 3.0)
        self.assertEqual(carryover, {"A": 200})
        self.assertEqual(meta["urgent_carryover"], {"A": 200})
        self.assertEqual(meta["safety_carryover"], {})
        self.assertEqual(meta["unserved_due"][0]["code"], "A")
        self.assertEqual(meta["unserved_due"][0]["qty"], 200)
        self.assertEqual(meta["unserved_due"][0]["due_date"], headers[0])


if __name__ == "__main__":
    unittest.main()
