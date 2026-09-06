import unittest
from datetime import date, timedelta
from unittest import mock

from planning_resource_timeline import validate_resource_timeline
from review_rgb_service import _day1_protection_bound
import sync_planning_schedule as base
import sync_planning_schedule_production as production
import sync_planning_schedule_rgb_v9 as v9
import sync_planning_schedule_rgb_v10 as v10


class RGBV10Tests(unittest.TestCase):
    @staticmethod
    def _products():
        headers = [date(2026, 9, 1) + timedelta(days=i) for i in range(3)]
        products = []
        for row, code in enumerate(("A", "B"), start=2):
            products.append(
                {
                    "row": row,
                    "code": code,
                    "uom": "Thùng",
                    "batch": 100.0,
                    "per_shift": 200.0,
                    "line": "RGB",
                    "product_group": code,
                    "classification": "Không đường",
                    "is_sugar": False,
                    "max_shifts_per_day": 2.0,
                    "actual_stock": 0.0,
                    "fc": 300.0,
                    "target_stock": 0.0,
                    "debt": 0.0,
                    "planned_qty": 300.0,
                    "required_units": 3,
                    "quantum_qty": 100.0,
                    "quantum_shift": 0.5,
                    "earliest_date": headers[0],
                    "preferred_date": headers[0],
                    "demand_by_day": {day: 100.0 for day in headers},
                    "existing_daily": [None] * 31,
                }
            )
        return headers, products

    def test_quantum_lookahead_can_remove_stockout_that_full_campaign_cannot(self):
        headers, products = self._products()

        schedule_v9, _, carry_v9, _ = v9.allocate_rgb_quantized_service(
            headers, products, 2.0
        )
        schedule_v10, _, carry_v10, meta_v10 = v10.allocate_rgb_quantum_lookahead(
            headers, products, 2.0
        )

        inv_v9 = base.simulate_inventory(headers, products, schedule_v9)
        inv_v10 = base.simulate_inventory(headers, products, schedule_v10)
        stockout_v9 = sum(1 for item in inv_v9.values() if item["first_stockout"])
        stockout_v10 = sum(1 for item in inv_v10.values() if item["first_stockout"])

        self.assertEqual(carry_v9, {})
        self.assertEqual(carry_v10, {})
        self.assertGreater(stockout_v9, stockout_v10)
        self.assertEqual(stockout_v10, 0)
        self.assertEqual(meta_v10["mode"], "rgb_quantum_lookahead_v10")
        self.assertGreater(meta_v10["setup_shifts"], 0)

        validation = validate_resource_timeline(
            headers,
            products,
            schedule_v10,
            2.0,
            "RGB",
            meta_v10["timeline"],
        )
        self.assertTrue(validation["validated"])
        self.assertLessEqual(validation["max_daily_usage"], 2.0 + 1e-7)

    def test_search_scores_incrementally_without_replaying_each_timeline(self):
        headers, products = self._products()

        with mock.patch.object(
            v10,
            "_schedule_from_timeline",
            wraps=v10._schedule_from_timeline,
        ) as replay:
            _, _, _, meta = v10.allocate_rgb_quantum_lookahead(
                headers,
                products,
                2.0,
            )

        self.assertEqual(replay.call_count, 1)
        self.assertTrue(meta["incremental_scoring"])

    def test_day1_lower_bound_accounts_for_quantum_capacity_and_setup(self):
        headers = [date(2026, 9, 1)]
        products = [
            {
                "code": "UNAVOIDABLE",
                "actual_stock": 0.0,
                "debt": 250.0,
                "demand_by_day": {headers[0]: 0.0},
                "quantum_qty": 100.0,
                "quantum_shift": 1.0,
                "earliest_date": headers[0],
            },
            {
                "code": "B",
                "actual_stock": 0.0,
                "debt": 100.0,
                "demand_by_day": {headers[0]: 0.0},
                "quantum_qty": 100.0,
                "quantum_shift": 1.0,
                "earliest_date": headers[0],
            },
            {
                "code": "C",
                "actual_stock": 0.0,
                "debt": 100.0,
                "demand_by_day": {headers[0]: 0.0},
                "quantum_qty": 100.0,
                "quantum_shift": 1.0,
                "earliest_date": headers[0],
            },
        ]

        bound = _day1_protection_bound(headers, products, 2.0)

        self.assertEqual(bound["at_risk_codes"], ["B", "C", "UNAVOIDABLE"])
        self.assertEqual(bound["inherently_unprotectable_codes"], ["UNAVOIDABLE"])
        self.assertEqual(bound["protectable_codes"], ["B", "C"])
        self.assertEqual(bound["max_protected_skus"], 1)
        self.assertEqual(bound["minimum_stockout_skus"], 2)
        self.assertEqual(bound["best_protected_sets"], [["B"], ["C"]])

    def test_production_installer_enables_v10_hook(self):
        production.install_production_output_cleanup()
        self.assertIs(base._allocate_weekly_rgb_line, v10.allocate_rgb_quantum_lookahead)


if __name__ == "__main__":
    unittest.main()
