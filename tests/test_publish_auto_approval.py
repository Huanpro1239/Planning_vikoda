import unittest

from planning.publish.auto_approval import AUTO_APPROVE_ENV, AUTO_APPROVER
from planning.publish.policy import publish_decision


ENABLED = {AUTO_APPROVE_ENV: "1"}


def _report(
    *,
    stockout=("130100011", "130100017"),
    balanced=("130100011", "130100017"),
    resource_ok=True,
    monthly_state="capacity_constrained_balanced",
):
    return {
        "plan_month": "2026-10",
        "proposal_id": "proposal-abc",
        "publish_status": "review_required",
        "mass_balance": {
            "130100011": {"service_carryover_qty": 29700, "uom": "Thùng"},
            "130100017": {"service_carryover_qty": 12000, "uom": "Thùng"},
        },
        "status": {
            "monthly_quantity": {"ok": False, "state": monthly_state},
            "resource_validation": {"ok": resource_ok},
            "service": {
                "ok": False,
                "state": "stockout_risk",
                "stockout_skus": list(stockout),
                "capacity_balanced_skus": list(balanced),
            },
        },
    }


class PublishAutoApprovalTests(unittest.TestCase):
    def test_capacity_shortfall_is_auto_approved_with_note(self):
        ok, decision = publish_decision(_report(), None, "daily-schedule", ENABLED)

        self.assertTrue(ok)
        self.assertEqual(decision["basis"], "review_approval")
        approval = decision["review_approval"]
        self.assertTrue(approval["auto"])
        self.assertEqual(approval["approved_by"], AUTO_APPROVER)
        self.assertEqual(approval["proposal_id"], "proposal-abc")
        self.assertIn("130100011 thiếu 29.700 Thùng", approval["reason"])
        self.assertIn("130100017 thiếu 12.000 Thùng", approval["reason"])
        self.assertIn("2026-10", approval["reason"])

    def test_disabled_by_default_stays_blocked(self):
        ok, decision = publish_decision(_report(), None, "daily-schedule", {})

        self.assertFalse(ok)
        self.assertEqual(decision["state"], "blocked")

    def test_shortfall_not_caused_by_capacity_stays_blocked(self):
        report = _report(balanced=("130100011",))

        ok, decision = publish_decision(report, None, "daily-schedule", ENABLED)

        self.assertFalse(ok)
        self.assertEqual(decision["state"], "blocked")

    def test_resource_failure_is_never_auto_approved(self):
        report = _report(resource_ok=False)

        ok, decision = publish_decision(report, None, "daily-schedule", ENABLED)

        self.assertFalse(ok)
        self.assertEqual(decision["basis"], "non_waivable_validation")

    def test_carryover_is_never_auto_approved(self):
        report = _report(monthly_state="carryover")

        ok, decision = publish_decision(report, None, "daily-schedule", ENABLED)

        self.assertFalse(ok)
        self.assertEqual(decision["basis"], "non_waivable_validation")

    def test_manual_approval_takes_precedence(self):
        manual = {
            "proposal_id": "proposal-abc",
            "reason": "Planner duyệt tay.",
            "approved_by": "planner",
        }

        ok, decision = publish_decision(_report(), manual, "planner", ENABLED)

        self.assertTrue(ok)
        self.assertFalse(decision["review_approval"]["auto"])
        self.assertEqual(decision["review_approval"]["reason"], "Planner duyệt tay.")


if __name__ == "__main__":
    unittest.main()
