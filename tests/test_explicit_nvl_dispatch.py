"""Contract tests for explicit GITHUB_TOKEN repository_dispatch pairing."""
import unittest
from pathlib import Path
from scripts.dispatch_planning_nvl import make_payload
from scripts.verify_planning_dispatch import validate_release


def _manifest():
    return {
        "release_id": "release-100",
        "commit": {"run_id": "100", "sha": "a" * 40},
        "publish_decision": {"state": "published"},
        "readiness": {"status": "passed"},
    }


class ExplicitNVLDispatchTests(unittest.TestCase):
    def test_valid_dispatch_contains_exact_provenance(self):
        payload = make_payload(
            _manifest(), run_id="100", head_sha="a" * 40,
            event="repository_dispatch",
        )
        self.assertEqual(payload["event_type"], "planning_nvl_published")
        self.assertEqual(payload["client_payload"]["planning_run_id"], "100")
        self.assertEqual(payload["client_payload"]["planning_head_sha"], "a" * 40)
        self.assertEqual(payload["client_payload"]["planning_event"], "repository_dispatch")
        self.assertEqual(payload["client_payload"]["planning_release_id"], "release-100")

    def test_invalid_release_fails_closed(self):
        for field, value in (("run_id", "101"), ("sha", "b" * 40)):
            with self.subTest(field=field):
                manifest = _manifest()
                manifest["commit"][field] = value
                with self.assertRaises(ValueError):
                    make_payload(
                        manifest, run_id="100", head_sha="a" * 40,
                        event="repository_dispatch",
                    )
        manifest = _manifest()
        manifest["publish_decision"]["state"] = "proposal_only"
        with self.assertRaises(ValueError):
            make_payload(
                manifest, run_id="100", head_sha="a" * 40,
                event="repository_dispatch",
            )

    def test_nvl_dispatch_requires_exact_immutable_release(self):
        validate_release(
            _manifest(), planning_run_id="100", head_sha="a" * 40,
            event="repository_dispatch", release_id="release-100",
        )
        for key, values in (
            ("planning_run_id", "101"),
            ("head_sha", "b" * 40),
            ("event", "manual"),
            ("release_id", "missing"),
        ):
            with self.subTest(key=key):
                args = {
                    "planning_run_id": "100", "head_sha": "a" * 40,
                    "event": "repository_dispatch", "release_id": "release-100",
                }
                args[key] = values
                with self.assertRaises(ValueError):
                    validate_release(_manifest(), **args)

    def test_workflow_guards_and_health_monitor_remain(self):
        root = Path(".github/workflows")
        planning = (root / "sync-stock.yml").read_text(encoding="utf-8")
        nvl = (root / "sync-nvl-stock.yml").read_text(encoding="utf-8")
        health = (root / "paired-run-health.yml").read_text(encoding="utf-8")
        self.assertIn("scripts/dispatch_planning_nvl.py", planning)
        self.assertIn("github.event_name == 'repository_dispatch'", planning)
        self.assertIn("planning_nvl_published", nvl)
        self.assertIn("github.event.workflow_run.event == 'schedule'", nvl)
        self.assertNotIn(
            "github.event.workflow_run.event == 'repository_dispatch'", nvl
        )
        self.assertIn("scripts/verify_planning_dispatch.py", nvl)
        self.assertIn("NVL_UPSTREAM_PLANNING_RELEASE_ID", nvl)
        self.assertIn("planning_nvl_completed", health)
        self.assertIn("scripts/check_paired_runs.py", health)


if __name__ == "__main__":
    unittest.main()
