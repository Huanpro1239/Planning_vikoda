"""Shared constants for Planning publish orchestration."""

from pathlib import Path

import stock


AUDIT_REVISION_FILE = Path("planning_input_revision.json")
PROPOSAL_WORKBOOK_FILE = Path("planning_proposal.xlsx")
PUBLISH_DECISION_FILE = Path("planning_publish_decision.json")
PRODUCTION_READINESS_REPORT_FILE = Path("production_readiness_report.json")
RELEASE_MANIFEST_FILE = Path("planning_release_manifest.json")
RELEASE_MANIFEST_DIR = Path("release_manifests")

READY_PUBLISH_STATUSES = {
    "ready_for_publish",
    "feasible",
}
REVIEW_REQUIRED_STATUS = "review_required"
VOLATILE_XLSX_PARTS = {"docProps/core.xml"}

SOURCES = {
    "actual_stock": stock.SOURCE_ACTUAL_PATH,
    "factory_vikoda": stock.SOURCE_FACTORY_VIKODA_PATH,
    "factory_vkd": stock.SOURCE_FACTORY_VKD_PATH,
    "accounting_vikoda": stock.SOURCE_ACCOUNTING_VIKODA_PATH,
    "accounting_vkd": stock.SOURCE_ACCOUNTING_VKD_PATH,
}
