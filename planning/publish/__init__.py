"""Canonical Planning publish facade."""

from .constants import (
    AUDIT_REVISION_FILE,
    PROPOSAL_WORKBOOK_FILE,
    PUBLISH_DECISION_FILE,
    READY_PUBLISH_STATUSES,
    REVIEW_REQUIRED_STATUS,
    SOURCES,
    VOLATILE_XLSX_PARTS,
)
from .policy import (
    blocked_decision,
    publish_decision,
)
from .proposal import (
    proposal_id,
    proposal_output_sha256,
    save_proposal_artifacts,
    with_proposal_identity,
)
from .runner import main, parse_args
from .service import run_pipeline_with_retry
from .snapshot import read_snapshot
from .state import (
    detect_input_changes,
    save_publish_decision,
    save_states_after_success,
)

# Compatibility aliases for existing callers/tests.
_blocked_decision = blocked_decision
_parse_args = parse_args
_proposal_id = proposal_id
_proposal_output_sha256 = proposal_output_sha256
_publish_decision = publish_decision
_read_snapshot = read_snapshot
_save_proposal_artifacts = save_proposal_artifacts
_save_publish_decision = save_publish_decision
_save_states_after_success = save_states_after_success
_with_proposal_identity = with_proposal_identity

__all__ = [
    "AUDIT_REVISION_FILE",
    "PROPOSAL_WORKBOOK_FILE",
    "PUBLISH_DECISION_FILE",
    "READY_PUBLISH_STATUSES",
    "REVIEW_REQUIRED_STATUS",
    "SOURCES",
    "VOLATILE_XLSX_PARTS",
    "blocked_decision",
    "detect_input_changes",
    "main",
    "parse_args",
    "proposal_id",
    "proposal_output_sha256",
    "publish_decision",
    "read_snapshot",
    "run_pipeline_with_retry",
    "save_proposal_artifacts",
    "save_publish_decision",
    "save_states_after_success",
    "with_proposal_identity",
]
