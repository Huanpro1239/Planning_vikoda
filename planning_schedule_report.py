"""Legacy compatibility facade for Planning schedule reports."""

from planning.schedule_report import (
    REPORT_PATH,
    attach_output_hash,
    json_safe,
    load_schedule_report,
    print_operational_report,
    save_schedule_report,
)

__all__ = [
    "REPORT_PATH",
    "json_safe",
    "attach_output_hash",
    "save_schedule_report",
    "load_schedule_report",
    "print_operational_report",
]
