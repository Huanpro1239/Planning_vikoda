import sync_planning_schedule_priority as priority
import sync_planning_schedule_priority_v8 as v8
from sync_planning_layout import canonicalize_planning_layout


def install_production_output_cleanup():
    """Install V8 scheduler and keep only real date columns in Ke_hoach_SX."""
    v8.install_priority_scheduler_v8()
    original_patch = priority.base.patch_schedule_workbook

    def patch_schedule_without_extra_columns(workbook_bytes, headers, products, schedule):
        updated, changed = original_patch(
            workbook_bytes,
            headers,
            products,
            schedule,
        )
        updated, layout_changes = canonicalize_planning_layout(
            updated,
            active_days=len(headers),
            reset_schedule=False,
        )
        return updated, changed + layout_changes

    priority.base.patch_schedule_workbook = patch_schedule_without_extra_columns


if __name__ == "__main__":
    install_production_output_cleanup()
    priority.base.main_with_retry()
