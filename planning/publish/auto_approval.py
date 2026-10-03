"""Tự duyệt stockout_risk do giới hạn công suất.

Bật bằng PLANNING_AUTO_APPROVE_CAPACITY_SHORTFALL (quyết định của chủ hệ
thống). Chỉ áp dụng khi mọi SKU thiếu Service đều do hết công suất.
"""

import os


AUTO_APPROVE_ENV = "PLANNING_AUTO_APPROVE_CAPACITY_SHORTFALL"
AUTO_APPROVER = "auto-policy:capacity-shortfall"


def auto_approval_enabled(environ=None):
    env = os.environ if environ is None else environ
    return str(env.get(AUTO_APPROVE_ENV, "")).strip().lower() in {
        "1",
        "true",
        "yes",
    }


def _fmt_qty(value):
    return f"{value:,.0f}".replace(",", ".")


def build_auto_approval(report):
    """Approval tự động cho thiếu hàng do giới hạn công suất.

    Chỉ trả về approval khi mọi SKU thiếu Service đều nằm trong
    capacity_balanced_skus, tức engine đã xếp hết công suất mà vẫn thiếu.
    """
    service = (report.get("status") or {}).get("service") or {}
    stockout = [str(code) for code in (service.get("stockout_skus") or [])]
    balanced = {
        str(code) for code in (service.get("capacity_balanced_skus") or [])
    }
    if not stockout or not set(stockout) <= balanced:
        return None

    mass_balance = report.get("mass_balance") or {}
    lines = []
    for code in sorted(stockout):
        values = mass_balance.get(code) or {}
        qty = float(values.get("service_carryover_qty", 0) or 0)
        uom = values.get("uom") or ""
        lines.append(f"{code} thiếu {_fmt_qty(qty)} {uom}".strip())

    reason = (
        f"[Tự động] Kế hoạch {report.get('plan_month') or ''}: đã dùng hết "
        "công suất máy nhưng vẫn thiếu hàng phục vụ bán: "
        + "; ".join(lines)
        + ". Publish sản lượng tối đa có thể; phần thiếu cần xử lý "
        "(tăng ca/công suất, điều chỉnh FC hoặc chuyển sang tháng sau)."
    )
    return {
        "proposal_id": report.get("proposal_id"),
        "reason": reason,
        "approved_by": AUTO_APPROVER,
    }
