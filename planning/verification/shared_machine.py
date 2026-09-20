import math

from sync_stock import normalize_code

SHARED_RESOURCE = "KHS + PET 9000"
# Tên resource máy chung mà planning.weekly_model ghi vào report. Verifier
# chấp nhận cả hai tên để độc lập với thay đổi nhãn của report.
SHARED_RESOURCE_ALIASES = (SHARED_RESOURCE, "KHS/PET 9000 shared machine")
SHARED_LINES = {"KHS", "PET 9000"}
SETUP_SHIFTS = 0.5


def _find_shared_resource_info(schedule_report):
    """Tìm block resource máy chung trong report theo nhiều tên khóa.

    Trả về ``(resource_info, meta)`` nếu tìm thấy, ngược lại ``(None, None)``.
    Ưu tiên các alias đã biết, sau đó tới bất kỳ resource nào có ``meta.timeline``.
    """
    resources = (schedule_report or {}).get("resources") or {}
    for name in SHARED_RESOURCE_ALIASES:
        info = resources.get(name)
        if isinstance(info, dict):
            return info, (info.get("meta") or {})
    for info in resources.values():
        if isinstance(info, dict) and isinstance((info.get("meta") or {}).get("timeline"), list):
            return info, info.get("meta") or {}
    return None, None


def _header_day(value):
    text = str(value or "").strip()
    return text.splitlines()[0] if text else "?"


def _finite_number(value, label):
    if value in (None, ""):
        return 0.0
    if isinstance(value, bool):
        raise RuntimeError(f"{label} chứa TRUE/FALSE, không phải số.")
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise RuntimeError(f"{label} không phải số: {value!r}") from exc
    if not math.isfinite(result):
        raise RuntimeError(f"{label} phải là số hữu hạn, hiện {value!r}.")
    return result


def _validate_report_provenance(schedule_report, plan_year, plan_month):
    if not isinstance(schedule_report, dict):
        raise RuntimeError("Thiếu schedule report có provenance để kiểm carryover/setup timeline.")
    revision = schedule_report.get("input_revision")
    if not isinstance(revision, dict) or not revision:
        raise RuntimeError("Schedule report thiếu input_revision/provenance.")
    reported_month = schedule_report.get("plan_month")
    expected_month = f"{plan_year:04d}-{plan_month:02d}"
    if reported_month and reported_month != expected_month:
        raise RuntimeError(
            f"Schedule report thuộc {reported_month}, không phải kỳ {expected_month}."
        )


def _validate_shared_timeline(
    schedule_report,
    headers,
    planning_data,
    capacity,
):
    _validate_report_provenance(
        schedule_report,
        planning_data["__plan_year"],
        planning_data["__plan_month"],
    )
    resource_info, meta = _find_shared_resource_info(schedule_report)
    if not isinstance(resource_info, dict):
        raise RuntimeError(
            f"Schedule report thiếu resource {SHARED_RESOURCE!r}; không chứng minh được setup/timeline."
        )

    # capacity có thể nằm ở cấp resource hoặc trong meta (shared_machine).
    capacity_source = (
        resource_info.get("capacity_shifts_per_day")
        if resource_info.get("capacity_shifts_per_day") is not None
        else meta.get("capacity_shifts_per_day")
    )
    report_capacity = _finite_number(
        capacity_source,
        f"report {SHARED_RESOURCE} capacity",
    )
    if not math.isclose(report_capacity, capacity, rel_tol=1e-9, abs_tol=1e-6):
        raise RuntimeError(
            f"Schedule report capacity {report_capacity} khác workbook capacity {capacity} cho {SHARED_RESOURCE}."
        )

    timeline = (meta or {}).get("timeline")
    if not isinstance(timeline, list) or not timeline:
        raise RuntimeError(
            f"Schedule report thiếu timeline production/setup cho {SHARED_RESOURCE}."
        )

    events = []
    horizon = len(headers) * capacity
    for index, raw in enumerate(timeline):
        if not isinstance(raw, dict):
            raise RuntimeError(f"Timeline event #{index + 1} không phải object.")
        event = dict(raw)
        kind = str(event.get("kind") or "").strip()
        if kind not in {"production", "setup"}:
            raise RuntimeError(f"Timeline event #{index + 1} có kind={kind!r} không hợp lệ.")
        start = _finite_number(event.get("start_shift"), f"timeline[{index}].start_shift")
        end = _finite_number(event.get("end_shift"), f"timeline[{index}].end_shift")
        if start < -1e-7 or end < start - 1e-7 or end > horizon + 1e-7:
            raise RuntimeError(
                f"Timeline event #{index + 1} ngoài horizon: start={start}, end={end}, horizon={horizon}."
            )
        event["_start"] = start
        event["_end"] = end
        events.append(event)

    events.sort(key=lambda item: (item["_start"], item["_end"]))
    previous_end = 0.0
    for event in events:
        if event["_start"] < previous_end - 1e-7:
            raise RuntimeError(
                f"Timeline {SHARED_RESOURCE} bị chồng lấn tại shift {event['_start']:.3f}."
            )
        previous_end = max(previous_end, event["_end"])

    daily_usage = [0.0] * len(headers)
    reconstructed = {
        code: [0.0] * len(headers)
        for code, item in planning_data.items()
        if not code.startswith("__") and item["line"] in SHARED_LINES
    }

    last_production_code = None
    setup_since_last_production = 0.0
    for event in events:
        duration = event["_end"] - event["_start"]
        kind = event["kind"]

        for day_index in range(len(headers)):
            day_start = day_index * capacity
            day_end = day_start + capacity
            overlap = max(
                0.0,
                min(event["_end"], day_end) - max(event["_start"], day_start),
            )
            if overlap > 1e-9:
                daily_usage[day_index] += overlap

        if kind == "setup":
            setup_since_last_production += duration
            continue

        code = normalize_code(event.get("code")) or str(event.get("code") or "").strip()
        if code not in reconstructed:
            raise RuntimeError(f"Timeline production chứa mã {code!r} không thuộc máy chung trong workbook.")

        if last_production_code is not None and code != last_production_code:
            if setup_since_last_production < SETUP_SHIFTS - 1e-7:
                raise RuntimeError(
                    f"Timeline đổi mã {last_production_code} -> {code} thiếu setup {SETUP_SHIFTS:g} ca; "
                    f"chỉ có {setup_since_last_production:g} ca."
                )
        setup_since_last_production = 0.0
        last_production_code = code

        per_shift = planning_data[code]["per_shift"]
        expected_qty = duration * per_shift
        event_qty = _finite_number(event.get("qty"), f"timeline production {code} qty")
        if not math.isclose(event_qty, expected_qty, rel_tol=1e-9, abs_tol=1e-5):
            raise RuntimeError(
                f"Timeline mã {code} qty={event_qty} không khớp duration*E={expected_qty}."
            )

        for day_index in range(len(headers)):
            day_start = day_index * capacity
            day_end = day_start + capacity
            overlap = max(
                0.0,
                min(event["_end"], day_end) - max(event["_start"], day_start),
            )
            if overlap > 1e-9:
                reconstructed[code][day_index] += overlap * per_shift

    for day_index, used in enumerate(daily_usage):
        if used > capacity + 1e-6:
            raise RuntimeError(
                f"Timeline {SHARED_RESOURCE} ngày {_header_day(headers[day_index])} dùng "
                f"{used:.3f} ca gồm production/setup > capacity {capacity:.3f}."
            )

    for code, values in reconstructed.items():
        workbook_values = planning_data[code]["daily_values"]
        for index, (expected, actual) in enumerate(zip(values, workbook_values)):
            if not math.isclose(expected, actual, rel_tol=1e-9, abs_tol=1e-5):
                raise RuntimeError(
                    f"Timeline mã {code} ngày {_header_day(headers[index])}={expected} "
                    f"khác workbook={actual}."
                )


def _validate_shared_from_workbook(headers, planning_data, capacity):
    """Kiểm tra khả thi máy chung KHS/PET 9000 chỉ từ dữ liệu workbook.

    Dùng khi report không kèm timeline production/setup. Ràng buộc capacity
    theo ngày đã được kiểm ở vòng lặp chính; ở đây kiểm thêm chặn dưới theo
    tháng có tính setup:

        tổng_ca_sản_xuất + (số_SKU_được_SX - 1) * setup <= capacity * số_ngày

    Đây là điều kiện cần: mọi lịch hợp lệ do engine sinh ra đều thỏa (vì đã xếp
    được trên một timeline chung có setup), nên không tạo lỗi giả; đồng thời bắt
    được trường hợp workbook bị chỉnh tay vượt tổng năng lực máy chung.
    """
    if capacity <= 0:
        raise RuntimeError(
            f"{SHARED_RESOURCE} có capacity {capacity} không hợp lệ."
        )

    production_shifts = 0.0
    produced_codes = 0
    for code, item in planning_data.items():
        if code.startswith("__") or item["line"] not in SHARED_LINES:
            continue
        scheduled = float(item["scheduled"])
        per_shift = float(item["per_shift"])
        if scheduled <= 1e-7:
            continue
        if per_shift <= 0:
            raise RuntimeError(
                f"Mã {code} máy chung có SL/ca <= 0 nhưng có lịch SX."
            )
        production_shifts += scheduled / per_shift
        produced_codes += 1

    if produced_codes <= 1:
        return

    setup_shifts = max(0, produced_codes - 1) * SETUP_SHIFTS
    total_capacity = capacity * len(headers)
    required = production_shifts + setup_shifts
    if required > total_capacity + 1e-6:
        raise RuntimeError(
            f"{SHARED_RESOURCE}: tổng nhu cầu {required:.3f} ca "
            f"(SX {production_shifts:.3f} + setup {setup_shifts:.3f}) vượt "
            f"năng lực tháng {total_capacity:.3f} ca ({capacity:g} ca/ngày x "
            f"{len(headers)} ngày)."
        )
