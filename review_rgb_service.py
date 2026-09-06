import itertools
import json
import math
from pathlib import Path

import sync_planning_schedule as base
from sync_planning_schedule_production import install_production_output_cleanup


ARTIFACT_DIR = Path("dry_run_artifacts")
WORKBOOK_PATH = ARTIFACT_DIR / "Sap_ke_hoach_dry_run.xlsx"
REPORT_PATH = ARTIFACT_DIR / "planning_schedule_report.json"
OUTPUT_JSON = ARTIFACT_DIR / "rgb_service_review.json"
OUTPUT_MD = ARTIFACT_DIR / "rgb_service_review.md"
FOCUS_CODES = ("130100013", "130100149")
EPS = 1e-7


def _current_schedule(headers, products):
    return {
        product["code"]: {
            day: (0.0 if product["existing_daily"][index] in (None, "") else float(product["existing_daily"][index]))
            for index, day in enumerate(headers)
        }
        for product in products
    }


def _empty_schedule(headers, products):
    return {
        product["code"]: {day: 0.0 for day in headers}
        for product in products
    }


def _effective_release_index(headers, product):
    earliest = product.get("earliest_date")
    if earliest is None or earliest <= headers[0]:
        return 0
    for index, day in enumerate(headers):
        if day >= earliest:
            return index
    return len(headers)


def _place_setup(remaining, headers, start_index, setup_shifts=0.5):
    for index in range(max(0, start_index), len(headers)):
        if remaining[index] + EPS >= setup_shifts:
            remaining[index] -= setup_shifts
            return index
    return None


def _place_quantum(remaining, headers, start_index, quantum_shift):
    for index in range(max(0, start_index), len(headers)):
        if remaining[index] + EPS >= quantum_shift:
            remaining[index] -= quantum_shift
            return index
    return None


def _schedule_one_campaign_sequence(headers, products, capacity, order):
    """Exact batch/shift, serial full-campaign comparator with 0.5 shift setup."""
    by_code = {product["code"]: product for product in products}
    schedule = _empty_schedule(headers, products)
    remaining = [float(capacity) for _ in headers]
    carryover = {}
    setup_shifts = 0.0
    cursor = 0
    last_code = None

    for code in order:
        product = by_code[code]
        units = int(product.get("required_units", 0) or 0)
        if units <= 0:
            continue
        cursor = max(cursor, _effective_release_index(headers, product))

        if last_code is not None and code != last_code:
            setup_day = _place_setup(remaining, headers, cursor)
            if setup_day is None:
                carryover[code] = units * float(product.get("quantum_qty", 0) or 0)
                continue
            setup_shifts += 0.5
            cursor = setup_day

        placed = 0
        quantum_shift = float(product.get("quantum_shift", 0) or 0)
        quantum_qty = float(product.get("quantum_qty", 0) or 0)
        while placed < units:
            day_index = _place_quantum(remaining, headers, cursor, quantum_shift)
            if day_index is None:
                break
            schedule[code][headers[day_index]] += quantum_qty
            cursor = day_index
            placed += 1

        if placed < units:
            carryover[code] = (units - placed) * quantum_qty
        if placed > 0:
            last_code = code

    return schedule, carryover, setup_shifts


def _service_metrics(headers, products, schedule, carryover=None, setup_shifts=0.0):
    carryover = dict(carryover or {})
    per_code = {}
    stockout_skus = 0
    stockout_sku_days = 0
    total_deficit = 0.0
    max_deficit = 0.0
    latest_stockout_index = -1

    for product in products:
        code = product["code"]
        balance = float(product.get("actual_stock", 0) or 0)
        debt = max(float(product.get("debt", 0) or 0), 0.0)
        days = []
        code_stockout_days = 0
        first_stockout = None
        last_stockout = None

        for index, day in enumerate(headers):
            production = float(schedule[code].get(day, 0) or 0)
            demand = float(product["demand_by_day"].get(day, 0) or 0)
            debt_due = debt if index == 0 else 0.0
            opening = balance
            balance += production - demand - debt_due
            deficit = max(-balance, 0.0)
            is_stockout = deficit > EPS
            if is_stockout:
                code_stockout_days += 1
                stockout_sku_days += 1
                total_deficit += deficit
                max_deficit = max(max_deficit, deficit)
                latest_stockout_index = max(latest_stockout_index, index)
                if first_stockout is None:
                    first_stockout = day
                last_stockout = day
            days.append(
                {
                    "date": day.isoformat(),
                    "opening_net": opening,
                    "production": production,
                    "demand": demand,
                    "debt_due": debt_due,
                    "ending_net": balance,
                    "deficit": deficit,
                    "stockout": is_stockout,
                }
            )

        if code_stockout_days:
            stockout_skus += 1
        per_code[code] = {
            "stockout_days": code_stockout_days,
            "first_stockout": first_stockout.isoformat() if first_stockout else None,
            "last_stockout": last_stockout.isoformat() if last_stockout else None,
            "ending_net": balance,
            "days": days,
        }

    carryover_qty = sum(max(float(value or 0), 0.0) for value in carryover.values())
    metrics = {
        "carryover_skus": len([value for value in carryover.values() if float(value or 0) > EPS]),
        "carryover_qty": carryover_qty,
        "stockout_skus": stockout_skus,
        "stockout_sku_days": stockout_sku_days,
        "total_deficit_sku_days": total_deficit,
        "max_deficit": max_deficit,
        "latest_stockout_index": latest_stockout_index,
        "setup_shifts": float(setup_shifts),
    }
    return metrics, per_code


def _objective(metrics):
    return (
        int(metrics["carryover_skus"]),
        round(float(metrics["carryover_qty"]), 6),
        int(metrics["stockout_skus"]),
        int(metrics["stockout_sku_days"]),
        round(float(metrics["total_deficit_sku_days"]), 6),
        round(float(metrics["max_deficit"]), 6),
        int(metrics["latest_stockout_index"]),
        round(float(metrics["setup_shifts"]), 6),
    )


def _search_best_sequence(headers, products, capacity):
    active = [product for product in products if int(product.get("required_units", 0) or 0) > 0]
    codes = [product["code"] for product in active]
    if len(codes) > 8:
        raise RuntimeError(
            f"RGB có {len(codes)} SKU active; exact permutation review giới hạn 8 SKU."
        )

    best = None
    evaluated = 0
    for order in itertools.permutations(codes):
        evaluated += 1
        schedule, carryover, setup_shifts = _schedule_one_campaign_sequence(
            headers,
            products,
            capacity,
            order,
        )
        metrics, per_code = _service_metrics(
            headers,
            products,
            schedule,
            carryover=carryover,
            setup_shifts=setup_shifts,
        )
        candidate = {
            "order": list(order),
            "schedule": schedule,
            "carryover": carryover,
            "metrics": metrics,
            "per_code": per_code,
        }
        if best is None or _objective(metrics) < _objective(best["metrics"]):
            best = candidate

    best["evaluated_sequences"] = evaluated
    return best


def _day1_protection_bound(headers, products, capacity):
    """Exact day-1 lower bound for stockout SKU count with setup and quanta."""
    if not headers:
        return {
            "at_risk_codes": [],
            "inherently_unprotectable_codes": [],
            "protectable_codes": [],
            "max_protected_skus": 0,
            "minimum_stockout_skus": 0,
            "best_protected_sets": [],
        }

    day = headers[0]
    protectable = []
    unprotectable = []
    at_risk = []
    details = {}

    for product in products:
        code = product["code"]
        stock = float(product.get("actual_stock", 0) or 0)
        debt = max(float(product.get("debt", 0) or 0), 0.0)
        demand = float(product["demand_by_day"].get(day, 0) or 0)
        need = max(debt + demand - stock, 0.0)
        if need <= EPS:
            continue

        at_risk.append(code)
        quantum_qty = float(product.get("quantum_qty", 0) or 0)
        quantum_shift = float(product.get("quantum_shift", 0) or 0)
        release_index = _effective_release_index(headers, product)
        if quantum_qty <= EPS or quantum_shift <= EPS:
            required_units = math.inf
            required_shifts = math.inf
        else:
            required_units = int(math.ceil(need / quantum_qty - EPS))
            required_shifts = required_units * quantum_shift

        details[code] = {
            "need": need,
            "required_units": None if not math.isfinite(required_units) else required_units,
            "required_shifts": None if not math.isfinite(required_shifts) else required_shifts,
            "release_index": release_index,
        }
        if (
            release_index > 0
            or not math.isfinite(required_shifts)
            or required_shifts > float(capacity) + EPS
        ):
            unprotectable.append(code)
        else:
            protectable.append(code)

    by_code = {product["code"]: product for product in products}
    best_count = 0
    best_sets = set()
    for subset_size in range(len(protectable) + 1):
        for subset in itertools.combinations(protectable, subset_size):
            if not subset:
                feasible = True
            else:
                feasible = False
                for order in itertools.permutations(subset):
                    production_shifts = sum(
                        float(details[code]["required_shifts"] or 0)
                        for code in order
                    )
                    setup_shifts = base.SETUP_SHIFTS * max(0, len(order) - 1)
                    if production_shifts + setup_shifts <= float(capacity) + EPS:
                        feasible = True
                        break
            if not feasible:
                continue
            if subset_size > best_count:
                best_count = subset_size
                best_sets = {tuple(sorted(subset))}
            elif subset_size == best_count:
                best_sets.add(tuple(sorted(subset)))

    return {
        "at_risk_codes": sorted(at_risk),
        "inherently_unprotectable_codes": sorted(unprotectable),
        "protectable_codes": sorted(protectable),
        "max_protected_skus": best_count,
        "minimum_stockout_skus": max(0, len(at_risk) - best_count),
        "best_protected_sets": [list(item) for item in sorted(best_sets)],
        "details": details,
        "capacity_shifts": float(capacity),
        "setup_shifts_per_switch": float(base.SETUP_SHIFTS),
    }


def _solo_day1_floor(headers, product, capacity):
    if not headers:
        return {}
    demand = float(product["demand_by_day"].get(headers[0], 0) or 0)
    debt = max(float(product.get("debt", 0) or 0), 0.0)
    stock = float(product.get("actual_stock", 0) or 0)
    need = max(debt + demand - stock, 0.0)
    quantum_shift = float(product.get("quantum_shift", 0) or 0)
    quantum_qty = float(product.get("quantum_qty", 0) or 0)
    if quantum_shift <= 0 or quantum_qty <= 0:
        max_solo = 0.0
    else:
        units = int(math.floor((float(capacity) + EPS) / quantum_shift))
        max_solo = units * quantum_qty
    return {
        "opening_stock": stock,
        "debt_due": debt,
        "demand_day1": demand,
        "gross_need_day1": need,
        "max_valid_solo_production_day1": max_solo,
        "minimum_unavoidable_day1_deficit_if_solo": max(need - max_solo, 0.0),
        "quantum_qty": quantum_qty,
        "quantum_shift": quantum_shift,
        "capacity_shifts": float(capacity),
        "effective_release_date": (
            headers[_effective_release_index(headers, product)].isoformat()
            if _effective_release_index(headers, product) < len(headers)
            else None
        ),
    }


def _json_schedule(schedule):
    return {
        code: {day.isoformat(): qty for day, qty in values.items() if abs(float(qty or 0)) > EPS}
        for code, values in schedule.items()
    }


def _build_review(workbook_bytes, plan_year, report=None):
    install_production_output_cleanup()
    headers, products = base.read_schedule_inputs(workbook_bytes, plan_year=plan_year)
    rgb_products = [product for product in products if product.get("line") == "RGB"]
    if not rgb_products:
        raise RuntimeError("Không có SKU RGB trong workbook dry-run.")
    capacity = max(float(product.get("max_shifts_per_day", 0) or 0) for product in rgb_products)

    report = report or {}
    rgb_report = ((report.get("resources") or {}).get("RGB") or {})
    rgb_meta = rgb_report.get("meta") or {}
    current_setup = float(rgb_meta.get("setup_shifts", 0) or 0)
    current_mode = str(rgb_meta.get("mode") or "unknown")

    current_schedule = _current_schedule(headers, rgb_products)
    current_metrics, current_per_code = _service_metrics(
        headers,
        rgb_products,
        current_schedule,
        setup_shifts=current_setup,
    )
    best = _search_best_sequence(headers, rgb_products, capacity)
    day1_bound = _day1_protection_bound(headers, rgb_products, capacity)
    current_day1_stockout_codes = sorted(
        code
        for code, item in current_per_code.items()
        if item.get("days") and item["days"][0].get("stockout")
    )
    day1_bound["current_stockout_codes"] = current_day1_stockout_codes
    day1_bound["current_stockout_skus"] = len(current_day1_stockout_codes)
    day1_bound["current_meets_lower_bound"] = (
        len(current_day1_stockout_codes) == day1_bound["minimum_stockout_skus"]
    )
    day1_bound["monthly_stockout_sku_lower_bound"] = day1_bound["minimum_stockout_skus"]
    day1_bound["current_monthly_stockout_skus"] = current_metrics["stockout_skus"]
    day1_bound["current_primary_objective_optimal"] = (
        current_metrics["stockout_skus"] == day1_bound["minimum_stockout_skus"]
    )

    floors = {
        code: _solo_day1_floor(headers, next(product for product in rgb_products if product["code"] == code), capacity)
        for code in FOCUS_CODES
        if any(product["code"] == code for product in rgb_products)
    }

    return {
        "resource": "RGB",
        "capacity_shifts_per_day": capacity,
        "active_codes": [product["code"] for product in rgb_products if product.get("planned_qty", 0) > EPS],
        "current": {
            "optimizer_mode": current_mode,
            "metrics": current_metrics,
            "objective": list(_objective(current_metrics)),
            "per_code": current_per_code,
            "schedule": _json_schedule(current_schedule),
        },
        "best_full_campaign_alternative": {
            "order": best["order"],
            "evaluated_sequences": best["evaluated_sequences"],
            "metrics": best["metrics"],
            "objective": list(_objective(best["metrics"])),
            "per_code": best["per_code"],
            "schedule": _json_schedule(best["schedule"]),
            "carryover": best["carryover"],
        },
        "alternative_strictly_better": _objective(best["metrics"]) < _objective(current_metrics),
        "day1_capacity_lower_bound": day1_bound,
        "day1_solo_lower_bounds": floors,
    }


def _fmt(value):
    if isinstance(value, float):
        return f"{value:.3f}"
    return str(value)


def _markdown(review):
    cur = review["current"]
    alt = review["best_full_campaign_alternative"]
    bound = review["day1_capacity_lower_bound"]
    lines = [
        "# RGB service / stockout review",
        "",
        f"- Capacity: {review['capacity_shifts_per_day']} ca/ngày",
        f"- Active RGB SKU: {', '.join(review['active_codes'])}",
        f"- Current optimizer: `{cur.get('optimizer_mode')}`",
        f"- Exact full-campaign sequences evaluated: {alt['evaluated_sequences']}",
        f"- Best full-campaign alternative strictly improves current objective: **{'YES' if review['alternative_strictly_better'] else 'NO'}**",
        "",
        "## Whole-resource comparison",
        "| Metric | Current production schedule | Best full-campaign alternative |",
        "|---|---:|---:|",
    ]
    for key in (
        "carryover_skus",
        "carryover_qty",
        "stockout_skus",
        "stockout_sku_days",
        "total_deficit_sku_days",
        "max_deficit",
        "latest_stockout_index",
        "setup_shifts",
    ):
        lines.append(
            f"| {key} | {_fmt(cur['metrics'][key])} | {_fmt(alt['metrics'][key])} |"
        )

    best_sets = [" + ".join(item) if item else "(none)" for item in bound["best_protected_sets"]]
    lines.extend([
        "",
        "## Day-1 capacity/setup lower bound",
        f"- SKU sẽ âm nếu ngày 1 không sản xuất: {', '.join(bound['at_risk_codes']) or '(none)' }",
        f"- SKU không thể cứu hết ngày 1 dù chạy một mình: {', '.join(bound['inherently_unprotectable_codes']) or '(none)' }",
        f"- Tối đa SKU có thể bảo vệ ngày 1 với capacity + setup: {bound['max_protected_skus']}",
        f"- Số SKU stockout tối thiểu bắt buộc ngày 1: **{bound['minimum_stockout_skus']}**",
        f"- Các tập SKU có thể bảo vệ tối đa: {' | '.join(best_sets) or '(none)'}",
        f"- Lịch V10 hiện tại stockout ngày 1: {bound['current_stockout_skus']} SKU ({', '.join(bound['current_stockout_codes']) or 'none'})",
        f"- V10 đạt lower bound ngày 1: **{'YES' if bound['current_meets_lower_bound'] else 'NO'}**",
        f"- Vì stockout ngày 1 đã tính vào stockout theo tháng, lower bound stockout SKU tháng cũng là {bound['monthly_stockout_sku_lower_bound']}; V10 hiện có {bound['current_monthly_stockout_skus']}.",
        f"- Primary objective (số SKU từng stockout) đạt tối ưu theo lower bound: **{'YES' if bound['current_primary_objective_optimal'] else 'NO'}**",
        "",
        f"- Alternative order: {' -> '.join(alt['order'])}",
        "",
        "## Day-1 lower bound if each focus SKU owned the whole RGB line",
    ])
    for code, floor in review["day1_solo_lower_bounds"].items():
        lines.append(
            f"- **{code}**: need={floor['gross_need_day1']:.3f}; "
            f"max valid solo production={floor['max_valid_solo_production_day1']:.3f}; "
            f"minimum unavoidable deficit={floor['minimum_unavoidable_day1_deficit_if_solo']:.3f}; "
            f"quantum={floor['quantum_qty']:.3f} ({floor['quantum_shift']:.3f} ca); "
            f"effective release={floor['effective_release_date']}"
        )

    for code in FOCUS_CODES:
        current = cur["per_code"].get(code)
        alternative = alt["per_code"].get(code)
        if not current or not alternative:
            continue
        floor = review["day1_solo_lower_bounds"].get(code, {})
        unavoidable = float(floor.get("minimum_unavoidable_day1_deficit_if_solo", 0) or 0)
        current_day1 = float(current["days"][0]["deficit"] or 0)
        alt_day1 = float(alternative["days"][0]["deficit"] or 0)
        lines.extend([
            "",
            f"## {code}",
            f"- Current stockout days: {current['stockout_days']}; first={current['first_stockout']}; last={current['last_stockout']}",
            f"- Alternative stockout days: {alternative['stockout_days']}; first={alternative['first_stockout']}; last={alternative['last_stockout']}",
            f"- Day1 current deficit={current_day1:.3f}; alternative={alt_day1:.3f}; solo lower bound={unavoidable:.3f}",
            "",
            "| Date | Current prod | Current net | Alt prod | Alt net |",
            "|---|---:|---:|---:|---:|",
        ])
        for current_day, alt_day in zip(current["days"], alternative["days"]):
            lines.append(
                f"| {current_day['date']} | {current_day['production']:.3f} | "
                f"{current_day['ending_net']:.3f} | {alt_day['production']:.3f} | "
                f"{alt_day['ending_net']:.3f} |"
            )

    lines.extend([
        "",
        "## Interpretation rule",
        "- Day-1 capacity/setup lower bound is an exact combinatorial bound for the first-day obligations under current quantum sizes, release dates, line capacity and 0.5-ca SKU setup.",
        "- Solo day-1 lower bound proves only the unavoidable minimum for that SKU if it monopolized RGB that day.",
        "- Whole-resource comparison is used to decide whether current delay is avoidable without merely moving shortage to another RGB SKU.",
        "- The comparator keeps exact P quantum/batch, 2-ca/day resource capacity, 0.5-ca setup on every SKU switch, and the effective release dates used by the production scheduler.",
    ])
    return "\n".join(lines) + "\n"


def main():
    if not WORKBOOK_PATH.exists() or not REPORT_PATH.exists():
        raise RuntimeError("Chưa có dry-run artifacts để review RGB.")
    report = json.loads(REPORT_PATH.read_text(encoding="utf-8"))
    plan_month = str(report.get("plan_month") or "")
    try:
        plan_year = int(plan_month.split("-", 1)[0])
    except Exception as exc:
        raise RuntimeError(f"plan_month không hợp lệ: {plan_month!r}") from exc

    review = _build_review(WORKBOOK_PATH.read_bytes(), plan_year, report=report)
    OUTPUT_JSON.write_text(
        json.dumps(review, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    markdown = _markdown(review)
    OUTPUT_MD.write_text(markdown, encoding="utf-8")
    print(markdown)


if __name__ == "__main__":
    main()
