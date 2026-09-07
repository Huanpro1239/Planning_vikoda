"""Thuật toán lập kế hoạch tuần/tháng theo mô hình ``Ke hoach SX tuan``.

Module này chỉ chứa business rules. Không chứa dữ liệu tháng, SKU hoặc sản lượng
thực tế của workbook tham chiếu. Mọi dữ liệu được truyền vào lúc runtime.
"""

from __future__ import annotations

from calendar import monthrange
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from math import ceil, floor
from typing import Literal

TOLERANCE = 1e-6
DebtFormulaMode = Literal["SUBTRACT_BOOK_ON_DEBT", "IGNORE_BOOK_ON_DEBT"]


@dataclass(frozen=True, slots=True)
class PlannerPolicy:
    """Cấu hình rule theo nhãn nghiệp vụ; không chứa dữ liệu kế hoạch.

    ``spread_product_codes`` là override runtime cho các sản phẩm Galon dùng
    profile rải đều. Reader có thể tự phát hiện profile này từ công thức workbook,
    vì vậy code dự án không cần hardcode mã sản phẩm.
    """

    serialized_lines: frozenset[str] = field(
        default_factory=lambda: frozenset({"KHS", "PET 9000"})
    )
    setup_shifts: float = 0.5
    galon_line: str = "Galon"
    rgb_line: str = "RGB"
    rgb_gas_group: str = "RGB có gas"
    rgb_nogas_group: str = "RGB không gas"
    pet_blocking_line: str = "PET 9000"
    spread_product_codes: frozenset[int] = field(default_factory=frozenset)


@dataclass(frozen=True, slots=True)
class WeeklyInputRow:
    source_row: int
    ma_sp: int
    ten_sp: str
    don_vi_tinh: str
    sl_me: float
    sl_ca: float
    chuyen: str
    nhom_sp: str
    phan_loai: str
    quy_cach: float
    shifts_per_day: float
    ton_dau_thuc_te: float
    ton_dau_so_sach: float
    fc: float
    ton_cuoi_du_kien: float
    no_kho: float
    avg_daily_sales: float
    leadtime: float
    debt_formula_mode: DebtFormulaMode


@dataclass(frozen=True, slots=True)
class WeeklyCalculatedRow:
    input: WeeklyInputRow
    p_need: float
    q_rounded: float
    production_days: float
    start_datetime: datetime | None
    p_service_need: float
    q_service_rounded: float

    @property
    def start_date(self) -> date | None:
        return None if self.start_datetime is None else self.start_datetime.date()

    @property
    def schedulable_qty(self) -> float:
        return max(0.0, self.q_rounded)

    @property
    def service_qty(self) -> float:
        return max(0.0, self.q_service_rounded)

    @property
    def buffer_qty(self) -> float:
        return max(0.0, self.schedulable_qty - self.service_qty)


@dataclass(frozen=True, slots=True)
class DailyPlanRow:
    ma_sp: int
    source_row: int
    chuyen: str
    date: date
    qty: float
    phase: str = "full"

SUGAR_CLASSIFICATION = "Có đường"


def is_sugar_classification(classification: str) -> bool:
    """Nhận diện SKU 'Có đường' bất kể hoa/thường/khoảng trắng.

    Trước đây engine so khớp chính xác chuỗi ``"Có đường"`` trong khi
    ``sync_planning_metrics`` và ``verify_planning_month`` dùng ``casefold()``.
    Sự lệch pha này khiến một SKU khai báo ``"có đường"`` (viết thường) bị
    engine tính theo SL/ca còn verifier lại kỳ vọng SL/mẻ, dẫn tới verify sai.
    Chuẩn hóa về một cách so khớp duy nhất cho cả ba nơi.
    """
    return str(classification or "").strip().casefold() == SUGAR_CLASSIFICATION.casefold()


def _roundup_away_from_zero(value: float) -> int:
    """Semantics tương đương Excel ``ROUNDUP(value, 0)``."""
    if abs(value) <= TOLERANCE:
        return 0
    return ceil(value) if value > 0 else -ceil(abs(value))


def _round_production(value: float, row: WeeklyInputRow) -> float:
    if abs(value) <= TOLERANCE:
        return 0.0
    basis = row.sl_me if is_sugar_classification(row.phan_loai) else row.sl_ca
    if basis <= 0:
        raise ValueError(f"Basis làm tròn <= 0 cho SKU {row.ma_sp}")
    return _roundup_away_from_zero(value / basis) * basis


def calculate_row(
    row: WeeklyInputRow,
    *,
    period_year: int,
    period_month: int,
) -> WeeklyCalculatedRow:
    """Tính mục tiêu đầy đủ và nhu cầu bắt buộc theo Service First.

    Mục tiêu đầy đủ vẫn dùng công thức workbook hiện hành. Khi không có nợ,
    phần tồn cuối dự kiến được tách thành safety-stock buffer. Phần service
    phải được ưu tiên trước; buffer chỉ dùng capacity còn lại.
    """

    if row.no_kho > 0:
        if row.debt_formula_mode == "SUBTRACT_BOOK_ON_DEBT":
            p_service_need = row.fc + row.no_kho - row.ton_dau_so_sach
        elif row.debt_formula_mode == "IGNORE_BOOK_ON_DEBT":
            p_service_need = row.fc + row.no_kho
        else:  # pragma: no cover
            raise ValueError(f"debt_formula_mode không hợp lệ: {row.debt_formula_mode}")
        # Giữ nguyên rule nợ đã kiểm chứng: nhánh nợ không cộng tồn cuối.
        p_need = p_service_need
    else:
        p_service_need = row.fc - row.ton_dau_so_sach + row.no_kho
        p_need = p_service_need + row.ton_cuoi_du_kien

    q_service_rounded = _round_production(p_service_need, row)
    q_rounded = _round_production(p_need, row)

    if row.sl_ca <= 0 or row.shifts_per_day <= 0:
        raise ValueError(f"SL/ca và ca/ngày phải > 0 cho SKU {row.ma_sp}")
    production_days = q_rounded / row.sl_ca / row.shifts_per_day

    start_datetime: datetime | None
    if row.avg_daily_sales <= 0:
        start_datetime = None
    else:
        first = datetime(period_year, period_month, 1)
        candidate = first + timedelta(
            days=row.ton_dau_thuc_te / row.avg_daily_sales - row.leadtime
        )
        start_datetime = max(first, candidate)

    return WeeklyCalculatedRow(
        input=row,
        p_need=p_need,
        q_rounded=q_rounded,
        production_days=production_days,
        start_datetime=start_datetime,
        p_service_need=p_service_need,
        q_service_rounded=q_service_rounded,
    )

def calculate_rows(
    rows: list[WeeklyInputRow],
    *,
    period_year: int,
    period_month: int,
) -> list[WeeklyCalculatedRow]:
    return [
        calculate_row(row, period_year=period_year, period_month=period_month)
        for row in rows
    ]


def _month_dates(year: int, month: int) -> list[date]:
    return [date(year, month, d) for d in range(1, monthrange(year, month)[1] + 1)]


def _schedule_serialized_lines(
    rows: list[WeeklyCalculatedRow],
    policy: PlannerPolicy,
) -> list[DailyPlanRow]:
    """Xếp KHS/PET 9000 trên một máy chung theo hai tầng ưu tiên.

    Tầng 1 chạy toàn bộ phần bắt buộc để đáp ứng bán hàng/nợ. Tầng 2 chỉ dùng
    capacity còn lại để xây tồn cuối dự kiến. Một timeline duy nhất được dùng
    cho cả KHS và PET 9000 nên hai line không thể chạy song song.
    """

    selected = [
        row
        for row in rows
        if row.input.chuyen in policy.serialized_lines
        and row.schedulable_qty > TOLERANCE
        and row.start_datetime is not None
    ]
    if not selected:
        return []

    shared_shifts_per_day = selected[0].input.shifts_per_day
    if shared_shifts_per_day <= 0:
        raise ValueError("Máy chung KHS/PET 9000 phải có số ca/ngày > 0.")
    if any(
        abs(row.input.shifts_per_day - shared_shifts_per_day) > TOLERANCE
        for row in selected[1:]
    ):
        raise ValueError(
            "KHS và PET 9000 chung một máy nên phải dùng cùng lịch số ca/ngày."
        )

    selected.sort(
        key=lambda r: (r.start_datetime, r.input.source_row, r.input.chuyen)
    )

    output: list[DailyPlanRow] = []
    previous_quy_cach: float | None = None
    previous_end_shift = 0.0

    def schedule_phase(phase: str):
        nonlocal previous_quy_cach, previous_end_shift
        for row in selected:
            item = row.input
            assert row.start_datetime is not None
            qty_to_schedule = (
                row.service_qty if phase == "service" else row.buffer_qty
            )
            if qty_to_schedule <= TOLERANCE:
                continue

            start_day = row.start_datetime.date()
            first = date(start_day.year, start_day.month, 1)
            earliest_shift = (start_day - first).days * shared_shifts_per_day
            setup_shifts = (
                policy.setup_shifts
                if previous_quy_cach is not None
                and abs(previous_quy_cach - item.quy_cach) > TOLERANCE
                else 0.0
            )

            start_shift = max(earliest_shift, previous_end_shift) + setup_shifts
            end_shift = start_shift + qty_to_schedule / item.sl_ca

            for day_no, current in enumerate(
                _month_dates(start_day.year, start_day.month), start=1
            ):
                day_start_shift = (day_no - 1) * shared_shifts_per_day
                day_end_shift = day_no * shared_shifts_per_day
                overlap = max(
                    0.0,
                    min(end_shift, day_end_shift) - max(start_shift, day_start_shift),
                )
                qty = overlap * item.sl_ca
                if qty > TOLERANCE:
                    output.append(
                        DailyPlanRow(
                            item.ma_sp,
                            item.source_row,
                            item.chuyen,
                            current,
                            qty,
                            phase,
                        )
                    )

            previous_quy_cach = item.quy_cach
            previous_end_shift = end_shift

    # Không cho safety stock của SKU A chiếm máy trước phần bán hàng của SKU B.
    schedule_phase("service")
    schedule_phase("buffer")
    return output

def _schedule_galon(
    rows: list[WeeklyCalculatedRow],
    policy: PlannerPolicy,
) -> list[DailyPlanRow]:
    output: list[DailyPlanRow] = []
    for row in rows:
        item = row.input
        if (
            item.chuyen != policy.galon_line
            or row.q_rounded <= TOLERANCE
            or row.start_datetime is None
        ):
            continue
        start = row.start_datetime.date()
        dates = _month_dates(start.year, start.month)

        if item.ma_sp in policy.spread_product_codes:
            working = [d for d in dates if d.weekday() != 6]
            if not working:
                continue
            work_days = len(working)
            base = (
                row.q_rounded / work_days
                if row.q_rounded <= work_days * item.sl_ca
                else item.sl_ca
            )
            extra_qty = max(0.0, row.q_rounded - work_days * item.sl_ca)
            for rank, current in enumerate(working, start=1):
                extra_today = 0.0
                if extra_qty > TOLERANCE:
                    extra_today = (
                        floor(extra_qty / item.sl_ca * rank / work_days)
                        - floor(extra_qty / item.sl_ca * (rank - 1) / work_days)
                    ) * item.sl_ca
                qty = base + extra_today
                if qty > TOLERANCE:
                    output.append(
                        DailyPlanRow(
                            item.ma_sp, item.source_row, item.chuyen, current, qty
                        )
                    )
        else:
            start_day_no = start.day
            for day_no, current in enumerate(dates, start=1):
                run_day = day_no - start_day_no
                qty = max(
                    0.0,
                    min(
                        row.q_rounded,
                        (run_day + 1) * item.shifts_per_day * item.sl_ca,
                    )
                    - max(0.0, run_day * item.shifts_per_day * item.sl_ca),
                )
                if qty > TOLERANCE:
                    output.append(
                        DailyPlanRow(
                            item.ma_sp, item.source_row, item.chuyen, current, qty
                        )
                    )
    return output


def _daily_lookup(plan: list[DailyPlanRow]) -> dict[tuple[int, date], float]:
    result: dict[tuple[int, date], float] = {}
    for row in plan:
        key = (row.ma_sp, row.date)
        result[key] = result.get(key, 0.0) + row.qty
    return result


def _schedule_rgb_gas(
    rows: list[WeeklyCalculatedRow],
    existing: list[DailyPlanRow],
    policy: PlannerPolicy,
) -> list[DailyPlanRow]:
    output: list[DailyPlanRow] = []
    gas_rows = [
        row
        for row in rows
        if row.input.chuyen == policy.rgb_line
        and row.input.nhom_sp == policy.rgb_gas_group
        and row.q_rounded > TOLERANCE
        and row.start_datetime is not None
    ]
    gas_rows.sort(key=lambda r: r.input.source_row)

    for row in gas_rows:
        item = row.input
        assert row.start_datetime is not None
        first = date(row.start_datetime.year, row.start_datetime.month, 1)
        dates = _month_dates(first.year, first.month)
        if item.sl_me <= 0:
            raise ValueError(f"SL/mẻ phải > 0 cho SKU RGB {item.ma_sp}")
        total_units = ceil(row.q_rounded / item.sl_me - TOLERANCE)
        previous = _daily_lookup(existing + output)

        previous_gas_rows = [
            r for r in gas_rows if r.input.source_row < item.source_row
        ]
        avail_by_day: dict[date, int] = {}
        for current in dates:
            prev_units = 0.0
            for prev in previous_gas_rows:
                if prev.input.sl_me <= 0:
                    continue
                prev_units += (
                    previous.get((prev.input.ma_sp, current), 0.0) / prev.input.sl_me
                )
            avail_by_day[current] = max(
                0, floor(item.shifts_per_day - prev_units + TOLERANCE)
            )

        normal_cap = sum(avail_by_day[d] for d in dates if d.weekday() != 6)
        allow_sun = total_units > normal_cap
        eligible = [
            d
            for d in dates
            if avail_by_day[d] > 0 and (d.weekday() != 6 or allow_sun)
        ]
        active_weeks = max(1, ceil(len(dates) / 7))
        done_units = 0

        for current in dates:
            if current not in eligible:
                continue
            wk = min(active_weeks, ceil(current.day / 7))
            need_units = max(0, total_units - done_units)
            future = [d for d in eligible if d > current]
            future_cap = sum(avail_by_day[d] for d in future)
            target_units = (
                total_units if not future else ceil(total_units * wk / active_weeks)
            )
            must_now = max(0, need_units - future_cap)
            want_units = max(must_now, max(0, target_units - done_units))
            run_units = min(avail_by_day[current], need_units, want_units)
            if run_units > 0:
                output.append(
                    DailyPlanRow(
                        item.ma_sp,
                        item.source_row,
                        item.chuyen,
                        current,
                        run_units * item.sl_me,
                    )
                )
                done_units += run_units
    return output


def _schedule_rgb_nogas(
    rows: list[WeeklyCalculatedRow],
    existing: list[DailyPlanRow],
    policy: PlannerPolicy,
) -> list[DailyPlanRow]:
    output: list[DailyPlanRow] = []
    nogas_rows = [
        row
        for row in rows
        if row.input.chuyen == policy.rgb_line
        and row.input.nhom_sp == policy.rgb_nogas_group
        and row.q_rounded > TOLERANCE
        and row.start_datetime is not None
    ]
    nogas_rows.sort(key=lambda r: r.input.source_row)

    for row in nogas_rows:
        item = row.input
        assert row.start_datetime is not None
        if item.sl_me <= 0:
            raise ValueError(f"SL/mẻ phải > 0 cho SKU RGB {item.ma_sp}")
        start = row.start_datetime.date()
        dates = _month_dates(start.year, start.month)
        total_units = ceil(row.q_rounded / item.sl_me - TOLERANCE)
        lookup = _daily_lookup(existing + output)

        avail_by_day: dict[date, int] = {}
        for current in dates:
            pet_busy = any(
                r.input.chuyen == policy.pet_blocking_line
                and lookup.get((r.input.ma_sp, current), 0.0) > TOLERANCE
                for r in rows
            )
            rgb_other_units = 0.0
            for other in rows:
                if (
                    other.input.chuyen == policy.rgb_line
                    and other.input.nhom_sp != policy.rgb_nogas_group
                    and other.input.sl_ca > 0
                ):
                    rgb_other_units += (
                        lookup.get((other.input.ma_sp, current), 0.0)
                        / other.input.sl_ca
                    )
            raw = floor(item.shifts_per_day - rgb_other_units + TOLERANCE)
            avail_by_day[current] = 0 if pet_busy else max(0, raw)

        active = [d for d in dates if d >= start]
        normal_cap = sum(avail_by_day[d] for d in active if d.weekday() != 6)
        allow_sun = total_units > normal_cap
        eligible = [
            d
            for d in active
            if avail_by_day[d] > 0 and (d.weekday() != 6 or allow_sun)
        ]
        active_weeks = max(1, ceil((len(dates) - start.day + 1) / 7))
        done_units = 0
        for current in active:
            if current not in eligible:
                continue
            wk = min(active_weeks, ceil((current.day - start.day + 1) / 7))
            need_units = max(0, total_units - done_units)
            future = [d for d in eligible if d > current]
            future_cap = sum(avail_by_day[d] for d in future)
            target_units = (
                total_units if not future else ceil(total_units * wk / active_weeks)
            )
            must_now = max(0, need_units - future_cap)
            want_units = max(must_now, max(0, target_units - done_units))
            run_units = min(avail_by_day[current], need_units, want_units)
            if run_units > 0:
                output.append(
                    DailyPlanRow(
                        item.ma_sp,
                        item.source_row,
                        item.chuyen,
                        current,
                        run_units * item.sl_me,
                    )
                )
                done_units += run_units
    return output


def build_daily_plan(
    rows: list[WeeklyCalculatedRow],
    *,
    policy: PlannerPolicy | None = None,
) -> list[DailyPlanRow]:
    """Tạo lịch ngày bằng các rule tổng quát của mô hình workbook."""

    active_policy = policy or PlannerPolicy()
    serialized = _schedule_serialized_lines(rows, active_policy)
    galon = _schedule_galon(rows, active_policy)
    rgb_gas = _schedule_rgb_gas(rows, serialized + galon, active_policy)
    rgb_nogas = _schedule_rgb_nogas(
        rows, serialized + galon + rgb_gas, active_policy
    )
    result = serialized + galon + rgb_gas + rgb_nogas
    result.sort(key=lambda r: (r.date, r.source_row, r.ma_sp))
    return result
