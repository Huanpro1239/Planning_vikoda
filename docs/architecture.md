# Kiến trúc mục tiêu

Repo đã hoàn tất các bước refactor chính theo hướng tách nghiệp vụ khỏi hạ tầng; workflow production dùng trực tiếp các entrypoint hiện hành.

```text
planning/     # planning engine, scheduling, proposal/publish orchestration
nvl/          # tồn NVL, PO mở, purchase planning
sharepoint/   # Microsoft Graph / SharePoint client, identity, retry, ETag
excel/        # workbook/XML read-write-verify utilities
domain/       # business rules thuần, không I/O
scripts/      # công cụ vận hành/audit
tests/        # regression/integration tests
docs/         # architecture, review history, runbooks
```

## Quy tắc migrate

1. Không đổi nghiệp vụ trong PR cấu trúc.
2. Compatibility shim chỉ được giữ khi còn consumer thực tế; shim không còn consumer phải được xóa.
3. Code mới ưu tiên import qua package canonical (`sharepoint.client`, `nvl.*`, `planning.*`).
4. Mỗi bước migrate phải qua toàn bộ test suite trước khi merge.
5. Sau khi toàn bộ caller đã chuyển sang package mới, compatibility shim phải được xóa và CI chặn tái xuất hiện.

## Thứ tự refactor

- Phase A: package layout + di chuyển docs.
- Phase B: hợp nhất SharePoint/Graph client.
- Phase C: migrate NVL modules.
- Phase D: migrate planning modules.
- Phase E: xóa compatibility shim sau khi workflow/tests không còn phụ thuộc.


## Repository hygiene

- `scripts/` chỉ chứa công cụ vận hành/audit dùng cho production hoặc support.
- `scripts/dev/` chứa công cụ chẩn đoán thủ công dành cho developer.
- `docs/archive/` chứa kế hoạch/review lịch sử; không để tài liệu tạm ở repository root.
- Package canonical không dùng `import *`; chỉ export API ổn định qua `__all__`.
- CI compile toàn bộ repository trước khi chạy unit tests.


## NVL module boundaries

`sync_nvl_stock.py` is a compatibility CLI only. Canonical responsibilities:

- `nvl/models.py`: immutable data contracts/state containers.
- `nvl/config.py`: JSON configuration loading.
- `nvl/values.py`: pure code/quantity normalization.
- `nvl/reconcile.py`: source reading and target reconciliation.
- `nvl/workbook.py`: OpenXML patching and integrity verification.
- `nvl/reporting.py`: audit/error report serialization.
- `nvl/service.py`: SharePoint orchestration, retries and publish boundary.
- `excel/workbook_xml.py`: reusable low-level OpenXML helpers.
- `excel/openpyxl_io.py`: deterministic workbook/ZipFile lifecycle cleanup.

NVL runtime modules must not import `sync_stock.py` directly. SharePoint access goes
through `sharepoint.client`, while workbook XML utilities go through `excel/`.


## Canonical SharePoint dependency

`sharepoint/client.py` owns `GraphClient`, `GraphRequestError`, Graph constants,
retry classification and SharePoint identity/read helpers. Authentication comes from
`sharepoint/auth.py`.

Legacy `sync_stock.py` re-exports the canonical Graph API for compatibility, but
canonical/runtime callers import Graph/auth symbols from `sharepoint.client` and
finished-goods stock business logic from `stock/`.
The dependency direction is therefore:

```text
planning/ scripts/ -> stock/ -> excel/
       |              |
       +------------> sharepoint.client -> sharepoint.auth

sync_stock.py -> stock/ + sharepoint.client  (legacy CLI/facade only)
```

Production workflows call their entrypoints directly. `scripts/auth_runner.py` has been removed; authentication is resolved by `sharepoint.auth` through canonical imports.


## Canonical Planning dependency

Planning implementations now live under the `planning/` package:

- `planning/metrics/`: static metrics package:
  - `state.py`: runtime-state serialization and period helpers.
  - `readers.py`: actual stock, system receipts, No kho and Danh_muc readers.
  - `calculation.py`: explicit all-month M:R calculations.
  - `workbook.py`: M:R OpenXML patch/diff.
  - `service.py`: standalone SharePoint orchestration.
- `planning/khsx_ki/`: KHSX_ki package:
  - `calendar.py`: Monday-Sunday month partitioning.
  - `layout.py`: week/total column layout, styles and safe merge handling.
  - `workbook.py`: SKU reconciliation and workbook patching.
  - `verification.py`: independent read-only integrity verification.
- `planning/pipeline.py`: pure proposal-building pipeline.

Legacy metrics patchers `sync_planning_metrics_compat.py`, `sync_planning_metrics_direct.py` and `sync_planning_metrics_all_months.py` are removed. No module may rebind `planning.metrics` functions or policy state at import/runtime. No kho debt, Danh_muc leadtime and all-month calculations are wired explicitly through function arguments.


### KHSX_ki dependency boundary

```text
calendar
   ↓
 layout
   ↓
workbook
   ↓
verification
```

`planning.khsx_ki.__init__` is a stable facade only. The former
`planning/khsx_ki.py` monolith is removed. Verification stays read-only and
must not become a workbook writer.


## Finished-goods stock boundaries

Finished-goods stock synchronization is canonical under `stock/`:

- `stock/constants.py`: SharePoint paths and workbook sheet names.
- `stock/values.py`: product-code and numeric normalization.
- `stock/state.py`: `state.json` fingerprint persistence.
- `stock/readers.py`: actual/factory/accounting/master-data readers and master fingerprint.
- `stock/workbook.py`: `Ton_kho` OpenXML patching.
- `stock/service.py`: standalone Graph orchestration.
- `stock/__init__.py`: stable business facade.

Root `sync_stock.py` is a thin compatibility CLI/facade only. Production runtime
must not import business logic from it. The obsolete `sync_stock_compat.py`
import-time patcher is removed; the dimension-tolerant Danh_muc reader is now the
single canonical implementation in `stock.readers`.

The `stock/` package may depend on `excel/` and `sharepoint/`, but must not
depend on `planning/` or the legacy root `sync_stock.py`.


## Canonical Planning entrypoints

The remaining root `sync_planning_*.py` implementation modules were migrated
into the `planning/` package:

- `planning/fc.py`: FC selection, robust target reader, hash and workbook patch.
- `planning/layout.py`: canonical daily schedule column layout.
- `planning/calendar.py`: date headers, cross-year resolution and all-month calendar update.
- `planning/stock_inputs.py`: Ton_kho -> Ke_hoach_SX J/K synchronization.
- `planning/publish.py`: proposal/publish policy, artifacts, retries and CLI behavior.

Root files `sync_planning_fc.py`, `sync_planning_calendar.py`,
`sync_planning_stock_inputs.py` and `sync_planning_pipeline.py` are CLI
entrypoints only. Production/runtime/tests must import `planning.*` directly.

Obsolete side-effect modules are removed:
- `sync_planning_fc_compat.py`
- `sync_planning_calendar_all_months.py`
- `sync_planning_layout.py`

The dimension-tolerant FC reader is the canonical implementation in
`planning.fc`. The all-month calendar behavior is explicit through
`planning.calendar.prepare_calendar_update_all_months`; no module rebinds
another module's function at import time.


## Planning publish boundaries

The former `planning/publish.py` orchestration monolith is split into:

- `planning/publish/snapshot.py`: SharePoint snapshot acquisition and provenance.
- `planning/publish/proposal.py`: stable XLSX proposal identity and review artifacts.
- `planning/publish/policy.py`: publish/review authorization only.
- `planning/publish/state.py`: input-diff detection and state persistence.
- `planning/publish/service.py`: retry/recompute/upload orchestration.
- `planning/publish/runner.py`: CLI parsing, authentication and Graph bootstrap.
- `planning/publish/__init__.py`: stable compatibility facade.

The dependency direction is intentionally one-way:

```text
snapshot   proposal   policy   state
    \        |         |       /
             service
                |
              runner
```

Policy remains independent of Graph, workbook mutation and artifact persistence.
Snapshot acquisition does not decide publish eligibility. Proposal identity does
not save runtime state. Service owns orchestration but not the individual policy
or hashing rules.


## Weekly model boundaries

The pure scheduling/calculation engine remains `planning/weekly_engine.py`.
The workbook adapter formerly in `planning/weekly_model.py` is split into:

- `planning/weekly_model/policy.py`: Debt mode and Schedule profile normalization.
- `planning/weekly_model/inputs.py`: Danh_muc/Ke_hoach_SX input extraction and planning-input fingerprint.
- `planning/weekly_model/schedule.py`: adapter-level analysis around the pure weekly engine.
- `planning/weekly_model/workbook.py`: O:R and daily schedule workbook patching.
- `planning/weekly_model/report.py`: mass balance, shared-machine validation and report construction.
- `planning/weekly_model/verification.py`: independent report/workbook verification.
- `planning/weekly_model/service.py`: thin analyze → patch → report orchestration.
- `planning/weekly_model/__init__.py`: stable facade.

Dependency direction:

```text
policy
  |
inputs
  |
schedule
 /     \
workbook report
   \     |
     service
       |
 verification (reads schedule + report)

weekly_engine  <-- imported by adapter modules only
```

`planning/weekly_engine.py` must never import `planning.weekly_model`; this
keeps the domain scheduling engine independent from Excel/workbook/report concerns.


## Planning package dependency boundary

Package-level architecture is enforced by
`tests/test_planning_package_architecture.py`. The test resolves absolute and
relative Python imports with AST, maps them to Planning domains, validates the
allowed direction and rejects cycles.

```text
                 weekly_engine
                       ^
                       |
          +------------+------------+
          |            |            |
       metrics    weekly_model    khsx_ki
          \            |            /
           \           |           /
                    pipeline
                   /        \
              publish     verification
```

Workbook-support modules are intentionally below orchestration:

```text
fc
^ \
|  \
layout  stock_inputs
^
|
calendar
```

The package rules are:

- `weekly_engine` is the bottom domain engine and imports no other Planning layer.
- `metrics`, `weekly_model` and `khsx_ki` are sibling domains and must not
  import one another.
- `weekly_model` may depend on `weekly_engine`; cross-domain verification is
  coordinated by `pipeline`, not by a sibling domain.
- `pipeline` may orchestrate support modules plus
  `metrics/weekly_model/khsx_ki`, but may not import `publish` or
  `verification`.
- `publish` is an outer online orchestration shell and may call `pipeline`
  and metrics runtime-state APIs.
- `verification` is an outer independent read-only verifier and may use only
  workbook support plus `weekly_engine` calculations.
- lower layers may never import `pipeline`, `publish` or `verification`.
- the actual domain graph must remain acyclic.

The former direct dependency
`weekly_model.verification -> khsx_ki` was removed. `planning.pipeline`
now invokes the weekly-model verifier and KHSX_ki verifier separately and
combines their results at the orchestration boundary.
