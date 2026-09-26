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

`nvl.stock` owns the canonical stock API and CLI (`python -m nvl.stock`).
The legacy root `sync_nvl_stock.py` facade has been removed. Canonical responsibilities:

- `nvl/models.py`: immutable data contracts/state containers.
- `nvl/config.py`: JSON configuration loading.
- `nvl/values.py`: pure code/quantity normalization.
- `nvl/reconcile.py`: source reading and target reconciliation.
- `nvl/workbook.py`: OpenXML patching and integrity verification.
- `nvl/reporting.py`: audit/error report serialization.
- `nvl/service.py`: SharePoint orchestration, retries and publish boundary.
- `nvl/open_po.py`: canonical Open-PO reader, reconciliation, XML patch/verify and online orchestration for `Ton_NVL!E`.
- `excel/workbook_xml.py`: reusable low-level OpenXML helpers.
- `excel/openpyxl_io.py`: deterministic workbook/ZipFile lifecycle cleanup.

Legacy root NVL facades `sync_nvl_stock.py`, `sync_nvl_open_po.py` and
`sync_nvl_open_po_safe.py` have been removed. Runtime/tests use only canonical
`nvl.*` modules. The safe resolver lives at `nvl.open_po_safe`, and production
invokes it with `python -m nvl.open_po_safe`. SharePoint access goes through
`sharepoint.client`, while workbook XML utilities go through `excel/`.


## Canonical SharePoint dependency

`sharepoint/client.py` owns `GraphClient`, `GraphRequestError`, Graph constants,
retry classification and SharePoint identity/read helpers. Authentication comes from
`sharepoint/auth.py`.

Canonical/runtime callers import Graph/auth symbols from `sharepoint.client` and
finished-goods stock business logic from `stock/`. The legacy root
`sync_stock.py` facade has been removed; standalone stock synchronization uses
`python -m stock`.
The dependency direction is therefore:

```text
planning/ scripts/ -> stock/ -> excel/
       |              |
       +------------> sharepoint.client -> sharepoint.auth
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
- `planning/cleanup.py`: Planning-specific workbook cleanup using shared OpenXML helpers.
- `planning/schedule_report.py`: report serialization, hashing and console rendering.
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

Legacy root entrypoints `sync_stock.py` and `sync_stock_compat.py` are removed.
Use the package-native CLI `python -m stock`. The dimension-tolerant Danh_muc
reader is the single canonical implementation in `stock.readers`.

The `stock/` package may depend on `excel/` and `sharepoint/`, but must not
depend on `planning/` or any root compatibility module.


## Canonical Planning entrypoints

The remaining root `sync_planning_*.py` implementation modules were migrated
into the `planning/` package:

- `planning/fc.py`: FC selection, robust target reader, hash and workbook patch.
- `planning/layout.py`: canonical daily schedule column layout.
- `planning/calendar.py`: date headers, cross-year resolution and all-month calendar update.
- `planning/stock_inputs.py`: Ton_kho -> Ke_hoach_SX J/K synchronization.
- `planning/publish/`: proposal/publish policy, artifacts, retries and CLI behavior.

Legacy root CLIs `sync_planning_fc.py`, `sync_planning_calendar.py`,
`sync_planning_stock_inputs.py` and `sync_planning_pipeline.py` have been
removed. Use package-native entrypoints instead:

- `python -m planning.fc`
- `python -m planning.calendar`
- `python -m planning.stock_inputs`
- `python -m planning.publish`

Production/runtime/tests import `planning.*` directly.

Legacy Planning support facades `planning_cleanup.py` and
`planning_schedule_report.py` have been removed. Canonical callers import
`planning.cleanup` and `planning.schedule_report` directly.

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


## Business invariant contracts

System-level business invariants are executable in
`tests/contracts/test_planning_business_invariants.py` and documented in
`docs/business_contracts.md`. Architecture refactors are not complete unless
this contract suite remains green. These tests intentionally cross
`weekly_engine -> weekly_model -> verification -> publish policy` boundaries
to protect business behavior rather than module shape.


## Production readiness gate

Release quality is enforced by one executable command:
`python -X utf8 scripts/production_readiness.py`.

The gate composes package architecture tests, business contracts, workbook
round-trip verification, deterministic proposal output, proposal-only dry-run
safety and the full regression suite. Both pull-request CI and the production
Planning workflow invoke this exact command. See
`docs/runbooks/production-readiness.md`.


## NVL production orchestration

Automatic NVL production no longer relies on an independent delayed cron. The
`Sync SharePoint NVL Stock` workflow listens to `workflow_run` completion of
`Sync SharePoint Stock` and starts only when the upstream Planning run
succeeded and its original event was `schedule` or `repository_dispatch`.
This makes NVL sequencing deterministic: Planning finishes first, then NVL
starts, so the two SharePoint writers do not overlap.

The NVL checkout is pinned to the upstream Planning `head_sha`, preserving
commit-level provenance across the paired production runs. A manual Planning
`workflow_dispatch` does not automatically publish NVL because the upstream
payload does not expose the Planning publish input safely. Manual NVL
`workflow_dispatch` remains available and is proposal-only unless
`publish=true`.

For paired automatic runs, the NVL manifest stores
`upstream_planning.workflow/run_id/head_sha/event`. The upstream `head_sha`
is also the canonical NVL release/readiness commit SHA, so the release record
can be traced directly back to the exact Planning run and code revision.
Direct/manual NVL runs keep this block present with null values rather than
fabricating an upstream Planning relationship.

## NVL production readiness gate

NVL production quality is enforced by one executable command:
`python -X utf8 scripts/nvl_production_readiness.py`.

The gate is intentionally separate from Planning release readiness and runs
seven ordered phases: architecture, business contracts, workbook round-trip,
deterministic proposal, ETag/concurrency safety, publish dry-run safety and
full regression. The production NVL workflow invokes this command before any
staging-copy, survey, proposal or publish operation and uploads
`nvl_production_readiness_report.json` with the audit artifacts.

## NVL release manifest and versioning

Successful NVL production publishes create `nvl_release_manifest.json` with
schema `nvl_release_manifest_v1`. One release represents the final
`Ton_NVL` state after both stock (column D) and Open PO (column E) publish
steps succeed. The manifest binds commit SHA, NVL gate version, source/target
revisions, proposal/workbook hashes and post-upload verification evidence.

History is persisted on the state-only `runtime-state` branch under
`nvl/latest_release.json` and `nvl/releases/<release_id>.json`, separate
from Planning release history. See `docs/runbooks/nvl-release-manifest.md`.

Historical NVL releases are independently re-verifiable with
`python -X utf8 scripts/verify_nvl_release.py <manifest>`. The verifier checks
manifest identity, commit/readiness evidence, input revisions, proposal/final
workbook hashes and Stock/Open-PO audit consistency; strict mode requires the
full evidence bundle.


NVL historical recovery is exposed through
`scripts/restore_nvl_release.py`. Recovery is proposal-first and restores only
`Ton_NVL!D:E` from a verified historical release. It requires a healthy ledger,
matching historical artifact hashes and stable target identity/layout; publish
requires exact release-ID approval plus an unchanged target ETag and performs a
post-upload integrity check. Recovery does not overwrite the whole historical
workbook and does not bypass normal release provenance.

NVL release history is also hash-linked: each new manifest records the previous
release ID and SHA-256 of the exact predecessor manifest bytes. The ledger audit
command `scripts/audit_nvl_releases.py` verifies every manifest, continuity,
fork/gap/reset conditions and the latest pointer, then emits JSON/CSV indexes.
Production runs this audit before committing release history to
`runtime-state`.

## NVL operational summary

Every NVL workflow run emits `nvl_operational_summary.json` and appends a
compact Markdown status table to GitHub Step Summary, even when an earlier step
fails. The summary surfaces readiness status, Stock D/Open-PO E changes,
release ID, published workbook hash, ledger status and the best available
failed phase. It is observational only and does not change publish behavior.

## Release manifest and versioning

Successful production publishes create
`planning_release_manifest.json` using schema
`planning_release_manifest_v1`. The manifest binds the published workbook to
its commit SHA, readiness gate, proposal identity, input revision, artifact
hashes and publish authorization. Production stores an immutable history under
`runtime-state/releases/` plus `latest_release.json`.

Manifest validation occurs before upload; persistence occurs only after the
authorized publish/state/decision path succeeds. See
`docs/runbooks/release-manifest.md`.


## Release verification

`planning.publish.verify_release` provides reusable offline verification for
historical release manifests, while `scripts/verify_release.py` is the CLI
boundary. Verification recomputes proposal artifact hashes/identity when the
workbook is available, checks readiness evidence and audit consistency, and can
require the recorded commit to exist in local Git history. Tamper detection is
part of the production-readiness release-safety suite.

## Cross-domain release overview

A read-only outer operations layer lives under `ops/`. The command
`scripts/release_overview.py` reads the state-only `runtime-state` checkout,
uses the canonical Planning verifier and NVL ledger auditor, and emits a unified
JSON/CSV/Markdown release history. It never participates in publish execution
and is deliberately not persisted by either production workflow, avoiding a
cross-domain write race between Planning and NVL.

## Planning → NVL paired-run health

The outer operations layer also owns a read-only paired-run monitor. Automatic
NVL runs encode their upstream Planning run ID, head SHA and trigger event in
the GitHub Actions `run-name`. `scripts/check_paired_runs.py` combines
Actions runtime evidence with immutable NVL release provenance and requires a
1:1 mapping for successful Planning production runs.

The monitor detects missing/duplicate NVL downstream runs, failed downstream
runs, SHA/event mismatches and missing/duplicate release provenance. It runs
after NVL workflow completion and on a daily safety-net schedule, with read-only
GitHub permissions. It never participates in either publish path.
