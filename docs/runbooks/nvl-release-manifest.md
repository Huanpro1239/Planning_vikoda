# NVL release manifest

Mỗi production publish NVL thành công tạo một immutable release record sau khi cả
hai bước `nvl.stock` (Ton_NVL!D) và `nvl.open_po_safe` (Ton_NVL!E) hoàn tất.

## Manifest

Artifact local:

```text
nvl_release_manifest.json
```

Schema hiện hành:

```text
nvl_release_manifest_v1
```

Manifest liên kết một release với:

- Git commit SHA, ref, workflow run ID/attempt và actor.
- NVL readiness gate version + SHA-256 của readiness report.
- Source/target ETag của stock và Open PO.
- Raw + stable SHA-256 của stock proposal và final Open-PO proposal.
- Final published workbook hash.
- Publish evidence: changed count, upload response, post-upload verification,
  expected ETag và server hash.

Production yêu cầu `NVL_REQUIRE_READINESS=1`; readiness report phải PASS và
Git SHA trong report phải trùng commit đang publish.

## Runtime-state history

Workflow production lưu manifest dưới namespace riêng:

```text
runtime-state
├── state.json
├── planning_runtime.json
├── latest_release.json
├── releases/
└── nvl/
    ├── latest_release.json
    └── releases/
        └── <release_id>.json
```

`nvl/latest_release.json` là con trỏ tiện dụng. Mỗi file trong
`nvl/releases/` là immutable history theo release ID.

## Production ordering

```text
NVL readiness PASS
    ↓
checkout runtime-state
    ↓
publish Ton_NVL!D
    ↓
publish Ton_NVL!E
    ↓
build nvl_release_manifest.json
    ↓
persist runtime-state/nvl/releases/<release_id>.json
    ↓
upload GitHub Actions audit artifacts
```

Proposal-only, staging-only hoặc publish thất bại không được tạo release record.

## Verify an old NVL release

Basic verification of a historical manifest:

```bash
python -X utf8 scripts/verify_nvl_release.py \
  runtime-state/nvl/releases/<release_id>.json
```

This always validates the manifest contract itself: schema/version, release ID
derived from the final stable workbook hash, successful Stock/Open-PO publish
evidence, readiness status, commit consistency, revision shape and internal hash
consistency. If external evidence is not present beside the manifest, basic mode
reports warnings rather than claiming those files were rechecked.

For complete evidence verification, place/download the per-run artifacts beside
the manifest and run:

```bash
python -X utf8 scripts/verify_nvl_release.py nvl_release_manifest.json \
  --strict \
  --stock-proposal nvl_stock_proposal.xlsx \
  --final-workbook nvl_open_po_proposal.xlsx \
  --readiness nvl_production_readiness_report.json \
  --stock-report nvl_stock_report.json \
  --open-po-report nvl_open_po_report.json
```

Strict mode recomputes or cross-checks:

- raw and stable SHA-256 of the stock proposal,
- raw and stable SHA-256 of the final Open-PO proposal,
- final published-workbook/server SHA against Open-PO post-upload evidence,
- readiness report SHA-256, PASS status, gate version and Git SHA,
- Stock source/target revisions, status, changed count and verification evidence,
- Open-PO source/target revisions, changed cells and exact publish evidence,
- release ID/version identity,
- release commit existence in local Git history.

Machine-readable output:

```bash
python -X utf8 scripts/verify_nvl_release.py <manifest> --json
```

Local Make target:

```bash
make verify-nvl-release MANIFEST=runtime-state/nvl/releases/<release_id>.json
make verify-nvl-release MANIFEST=nvl_release_manifest.json STRICT=1
```

For a shallow clone that does not contain an old release commit, fetch the
relevant Git history before strict verification. `--no-git` skips only the
local Git-history lookup; use it when commit existence is verified by another
trusted mechanism.

## Hash-linked release chain

New NVL manifests contain:

```json
{
  "chain": {
    "previous_release_id": "...",
    "previous_manifest_sha256": "..."
  }
}
```

The first NVL release is the chain anchor and stores both values as `null`.
Every later release hashes the exact immutable JSON bytes of the previous
manifest. This makes mutation or deletion of an older release detectable by the
next release in the ledger.

The production workflow refreshes `runtime-state` immediately before manifest
creation, passes `nvl/latest_release.json` as the predecessor, audits the
updated ledger, and only then commits/pushes release history.

## Audit the whole NVL release ledger

Run:

```bash
python -X utf8 scripts/audit_nvl_releases.py runtime-state/nvl/releases
```

The command scans all immutable manifests and checks:

- each manifest with the normal NVL release verifier,
- filename ↔ `release_id` identity,
- duplicate release IDs,
- chronological order,
- predecessor existence,
- predecessor manifest SHA-256,
- chain gaps, resets and forks,
- `latest_release.json` identity and byte equality with the newest release.

It writes two derived indexes:

```text
nvl_release_index.json
nvl_release_index.csv
```

The CSV is a compact history table with release time, release ID, commit,
readiness gate, source revisions, workbook hash, chain status and issues.

Production persists the derived indexes in the state-only branch:

```text
runtime-state/nvl/release_index.json
runtime-state/nvl/release_index.csv
```

If the ledger is missing a referenced predecessor, a predecessor hash no longer
matches, the chain forks/resets, a manifest fails verification, or
`latest_release.json` is stale, the audit exits non-zero and production does
not persist the new history commit.

For a local checkout with shallow Git history:

```bash
python -X utf8 scripts/audit_nvl_releases.py runtime-state/nvl/releases --no-git
```

Make target:

```bash
make audit-nvl-releases
make audit-nvl-releases RELEASES_DIR=.runtime-state/nvl/releases NO_GIT=1
```

A legacy manifest without a `chain` block is reported as
`legacy_unlinked` with a warning rather than falsely classified as tampered.
