# Planning Release Manifest

Every successful production Planning publish creates:

```text
planning_release_manifest.json
```

The manifest is an immutable release record for the exact proposal that was
authorized and published.

## Required traceability

Current schema: `planning_release_manifest_v2`.

Each manifest records:

- `release_id` and `release_version`
- publish timestamp
- plan month, algorithm and engine version
- Git commit SHA, ref, repository, workflow, run ID/attempt and actor
- production-readiness status, gate version and readiness-report SHA-256
- proposal ID
- raw workbook SHA-256
- stable workbook/proposal SHA-256
- full input revision/provenance
- publish decision/approval basis
- input/master fingerprints used by the pipeline

## Production strict mode

The production workflow sets:

```text
PLANNING_REQUIRE_READINESS=1
```

At the publish boundary this requires:

1. `production_readiness_report.json` exists.
2. readiness status is `passed`.
3. a commit SHA is available.
4. the publish decision is `published` or `published_no_change`.

The manifest is validated in memory before SharePoint upload. It is written to
disk only after upload/state/decision persistence succeeds. A failed or retried
publish therefore must not leave a release record that looks successful.

## Persistent history

After a successful production publish, `sync-stock.yml` stores the manifest on
the dedicated `runtime-state` branch:

```text
latest_release.json
releases/<release_id>.json
```

`latest_release.json` is a convenience pointer. Files under `releases/`
form the immutable historical ledger and are committed together with runtime
state.

Release v2 also persists a JSON evidence bundle:

```text
release_evidence/<release_id>/
├── manifest.json
├── production_readiness_report.json
├── planning_schedule_report.json
├── planning_input_revision.json
└── planning_publish_decision.json
```

The v2 manifest stores canonical SHA-256 fingerprints for the schedule report,
input revision and publish decision. This lets the historical verifier detect
audit tampering independently of JSON whitespace/key ordering.

The same manifest is included in the per-run `planning-audit-<run_id>`
GitHub Actions artifact.


## Historical verification

Verify any stored release:

```bash
python scripts/verify_release.py runtime-state/releases/<release_id>.json
```

When the manifest is inside `runtime-state/releases/`, the command
automatically discovers the matching `release_evidence/<release_id>/` bundle.

For byte-level proposal verification:

```bash
python scripts/verify_release.py \
  runtime-state/releases/<release_id>.json \
  --proposal planning_proposal.xlsx
```

For a complete audit where every evidence source must be available:

```bash
python scripts/verify_release.py \
  runtime-state/releases/<release_id>.json \
  --proposal planning_proposal.xlsx \
  --strict
```

Without `--strict`, unavailable evidence is reported as `SKIPPED`; it is
never reported as passed. Any actual mismatch returns a failing verification
status and a non-zero CLI exit code.

The verifier supports both manifest v1 and v2. A v2 release additionally checks
canonical audit hashes. Commit verification uses the local Git object database
and checks that the release commit is an ancestor of current HEAD. For a shallow
clone, fetch repository history before using `--strict`.
