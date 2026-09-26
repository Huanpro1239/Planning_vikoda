# Planning → NVL Paired-Run Health

The paired-run health monitor verifies that every successful Planning production
run has exactly one corresponding NVL downstream execution and consistent
release provenance.

## What is checked

For Planning runs whose original event is `schedule` or
`repository_dispatch` and conclusion is `success`, the checker validates:

- one and only one paired NVL downstream run,
- NVL run-name upstream Planning run ID,
- Planning `head_sha` equals the SHA encoded in the NVL run,
- NVL downstream conclusion is `success`,
- successful NVL runs have paired release provenance in
  `runtime-state/nvl/releases/`,
- release `upstream_planning.head_sha` equals the Planning SHA,
- release `commit.sha` equals the Planning SHA,
- duplicate release provenance is rejected.

Primary failure codes are:

```text
missing_nvl
duplicate_nvl
head_sha_mismatch
nvl_downstream_failed
missing_release_provenance
duplicate_release_provenance
release_head_sha_mismatch
release_commit_sha_mismatch
```

## Run locally

With a checked-out `runtime-state` branch and a GitHub token:

```bash
GITHUB_TOKEN=... python -X utf8 scripts/check_paired_runs.py \
  --runtime-state .runtime-state \
  --enforce-after 2026-09-26T03:30:00Z
```

Or:

```bash
make paired-run-health \
  RUNTIME_STATE=.runtime-state \
  ENFORCE_AFTER=2026-09-26T03:30:00Z
```

The command emits:

```text
paired_run_health.json
paired_run_health.csv
paired_run_health.md
```

It returns non-zero if any enforced pair is unhealthy.

## GitHub Actions monitor

`.github/workflows/paired-run-health.yml` runs in two modes:

1. immediately after every `Sync SharePoint NVL Stock` completion,
2. daily at 06:45 Vietnam time as a safety net.

The event-driven run catches failed NVL downstream workflows quickly. The daily
run detects the harder case where Planning succeeded but no NVL run was created
at all.

The monitor is read-only. It has only `actions: read` and `contents: read`
permissions and never writes SharePoint or `runtime-state`.

## Grace window

A Planning run is not evaluated until 20 minutes after completion. This avoids
false alarms while the paired NVL run is still starting or finishing.

## Cutover compatibility

Enforcement begins at the configured `--enforce-after` timestamp. This avoids
retroactively treating legacy runs as failures.

During migration, a historical paired NVL release manifest is accepted as
evidence even if its older Actions run does not yet have the new structured
`run-name`. That case is surfaced as warning
`actions_pair_evidence_missing`, not as `missing_nvl`.

## Run-name contract

Paired automatic NVL runs use:

```text
NVL paired planning=<planning_run_id> sha=<planning_head_sha> event=<event>
```

Direct/manual NVL runs use a different `NVL direct ...` run-name and are
ignored by the pairing matcher.
