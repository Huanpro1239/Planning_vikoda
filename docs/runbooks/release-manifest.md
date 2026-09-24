# Planning Release Manifest

Every successful production Planning publish creates:

```text
planning_release_manifest.json
```

The manifest is an immutable release record for the exact proposal that was
authorized and published.

## Required traceability

Schema: `planning_release_manifest_v1`.

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

The same manifest is included in the per-run `planning-audit-<run_id>`
GitHub Actions artifact.


## Verify an old release

Basic manifest verification:

```bash
python -X utf8 scripts/verify_release.py runtime-state/releases/<release_id>.json
```

This verifies the manifest schema, release/proposal identity, successful publish
state, embedded readiness status, commit consistency and local Git history when
the commit is available. If external evidence is not beside the manifest, the
command reports warnings rather than pretending those hashes were rechecked.

For complete evidence verification, download/copy the release audit files into
one directory and run:

```bash
python -X utf8 scripts/verify_release.py planning_release_manifest.json \
  --strict \
  --proposal planning_proposal.xlsx \
  --readiness production_readiness_report.json \
  --report planning_schedule_report.json \
  --decision planning_publish_decision.json
```

Strict mode recomputes and verifies:

When the evidence files use their standard names and sit beside the manifest,
the verifier discovers them automatically; explicit `--proposal`,
`--readiness`, `--report` and `--decision` paths are only needed when the
files live elsewhere.

- raw workbook SHA-256
- stable workbook/proposal SHA-256
- proposal ID from algorithm + plan month + input revision + stable hash
- readiness report SHA-256, status, gate version and Git SHA
- planning report plan month, algorithm, input revision and proposal ID
- publish decision exact equality
- release commit existence in local Git history

Machine-readable output:

```bash
python -X utf8 scripts/verify_release.py <manifest> --json
```

Local Make target:

```bash
make verify-release MANIFEST=runtime-state/releases/<release_id>.json
make verify-release MANIFEST=planning_release_manifest.json STRICT=1
```

If a clone is shallow and does not contain an old commit, fetch the relevant Git
history before strict verification. `--no-git` explicitly skips local commit
history lookup; use it only when commit existence is being verified by another
trusted mechanism.


The verifier itself is covered by the production-readiness release-safety phase,
so changes to release verification must pass the same gate used before merge and
production publish.
