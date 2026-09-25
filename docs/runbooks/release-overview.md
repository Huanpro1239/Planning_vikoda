# Planning + NVL Release Overview

This is a read-only operational view over the state-only `runtime-state`
branch. It does not publish workbooks and does not mutate release history.

## Run

Checkout or expose the `runtime-state` branch as a directory, for example
`.runtime-state`, then run:

```bash
python -X utf8 scripts/release_overview.py .runtime-state
```

Or:

```bash
make release-overview RUNTIME_STATE=.runtime-state
```

For shallow/offline checkouts where commit existence is verified elsewhere:

```bash
python -X utf8 scripts/release_overview.py .runtime-state --no-git
```

## Outputs

The command writes:

```text
release_overview.json
release_overview.csv
release_overview.md
```

The JSON contains both domain ledgers plus one combined timeline. The CSV is a
flat history table suitable for Excel/Power BI. The Markdown file is a compact
human-readable status page.

## Planning checks

For `runtime-state/releases/*.json`, the overview:

- runs the existing Planning release verifier in basic mode,
- checks filename ↔ release ID,
- checks publish timestamp,
- verifies `latest_release.json` points to the newest immutable release and is
  byte-identical to it.

## NVL checks

The overview delegates to the existing NVL ledger auditor, so it includes:

- manifest verification,
- hash-linked predecessor continuity,
- missing predecessor/tamper/fork/reset detection,
- latest pointer validation.

## Combined timeline

Each row contains:

- global sequence,
- domain: `planning` or `nvl`,
- publish timestamp,
- release ID,
- commit SHA,
- readiness gate version,
- Planning plan month/proposal ID when applicable,
- final proposal/workbook hash,
- verification status and issues.

The command returns non-zero if either domain release history is internally
invalid. An entirely empty release history is reported as `empty`, not as a
failure.

This command intentionally stays separate from both production publish
workflows. That avoids making Planning and NVL releases contend on a shared
derived overview file during concurrent runs.
