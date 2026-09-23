# Production Readiness Gate

Planning releases use one executable gate everywhere:

```bash
python -X utf8 scripts/production_readiness.py
```

For local use:

```bash
make readiness
```

## Gate phases

The gate fails fast. A release is ready only if every phase passes:

1. **compile** — compile all Python modules.
2. **architecture** — package dependency boundaries and repository hygiene.
3. **business_contracts** — Service First, debt, Safety Stock, KHS/PET9000 and publish-policy contracts.
4. **release_safety** — workbook round-trip, deterministic stable output, proposal/dry-run no-upload behavior, and workflow wiring.
5. **full_regression** — the complete test suite.

The gate writes:

```text
production_readiness_report.json
```

The report contains gate version, Git SHA when available, phase status,
return code and elapsed time. CI uploads it as an artifact.

## Merge boundary

`.github/workflows/test.yml` keeps the existing workflow/job identity but runs
the readiness command instead of separate compile/test commands. This preserves
the protected-branch check surface while strengthening what the check means.

A pull request must therefore pass the same readiness command used for release.

## Production publish boundary

`.github/workflows/sync-stock.yml` runs the same readiness command before:

- production runtime-state restore is used by the planner;
- proposal generation;
- any controlled SharePoint publish.

If the readiness command exits non-zero, GitHub Actions stops before the
Planning publish command. The readiness report is included in the planning
audit artifact.

## Release-safety semantics

The release-specific tests guarantee:

- the generated workbook can be opened, saved again and independently verified;
- identical snapshots produce identical stable proposal hashes and proposal IDs;
- proposal/dry-run mode never uploads SharePoint and never persists runtime state;
- PR CI and production publish workflow are wired to the exact same gate command;
- production workflow cannot conditionally skip the gate.

## Change control

Do not weaken or skip a gate phase merely to make a release pass. If a business
contract or release rule is intentionally changed, update the corresponding
contract/readiness test and documentation in the same reviewed pull request.
