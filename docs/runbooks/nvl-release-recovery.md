# NVL release recovery

Recovery is a controlled forward operation that restores only `Ton_NVL!D:E`
from a verified historical NVL release. It never overwrites the whole workbook
with an old file.

## Dry-run first

Provide the immutable release ID (or manifest path) and the historical
`nvl_open_po_proposal.xlsx` from that release's GitHub Actions artifact:

```bash
python -X utf8 scripts/restore_nvl_release.py <release_id> \
  --historical-workbook nvl_open_po_proposal.xlsx
```

Default mode is dry-run. It creates:

```text
nvl_recovery_current_backup.xlsx
nvl_recovery_proposal.xlsx
nvl_recovery_report.json
```

No SharePoint upload occurs.

## Fail-closed guards

Before a proposal is produced the command requires:

1. the release exists in the audited NVL ledger,
2. the whole release ledger passes,
3. the historical workbook hash matches the selected release manifest,
4. current SharePoint target identity matches configured name + sourcedoc,
5. historical/current `Ton_NVL` code-to-row layout is identical.

The proposal copies only cells in columns D/E from the historical release.
Every other cell in `Ton_NVL` and every unrelated ZIP part must stay equal to
the current workbook.

## Controlled publish

Publish requires an explicit approval token equal to the exact historical
release ID:

```bash
python -X utf8 scripts/restore_nvl_release.py <release_id> \
  --historical-workbook nvl_open_po_proposal.xlsx \
  --publish \
  --approve-release-id <release_id>
```

Immediately before upload the target metadata is fetched again. If the ETag
changed, recovery fails rather than recomputing or retrying silently. Upload
uses `If-Match` with the original ETag.

After upload, the workbook is downloaded again and verified:

- D/E equal the historical release,
- all non-D/E cells remain equal to the pre-recovery current snapshot,
- unrelated workbook ZIP parts remain unchanged except specifically validated
  SharePoint/Office server metadata.

A failed post-upload verification is reported as failure; recovery does not
silently upload a second time.

## Make target

Dry-run:

```bash
make restore-nvl-release \
  RELEASE=<release_id> \
  HISTORICAL=nvl_open_po_proposal.xlsx
```

Controlled publish:

```bash
make restore-nvl-release \
  RELEASE=<release_id> \
  HISTORICAL=nvl_open_po_proposal.xlsx \
  PUBLISH=1 \
  APPROVE=<release_id>
```

Recovery intentionally does not fabricate a normal stock/Open-PO release
manifest. A recovery-specific ledger record should be added only with explicit
recovery provenance in a separate release-history contract.
