# SharePoint Planning change watch

Planning still keeps the daily production schedule as a safety net. In addition,
`.github/workflows/sharepoint-watch.yml` polls SharePoint every 10 minutes and
compares the exact production input fingerprints stored in `runtime-state/state.json`.

The watcher is read-only. It does not publish the workbook and does not update
runtime state. When one or more production inputs have changed, it emits:

`repository_dispatch: sharepoint_stock_updated`

The existing `Sync SharePoint Stock` workflow receives that event and runs with
`--skip-if-unchanged`. Planning and the watcher share the
`sharepoint-stock-sync` concurrency group, so a watcher cannot create a second
production cycle while Planning is already running.

Inputs compared are the same ones used by the production publish boundary:

- five SharePoint stock source ETags;
- `Danh_muc` fingerprint;
- `FC` fingerprint;
- `No_kho` fingerprint;
- `Ke_hoach_SX` planning-input fingerprint;
- planning engine version.

If SharePoint keeps the target workbook locked and publish fails with HTTP 423,
runtime state remains unchanged. The next watcher cycle will therefore detect
the still-pending input change and emit the event again. After a successful
publish updates runtime state, later watcher runs become no-ops.

External Power Automate or another sender may still emit the same
`sharepoint_stock_updated` event; the receiver remains compatible.
