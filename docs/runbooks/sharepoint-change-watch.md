# SharePoint Planning change watch

Planning has no fixed-time schedule. It runs only when production inputs change
(or when started manually). Two senders emit the change event:

1. **Power Automate** (primary, near real-time) — see the setup section below.
2. **Watcher** (fallback) — `.github/workflows/sharepoint-watch.yml` polls SharePoint every 10 minutes and
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

GitHub may delay scheduled workflows by hours on busy runners, so the watcher
alone is not real-time. Use Power Automate for immediate runs.

## Power Automate setup (chạy ngay khi file tồn thay đổi)

1. GitHub → Settings → Developer settings → Fine-grained personal access token:
   repository `Huanpro1239/Planning_vikoda`, permission **Contents: Read and write**.
2. Power Automate → Automated cloud flow, trigger **SharePoint – When a file is
   created or modified (properties only)**:
   - Site: `https://vikodacomvn.sharepoint.com/sites/Planning`
   - Library: Documents
   - Folder: `/Tinh san xuat Mua hang 2027`
     (gồm `Ton thuc te/` và `Ton He thong/Ton Nha May|Ton Ke Toan/`).
   - Trigger settings → Concurrency control: On, degree of parallelism = 1.
3. (Khuyên dùng) Action **Delay** 2 phút để người dùng lưu xong file.
4. Action **HTTP**:
   - Method: `POST`
   - URI: `https://api.github.com/repos/Huanpro1239/Planning_vikoda/dispatches`
   - Headers: `Accept: application/vnd.github+json`,
     `Authorization: Bearer <PAT>`, `X-GitHub-Api-Version: 2022-11-28`
   - Body:

     ```json
     {"event_type": "sharepoint_stock_updated",
      "client_payload": {"source": "power-automate",
                         "file": "@{triggerOutputs()?['body/{FilenameWithExtension}']}"}}
     ```
5. Lưu flow, sửa thử 1 file XNT, kiểm tra Actions có run
   `Sync SharePoint Stock` với event `repository_dispatch`.

Sửa file ngoài 5 file nguồn trong thư mục cũng gửi event, nhưng Planning sẽ
bỏ qua nhờ `--skip-if-unchanged`. Các thay đổi trên workbook kế hoạch (FC,
`No kho`, `Danh_muc`, số ca) vẫn được watcher dự phòng phát hiện.
