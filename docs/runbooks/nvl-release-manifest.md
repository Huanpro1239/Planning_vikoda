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
