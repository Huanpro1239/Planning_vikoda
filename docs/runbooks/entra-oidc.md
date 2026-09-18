# Chuyển GitHub Actions sang Microsoft Entra OIDC

Mục tiêu: GitHub Actions lấy Microsoft Graph access token bằng workload identity federation, không cần lưu `MS_CLIENT_SECRET` dài hạn.

> Trạng thái hiện tại: production workflows dùng OIDC mặc định (`MS_AUTH_MODE=oidc`); không còn phụ thuộc `MS_CLIENT_SECRET` trong workflow.

## 1. Identity của repository

Repository này được tạo sau khi GitHub áp dụng immutable OIDC subject cho repository mới.

- Owner: `Huanpro1239`
- Owner ID: `213777839`
- Repository: `Planning_vikoda`
- Repository ID: `1354289069`
- Production branch: `main`
- Issuer: `https://token.actions.githubusercontent.com`
- Audience: `api://AzureADTokenExchange`
- Immutable subject cho `main`:

```text
repo:Huanpro1239@213777839/Planning_vikoda@1354289069:ref:refs/heads/main
```

Không thay subject này bằng `repo:Huanpro1239/Planning_vikoda:ref:refs/heads/main` nếu GitHub OIDC token của repo đang dùng immutable subject.

## 2. Tạo Federated Credential trong Microsoft Entra

Dùng đúng App Registration đang có `Application (client) ID` bằng GitHub secret `MS_CLIENT_ID`.

Trong Microsoft Entra admin center:

1. `App registrations` → mở app hiện tại.
2. `Certificates & secrets` → `Federated credentials` → `Add credential`.
3. Chọn kiểu cho phép nhập issuer/subject trực tiếp (hoặc tạo bằng Microsoft Graph/Azure CLI).
4. Nhập:
   - Issuer: `https://token.actions.githubusercontent.com`
   - Subject: `repo:Huanpro1239@213777839/Planning_vikoda@1354289069:ref:refs/heads/main`
   - Audience: `api://AzureADTokenExchange`
   - Name gợi ý: `github-planning-vikoda-main`
5. Lưu credential.

Các Microsoft Graph application permissions/admin consent hiện có của app giữ nguyên. Federation chỉ thay cách app chứng minh danh tính, không tự cấp thêm Graph permission.

### Azure CLI tương đương

Nếu dùng CLI, `--id` nên là Object ID của App Registration (không nhầm với client ID):

```json
{
  "name": "github-planning-vikoda-main",
  "issuer": "https://token.actions.githubusercontent.com",
  "subject": "repo:Huanpro1239@213777839/Planning_vikoda@1354289069:ref:refs/heads/main",
  "audiences": ["api://AzureADTokenExchange"]
}
```

```bash
az ad app federated-credential create \
  --id <APP_OBJECT_ID> \
  --parameters credential.json
```

## 3. Cấu hình OIDC trên GitHub

Hai production workflows đã có:

```yaml
permissions:
  id-token: write
```

và gọi trực tiếp các entrypoint production; các module lấy token qua `sharepoint.auth` / `sharepoint.client`.

Trong GitHub repository:

1. `Settings` → `Secrets and variables` → `Actions` → `Variables`.
2. Tạo repository variable:

```text
MS_AUTH_MODE = oidc
```

Production workflow không đọc `MS_CLIENT_SECRET`. OIDC mode **không fallback âm thầm** sang secret; nếu federation sai, workflow phải đỏ để lỗi cấu hình hiện rõ.

## 4. Nghiệm thu trước khi bỏ secret

Thực hiện theo thứ tự:

1. Chạy `Sync SharePoint Stock` thủ công với `publish=false`.
2. Xác nhận bước Planning đọc SharePoint và tạo proposal thành công.
3. Chạy `Sync SharePoint NVL Stock` thủ công với `publish=false`.
4. Xác nhận cả tồn NVL và open PO resolve/đọc đúng nguồn.
5. Sau hai dry-run xanh, chạy controlled publish theo quy trình hiện tại và xác minh workbook SharePoint.
6. Theo dõi ít nhất một lần schedule production.

Nếu lỗi Entra có dạng `AADSTS70021`, `AADSTS700213` hoặc tương tự về federated identity, kiểm tra trước tiên issuer, audience và subject có khớp tuyệt đối token GitHub hay không.

## 5. Trạng thái cutover production

Production workflows hiện đã hoàn tất cutover:
1. Không khai báo hoặc sử dụng `MS_CLIENT_SECRET`.
2. Giữ `MS_TENANT_ID`, `MS_CLIENT_ID` và `MS_AUTH_MODE=oidc`.
3. Client secret cũ trong App Registration chỉ nên giữ nếu còn workload khác ngoài repository sử dụng.

Không xóa client secret bên Entra nếu cùng secret/app đang được một hệ thống khác ngoài repository này sử dụng mà chưa migrate.

## 6. Rollback

Trong giai đoạn chuyển tiếp, nếu OIDC có sự cố:

1. Đổi repository variable `MS_AUTH_MODE` từ `oidc` về `secret`.
2. Giữ `MS_CLIENT_SECRET` tồn tại cho tới khi thời gian quan sát OIDC hoàn tất.

Sau cutover cuối và secret đã bị revoke, rollback phải là sửa federated credential/OIDC, không quay lại secret mới trừ khi có quyết định bảo mật riêng.
