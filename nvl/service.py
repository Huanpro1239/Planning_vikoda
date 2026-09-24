"""NVL synchronization use-case orchestration.

SharePoint/Graph I/O stays here; parsing, reconciliation and workbook mutation
remain in smaller domain modules.
"""

from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import socket
import time
from typing import Any

from sharepoint.retry import install_retry_after_support, retry_delay_seconds
from nvl.models import NVLConfig
from nvl.reconcile import read_nvl_source_stock, reconcile_nvl_target
from nvl.reporting import generate_nvl_error_report, generate_nvl_report
from nvl.workbook import patch_nvl_destination_workbook, verify_nvl_patched_workbook
from sharepoint.client import (
    GraphClient,
    GraphRequestError,
    get_access_token,
    is_retryable_graph_error,
)

def _is_network_timeout_or_reset(exc: Exception) -> bool:
    """Xác định các lỗi timeout hoặc ngắt kết nối mạng tạm thời."""
    if isinstance(exc, (TimeoutError, socket.timeout, ConnectionResetError, ConnectionRefusedError)):
        return True
    name = type(exc).__name__.lower()
    msg = str(exc).lower()
    if any(k in name for k in ("timeout", "connectionerror", "urlerror", "readtimeouterror")):
        return True
    if any(k in msg for k in ("timed out", "timeout", "connection reset", "connection refused", "remotely closed")):
        return True
    return False


def run_nvl_sync(
    config: NVLConfig,
    *,
    source_file: str | Path | None = None,
    target_file: str | Path | None = None,
    out_dir: str | Path | None = None,
    publish: bool = False,
    graph: GraphClient | None = None,
    max_publish_attempts: int = 3,
) -> dict[str, Any]:
    """Chạy quy trình đồng bộ tồn NVL ở chế độ offline hoặc online."""
    install_retry_after_support()

    if out_dir is None:
        out_dir = Path("offline_out/nvl") if (source_file or target_file) else Path(".")
    else:
        out_dir = Path(out_dir)

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass

    proposal_path = out_dir / "nvl_stock_proposal.xlsx"
    report_path = out_dir / "nvl_stock_report.json"

    # Dọn dẹp artifact cũ để không gây hiểu nhầm nếu lần chạy hiện tại lỗi sớm
    if proposal_path.exists():
        try:
            proposal_path.unlink()
        except OSError:
            pass
    if report_path.exists():
        try:
            report_path.unlink()
        except OSError:
            pass

    mode = "publish" if publish else ("offline" if (source_file or target_file) else "dry_run")
    source_rev_final = None
    target_rev_final = None
    current_phase = "init"
    last_error: Exception | None = None

    def _emit_error_and_raise(
        phase: str,
        err: Exception,
        attempt: int = 1,
        src_rev: str | None = None,
        tgt_rev: str | None = None,
        extra: dict[str, Any] | None = None,
    ) -> None:
        if proposal_path.exists():
            try:
                proposal_path.unlink()
            except OSError:
                pass
        err_rep = generate_nvl_error_report(
            config,
            mode=mode,
            phase=phase,
            attempt=attempt,
            error=err,
            source_revision=src_rev or source_rev_final,
            target_revision=tgt_rev or target_rev_final,
            extra=extra,
        )
        try:
            report_path.write_text(json.dumps(err_rep, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError:
            pass
        raise err

    # 1. Kiểm tra cấu hình và tham số đầu vào
    if (source_file or target_file) and publish:
        _emit_error_and_raise(
            "validate_input",
            ValueError("Chế độ offline (--source-file/--target-file) không thể kết hợp với cờ --publish."),
        )

    # 2. Chế độ OFFLINE (chạy từ file cục bộ)
    if source_file or target_file:
        if not (source_file and target_file):
            _emit_error_and_raise(
                "validate_input",
                ValueError("Phải cung cấp đồng thời cả --source-file và --target-file."),
            )
        src_path = Path(source_file)
        tgt_path = Path(target_file)
        if not src_path.exists():
            _emit_error_and_raise("offline_input", FileNotFoundError(f"Không tìm thấy file nguồn: {src_path}"))
        if not tgt_path.exists():
            _emit_error_and_raise("offline_input", FileNotFoundError(f"Không tìm thấy file đích: {tgt_path}"))

        current_phase = "offline_read"
        try:
            print(f"[OFFLINE] Đọc nguồn: {src_path}")
            source_bytes = src_path.read_bytes()
            print(f"[OFFLINE] Đọc đích: {tgt_path}")
            target_bytes = tgt_path.read_bytes()

            current_phase = "offline_reconcile"
            source_stock, source_metadata = read_nvl_source_stock(source_bytes, config)
            print(f"[OFFLINE] Đọc được {len(source_stock)} mã vật tư từ nguồn.")

            reconcile_res = reconcile_nvl_target(target_bytes, source_stock, config)
            print(
                f"[OFFLINE] Đối soát đích: {len(reconcile_res.changes)} ô đổi, "
                f"{len(reconcile_res.unchanged)} ô không đổi, "
                f"{len(reconcile_res.missing_in_source)} mã thiếu ở nguồn."
            )

            current_phase = "offline_patch"
            patched_bytes = patch_nvl_destination_workbook(target_bytes, reconcile_res, config)

            current_phase = "offline_verify"
            verify_nvl_patched_workbook(target_bytes, patched_bytes, reconcile_res, config)

            proposal_path.write_bytes(patched_bytes)
            print(f"[OFFLINE] Đã ghi proposal: {proposal_path}")

            report = generate_nvl_report(
                reconcile_res,
                config,
                mode="offline",
                reporting_period=source_metadata.get("reporting_period"),
            )
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            print(f"[OFFLINE] Đã ghi báo cáo: {report_path}")
            return report
        except Exception as exc:
            _emit_error_and_raise(current_phase, exc, attempt=1)

    # 3. Chế độ ONLINE (kết nối Microsoft Graph)
    if not getattr(config, "source_path", None):
        _emit_error_and_raise("validate_config", ValueError("Chưa cấu hình 'source.sharepoint_path' trong file cấu hình."))
    if not getattr(config, "target_path", None):
        _emit_error_and_raise(
            "validate_config",
            ValueError(
                f"Chưa cấu hình đường dẫn 'target.sharepoint_path' cho '{config.target_name}' trên SharePoint.\n"
                f"(Sourcedoc ID đã biết: {config.target_sourcedoc}). Vui lòng điền đường dẫn thư mục chính xác."
            ),
        )

    current_phase = "init_graph"
    try:
        if graph is None:
            token = get_access_token()
            graph = GraphClient(token)
        site_id = graph.get_site_id()
        drive_id = graph.get_default_drive_id(site_id)
    except Exception as exc:
        _emit_error_and_raise("init_graph", exc, attempt=1)

    for attempt in range(1, max_publish_attempts + 1):
        try:
            # 3.1. Đọc metadata nguồn trước khi download
            current_phase = "fetch_source_metadata"
            print(f"[ONLINE] Lượt {attempt}/{max_publish_attempts}: Đọc metadata nguồn...")
            source_item_before = graph.get_item_by_path(drive_id, config.source_path)
            source_etag_before = source_item_before.get("eTag")

            # 3.2. Tải snapshot nguồn
            current_phase = "download_source"
            source_bytes = graph.download_file(drive_id, source_item_before["id"])

            # 3.3. Kiểm tra tính tươi mới của nguồn ngay sau download
            current_phase = "verify_source_freshness"
            source_item_after = graph.get_item_by_path(drive_id, config.source_path)
            if source_item_after.get("eTag") != source_etag_before:
                print(f"[ONLINE] Nguồn đã thay đổi trong lúc download snapshot (lượt {attempt}/{max_publish_attempts}); tải lại...")
                last_error = RuntimeError(
                    f"Nguồn đã thay đổi trong lúc download snapshot (eTag trước: {source_etag_before}, sau: {source_item_after.get('eTag')})."
                )
                source_rev_final = source_item_after.get("eTag")
                if proposal_path.exists():
                    try:
                        proposal_path.unlink()
                    except OSError:
                        pass
                if attempt < max_publish_attempts:
                    time.sleep(1.0)
                    continue
                else:
                    _emit_error_and_raise(
                        "verify_source_freshness",
                        last_error,
                        attempt=attempt,
                        src_rev=source_rev_final,
                        tgt_rev=target_rev_final,
                    )
            source_rev_final = source_item_after.get("eTag")

            # 3.4. Tải snapshot đích
            current_phase = "download_target"
            target_item = graph.get_item_by_path(drive_id, config.target_path)
            target_rev_final = target_item.get("eTag")
            target_bytes = graph.download_file(drive_id, target_item["id"])
            target_sha256 = hashlib.sha256(target_bytes).hexdigest()

            # 3.4.1. Lưu và xác minh bản backup thực tế đích trước khi patch/upload
            current_phase = "pre_upload_backup"
            backup_raw_path = out_dir / "official_backup_target_raw.xlsx"
            backup_info_path = out_dir / "official_target_backup_info.json"
            try:
                backup_raw_path.write_bytes(target_bytes)
                saved_raw_bytes = backup_raw_path.read_bytes()
                saved_raw_sha = hashlib.sha256(saved_raw_bytes).hexdigest()
                if saved_raw_sha != target_sha256:
                    raise RuntimeError(
                        f"Xác minh SHA-256 backup đích thất bại: kỳ vọng {target_sha256}, thực tế {saved_raw_sha}"
                    )
                backup_meta = {
                    "name": target_item.get("name", config.target_name),
                    "sharepoint_path": config.target_path,
                    "item_id": target_item.get("id"),
                    "eTag": target_rev_final,
                    "sourcedoc": config.target_sourcedoc,
                    "sha256": target_sha256,
                    "size_bytes": len(target_bytes),
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                }
                backup_info_path.write_text(
                    json.dumps(backup_meta, indent=2, ensure_ascii=False), encoding="utf-8"
                )
                saved_meta = json.loads(backup_info_path.read_text(encoding="utf-8"))
                if saved_meta.get("sha256") != target_sha256:
                    raise RuntimeError(
                        "Xác minh metadata backup thất bại: SHA-256 trong file JSON không khớp với target_sha256."
                    )
                print(f"[ONLINE] Đã lưu và xác minh bản backup thực tế đích trước khi patch: {backup_raw_path} (SHA-256: {target_sha256})")
            except Exception as backup_exc:
                err = RuntimeError(f"Lưu hoặc xác minh backup thực tế đích thất bại trước khi upload: {backup_exc}")
                _emit_error_and_raise(
                    "pre_upload_backup",
                    err,
                    attempt=attempt,
                    src_rev=source_rev_final,
                    tgt_rev=target_rev_final,
                )

            # 3.5. Đối soát và lập kế hoạch patch
            current_phase = "reconcile"
            source_stock, source_metadata = read_nvl_source_stock(source_bytes, config)
            reconcile_res = reconcile_nvl_target(target_bytes, source_stock, config)

            # 3.6. Patch và xác minh proposal cục bộ
            current_phase = "patch"
            patched_bytes = patch_nvl_destination_workbook(target_bytes, reconcile_res, config)

            current_phase = "verify_proposal"
            verify_nvl_patched_workbook(target_bytes, patched_bytes, reconcile_res, config)
            proposal_path.write_bytes(patched_bytes)

            # Trường hợp dry-run: dừng tại đây và ghi nhận báo cáo đề xuất
            if not publish:
                current_phase = "finalize_dry_run"
                report = generate_nvl_report(
                    reconcile_res,
                    config,
                    mode="dry_run",
                    source_revision=source_rev_final,
                    target_revision=target_rev_final,
                    reporting_period=source_metadata.get("reporting_period"),
                )
                report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
                print(f"[ONLINE] Dry-run hoàn tất. Proposal: {proposal_path}, Report: {report_path}")
                return report

            # Trường hợp publish nhưng không có thay đổi (dữ liệu đã khớp)
            if not reconcile_res.changes:
                current_phase = "finalize_unchanged"
                report = generate_nvl_report(
                    reconcile_res,
                    config,
                    mode="publish",
                    source_revision=source_rev_final,
                    target_revision=target_rev_final,
                    reporting_period=source_metadata.get("reporting_period"),
                )
                report["status"] = "unchanged"
                report["message"] = "Dữ liệu tồn kho khớp hoàn toàn, không có ô nào cần upload."
                report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
                print("[ONLINE] Dữ liệu khớp 100%, không có ô nào cần upload.")
                return report

            # 3.7. Chuẩn bị publish: Kiểm tra lại nguồn trước khi upload đích
            current_phase = "pre_upload_source_check"
            source_item_pre = graph.get_item_by_path(drive_id, config.source_path)
            if source_item_pre.get("eTag") != source_rev_final:
                print(
                    f"[ONLINE] Nguồn đã bị thay đổi trước khi upload (lượt {attempt}/{max_publish_attempts}); "
                    "hủy lượt tải lên và làm mới snapshot..."
                )
                last_error = RuntimeError(
                    f"Nguồn đã bị thay đổi trước khi upload (eTag snapshot: {source_rev_final}, hiện tại: {source_item_pre.get('eTag')})."
                )
                source_rev_final = source_item_pre.get("eTag")
                if proposal_path.exists():
                    try:
                        proposal_path.unlink()
                    except OSError:
                        pass
                if attempt < max_publish_attempts:
                    time.sleep(1.0)
                    continue
                else:
                    _emit_error_and_raise(
                        "pre_upload_source_check",
                        last_error,
                        attempt=attempt,
                        src_rev=source_rev_final,
                        tgt_rev=target_rev_final,
                    )

            # 3.8. Upload file đích với If-Match ETag
            current_phase = "upload_target"
            print(f"[ONLINE] Đang upload file đích với ETag {target_rev_final}...")
            upload_response = None
            is_timeout_upload = False
            try:
                upload_response = graph.upload_file(
                    drive_id,
                    target_item["id"],
                    patched_bytes,
                    expected_etag=target_rev_final,
                )
                print(f"[ONLINE] Upload hoàn tất: {upload_response.get('name')} (Lượt {attempt})")
            except Exception as up_exc:
                if _is_network_timeout_or_reset(up_exc):
                    print(f"[ONLINE] Upload bị ngắt kết nối/timeout ({up_exc}). Chuyển sang xác minh trạng thái server...")
                    is_timeout_upload = True
                    upload_response = {
                        "id": target_item["id"],
                        "name": target_item.get("name"),
                        "status": "timeout_unconfirmed_ack",
                        "raw_timeout_error": str(up_exc),
                    }
                else:
                    # Lỗi không phải timeout (ví dụ 412 ETag conflict, 403 Forbidden...):
                    # Chắc chắn server chưa ghi nhận nội dung này. Ném lỗi ra ngoài xử lý.
                    raise up_exc

            # 3.9. Tải lại file đích sau upload để xác minh toàn vẹn trên server
            # LƯU Ý BẢO TOÀN VÀ ĐỐI SOÁT (Áp dụng cho CẢ trường hợp có ACK và trường hợp timeout):
            # - Khi server đã nhận file hoặc gặp timeout chờ phản hồi: tuyệt đối KHÔNG tự ý upload lại toàn bộ
            #   với ETag mới hoặc rollback ghi đè file của người khác nếu phát hiện xung đột.
            # - Nếu GET tải lại gặp lỗi mạng tạm thời (429, 503, timeout): retry đọc tối đa 3 lần.
            # - Nếu dữ liệu trên server không khớp với patch hoặc bị chỉnh sửa đồng thời ngoài cột D: dừng ngay,
            #   báo lỗi, ghi report failed với phase='post_upload_verify' và giữ upload_acknowledged=True.
            # - Nếu GET thất bại sau 3 lần thử: dừng ngay với trạng thái unverified.
            current_phase = "post_upload_verify"
            max_verify_download_attempts = 3
            post_upload_err = None
            verify_success = False
            verify_res: dict[str, Any] | None = None

            for v_attempt in range(1, max_verify_download_attempts + 1):
                try:
                    server_bytes = graph.download_file(drive_id, target_item["id"])
                    verify_res = verify_nvl_patched_workbook(
                        target_bytes, server_bytes, reconcile_res, config, is_server_comparison=True
                    )
                    verify_success = True
                    if is_timeout_upload:
                        print("[ONLINE] Server đã nhận đủ dữ liệu trước khi timeout; xác nhận thành công 100%.")
                        upload_response["status"] = "verified_post_timeout"
                    else:
                        print("[ONLINE] Đã tải lại file đích từ SharePoint và đối soát thành công 100%.")
                    break
                except Exception as verr:
                    post_upload_err = verr
                    # Kiểm tra nếu là lỗi đối soát dữ liệu (mismatch/conflict): dừng ngay, không retry GET!
                    is_data_mismatch = isinstance(verr, RuntimeError) and (
                        any(k in str(verr).lower() for k in ("xác minh thất bại", "bị thay đổi", "khác biệt", "không khớp", "biến đổi", "bị mất"))
                        or (not _is_network_timeout_or_reset(verr) and not isinstance(verr, GraphRequestError))
                    )
                    if is_data_mismatch:
                        prefix = "LỖI XÁC MINH SAU TIMEOUT" if is_timeout_upload else "LỖI XÁC MINH SAU UPLOAD"
                        print(f"[ONLINE] {prefix} (Xung đột dữ liệu trên server): {verr}")
                        break

                    # Nếu là lỗi Graph hoặc timeout tạm thời khi download: retry GET
                    is_temp = (
                        isinstance(verr, GraphRequestError) and is_retryable_graph_error(verr)
                    ) or _is_network_timeout_or_reset(verr)
                    if is_temp and v_attempt < max_verify_download_attempts:
                        delay = retry_delay_seconds(verr, 1.0) if isinstance(verr, GraphRequestError) else 1.0
                        print(
                            f"[ONLINE] Lỗi tạm thời khi tải lại để xác minh ({verr}). "
                            f"Thử lại lượt {v_attempt + 1}/{max_verify_download_attempts} sau {delay}s..."
                        )
                        time.sleep(delay)
                        continue
                    break

            if not verify_success:
                is_data_mismatch = isinstance(post_upload_err, RuntimeError) and (
                    any(k in str(post_upload_err).lower() for k in ("xác minh thất bại", "bị thay đổi", "khác biệt", "không khớp", "biến đổi", "bị mất"))
                    or (not _is_network_timeout_or_reset(post_upload_err) and not isinstance(post_upload_err, GraphRequestError))
                )
                v_status = "conflict_or_mismatch" if is_data_mismatch else "unverified"
                if is_timeout_upload:
                    err_msg = (
                        f"Upload bị timeout/mất kết nối và xác minh sau đó thất bại do dữ liệu trên server bị chỉnh sửa "
                        f"đồng thời hoặc không khớp với bản patch: {post_upload_err}. Dữ liệu trên server được giữ nguyên, "
                        f"tuyệt đối không tự ý upload lại."
                        if is_data_mismatch
                        else (
                            f"Upload bị timeout/mất kết nối và không thể tải lại file để xác minh sau "
                            f"{max_verify_download_attempts} lần thử: {post_upload_err}. Upload có thể đã được server ghi nhận."
                        )
                    )
                else:
                    err_msg = (
                        f"Upload đã được SharePoint ghi nhận nhưng xác minh sau upload thất bại do dữ liệu bị sửa đổi "
                        f"đồng thời hoặc không khớp: {post_upload_err}"
                        if is_data_mismatch
                        else (
                            f"Upload đã gửi thành công nhưng không thể tải lại file để xác minh sau "
                            f"{max_verify_download_attempts} lần thử: {post_upload_err}"
                        )
                    )
                final_err = RuntimeError(err_msg)
                _emit_error_and_raise(
                    "post_upload_verify",
                    final_err,
                    attempt=attempt,
                    src_rev=source_rev_final,
                    tgt_rev=target_rev_final,
                    extra={
                        "upload_acknowledged": not is_timeout_upload,
                        "upload_may_have_committed": True,
                        "upload_result": upload_response,
                        "verification_status": v_status,
                        "is_timeout_upload": is_timeout_upload,
                        "raw_verification_error": str(post_upload_err),
                    },
                )

            # 3.10. Ghi nhận báo cáo thành công (Chỉ khi xác minh thành công!)
            current_phase = "finalize_published"
            exempted_list = verify_res.get("exempted_parts", []) if verify_res else []
            report = generate_nvl_report(
                reconcile_res,
                config,
                mode="publish",
                source_revision=source_rev_final,
                target_revision=target_rev_final,
                reporting_period=source_metadata.get("reporting_period"),
                exempted_parts=exempted_list,
            )
            report["status"] = "published" if not reconcile_res.missing_in_source else "published_with_warnings"
            report["message"] = f"Đồng bộ và publish thành công: {len(reconcile_res.changes)} ô đã cập nhật lên SharePoint."
            report["upload_result"] = upload_response
            report["post_upload_verified"] = True
            report["exempted_server_metadata_parts"] = exempted_list
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            return report

        except Exception as exc:
            if current_phase in ("pre_upload_backup", "post_upload_verify"):
                raise

            last_error = exc
            is_retryable = False
            delay = 1.0

            if isinstance(exc, GraphRequestError):
                if exc.status_code == 412:
                    print(
                        f"[ONLINE] HTTP 412: File đích đã thay đổi đồng thời trên SharePoint "
                        f"(Lượt {attempt}/{max_publish_attempts})."
                    )
                    is_retryable = True
                    delay = 1.0
                elif is_retryable_graph_error(exc):
                    is_retryable = True
                    delay = retry_delay_seconds(exc, 3.0)
                    print(f"[ONLINE] Lỗi tạm thời Graph HTTP {exc.status_code} ({exc}). Chờ {delay}s...")
            elif _is_network_timeout_or_reset(exc):
                is_retryable = True
                delay = 2.0
                print(f"[ONLINE] Lỗi mạng tạm thời ({exc}). Chờ {delay}s...")

            if is_retryable and attempt < max_publish_attempts:
                if proposal_path.exists():
                    try:
                        proposal_path.unlink()
                    except OSError:
                        pass
                time.sleep(delay)
                continue

            # Fail fast cho lỗi không retryable hoặc khi đã hết lượt thử
            _emit_error_and_raise(
                current_phase,
                exc,
                attempt=attempt,
                src_rev=source_rev_final,
                tgt_rev=target_rev_final,
            )

    # Thoát vòng lặp do vượt quá số lần thử
    _emit_error_and_raise(
        current_phase,
        last_error or RuntimeError("Vượt quá số lần thử tải lên SharePoint do xung đột hoặc lỗi tạm thời liên tục."),
        attempt=max_publish_attempts,
        src_rev=source_rev_final,
        tgt_rev=target_rev_final,
    )


__all__ = ["run_nvl_sync"]
