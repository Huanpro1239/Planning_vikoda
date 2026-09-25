.PHONY: install test readiness nvl-readiness verify-release verify-nvl-release audit-nvl-releases restore-nvl-release nvl-ops-summary compile clean

PYTHON ?= python3

install:
	$(PYTHON) -m pip install --upgrade pip
	$(PYTHON) -m pip install -r requirements.txt

# Chạy toàn bộ test (offline, không cần SharePoint).
test:
	$(PYTHON) -m unittest discover -s tests -v

readiness:
	$(PYTHON) -X utf8 scripts/production_readiness.py

nvl-readiness:
	$(PYTHON) -X utf8 scripts/nvl_production_readiness.py

verify-release:
	@test -n "$(MANIFEST)" || (echo "Usage: make verify-release MANIFEST=<manifest.json> [STRICT=1]" && exit 2)
	$(PYTHON) -X utf8 scripts/verify_release.py "$(MANIFEST)" $(if $(STRICT),--strict,)

verify-nvl-release:
	@test -n "$(MANIFEST)" || (echo "Usage: make verify-nvl-release MANIFEST=<manifest.json> [STRICT=1]" && exit 2)
	$(PYTHON) -X utf8 scripts/verify_nvl_release.py "$(MANIFEST)" $(if $(STRICT),--strict,)

audit-nvl-releases:
	$(PYTHON) -X utf8 scripts/audit_nvl_releases.py $(if $(RELEASES_DIR),"$(RELEASES_DIR)",runtime-state/nvl/releases) $(if $(NO_GIT),--no-git,)

restore-nvl-release:
	@test -n "$(RELEASE)" || (echo "Usage: make restore-nvl-release RELEASE=<release_id> HISTORICAL=<workbook.xlsx> [PUBLISH=1 APPROVE=<release_id>]" && exit 2)
	@test -n "$(HISTORICAL)" || (echo "HISTORICAL=<nvl_open_po_proposal.xlsx> is required" && exit 2)
	$(PYTHON) -X utf8 scripts/restore_nvl_release.py "$(RELEASE)" --historical-workbook "$(HISTORICAL)" $(if $(PUBLISH),--publish --approve-release-id "$(APPROVE)",)

nvl-ops-summary:
	$(PYTHON) -X utf8 scripts/summarize_nvl_run.py

# Kiểm tra biên dịch mọi module (bắt lỗi cú pháp nhanh).
compile:
	$(PYTHON) -m compileall -q .

clean:
	rm -rf __pycache__ tests/__pycache__ .mypy_cache .pytest_cache
	rm -f planning_proposal.xlsx planning_schedule_report.json \
	      planning_input_revision.json planning_publish_decision.json \
	      production_readiness_report.json nvl_production_readiness_report.json planning_release_manifest.json nvl_release_manifest.json nvl_release_index.json nvl_release_index.csv nvl_recovery_report.json nvl_recovery_proposal.xlsx nvl_recovery_current_backup.xlsx
	rm -rf offline_out dry_run_artifacts
