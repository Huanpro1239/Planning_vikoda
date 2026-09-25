.PHONY: install test readiness nvl-readiness verify-release compile clean

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

# Kiểm tra biên dịch mọi module (bắt lỗi cú pháp nhanh).
compile:
	$(PYTHON) -m compileall -q .

clean:
	rm -rf __pycache__ tests/__pycache__ .mypy_cache .pytest_cache
	rm -f planning_proposal.xlsx planning_schedule_report.json \
	      planning_input_revision.json planning_publish_decision.json \
	      production_readiness_report.json nvl_production_readiness_report.json planning_release_manifest.json
	rm -rf offline_out dry_run_artifacts
