.PHONY: install test compile clean

PYTHON ?= python3

install:
	$(PYTHON) -m pip install --upgrade pip
	$(PYTHON) -m pip install -r requirements.txt

# Chạy toàn bộ test (offline, không cần SharePoint).
test:
	$(PYTHON) -m unittest discover -s tests -v

# Kiểm tra biên dịch mọi module (bắt lỗi cú pháp nhanh).
compile:
	$(PYTHON) -m compileall -q *.py tests

clean:
	rm -rf __pycache__ tests/__pycache__ .mypy_cache .pytest_cache
	rm -f planning_proposal.xlsx planning_schedule_report.json \
	      planning_input_revision.json planning_publish_decision.json
	rm -rf offline_out dry_run_artifacts
