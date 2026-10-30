.PHONY: install test sell-in sell-in-dry xuat-kho xuat-kho-dry clean

PYTHON ?= python3

install:
	$(PYTHON) -m pip install --upgrade pip
	$(PYTHON) -m pip install -r requirements.txt

# Test offline, không cần SharePoint.
test:
	$(PYTHON) -m unittest discover -s tests -v

# Chạy thật: cần MS_TENANT_ID, MS_CLIENT_ID (+ MS_AUTH_MODE=secret và MS_CLIENT_SECRET khi chạy ngoài GitHub Actions).
sell-in:
	$(PYTHON) -X utf8 -m sell_in

sell-in-dry:
	$(PYTHON) -X utf8 -m sell_in --dry-run

xuat-kho:
	$(PYTHON) -X utf8 -m xuat_kho

xuat-kho-dry:
	$(PYTHON) -X utf8 -m xuat_kho --dry-run

clean:
	find . -name "__pycache__" -type d -prune -exec rm -rf {} +
