.PHONY: install test fit-ban-hang fit-ban-hang-dry sell-in sell-in-dry xuat-kho xuat-kho-dry nxt-ke-toan nxt-ke-toan-dry clean

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

nxt-ke-toan:
	$(PYTHON) -X utf8 -m nxt_ke_toan

nxt-ke-toan-dry:
	$(PYTHON) -X utf8 -m nxt_ke_toan --dry-run

# FIT ĐH Bán trong kỳ NPP qua Worker erp-nhan-hang: cần FIT_CF_ACCESS_CLIENT_ID/SECRET (+ MS_* khi ghi SharePoint).
fit-ban-hang:
	$(PYTHON) -X utf8 -m fit_ban_hang

fit-ban-hang-dry:
	$(PYTHON) -X utf8 -m fit_ban_hang --dry-run

clean:
	find . -name "__pycache__" -type d -prune -exec rm -rf {} +
