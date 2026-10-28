"""Jalankan Data Flow Check dari terminal (cocok untuk cron / monitoring).

Pakai (dari folder utama repo):  python -m meta_ads.check_flow [--days 7]
Exit code 1 kalau ada pengecekan yang GAGAL. Butuh Python 3.11+.
"""
import argparse
import os
import sys
import tomllib
from datetime import date, timedelta
from pathlib import Path

from .checks import ICON, run_health_checks
from .meta_client import MetaClient

p = argparse.ArgumentParser()
p.add_argument("--days", type=int, default=7)
args = p.parse_args()

cfg = {}
f = Path(".streamlit/secrets.toml")
if f.exists():
    cfg = tomllib.loads(f.read_text())
get = lambda k, d=None: os.environ.get(k) or cfg.get(k, d)

token, acc = get("META_TOKEN"), str(get("AD_ACCOUNT_ID", ""))
if not token or not acc:
    sys.exit("META_TOKEN / AD_ACCOUNT_ID belum diisi.")
acc = acc if acc.startswith("act_") else f"act_{acc}"

client = MetaClient(token, get("API_VERSION", "v23.0"))
checks, _ = run_health_checks(client, acc, date.today() - timedelta(days=args.days), date.today(), get("BUSINESS_ID"))
for c in checks:
    print(f"{ICON[c.status]} [{c.group}] {c.name}: {c.detail}")
sys.exit(1 if any(c.status == "fail" for c in checks) else 0)
