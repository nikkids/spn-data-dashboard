"""Client tipis untuk Meta Graph API: log tiap request, pagination, error handling."""
from __future__ import annotations

import json
import time
from datetime import datetime

import pandas as pd
import requests

GRAPH = "https://graph.facebook.com"

HINTS = {
    190: "Token invalid / expired / di-revoke. Generate token baru di Business Settings → System users.",
    200: "Permission kurang. Cek scope yang dicentang saat generate token, dan pastikan asset (ad account / Page) sudah di-assign ke system user.",
    10: "Permission kurang. Cek scope token dan assignment asset ke system user.",
    100: "Parameter / field tidak valid, atau ID salah. Pastikan AD_ACCOUNT_ID pakai prefix act_.",
    17: "Kena rate limit. Tunggu beberapa menit lalu coba lagi.",
    4: "Kena rate limit aplikasi. Tunggu beberapa menit.",
    32: "Kena rate limit Page. Tunggu beberapa menit.",
    613: "Kena rate limit. Tunggu beberapa menit.",
    80004: "Kena rate limit Ads Insights. Kurangi frekuensi refresh / rentang tanggal.",
    2635: "Versi API sudah deprecated. Ubah API_VERSION di secrets ke versi terbaru.",
}


class MetaAPIError(Exception):
    def __init__(self, message, code=None, subcode=None, status=None, path=""):
        super().__init__(message)
        self.code, self.subcode, self.status, self.path = code, subcode, status, path

    @property
    def hint(self) -> str:
        if "api access blocked" in str(self).lower():
            return ("Meta memblokir akses API untuk app/token ini; ini BUKAN soal scope. Cek Alerts/Restrictions "
                    "di dashboard app (developers.facebook.com), status akun di Business Support Home, dan apa yang "
                    "berubah sejak terakhir hijau (token baru, setting app, assignment).")
        return HINTS.get(self.code, "")


class MetaClient:
    def __init__(self, token: str, version: str = "v23.0", timeout: int = 60):
        self.token, self.version, self.timeout = token, version, timeout
        self.log: list[dict] = []

    def _scrub(self, text: str, *tokens) -> str:
        for t in (self.token, *tokens):
            if t:
                text = text.replace(t, "[token]")
        return text

    def _record(self, **kw):
        kw["time"] = datetime.now().strftime("%H:%M:%S")
        self.log.append(kw)
        del self.log[:-300]

    def get(self, path, params=None, *, paginate=False, max_pages=30, token=None):
        """GET ke Graph API. paginate=True -> kembalikan list gabungan semua halaman."""
        url = f"{GRAPH}/{self.version}/{path.lstrip('/')}"
        p = dict(params or {})
        p["access_token"] = token or self.token
        t0, data, last, nxt, pages = time.perf_counter(), [], None, url, 0
        while nxt:
            try:
                r = requests.get(nxt, params=p if pages == 0 else None, timeout=self.timeout)
            except requests.RequestException as e:
                msg = self._scrub(str(e), token)
                self._record(path=path, status="NET", rows=0, ms=self._ms(t0), error=msg[:200])
                raise MetaAPIError(f"Koneksi gagal: {msg}", path=path) from None
            try:
                js = r.json()
            except ValueError:
                js = {}
            if r.status_code != 200 or "error" in js:
                err = js.get("error", {})
                msg = self._scrub(err.get("message", r.text[:200]), token)
                detail = f"[code {err.get('code')} / subcode {err.get('error_subcode')} / trace {err.get('fbtrace_id')}]"
                self._record(path=path, status=r.status_code, rows=0, ms=self._ms(t0),
                             error=f"{msg} {detail}"[:300])
                raise MetaAPIError(f"{msg} {detail}", err.get("code"), err.get("error_subcode"), r.status_code, path)
            pages += 1
            if paginate and isinstance(js.get("data"), list):
                data += js["data"]
                nxt = js.get("paging", {}).get("next") if pages < max_pages else None
            else:
                last, nxt = js, None
        if paginate:
            result, n = data, len(data)
        else:
            result = last
            n = len(last["data"]) if isinstance(last.get("data"), list) else 1
        self._record(path=path, status=200, rows=n, ms=self._ms(t0), error="")
        return result

    @staticmethod
    def _ms(t0):
        return round((time.perf_counter() - t0) * 1000)


# ---------- Insights helper ----------
BASE_FIELDS = ["spend", "impressions", "reach", "frequency", "clicks",
               "inline_link_clicks", "ctr", "cpc", "cpm"]
LEVEL_FIELDS = {
    "account": ["account_name"],
    "campaign": ["campaign_id", "campaign_name"],
    "adset": ["campaign_name", "adset_id", "adset_name"],
    "ad": ["campaign_name", "adset_name", "ad_id", "ad_name"],
}


def _num(v) -> float:
    """Angka dari field insights: skalar, string, atau list [{'value': ..}]."""
    if isinstance(v, list):
        return float(sum(_num(x.get("value", 0)) if isinstance(x, dict) else _num(x) for x in v))
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def get_insights(client, account_id, level, since, until, *, daily=False, breakdowns=(),
                 with_actions=False, fields=None, extra=()) -> pd.DataFrame:
    f = list(fields) if fields else BASE_FIELDS + LEVEL_FIELDS[level]
    if with_actions:
        f += ["actions", "cost_per_action_type"]
    f += list(extra)
    params = {
        "level": level,
        "fields": ",".join(f),
        "time_range": json.dumps({"since": str(since), "until": str(until)}),
        "limit": 500,
    }
    if daily:
        params["time_increment"] = 1
    if breakdowns:
        params["breakdowns"] = ",".join(breakdowns)
    df = pd.DataFrame(client.get(f"{account_id}/insights", params, paginate=True))
    if df.empty:
        return df
    for c in BASE_FIELDS:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        elif not fields:
            df[c] = 0.0
    for c in extra:
        # Keep a missing optional field absent so callers can distinguish "not returned"
        # from a real metric whose value is 0.0.
        if c in df:
            df[c] = df[c].map(_num)
    if "date_start" in df:
        df["date"] = pd.to_datetime(df["date_start"])
    return df
