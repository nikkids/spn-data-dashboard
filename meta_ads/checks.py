"""Pengecekan 'apakah data beneran masuk dan benar?' untuk Meta Ads API."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timezone

from .meta_client import MetaAPIError, get_insights

ICON = {"pass": "PASS", "warn": "WARN", "fail": "FAIL", "info": "INFO"}
ACCOUNT_STATUS = {1: "ACTIVE", 2: "DISABLED", 3: "UNSETTLED", 7: "PENDING_RISK_REVIEW",
                  8: "PENDING_SETTLEMENT", 9: "IN_GRACE_PERIOD", 100: "PENDING_CLOSURE",
                  101: "CLOSED", 201: "ANY_ACTIVE", 202: "ANY_CLOSED"}
REQUIRED_SCOPES = ["ads_read"]
OPTIONAL_SCOPES = ["business_management", "pages_read_engagement", "pages_show_list"]
WRITE_SCOPES = ["ads_management", "pages_manage_ads"]


@dataclass
class Check:
    group: str
    name: str
    status: str  # pass | warn | fail | info
    detail: str


def _close(a: float, b: float, tol=0.01) -> bool:
    return abs(a - b) <= max(tol * max(abs(a), abs(b)), 1.0)


def run_health_checks(client, account_id, since, until, business_id=None, *, include_business_pages=True):
    out: list[Check] = []
    ctx: dict = {}

    def add(group, name, status, detail):
        out.append(Check(group, name, status, str(detail)))

    def msg(e):
        return f"{e}" + (f" → {e.hint}" if getattr(e, "hint", "") else "")

    # ---- 1. Token ----
    try:
        me = client.get("me", {"fields": "id,name"})
        add("Token", "Token valid", "pass", f"Terhubung sebagai {me.get('name', '?')} (ID {me.get('id')})")
    except MetaAPIError as e:
        add("Token", "Token valid", "fail", msg(e))
        return out, ctx

    scopes = None
    try:
        d = client.get("debug_token", {"input_token": client.token}).get("data", {})
        scopes = d.get("scopes") or []
        if d.get("is_valid") is False:
            add("Token", "Status token", "fail", "Token ditandai tidak valid oleh Meta")
        exp = d.get("expires_at", 0) or 0
        if exp == 0:
            add("Token", "Masa berlaku", "pass", "Tidak ada expiry (never expires)")
        else:
            when = datetime.fromtimestamp(exp, tz=timezone.utc)
            days = (when - datetime.now(timezone.utc)).days
            add("Token", "Masa berlaku", "warn" if days < 7 else "pass",
                f"Expired {when:%Y-%m-%d} (sisa {days} hari)")
    except MetaAPIError:
        try:
            perms = client.get("me/permissions", paginate=True)
            scopes = [p["permission"] for p in perms if p.get("status") == "granted"]
        except MetaAPIError:
            scopes = None
        add("Token", "Masa berlaku", "info", "Tidak bisa dicek lewat debug_token")

    if scopes is None:
        add("Token", "Scope", "warn", "Scope token tidak bisa dibaca")
    else:
        ctx["scopes"] = scopes
        miss = [s for s in REQUIRED_SCOPES if s not in scopes]
        add("Token", "Scope wajib (ads_read)", "fail" if miss else "pass",
            "ads_read belum ada, generate ulang token" if miss else "ads_read ada")
        have = [s for s in OPTIONAL_SCOPES if s in scopes]
        lack = [s for s in OPTIONAL_SCOPES if s not in scopes]
        add("Token", "Scope tambahan", "info",
            f"Ada: {', '.join(have) or '-'} | Belum ada: {', '.join(lack) or '-'} (tab Business & Pages butuh ini)")
        risky = [s for s in WRITE_SCOPES if s in scopes]
        if risky:
            add("Token", "Scope berisiko", "warn",
                f"Token punya izin tulis ({', '.join(risky)}). Dashboard read-only nggak butuh ini; sebaiknya generate ulang tanpa scope tsb.")

    # ---- 2. Ad account ----
    try:
        acc = client.get(account_id, {"fields": "name,account_status,currency,timezone_name"})
        st_code = acc.get("account_status")
        ctx["account"] = acc
        add("Ad account", "Akses ad account", "pass" if st_code == 1 else "warn",
            f"{acc.get('name')} | status {ACCOUNT_STATUS.get(st_code, st_code)} | "
            f"{acc.get('currency')} | timezone {acc.get('timezone_name')}")
    except MetaAPIError as e:
        add("Ad account", "Akses ad account", "fail", msg(e))
        return out, ctx

    biz_id = business_id
    try:
        b = client.get(account_id, {"fields": "business"}).get("business") or {}
        biz_id = biz_id or b.get("id")
    except MetaAPIError:
        pass

    # ---- 3. Insights ----
    try:
        daily = get_insights(client, account_id, "account", since, until, daily=True)
    except MetaAPIError as e:
        add("Insights", "Endpoint insights", "fail", msg(e))
        daily = None
    if daily is not None:
        ctx["daily"] = daily
        if daily.empty:
            add("Insights", "Endpoint insights", "warn",
                "Request sukses tapi data kosong. Normal kalau tidak ada iklan tayang di periode ini; coba perlebar rentang tanggal.")
        else:
            add("Insights", "Endpoint insights", "pass", f"{len(daily)} baris harian diterima")
            n_days = (until - since).days + 1
            have_days = daily["date_start"].nunique()
            add("Insights", "Cakupan hari", "info",
                f"{have_days} dari {n_days} hari punya data (hari tanpa iklan tayang memang tidak muncul)")
            active = daily[daily["impressions"] > 0]
            if not active.empty:
                last = active["date_start"].max()
                lag = (date.today() - date.fromisoformat(last)).days
                add("Insights", "Kesegaran data", "pass" if lag <= 2 else "warn",
                    f"Data tayang terakhir: {last} ({lag} hari lalu)")

            # integritas
            problems = []
            if daily.duplicated("date_start").any():
                problems.append("ada tanggal duplikat")
            if (daily[["spend", "impressions", "clicks", "reach"]] < 0).any().any():
                problems.append("ada nilai negatif")
            if daily[["spend", "impressions", "clicks"]].isna().any().any():
                problems.append("ada nilai kosong")
            if (daily["clicks"] > daily["impressions"]).any():
                problems.append("clicks > impressions")
            add("Integritas", "Struktur & logika baris", "fail" if problems else "pass",
                "; ".join(problems) if problems else "Tidak ada duplikat, negatif, kosong, atau clicks > impressions")
            nz = daily[daily["impressions"] > 0]
            if not nz.empty:
                calc = nz["clicks"] / nz["impressions"] * 100
                gap = (calc - nz["ctr"]).abs().max()
                add("Integritas", "CTR konsisten", "pass" if gap <= 0.05 else "warn",
                    f"Selisih maks hitung-ulang vs CTR dari API: {gap:.3f} poin")

            # rekonsiliasi
            try:
                agg = get_insights(client, account_id, "account", since, until)
                ctx["agg"] = agg
                if not agg.empty:
                    a = agg.iloc[0]
                    diffs, bad = [], False
                    for m in ("spend", "impressions", "clicks"):
                        s, t = float(daily[m].sum()), float(a[m])
                        ok = _close(s, t)
                        bad |= not ok
                        diffs.append(f"{m}: harian {s:,.0f} vs total {t:,.0f} {'ok' if ok else 'SELISIH'}")
                    add("Rekonsiliasi", "Total harian = total periode", "fail" if bad else "pass", " | ".join(diffs))
            except MetaAPIError as e:
                add("Rekonsiliasi", "Total harian = total periode", "warn", msg(e))
            try:
                camp = get_insights(client, account_id, "campaign", since, until)
                ctx["camp"] = camp
                if not camp.empty and "agg" in ctx and not ctx["agg"].empty:
                    cs, t = float(camp["spend"].sum()), float(ctx["agg"].iloc[0]["spend"])
                    ok = _close(cs, t)
                    add("Rekonsiliasi", "Total campaign = total akun", "pass" if ok else "warn",
                        f"{len(camp)} campaign, spend {cs:,.0f} vs akun {t:,.0f}"
                        + ("" if ok else " (campaign yang sudah dihapus bisa tidak ikut ter-list)"))
            except MetaAPIError as e:
                add("Rekonsiliasi", "Total campaign = total akun", "warn", msg(e))

    # Dashboard iklan Instagram tidak memerlukan data Business Portfolio / Facebook Pages.
    if not include_business_pages:
        return out, ctx

    # ---- 4. Business & Pages (opsional) ----
    if not biz_id:
        add("Business", "Business portfolio", "info", "Business ID tidak ketemu. Isi BUSINESS_ID di secrets kalau mau cek bagian ini.")
        return out, ctx
    ctx["business_id"] = biz_id
    try:
        owned = client.get(f"{biz_id}/owned_ad_accounts", {"fields": "id,name", "limit": 100}, paginate=True)
        inc = any(x["id"] == account_id for x in owned)
        add("Business", "Ad account di portfolio", "pass" if inc else "info",
            f"{len(owned)} ad account; {account_id} {'termasuk' if inc else 'tidak termasuk (mungkin milik portfolio lain)'}")
    except MetaAPIError as e:
        add("Business", "Ad account di portfolio", "warn", msg(e))
    try:
        pages = client.get(f"{biz_id}/owned_pages", {"fields": "id,name", "limit": 100}, paginate=True)
    except MetaAPIError as e:
        add("Pages", "Daftar Page", "warn", msg(e))
        return out, ctx
    if not pages:
        add("Pages", "Daftar Page", "warn", "Belum ada Page di portfolio / belum di-assign ke system user")
        return out, ctx
    add("Pages", "Daftar Page", "pass", f"{len(pages)} Page: {', '.join(p['name'] for p in pages[:5])}")
    pid = pages[0]["id"]
    try:
        ptok = client.get(pid, {"fields": "access_token"}).get("access_token")
        info = client.get(pid, {"fields": "name,fan_count,followers_count"}, token=ptok)
        add("Pages", "Akses data Page", "pass",
            f"{info.get('name')} | fans {info.get('fan_count', '-')} | followers {info.get('followers_count', '-')}")
        posts = client.get(f"{pid}/posts", {"fields": "id,created_time", "limit": 5}, token=ptok)
        n = len(posts.get("data", []))
        add("Pages", "Akses post Page", "pass" if n else "info", f"{n} post terbaru terbaca")
    except MetaAPIError as e:
        add("Pages", "Akses data Page", "warn", msg(e))
    return out, ctx
