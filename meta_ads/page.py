"""Halaman Meta Ads untuk SPN Data Explorer.

Dipindahkan dari aplikasi Meta Ads yang berdiri sendiri (app.py) menjadi satu fungsi `render()`
yang dipanggil dari menu "Meta Ads" di app-publish.py. Perbedaan dengan app aslinya:

  * tidak memanggil st.set_page_config dan tidak menyuntik CSS global (tema SPN tetap dipakai)
  * tampilan ringkasan, Traffic, Leads, dan koneksi dibuat ringkas mengikuti tema SPN
  * laporan hanya menampilkan metrik iklan dan tidak mengambil data Facebook Pages
  * semua key session_state diawali "mx_" supaya tidak bentrok dengan filter dashboard SPN
  * tombol Refresh hanya menghapus cache milik Meta Ads, bukan seluruh cache aplikasi
  * laporan lengkap dapat diunduh sebagai workbook Excel dan PDF terformat
  * kalau secrets belum diisi, hanya halaman ini yang menampilkan pesan error
"""
from __future__ import annotations

import hmac
from io import BytesIO
import html as html_lib
import os
import re
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

from .checks import ACCOUNT_STATUS, run_health_checks
from .kpi import action_map, action_totals, detect, div, parse_keywords
from .meta_client import MetaAPIError, MetaClient, get_insights


def cfg(key, default=None):
    try:
        return st.secrets[key]
    except Exception:
        return os.environ.get(key, default)


METRICS = ["spend", "impressions", "reach", "clicks", "ctr", "cpc", "cpm", "frequency"]
AD_FIELDS = ["ad_name", "campaign_name", "impressions", "spend"]
VISIT_FIELD = "instagram_profile_visits"
FOLLOW_FIELD = "instagram_profile_follow"
MONTHS_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
             "Juli", "Agustus", "September", "Oktober", "November", "Desember"]

FOLLOW_KW = parse_keywords(cfg("FOLLOW_KEYWORDS", "get follow, getfollow, ig follow, instagram follow, follower, followers"))
TRAFFIC_KW = parse_keywords(cfg("TRAFFIC_KEYWORDS", "traffic, profile visit, profile visits"))
LEADS_KW = parse_keywords(cfg("LEADS_KEYWORDS", "leads, lead, ctwa, click to whatsapp, whatsapp lead, wa lead"))


@dataclass
class Ctx:
    """Koneksi dan periode yang sedang aktif; dioper ke semua fungsi render."""
    client: MetaClient
    account: str
    version: str
    business_id: str | None
    start: date
    end: date


# ---------------- format angka ----------------
money = lambda value: f"Rp {value:,.0f}"


def idn(value, dec=0):
    return f"{float(value or 0):,.{dec}f}".replace(",", "§").replace(".", ",").replace("§", ".")


def rp(value):
    return "-" if value is None else f"Rp {idn(value)}"


# ---------------- klasifikasi campaign ----------------
def norm_name(value):
    return re.sub(r"[\s_\-.]+", " ", str(value or "").lower()).strip()


def campaign_category(name):
    """Classify by campaign objective/name. Explicit leading label takes precedence.

    This intentionally gives names beginning with 'Leads' priority over words such as
    '(Traffic)' later in the name, e.g. 'Leads - SPN 2026 (Traffic)'.
    """
    n = norm_name(name)
    if not n:
        return "Lainnya"
    if re.match(r"^(leads?|ctwa)\b", n):
        return "Leads"
    if re.match(r"^traffic\b", n):
        return "Traffic"
    lead_terms = [norm_name(x) for x in LEADS_KW if norm_name(x)]
    traffic_terms = [norm_name(x) for x in TRAFFIC_KW if norm_name(x)]
    if any(term in n for term in lead_terms):
        return "Leads"
    if any(term in n for term in traffic_terms):
        return "Traffic"
    # Older Get Follow campaigns should not be silently mixed into Traffic or Leads.
    if any(norm_name(x) in n for x in FOLLOW_KW if norm_name(x)):
        return "Lainnya"
    return "Lainnya"


def tag_campaigns(df):
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.copy()
    out = df.copy()
    if "campaign_name" not in out:
        out["campaign_name"] = ""
    out["dashboard_category"] = out["campaign_name"].map(campaign_category)
    for col in ("spend", "impressions", "reach", "clicks", "ctr", "frequency"):
        if col not in out:
            out[col] = 0.0
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    return out


# ---------------- pengambilan data (cache) ----------------
@st.cache_data(ttl=600, show_spinner=False)
def load_insights(_client, account, version, level, since, until, daily=False,
                  breakdowns=(), actions=False, fields=None, extra=()):
    return get_insights(_client, account, level, since, until, daily=daily,
                        breakdowns=breakdowns, with_actions=actions, fields=fields, extra=extra)


@st.cache_data(ttl=600, show_spinner="Mengambil data campaign dari Meta...")
def load_campaigns(_client, account, version, since, until):
    """Retrieve campaign metrics with native IG fields, gracefully retrying unsupported fields."""
    attempts = ((VISIT_FIELD, FOLLOW_FIELD), (VISIT_FIELD,), (FOLLOW_FIELD,), ())
    last_error = None
    for optional in attempts:
        try:
            return get_insights(_client, account, "campaign", since, until,
                                with_actions=True, extra=optional)
        except MetaAPIError as exc:
            if getattr(exc, "code", None) != 100:
                raise
            last_error = exc
    if last_error:
        raise last_error
    return get_insights(_client, account, "campaign", since, until, with_actions=True)


@st.cache_data(ttl=600, show_spinner=False)
def graph(_client, version, path, params=(), paginate=True, token=None):
    return _client.get(path, dict(params), paginate=paginate, token=token)


def clear_meta_cache():
    """Hapus cache Meta Ads saja. Jangan pakai st.cache_data.clear(): itu ikut menghapus cache data SPN."""
    for fn in (load_insights, load_campaigns, graph, _spn_registrations):
        fn.clear()
    st.session_state.pop("mx_full_report_files", None)
    st.session_state.pop("mx_full_report_files", None)


def safe_run(fn, *args, **kwargs):
    try:
        return fn(*args, **kwargs)
    except MetaAPIError as exc:
        st.error(f"Meta API: {exc}")
        if getattr(exc, "hint", ""):
            st.caption(exc.hint)
        return None


def campaign_data(ctx: Ctx):
    result = safe_run(load_campaigns, ctx.client, ctx.account, ctx.version, str(ctx.start), str(ctx.end))
    if result is None:
        return None
    return tag_campaigns(result)


# ---------------- hasil / actions ----------------
def lead_action_choice(df):
    """Select one primary result action automatically (never sum overlapping result types)."""
    totals = action_totals(df)
    if totals.empty:
        return None
    candidates = []
    for action_type, count in totals.items():
        lower = str(action_type).lower()
        if "lead" in lower:
            rank = 0
        elif "messaging_conversation_started" in lower or "conversation_started" in lower:
            rank = 1
        elif "messaging_first_reply" in lower or "contact" in lower:
            rank = 2
        else:
            continue
        candidates.append((rank, -float(count), str(action_type)))
    if not candidates:
        return None
    candidates.sort()
    return candidates[0][2]


def add_action_counts(df):
    out = df.copy()
    if "actions" not in out:
        out["actions"] = [[] for _ in range(len(out))]
    maps = out["actions"].map(action_map)
    out["_action_map"] = maps
    return out


def prepare_results(df):
    """Attach stable native/action-based profile visit counts to campaign rows."""
    if df is None or df.empty:
        return df
    out = add_action_counts(df)
    native_visits = 0.0
    if VISIT_FIELD in out.columns:
        native_visits = pd.to_numeric(out[VISIT_FIELD], errors="coerce").fillna(0).sum()
    visit_types = detect(action_totals(out), "visit")
    if native_visits > 0:
        out["profile_visits"] = pd.to_numeric(out[VISIT_FIELD], errors="coerce").fillna(0.0)
    else:
        out["profile_visits"] = out["_action_map"].map(lambda m: sum(float(m.get(t, 0) or 0) for t in visit_types)).astype(float)
    return out


def action_rows(df):
    rows = []
    if df is None or df.empty or "actions" not in df:
        return pd.DataFrame(columns=["campaign_name", "action_type", "hasil", "spend"])
    for _, row in df.iterrows():
        actions = row.get("actions")
        if not isinstance(actions, list):
            continue
        for item in actions:
            try:
                count = float(item.get("value", 0) or 0)
            except (TypeError, ValueError):
                continue
            rows.append({
                "campaign_name": row.get("campaign_name", "-"),
                "action_type": item.get("action_type", "unknown"),
                "hasil": count,
                "spend": float(row.get("spend", 0) or 0),
            })
    if not rows:
        return pd.DataFrame(columns=["campaign_name", "action_type", "hasil", "spend"])
    result = pd.DataFrame(rows)
    return result.groupby(["campaign_name", "action_type"], as_index=False).agg(
        hasil=("hasil", "sum"), spend=("spend", "max")
    )


def primary_action_per_campaign(df, action_type):
    if df is None or df.empty or not action_type:
        return pd.Series([0.0] * (0 if df is None else len(df)), index=None if df is None else df.index)
    maps = df["actions"].map(action_map) if "actions" in df else pd.Series([{}] * len(df), index=df.index)
    return maps.map(lambda m: float(m.get(action_type, 0) or 0))


def headline_cards(items):
    """Kartu headline dengan gaya visual yang sama seperti halaman Ringkasan SPN."""
    accents = ["#D91A7D", "#4A7CD9", "#2D9C9A", "#E2A13A", "#5F64D8", "#4797D3", "#47A36B", "#D97757"]
    cells = []
    for i, item in enumerate(items):
        label, value = item[0], item[1]
        note = item[2] if len(item) > 2 else ""
        accent = item[3] if len(item) > 3 else accents[i % len(accents)]
        cells.append(
            f'<div class="mx-kpi" style="border-top-color:{html_lib.escape(str(accent), quote=True)}">'
            f'<div class="mx-kpi-label">{html_lib.escape(str(label))}</div>'
            f'<div class="mx-kpi-value">{html_lib.escape(str(value))}</div>'
            f'<div class="mx-kpi-note">{html_lib.escape(str(note))}</div></div>'
        )
    st.markdown('<div class="mx-kpis">' + ''.join(cells) + '</div>', unsafe_allow_html=True)


def category_frame(df, category):
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.copy()
    return df[df["dashboard_category"] == category].copy()


def _series_sum(df, column):
    if df is None or df.empty or column not in df:
        return 0.0
    return float(pd.to_numeric(df[column], errors="coerce").fillna(0).sum())


def _campaign_count(df):
    if df is None or df.empty or "campaign_name" not in df:
        return 0
    return int(df["campaign_name"].nunique())


def _daily_category_trend(ctx: Ctx, category=None):
    daily = safe_run(load_insights, ctx.client, ctx.account, ctx.version, "campaign",
                     str(ctx.start), str(ctx.end), daily=True)
    if daily is None or daily.empty or "date_start" not in daily.columns:
        return pd.DataFrame()
    daily = tag_campaigns(daily)
    if category:
        daily = category_frame(daily, category)
    else:
        daily = daily[daily["dashboard_category"].isin(["Traffic", "Leads"])].copy()
    if daily.empty:
        return pd.DataFrame()
    daily["date_start"] = pd.to_datetime(daily["date_start"], errors="coerce").dt.normalize()
    daily = daily.dropna(subset=["date_start"])
    if daily.empty:
        return pd.DataFrame()
    return daily.groupby("date_start", as_index=True).agg(
        spend=("spend", "sum"), impressions=("impressions", "sum"), clicks=("clicks", "sum")
    ).sort_index()


def render_overview(ctx: Ctx, raw_campaigns):
    """Ringkasan utama: headline, tren belanja, dan perbandingan tujuan iklan."""
    if raw_campaigns is None:
        st.info("Data iklan belum bisa dimuat. Periksa koneksi pada bagian Koneksi.")
        return
    if raw_campaigns.empty:
        st.info("Belum ada data iklan pada periode ini. Coba pilih periode yang lebih panjang.")
        return

    traffic = category_frame(raw_campaigns, "Traffic")
    leads = category_frame(raw_campaigns, "Leads")
    traffic = prepare_results(traffic) if not traffic.empty else traffic
    leads = prepare_results(leads) if not leads.empty else leads

    traffic_spend, leads_spend = _series_sum(traffic, "spend"), _series_sum(leads, "spend")
    traffic_visits = _series_sum(traffic, "profile_visits")
    result_action = lead_action_choice(leads) if not leads.empty else None
    lead_results = float(primary_action_per_campaign(leads, result_action).sum()) if result_action else None
    impressions = _series_sum(traffic, "impressions") + _series_sum(leads, "impressions")
    clicks = _series_sum(traffic, "clicks") + _series_sum(leads, "clicks")
    ctr = clicks / impressions * 100 if impressions else 0.0

    headline_cards([
        ("Total belanja iklan", rp(traffic_spend + leads_spend), f"Traffic {rp(traffic_spend)} · Leads {rp(leads_spend)}"),
        ("Kunjungan profil IG", idn(traffic_visits), "Hasil campaign Traffic"),
        ("Hasil Leads", idn(lead_results) if lead_results is not None else "-", "Hasil utama dari Meta"),
        ("Impresi", idn(impressions), "Traffic + Leads"),
        ("Klik", idn(clicks), "Traffic + Leads"),
        ("CTR gabungan", f"{idn(ctr, 2)}%", "Klik dibagi impresi"),
        ("Campaign Traffic", idn(_campaign_count(traffic)), "Campaign terklasifikasi"),
        ("Campaign Leads", idn(_campaign_count(leads)), "Campaign terklasifikasi"),
    ])

    left, right = st.columns([1.7, 1], gap="large")
    with left:
        with st.container(border=True):
            st.markdown('<div class="section-title">Tren belanja harian</div>', unsafe_allow_html=True)
            daily_source = safe_run(load_insights, ctx.client, ctx.account, ctx.version, "campaign",
                                    str(ctx.start), str(ctx.end), daily=True)
            if daily_source is not None and not daily_source.empty and "date_start" in daily_source:
                daily_source = tag_campaigns(daily_source)
                daily_source = daily_source[daily_source["dashboard_category"].isin(["Traffic", "Leads"])].copy()
                daily_source["date_start"] = pd.to_datetime(daily_source["date_start"], errors="coerce").dt.normalize()
                daily_source = daily_source.dropna(subset=["date_start"])
                if not daily_source.empty:
                    full_index = pd.date_range(pd.Timestamp(ctx.start), pd.Timestamp(ctx.end), freq="D")
                    plot = daily_source.pivot_table(index="date_start", columns="dashboard_category", values="spend", aggfunc="sum")
                    plot = plot.reindex(full_index, fill_value=0)
                    for cat in ("Traffic", "Leads"):
                        if cat not in plot.columns:
                            plot[cat] = 0.0
                    plot = plot[["Traffic", "Leads"]]
                    plot.index.name = "Tanggal"
                    st.line_chart(plot, height=300)
                else:
                    st.info("Belum ada tren harian untuk periode ini.")
            else:
                st.info("Belum ada tren harian untuk periode ini.")
    with right:
        with st.container(border=True):
            st.markdown('<div class="section-title">Belanja per tujuan</div>', unsafe_allow_html=True)
            spend_by_goal = pd.Series({"Traffic": traffic_spend, "Leads": leads_spend}, name="Belanja")
            st.bar_chart(spend_by_goal, horizontal=True, height=300)
            st.caption("Perbandingan belanja antara campaign Traffic dan Leads.")

    with st.container(border=True):
        st.markdown('<div class="section-title">Performa per tujuan</div>', unsafe_allow_html=True)
        lead_clicks = _series_sum(leads, "clicks")
        traffic_clicks = _series_sum(traffic, "clicks")
        rows = [
            {"Tujuan": "Traffic", "Campaign": _campaign_count(traffic), "Belanja": rp(traffic_spend),
             "Impresi": idn(_series_sum(traffic, "impressions")), "Klik": idn(traffic_clicks),
             "Hasil utama": idn(traffic_visits) + " kunjungan profil"},
            {"Tujuan": "Leads", "Campaign": _campaign_count(leads), "Belanja": rp(leads_spend),
             "Impresi": idn(_series_sum(leads, "impressions")), "Klik": idn(lead_clicks),
             "Hasil utama": idn(lead_results) + " leads" if lead_results is not None else "Belum tersedia"},
        ]
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")

    other = raw_campaigns[raw_campaigns["dashboard_category"] == "Lainnya"]
    if not other.empty:
        with st.expander(f"Campaign yang belum terkategori ({len(other)})", expanded=False):
            st.caption("Campaign ini tidak dimasukkan ke total Traffic atau Leads agar perbandingan tetap jelas.")
            cols = [c for c in ("campaign_name", "spend", "impressions", "clicks") if c in other.columns]
            st.dataframe(other[cols].rename(columns={"campaign_name": "Campaign", "spend": "Belanja",
                                                     "impressions": "Impresi", "clicks": "Klik"}),
                         hide_index=True, width="stretch")



@st.cache_data(ttl=600, show_spinner=False)
def _spn_registrations(start: str, end: str, csv_path: str) -> int | None:
    """Hitung pendaftar unik SPN dalam rentang tanggal yang sama dengan periode iklan."""
    path = Path(csv_path)
    if not path.is_absolute():
        candidates = [Path.cwd() / path, Path(__file__).resolve().parent.parent / path]
        path = next((candidate for candidate in candidates if candidate.is_file()), candidates[0])
    if not path.is_file():
        return None
    try:
        raw = pd.read_csv(path, encoding="utf-8-sig", usecols=["timestamp", "email"])
    except (ValueError, OSError, UnicodeError):
        return None
    dates = pd.to_datetime(raw["timestamp"], errors="coerce")
    emails = raw["email"].fillna("").astype(str).str.strip().str.lower()
    valid = (dates >= pd.Timestamp(start)) & (dates < pd.Timestamp(end) + pd.Timedelta(days=1))
    valid &= emails.ne("")
    return int(emails[valid].nunique())


def _ad_follow_series(frame: pd.DataFrame) -> pd.Series:
    """Jumlah follow Meta per campaign, memakai field native lalu fallback ke actions."""
    if frame is None:
        return pd.Series(dtype=float)
    if frame.empty:
        return pd.Series(0.0, index=frame.index)
    if FOLLOW_FIELD in frame:
        native = pd.to_numeric(frame[FOLLOW_FIELD], errors="coerce").fillna(0.0)
        if float(native.sum()) > 0:
            return native.astype(float)
    totals = action_totals(frame)
    follow_types = detect(totals, "follow")
    if not follow_types or "actions" not in frame:
        return pd.Series(0.0, index=frame.index)
    return frame["actions"].map(
        lambda acts: sum(action_map(acts).get(action, 0.0) for action in follow_types)
    ).astype(float)


def _ad_follow_count(frame: pd.DataFrame) -> float:
    """Jumlah follow yang dilaporkan Meta untuk campaign dalam kategori aktif."""
    return float(_ad_follow_series(frame).sum())


def _instagram_profile_and_trend(ctx: Ctx):
    """Ambil follower aktual dan jumlah follower baru harian dari Instagram Insights.

    Meta mengembalikan follows_and_unfollows sebagai total_value, bukan time series
    harian. Untuk grafik harian, gunakan metric follower_count yang sudah teruji
    mengembalikan values per hari di Graph API Explorer.
    """
    ig_user_id = cfg("IG_USER_ID", cfg("INSTAGRAM_USER_ID", ""))
    ig_token = cfg("IG_ACCESS_TOKEN", "")
    if not ig_user_id:
        return None, pd.DataFrame(), "Isi IG_USER_ID pada Streamlit Secrets untuk menghubungkan akun Instagram."
    if not ig_token:
        return None, pd.DataFrame(), "Isi IG_ACCESS_TOKEN pada Streamlit Secrets. Token Instagram dipisahkan dari META_TOKEN."

    # Pisahkan versi Instagram Graph API dari Marketing API agar perubahan versi
    # Insights tidak memengaruhi request Ads Insights yang memakai META_TOKEN.
    ig_version = str(cfg("IG_API_VERSION", "v26.0"))
    ig_client = MetaClient(str(ig_token), ig_version)

    profile = safe_run(
        graph, ig_client, ig_version, str(ig_user_id),
        (("fields", "username,followers_count"),), False, token=str(ig_token)
    )
    follower_total = None if not profile else profile.get("followers_count")

    # Instagram Insights only supports follower_count in a rolling recent window,
    # excluding the current day. Calculate the supported overlap on every run so
    # the limit moves forward automatically; use Jakarta's calendar date.
    today_jakarta = pd.Timestamp.now(tz="Asia/Jakarta").date()
    first_supported_day = today_jakarta - timedelta(days=30)
    last_supported_day = today_jakarta - timedelta(days=1)
    query_start = max(ctx.start, first_supported_day)
    query_end = min(ctx.end, last_supported_day)

    if query_start > query_end:
        return (
            follower_total,
            pd.DataFrame(),
            "Data follower tidak tersedia untuk periode yang dipilih. "
            "Instagram Insights hanya menyediakan data historis terbaru. "
            "Data Ads dan KPI lainnya tetap ditampilkan.",
        )

    partial_window = query_start != ctx.start or query_end != ctx.end
    raw = safe_run(
        graph, ig_client, ig_version, f"{ig_user_id}/insights",
        (("metric", "follower_count"),
         ("metric_type", "time_series"),
         ("period", "day"),
         ("since", str(query_start)),
         ("until", str(query_end))), False, token=str(ig_token)
    )

    rows = []
    if isinstance(raw, dict):
        for metric in raw.get("data", []):
            if metric.get("name") != "follower_count":
                continue
            for point in metric.get("values", []):
                stamp = pd.to_datetime(point.get("end_time"), errors="coerce", utc=True)
                if pd.isna(stamp):
                    continue
                # API mengembalikan end_time ber-offset UTC. Samakan ke tanggal WIB
                # sebelum membuat index agar tidak bentrok dengan date_range yang naive.
                local_day = stamp.tz_convert("Asia/Jakarta").tz_localize(None).normalize()
                try:
                    followers_added = float(point.get("value", 0) or 0)
                except (TypeError, ValueError):
                    continue
                rows.append({"Tanggal": local_day, "Follower baru": followers_added})

    trend = pd.DataFrame(rows)
    if not trend.empty:
        trend = trend.dropna(subset=["Tanggal"]).set_index("Tanggal").sort_index()
        trend = trend[~trend.index.duplicated(keep="last")]
    note = (
        "Grafik follower hanya menampilkan tanggal yang masih tersedia dari Instagram Insights."
        if partial_window and not trend.empty else None
    )
    return follower_total, trend, note


def render_stats(ctx: Ctx, category, raw_campaigns):
    """Traffic dan Leads memakai KPI jelas, funnel SPN, dan tren follower baru harian."""
    campaigns = category_frame(raw_campaigns, category) if raw_campaigns is not None else pd.DataFrame()
    if campaigns.empty:
        st.info(f"Belum ada campaign {category} pada periode ini.")

    campaigns = prepare_results(campaigns)
    spend = _series_sum(campaigns, "spend")
    impressions = _series_sum(campaigns, "impressions")
    reach = _series_sum(campaigns, "reach")
    ad_follows = _ad_follow_count(campaigns)
    cost_per_follow = div(spend, ad_follows)

    # Angka funnel dihitung lintas campaign Traffic + Leads untuk periode yang sama.
    traffic_campaigns = category_frame(raw_campaigns, "Traffic") if raw_campaigns is not None else pd.DataFrame()
    leads_campaigns = category_frame(raw_campaigns, "Leads") if raw_campaigns is not None else pd.DataFrame()
    traffic_campaigns = prepare_results(traffic_campaigns) if not traffic_campaigns.empty else traffic_campaigns
    leads_campaigns = prepare_results(leads_campaigns) if not leads_campaigns.empty else leads_campaigns
    traffic_visits = _series_sum(traffic_campaigns, "profile_visits")
    total_acquisition_spend = _series_sum(traffic_campaigns, "spend") + _series_sum(leads_campaigns, "spend")
    csv_path = str(cfg("SPN_CSV_PATH", "master_penting_3.csv"))
    registrations = _spn_registrations(str(ctx.start), str(ctx.end), csv_path)
    conversion_rate = (
        registrations / traffic_visits * 100
        if registrations is not None and traffic_visits > 0 else None
    )
    cac = (
        total_acquisition_spend / registrations
        if registrations is not None and registrations > 0 else None
    )

    if category == "Traffic":
        result = _series_sum(campaigns, "profile_visits")
        result_label = "Kunjungan profil"
        cost_per_result = div(spend, result)
        first_kpi_label = "Profile Visits"
        cost_kpi_label = "Cost per Profile Visit (CPV)"
        cost_kpi_note = "Spend Traffic dibagi kunjungan profil"
    else:
        result_action = lead_action_choice(campaigns)
        meta_lead_results = (
            float(primary_action_per_campaign(campaigns, result_action).sum())
            if result_action else 0.0
        )
        result = registrations
        result_label = "Email unik yang bergabung ke SPN"
        cost_per_result = cac
        first_kpi_label = "Total Join Community"
        cost_kpi_label = "CAC"
        cost_kpi_note = "Total spend Traffic + Leads dibagi member baru"

    follower_total, follower_trend, follower_error = _instagram_profile_and_trend(ctx)

    headline_cards([
        (first_kpi_label, idn(result) if result is not None else "-", result_label),
        (cost_kpi_label, rp(cost_per_result), cost_kpi_note),
        ("Follower By Ads", idn(ad_follows), "Follow yang dilaporkan Meta dari campaign ini"),
        ("Follower Asli IG", idn(follower_total) if follower_total is not None else "-",
         "Jumlah follower aktual akun Instagram"),
        ("CPF (Cost per Follow)", rp(cost_per_follow), "Spend kategori ini dibagi Follower By Ads"),
        ("Total Spend", rp(spend), f"{category} · {ctx.start:%d %b %Y} – {ctx.end:%d %b %Y}"),
        ("Impressions", idn(impressions), "Total tayangan dari campaign kategori ini"),
        ("Reach", idn(reach), "Akumulasi reach per campaign; bisa ada overlap audiens"),
    ])

    with st.container(border=True):
        st.markdown('<div class="section-title">Konversi ke Community</div>', unsafe_allow_html=True)
        funnel_cols = st.columns(4)
        funnel_cols[0].metric("Profile Visits (Traffic)", idn(traffic_visits))
        funnel_cols[1].metric("Total Join Community", idn(registrations) if registrations is not None else "-")
        funnel_cols[2].metric("CR", f"{idn(conversion_rate, 2)}%" if conversion_rate is not None else "-")
        funnel_cols[3].metric("CAC", rp(cac))
        if registrations is None:
            st.warning("Data pendaftar belum terbaca. Pastikan file CSV tersedia dan memiliki kolom timestamp serta email.")

    if category == "Leads" and registrations is not None and meta_lead_results <= 0:
        st.caption("Meta belum mengembalikan action Leads untuk campaign pada periode ini; Total Join Community tetap dihitung dari email unik SPN.")

    with st.container(border=True):
        st.markdown('<div class="section-title">Follower baru Instagram per hari</div>', unsafe_allow_html=True)
        if follower_trend.empty:
            st.info(follower_error or "Data follower tidak tersedia untuk periode ini.")
        else:
            if follower_error:
                st.caption(follower_error)
            full_index = pd.date_range(pd.Timestamp(ctx.start), pd.Timestamp(ctx.end), freq="D")
            # Biarkan tanggal tanpa respons sebagai kosong, bukan angka 0 palsu.
            plot = follower_trend.reindex(full_index)
            plot.index.name = "Tanggal"
            st.caption(
                "Grafik menunjukkan jumlah follower baru per hari. Data yang tidak tersedia tidak ditampilkan."
            )
            st.line_chart(plot[["Follower baru"]], height=280)
            with st.expander("Rincian follower baru harian", expanded=False):
                detail = plot.rename_axis("Tanggal").reset_index()
                detail["Tanggal"] = detail["Tanggal"].dt.strftime("%d %b %Y")
                detail["Follower baru"] = detail["Follower baru"].map(
                    lambda value: "-" if pd.isna(value) else idn(value)
                )
                st.dataframe(detail, hide_index=True, width="stretch")

    if campaigns.empty:
        return

    # Tabel dibuat konsisten di Traffic dan Leads. Result per campaign tetap memakai action Meta;
    # pendaftar SPN hanya dihitung di KPI agregat karena tidak tersedia atribusi per campaign.
    display_raw = campaigns.sort_values("spend", ascending=False).copy()
    top = display_raw.head(10).copy()
    result_by_campaign = (
        top["profile_visits"].fillna(0)
        if category == "Traffic"
        else primary_action_per_campaign(top, lead_action_choice(campaigns))
    )
    table = pd.DataFrame({
        "Campaign": top.get("campaign_name", pd.Series("-", index=top.index)).fillna("-"),
        "Spend": top.get("spend", pd.Series(0.0, index=top.index)).map(rp),
        "Result dari Meta": result_by_campaign.map(idn),
        "Follower By Ads": _ad_follow_series(top).map(idn),
    })
    with st.container(border=True):
        st.markdown('<div class="section-title">Performa campaign</div>', unsafe_allow_html=True)
        st.caption(
            "Result pada tabel adalah hasil yang dilaporkan Meta per campaign. "
            "Jumlah pendaftar SPN digunakan pada KPI Cost per Result di atas."
            if category == "Leads" else
            "Result pada tabel adalah kunjungan profil Instagram yang dilaporkan Meta."
        )
        st.dataframe(table, hide_index=True, width="stretch")

    with st.expander("Analisis lanjutan", expanded=False):
        analysis = st.radio(
            "Apa yang ingin dilihat?", ["Rincian iklan", "Breakdown audiens", "Hasil dari Meta"],
            horizontal=True, key=f"mx_advanced_kind_v4_{category}",
        )
        if analysis == "Rincian iklan":
            level_choice = st.selectbox(
                "Tingkat rincian", ["Campaign", "Ad set", "Iklan"],
                key=f"mx_level_v4_{category}"
            )
            level = {"Campaign": "campaign", "Ad set": "adset", "Iklan": "ad"}[level_choice]
            if level == "campaign":
                detail = campaigns.copy()
            else:
                detail_raw = safe_run(
                    load_insights, ctx.client, ctx.account, ctx.version,
                    level, str(ctx.start), str(ctx.end)
                )
                detail = tag_campaigns(detail_raw) if detail_raw is not None and not detail_raw.empty else pd.DataFrame()
                detail = category_frame(detail, category) if not detail.empty else detail
            if detail is not None and not detail.empty:
                names = [c for c in ("campaign_name", "adset_name", "ad_name") if c in detail.columns]
                detail_cols = names + [c for c in ("spend", FOLLOW_FIELD) if c in detail.columns]
                shown = detail.sort_values("spend", ascending=False)[detail_cols].copy()
                shown = shown.rename(columns={
                    "campaign_name": "Campaign", "adset_name": "Ad set",
                    "ad_name": "Iklan", "spend": "Spend", FOLLOW_FIELD: "Follower By Ads",
                })
                if "Spend" in shown:
                    shown["Spend"] = shown["Spend"].map(rp)
                if "Follower By Ads" in shown:
                    shown["Follower By Ads"] = shown["Follower By Ads"].map(idn)
                st.dataframe(shown, hide_index=True, width="stretch")
        elif analysis == "Breakdown audiens":
            breakdown = st.selectbox(
                "Breakdown", ["age", "gender", "country", "region", "publisher_platform"],
                key=f"mx_breakdown_v4_{category}"
            )
            detail = safe_run(
                load_insights, ctx.client, ctx.account, ctx.version, "campaign",
                str(ctx.start), str(ctx.end), breakdowns=(breakdown,)
            )
            if detail is not None and not detail.empty:
                detail = tag_campaigns(detail)
                detail = category_frame(detail, category)
                keep = [c for c in (breakdown, "spend", FOLLOW_FIELD) if c in detail.columns]
                st.dataframe(detail[keep].sort_values("spend", ascending=False),
                             hide_index=True, width="stretch")
        else:
            totals = action_totals(campaigns)
            if totals.empty:
                st.info("Meta tidak mengembalikan data actions pada periode ini.")
            else:
                st.dataframe(
                    totals.rename_axis("Action").reset_index(name="Jumlah"),
                    hide_index=True, width="stretch"
                )


def render_admin(ctx: Ctx):
    """Info akun iklan dan log API; tidak meminta atau menampilkan data Facebook Pages."""
    with st.container(border=True):
        st.markdown('<div class="section-title">Akun iklan Instagram</div>', unsafe_allow_html=True)
        acc = safe_run(
            graph, ctx.client, ctx.version, ctx.account,
            (("fields", "name,account_status,currency,timezone_name"),), False
        )
        if acc:
            cols = st.columns(4)
            cols[0].metric("Nama akun", acc.get("name", "-"))
            cols[1].metric("Status", ACCOUNT_STATUS.get(acc.get("account_status"), "-"))
            cols[2].metric("Mata uang", acc.get("currency", "-"))
            cols[3].metric("Zona waktu", acc.get("timezone_name", "-"))
        else:
            st.caption("Informasi akun belum tersedia. Jalankan pemeriksaan koneksi.")

    with st.expander("Log teknis API (opsional)", expanded=False):
        logs = pd.DataFrame(ctx.client.log[::-1])
        if logs.empty:
            st.info("Belum ada request pada sesi ini.")
        else:
            cols = st.columns(3)
            cols[0].metric("Request", len(logs))
            cols[1].metric("Error", int((logs["status"] != 200).sum()))
            cols[2].metric("Latensi rata-rata", f"{logs['ms'].mean():.0f} ms")
            visible_cols = [c for c in ["time", "path", "status", "rows", "ms", "error"] if c in logs.columns]
            st.dataframe(logs[visible_cols], hide_index=True, width="stretch")


# ---------------- filter periode (di badan halaman) ----------------
def period_filter(account: str, version: str):
    """Pilihan periode praktis dengan rentang tanggal kustom bila diperlukan."""
    today = date.today()
    presets = ["Bulan ini", "7 hari terakhir", "30 hari terakhir", "Bulan lalu", "Pilih tanggal"]
    with st.container(border=True):
        c_period, c_dates, c_refresh = st.columns([1.2, 2.6, 0.9], vertical_alignment="bottom")
        preset = c_period.selectbox("Periode", presets, index=0, key="mx_period_preset_v3")
        if preset == "7 hari terakhir":
            start, end = today - timedelta(days=6), today
        elif preset == "30 hari terakhir":
            start, end = today - timedelta(days=29), today
        elif preset == "Bulan lalu":
            last_day = today.replace(day=1) - timedelta(days=1)
            start, end = last_day.replace(day=1), last_day
        elif preset == "Pilih tanggal":
            default_range = (today - timedelta(days=29), today)
            values = c_dates.date_input("Rentang tanggal", value=default_range, key="mx_custom_date_range_v3")
            if not isinstance(values, (tuple, list)) or len(values) != 2:
                st.caption("Pilih tanggal mulai dan selesai.")
                return None
            start, end = values
            if start > end:
                st.error("Tanggal mulai harus sebelum tanggal selesai.")
                return None
        else:
            start, end = today.replace(day=1), today

        c_dates.caption(f"{start:%d %b %Y} – {end:%d %b %Y}")
        if c_refresh.button("Refresh", icon=":material/refresh:", width="stretch", key="mx_refresh_data_v3"):
            clear_meta_cache()
            st.rerun()
    return start, end


def _password_ok() -> bool:
    """Kunci opsional: aktif hanya kalau META_ADS_PASSWORD diisi di secrets."""
    pw = cfg("META_ADS_PASSWORD")
    if not pw or st.session_state.get("mx_auth"):
        return True
    st.info("Halaman Meta Ads dikunci. Masukkan password untuk membuka.")
    entered = st.text_input("Password", type="password", key="mx_pw_input")
    if entered:
        if hmac.compare_digest(str(entered), str(pw)):
            st.session_state.mx_auth = True
            st.rerun()
        st.error("Password salah.")
    return False


# ---------------- laporan lengkap: Excel dan PDF ----------------
def _report_number(value):
    try:
        return float(value) if value is not None and pd.notna(value) else None
    except (TypeError, ValueError):
        return None


def _report_payload(ctx: Ctx, raw_campaigns):
    """Siapkan satu paket data lengkap agar export konsisten dengan filter aktif."""
    raw = tag_campaigns(raw_campaigns) if raw_campaigns is not None and not raw_campaigns.empty else pd.DataFrame()
    traffic = category_frame(raw, "Traffic") if not raw.empty else pd.DataFrame()
    leads = category_frame(raw, "Leads") if not raw.empty else pd.DataFrame()
    traffic = prepare_results(traffic) if not traffic.empty else traffic
    leads = prepare_results(leads) if not leads.empty else leads

    traffic_spend = _series_sum(traffic, "spend")
    leads_spend = _series_sum(leads, "spend")
    traffic_visits = _series_sum(traffic, "profile_visits")
    registrations = _spn_registrations(
        str(ctx.start), str(ctx.end), str(cfg("SPN_CSV_PATH", "master_penting_3.csv"))
    )
    total_spend = traffic_spend + leads_spend
    cr = (registrations / traffic_visits * 100) if registrations is not None and traffic_visits > 0 else None
    cac = (total_spend / registrations) if registrations is not None and registrations > 0 else None
    traffic_follows = _ad_follow_count(traffic)
    leads_follows = _ad_follow_count(leads)
    traffic_cpf = div(traffic_spend, traffic_follows)
    leads_cpf = div(leads_spend, leads_follows)
    follower_total, follower_trend, follower_error = _instagram_profile_and_trend(ctx)

    # Lengkapi tabel campaign: seluruh campaign yang berhasil diambil, bukan hanya top 10.
    prepared = pd.concat([traffic, leads], ignore_index=True) if (not traffic.empty or not leads.empty) else pd.DataFrame()
    if not prepared.empty:
        campaign_columns = [
            ("dashboard_category", "Kategori"), ("campaign_name", "Campaign"),
            ("spend", "Spend (Rp)"), ("impressions", "Impressions"),
            ("reach", "Reach"), ("clicks", "Clicks"), ("ctr", "CTR (%)"),
            ("profile_visits", "Profile Visits"), (FOLLOW_FIELD, "Follower By Ads"),
            ("frequency", "Frequency"),
        ]
        available = [(src, dst) for src, dst in campaign_columns if src in prepared.columns]
        campaigns_export = prepared[[src for src, _ in available]].rename(columns=dict(available))
        if "Campaign" in campaigns_export:
            campaigns_export = campaigns_export.sort_values("Spend (Rp)" if "Spend (Rp)" in campaigns_export else "Campaign", ascending=False)
    else:
        campaigns_export = pd.DataFrame(columns=[
            "Kategori", "Campaign", "Spend (Rp)", "Impressions", "Reach", "Clicks",
            "CTR (%)", "Profile Visits", "Follower By Ads", "Frequency"
        ])

    # Daily spend by objective; the query is already cached by load_insights.
    daily_source = safe_run(
        load_insights, ctx.client, ctx.account, ctx.version, "campaign",
        str(ctx.start), str(ctx.end), daily=True
    )
    if daily_source is not None and not daily_source.empty and "date_start" in daily_source.columns:
        daily_source = tag_campaigns(daily_source)
        daily_source = daily_source[daily_source["dashboard_category"].isin(["Traffic", "Leads"])].copy()
        daily_source["date_start"] = pd.to_datetime(daily_source["date_start"], errors="coerce").dt.normalize()
        daily_source = daily_source.dropna(subset=["date_start"])
        daily_spend = daily_source.pivot_table(
            index="date_start", columns="dashboard_category", values="spend", aggfunc="sum"
        )
    else:
        daily_spend = pd.DataFrame()
    full_index = pd.date_range(pd.Timestamp(ctx.start), pd.Timestamp(ctx.end), freq="D")
    daily_spend = daily_spend.reindex(full_index, fill_value=0).fillna(0)
    for col in ("Traffic", "Leads"):
        if col not in daily_spend.columns:
            daily_spend[col] = 0.0
    daily_spend = daily_spend[["Traffic", "Leads"]]
    daily_spend.index.name = "Tanggal"

    daily_report = daily_spend.reset_index()
    if not follower_trend.empty:
        daily_report = daily_report.merge(
            follower_trend.reset_index()[["Tanggal", "Follower baru"]], on="Tanggal", how="left"
        )
    else:
        daily_report["Follower baru"] = pd.NA

    metrics = pd.DataFrame([
        {"Metrik": "Total Spend (Rp)", "Traffic": traffic_spend, "Leads": leads_spend, "Total / Global": total_spend},
        {"Metrik": "Result", "Traffic": traffic_visits, "Leads": registrations, "Total / Global": None},
        {"Metrik": "Cost per Result (Rp)", "Traffic": div(traffic_spend, traffic_visits), "Leads": cac, "Total / Global": None},
        {"Metrik": "Follower By Ads", "Traffic": traffic_follows, "Leads": leads_follows, "Total / Global": traffic_follows + leads_follows},
        {"Metrik": "CPF (Rp)", "Traffic": traffic_cpf, "Leads": leads_cpf, "Total / Global": None},
        {"Metrik": "Follower Asli IG", "Traffic": None, "Leads": None, "Total / Global": follower_total},
        {"Metrik": "Impressions", "Traffic": _series_sum(traffic, "impressions"), "Leads": _series_sum(leads, "impressions"), "Total / Global": _series_sum(traffic, "impressions") + _series_sum(leads, "impressions")},
        {"Metrik": "Reach (akumulasi campaign)", "Traffic": _series_sum(traffic, "reach"), "Leads": _series_sum(leads, "reach"), "Total / Global": _series_sum(traffic, "reach") + _series_sum(leads, "reach")},
    ])
    funnel = pd.DataFrame([
        {"Metrik": "Profile Visits (Traffic)", "Nilai": traffic_visits},
        {"Metrik": "Total Join Community (email unik)", "Nilai": registrations},
        {"Metrik": "CR (%)", "Nilai": cr},
        {"Metrik": "CAC (Rp / member baru)", "Nilai": cac},
    ])
    objective = pd.DataFrame([
        {"Objective": "Traffic", "Campaign": _campaign_count(traffic), "Spend (Rp)": traffic_spend,
         "Result": traffic_visits, "Cost per Result (Rp)": div(traffic_spend, traffic_visits),
         "Follower By Ads": traffic_follows, "CPF (Rp)": traffic_cpf,
         "Impressions": _series_sum(traffic, "impressions"), "Reach": _series_sum(traffic, "reach")},
        {"Objective": "Leads", "Campaign": _campaign_count(leads), "Spend (Rp)": leads_spend,
         "Result": registrations, "Cost per Result (Rp)": cac,
         "Follower By Ads": leads_follows, "CPF (Rp)": leads_cpf,
         "Impressions": _series_sum(leads, "impressions"), "Reach": _series_sum(leads, "reach")},
    ])
    return {
        "start": ctx.start, "end": ctx.end, "metrics": metrics, "funnel": funnel,
        "objective": objective, "campaigns": campaigns_export,
        "daily": daily_report, "follower_error": follower_error,
        "follower_total": follower_total,
    }


def _excel_report(payload) -> bytes:
    """Workbook terformat: executive summary, funnel, campaign detail, dan tren harian."""
    import xlsxwriter  # noqa: F401 - memastikan engine tersedia di deployment

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd mmm yyyy") as writer:
        payload["metrics"].to_excel(writer, sheet_name="KPI Summary", index=False, startrow=4)
        payload["objective"].to_excel(writer, sheet_name="KPI Summary", index=False, startrow=17)
        payload["funnel"].to_excel(writer, sheet_name="Community Funnel", index=False, startrow=4)
        payload["campaigns"].to_excel(writer, sheet_name="Campaign Detail", index=False, startrow=4)
        payload["daily"].to_excel(writer, sheet_name="Daily Trend", index=False, startrow=4)

        workbook = writer.book
        navy = "1A1464"
        blue = "497CD6"
        pale = "EAF1FF"
        muted = "65718F"
        title_fmt = workbook.add_format({"bold": True, "font_size": 18, "font_color": navy})
        subtitle_fmt = workbook.add_format({"font_size": 10, "font_color": muted})
        section_fmt = workbook.add_format({"bold": True, "font_size": 11, "font_color": navy, "bottom": 2, "bottom_color": blue})
        header_fmt = workbook.add_format({"bold": True, "font_color": "FFFFFF", "bg_color": navy, "text_wrap": True, "valign": "vcenter", "border": 0})
        money_fmt = workbook.add_format({"num_format": '"Rp" #,##0;[Red]("Rp" #,##0);-'})
        number_fmt = workbook.add_format({"num_format": "#,##0;[Red](#,##0);-"})
        percent_fmt = workbook.add_format({"num_format": '0.00"%"'})
        date_fmt = workbook.add_format({"num_format": "dd mmm yyyy"})
        text_fmt = workbook.add_format({"font_color": "24304F"})
        note_fmt = workbook.add_format({"font_color": muted, "italic": True, "text_wrap": True, "valign": "top"})
        sheet_widths = {
            "KPI Summary": 9,
            "Community Funnel": max(2, len(payload["funnel"].columns)),
            "Campaign Detail": max(1, len(payload["campaigns"].columns)),
            "Daily Trend": max(1, len(payload["daily"].columns)),
        }
        for name in writer.sheets:
            ws = writer.sheets[name]
            last_col = sheet_widths.get(name, 4) - 1
            ws.hide_gridlines(2)
            ws.set_default_row(20)
            ws.set_tab_color(blue)
            ws.set_landscape()
            ws.fit_to_pages(1, 0)
            ws.set_margins(0.25, 0.25, 0.5, 0.5)
            ws.freeze_panes(5 if name in ("Campaign Detail", "Daily Trend") else 4, 0)
            ws.merge_range(0, 0, 0, last_col, "SPN | INSTAGRAM ADS PERFORMANCE REPORT", title_fmt)
            ws.merge_range(1, 0, 1, min(last_col, 3), f"Periode: {payload['start']:%d %b %Y} – {payload['end']:%d %b %Y}", subtitle_fmt)
            ws.set_row(0, 30)
            ws.set_row(2, 22)
        # KPI Summary: table utama dan per objective.
        ws = writer.sheets["KPI Summary"]
        ws.merge_range("A3:D3", "RINGKASAN KPI", section_fmt)
        for col_idx, col_name in enumerate(payload["metrics"].columns):
            ws.write(4, col_idx, col_name, header_fmt)
        ws.set_row(4, 32)
        ws.set_row(15, 22)
        ws.set_row(16, 8)
        ws.set_column("A:A", 32, text_fmt)
        ws.set_column("B:D", 20, number_fmt)
        ws.set_column("E:I", 20, number_fmt)
        for row in (4,):
            for col in (1, 2, 3):
                ws.set_column(col, col, 20, number_fmt)
        ws.merge_range("A16:I16", "PERFORMA PER OBJECTIVE", section_fmt)
        for col_idx, col_name in enumerate(payload["objective"].columns):
            ws.write(17, col_idx, col_name, header_fmt)
        ws.set_row(17, 32)
        # Format sel tabel objective per baris agar format kolom Leads di KPI utama
        # tidak ikut berubah menjadi format mata uang.
        objective_money_fields = {"Spend (Rp)", "Cost per Result (Rp)", "CPF (Rp)"}
        for row_idx, row_data in payload["objective"].iterrows():
            for col_idx, col_name in enumerate(payload["objective"].columns):
                value = row_data[col_name]
                if pd.notna(value) and isinstance(value, (int, float)):
                    cell_fmt = money_fmt if col_name in objective_money_fields else number_fmt
                    ws.write_number(18 + row_idx, col_idx, float(value), cell_fmt)
        # Format kolom uang khusus pada KPI utama.
        for row in [5, 7, 9]:
            for col in [1, 2, 3]:
                metric_idx = row - 5
                if col < len(payload["metrics"].columns) and pd.notna(payload["metrics"].iloc[metric_idx, col]):
                    ws.write_number(row, col, float(payload["metrics"].iloc[metric_idx, col]), money_fmt)
        # Funnel
        ws = writer.sheets["Community Funnel"]
        ws.merge_range("A3:B3", "COMMUNITY CONVERSION", section_fmt)
        for col_idx, col_name in enumerate(payload["funnel"].columns):
            ws.write(4, col_idx, col_name, header_fmt)
        ws.set_row(4, 32)
        ws.set_column("A:A", 42, text_fmt)
        ws.set_column("B:B", 24, number_fmt)
        for idx, row in payload["funnel"].iterrows():
            value = row["Nilai"]
            if pd.notna(value):
                fmt = money_fmt if "CAC" in row["Metrik"] else (percent_fmt if "CR" in row["Metrik"] else number_fmt)
                ws.write_number(5 + idx, 1, float(value), fmt)
        # Campaign detail
        ws = writer.sheets["Campaign Detail"]
        ws.merge_range(
            3, 0, 3, max(0, len(payload["campaigns"].columns) - 1),
            "SELURUH CAMPAIGN DALAM PERIODE", section_fmt
        )
        for col_idx, col_name in enumerate(payload["campaigns"].columns):
            ws.write(4, col_idx, col_name, header_fmt)
        ws.set_row(4, 34)
        ws.freeze_panes(5, 2)
        ws.autofilter(4, 0, 4 + len(payload["campaigns"]), max(0, len(payload["campaigns"].columns) - 1))
        widths = {"Kategori": 13, "Campaign": 42, "Spend (Rp)": 16, "Impressions": 14, "Reach": 14, "Clicks": 12, "CTR (%)": 12, "Profile Visits": 16, "Follower By Ads": 16, "Frequency": 12}
        for col_idx, col_name in enumerate(payload["campaigns"].columns):
            ws.set_column(col_idx, col_idx, widths.get(col_name, 16))
            if col_name == "Spend (Rp)":
                ws.set_column(col_idx, col_idx, widths.get(col_name, 16), money_fmt)
            elif col_name == "CTR (%)":
                ws.set_column(col_idx, col_idx, widths.get(col_name, 16), percent_fmt)
            elif col_name != "Campaign" and col_name != "Kategori":
                ws.set_column(col_idx, col_idx, widths.get(col_name, 16), number_fmt)
        if len(payload["campaigns"]):
            ws.conditional_format(5, 2, 4 + len(payload["campaigns"]), 2, {"type": "data_bar", "bar_color": blue})
        # Daily trend
        ws = writer.sheets["Daily Trend"]
        ws.merge_range("A3:D3", "TREN HARIAN", section_fmt)
        for col_idx, col_name in enumerate(payload["daily"].columns):
            ws.write(4, col_idx, col_name, header_fmt)
        ws.set_row(4, 32)
        ws.set_column("A:A", 18, date_fmt)
        ws.set_column("B:C", 18, money_fmt)
        ws.set_column("D:D", 18, number_fmt)
        ws.autofilter(4, 0, 4 + len(payload["daily"]), max(0, len(payload["daily"].columns) - 1))
        for row_idx, day_value in enumerate(payload["daily"]["Tanggal"], start=5):
            if pd.notna(day_value):
                ws.write_datetime(row_idx, 0, pd.Timestamp(day_value).to_pydatetime(), date_fmt)
        if len(payload["daily"]):
            chart = workbook.add_chart({"type": "line"})
            n = len(payload["daily"])
            for col, label in [(1, "Traffic"), (2, "Leads")]:
                chart.add_series({"name": label, "categories": ["Daily Trend", 5, 0, 4 + n, 0], "values": ["Daily Trend", 5, col, 4 + n, col]})
            chart.set_title({"name": "Daily Ad Spend"})
            chart.set_legend({"position": "bottom"})
            chart.set_y_axis({"name": "Spend (Rp)", "num_format": '"Rp" #,##0'})
            chart.set_size({"width": 700, "height": 300})
            ws.insert_chart("F4", chart)
        if payload.get("follower_error"):
            note_row = 6 + len(payload["daily"])
            ws.merge_range(note_row, 0, note_row + 1, max(0, len(payload["daily"].columns) - 1),
                           "Catatan Instagram Insights: " + str(payload["follower_error"]), note_fmt)
            ws.set_row(note_row, 30)
            ws.set_row(note_row + 1, 24)
    return output.getvalue()


def _pdf_report(payload) -> bytes:
    """PDF landscape dengan ringkasan, funnel, tabel campaign, dan tren harian."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, LongTable, Flowable

    navy = colors.HexColor("#1A1464")
    blue = colors.HexColor("#497CD6")
    pale = colors.HexColor("#EAF1FF")
    muted = colors.HexColor("#65718F")
    ink = colors.HexColor("#24304F")

    class ReportLineChart(Flowable):
        def __init__(self, series, width=245*mm, height=32*mm, title="Tren harian"):
            super().__init__()
            self.series, self.width, self.height, self.title = series, width, height, title
        def draw(self):
            c = self.canv
            c.setFont("Helvetica-Bold", 9)
            c.setFillColor(navy)
            c.drawString(0, self.height - 10, self.title)
            left, bottom, top = 28, 17, self.height - 20
            right = self.width - 8
            points = [v for _, vals, _ in self.series for v in vals if v is not None and pd.notna(v)]
            if not points:
                c.setFillColor(muted); c.setFont("Helvetica", 8); c.drawString(left, bottom + 15, "Tidak ada data harian.")
                return
            ymin, ymax = min(0, min(points)), max(points)
            if ymax <= ymin: ymax = ymin + 1
            c.setStrokeColor(colors.HexColor("#D9E1F2")); c.setLineWidth(.5)
            for i in range(4):
                y = bottom + (top-bottom)*i/3
                c.line(left, y, right, y)
            c.setFont("Helvetica", 6.5); c.setFillColor(muted)
            c.drawString(0, bottom, f"{ymin:,.0f}")
            c.drawString(0, top-2, f"{ymax:,.0f}")
            palette = [blue, colors.HexColor("#D91A7D"), colors.HexColor("#2D9C9A")]
            chart_dates = next((dates for _, _, dates in self.series if dates), [])
            if chart_dates:
                c.setFillColor(muted); c.setFont("Helvetica", 6.5)
                first = pd.to_datetime(chart_dates[0], errors="coerce")
                last = pd.to_datetime(chart_dates[-1], errors="coerce")
                if pd.notna(first) and pd.notna(last):
                    c.drawRightString(right, self.height - 10, f"{first:%d %b} - {last:%d %b}")
            for si, (label, values, dates) in enumerate(self.series):
                valid = [(i, v) for i, v in enumerate(values) if v is not None and pd.notna(v)]
                if not valid: continue
                c.setStrokeColor(palette[si % len(palette)]); c.setLineWidth(1.6)
                prev = None
                n = max(len(values)-1, 1)
                for i, v in valid:
                    x = left + (right-left)*i/n
                    y = bottom + (top-bottom)*(float(v)-ymin)/(ymax-ymin)
                    if prev: c.line(prev[0], prev[1], x, y)
                    c.setFillColor(palette[si % len(palette)]); c.circle(x, y, 1.5, stroke=0, fill=1)
                    prev = (x, y)
                c.setFillColor(palette[si % len(palette)]); c.setFont("Helvetica", 7)
                c.drawString(left + si*55, 2, label)

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=13*mm, leftMargin=13*mm, topMargin=17*mm, bottomMargin=15*mm, title="SPN Instagram Ads Performance Report")
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="ReportTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=19, leading=23, textColor=navy, alignment=TA_LEFT, spaceAfter=4))
    styles.add(ParagraphStyle(name="ReportSub", parent=styles["Normal"], fontSize=9, textColor=muted, spaceAfter=10))
    styles.add(ParagraphStyle(name="Section", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=11, leading=14, textColor=navy, spaceBefore=8, spaceAfter=6))
    styles.add(ParagraphStyle(name="Small", parent=styles["Normal"], fontSize=7, leading=9, textColor=ink))

    def fmt_num(v):
        return "-" if v is None or pd.isna(v) else f"{float(v):,.0f}".replace(",", ".")
    def fmt_rp(v):
        return "-" if v is None or pd.isna(v) else "Rp " + fmt_num(v)
    def styled_table(data, widths=None, header=True, font=7):
        t = Table(data, colWidths=widths, repeatRows=1 if header else 0, hAlign="LEFT")
        cmds = [
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING", (0,0), (-1,-1), 5), ("RIGHTPADDING", (0,0), (-1,-1), 5),
            ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("GRID", (0,0), (-1,-1), .35, colors.HexColor("#DFE6F2")),
            ("FONTSIZE", (0,0), (-1,-1), font), ("TEXTCOLOR", (0,0), (-1,-1), ink),
        ]
        if header:
            cmds += [("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold")]
            if len(data) > 1:
                cmds += [("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, pale])]
        t.setStyle(TableStyle(cmds))
        return t

    story = [
        Paragraph("SPN | INSTAGRAM ADS PERFORMANCE REPORT", styles["ReportTitle"]),
        Paragraph(f"Periode laporan: {payload['start']:%d %b %Y} – {payload['end']:%d %b %Y}  ·  Dibuat: {datetime.now():%d %b %Y, %H:%M}", styles["ReportSub"]),
        Paragraph("Ringkasan KPI", styles["Section"]),
    ]
    m = payload["metrics"]
    metric_rows = [["Metrik", "Traffic", "Leads", "Total / Global"]]
    for _, row in m.iterrows():
        vals=[]
        for col in ("Traffic", "Leads", "Total / Global"):
            v = row[col]
            vals.append(fmt_rp(v) if "Spend" in row["Metrik"] or "Cost" in row["Metrik"] or "CPF" in row["Metrik"] else fmt_num(v))
        metric_rows.append([row["Metrik"], *vals])
    story.append(styled_table(metric_rows, [70*mm, 42*mm, 42*mm, 48*mm], font=7.5))

    story.append(Paragraph("Konversi ke Community", styles["Section"]))
    f = payload["funnel"]
    frow = [["Profile Visits", "Total Join Community", "CR", "CAC"]]
    vals = [f.iloc[0]["Nilai"], f.iloc[1]["Nilai"], f.iloc[2]["Nilai"], f.iloc[3]["Nilai"]]
    frow.append([fmt_num(vals[0]), fmt_num(vals[1]), ("-" if vals[2] is None or pd.isna(vals[2]) else f"{float(vals[2]):,.2f}%"), fmt_rp(vals[3])])
    story.append(styled_table(frow, [60*mm, 60*mm, 45*mm, 45*mm], font=8))

    story.append(Paragraph("Performa per Objective", styles["Section"]))
    obj = payload["objective"]
    obj_rows = [["Objective", "Campaign", "Spend", "Result", "Cost / Result", "Follower By Ads", "CPF", "Impressions", "Reach"]]
    for _, row in obj.iterrows():
        obj_rows.append([row["Objective"], fmt_num(row["Campaign"]), fmt_rp(row["Spend (Rp)"]), fmt_num(row["Result"]), fmt_rp(row["Cost per Result (Rp)"]), fmt_num(row["Follower By Ads"]), fmt_rp(row["CPF (Rp)"]), fmt_num(row["Impressions"]), fmt_num(row["Reach"])])
    story.append(styled_table(obj_rows, [24*mm, 18*mm, 30*mm, 22*mm, 30*mm, 26*mm, 25*mm, 24*mm, 22*mm], font=6.3))

    daily = payload["daily"]
    if not daily.empty:
        dates = daily["Tanggal"].tolist()
        spend_series = [
            ("Traffic Spend", daily["Traffic"].tolist(), dates),
            ("Leads Spend", daily["Leads"].tolist(), dates),
        ]
        story += [Spacer(1, 4), ReportLineChart(spend_series, title="Daily Ad Spend (Rp)")]

    story.append(Spacer(1, 8))
    story.append(Paragraph("Detail Campaign", styles["ReportTitle"]))
    story.append(Paragraph("Seluruh campaign yang dikembalikan oleh Meta pada periode terpilih.", styles["ReportSub"]))
    campaign_df = payload["campaigns"]
    ccols = [c for c in ["Kategori", "Campaign", "Spend (Rp)", "Impressions", "Reach", "Clicks", "CTR (%)", "Profile Visits", "Follower By Ads"] if c in campaign_df.columns]
    cdata = [["Kategori", "Campaign", "Spend", "Impr.", "Reach", "Clicks", "CTR", "Profile Visits", "Follows"]]
    for _, r in campaign_df[ccols].iterrows():
        cdata.append([
            str(r.get("Kategori", "")), Paragraph(html_lib.escape(str(r.get("Campaign", ""))[:120]), styles["Small"]),
            fmt_rp(r.get("Spend (Rp)")), fmt_num(r.get("Impressions")), fmt_num(r.get("Reach")),
            fmt_num(r.get("Clicks")), ("-" if pd.isna(r.get("CTR (%)")) else f"{float(r.get('CTR (%)')):,.2f}%"),
            fmt_num(r.get("Profile Visits")), fmt_num(r.get("Follower By Ads")),
        ])
    if len(cdata) == 1:
        cdata.append(["-", "Tidak ada data campaign", "-", "-", "-", "-", "-", "-", "-"])
    campaign_table = LongTable(cdata, colWidths=[22*mm, 66*mm, 27*mm, 21*mm, 20*mm, 19*mm, 18*mm, 25*mm, 22*mm], repeatRows=1)
    campaign_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 6.3),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"), ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#DFE6F2")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, pale]),
        ("LEFTPADDING", (0,0), (-1,-1), 4), ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(campaign_table)

    story.append(Spacer(1, 8))
    story.append(Paragraph("Daily Data", styles["Section"]))
    story.append(Paragraph("Belanja harian berdasarkan objective dan data follower baru dari Instagram Insights jika tersedia.", styles["ReportSub"]))
    if payload.get("follower_error"):
        story.append(Paragraph("Catatan Instagram Insights: " + html_lib.escape(str(payload["follower_error"])), styles["Small"]))
        story.append(Spacer(1, 4))
    ddata = [["Tanggal", "Traffic Spend", "Leads Spend", "Follower baru"]]
    for _, row in daily.iterrows():
        d = row["Tanggal"]
        ddata.append([d.strftime("%d %b %Y") if pd.notna(d) else "-", fmt_rp(row.get("Traffic")), fmt_rp(row.get("Leads")), fmt_num(row.get("Follower baru"))])
    if len(ddata) == 1:
        ddata.append(["-", "-", "-", "-"])
    story.append(LongTable(ddata, colWidths=[45*mm, 48*mm, 48*mm, 45*mm], repeatRows=1, style=TableStyle([
        ("BACKGROUND", (0,0), (-1,0), navy), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTSIZE", (0,0), (-1,-1), 7.5),
        ("GRID", (0,0), (-1,-1), .3, colors.HexColor("#DFE6F2")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, pale]),
        ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ])))

    def page_decor(canvas, doc_):
        canvas.saveState()
        w, h = landscape(A4)
        canvas.setStrokeColor(blue); canvas.setLineWidth(1.2); canvas.line(13*mm, h-10*mm, w-13*mm, h-10*mm)
        canvas.setFont("Helvetica", 7); canvas.setFillColor(muted)
        canvas.drawString(13*mm, 7*mm, "SPN · Internal Performance Report")
        canvas.drawRightString(w-13*mm, 7*mm, f"Page {doc_.page}")
        canvas.restoreState()
    doc.build(story, onFirstPage=page_decor, onLaterPages=page_decor)
    return buffer.getvalue()


def render_export_panel(ctx: Ctx, raw_campaigns):
    """Visible export panel with recognizable PDF and Excel file badges."""
    period_key = (str(ctx.start), str(ctx.end))
    cached = st.session_state.get("mx_full_report_files")
    if cached and cached.get("period") != period_key:
        st.session_state.pop("mx_full_report_files", None)
        cached = None

    st.markdown("""
    <style>
      .mx-export-format {
        display:flex; align-items:center; gap:.8rem; min-height:66px;
        padding:.72rem .8rem; border:1px solid var(--line,#DCE4F5);
        border-radius:12px; background:var(--panel,#F7F8FC); margin:.15rem 0 .55rem;
      }
      .mx-export-filemark {
        width:48px; height:52px; flex:0 0 48px; border-radius:9px;
        display:flex; align-items:center; justify-content:center;
        font-weight:850; font-size:.78rem; letter-spacing:-.02em;
        box-shadow:inset 0 0 0 1px rgba(0,0,0,.04);
      }
      .mx-export-filemark--pdf { color:#B42318; background:#FDE8E7; }
      .mx-export-filemark--xlsx { color:#167344; background:#E3F4EA; }
      .mx-export-filemark svg { width:27px; height:29px; display:block; margin:auto auto 1px; }
      .mx-export-filemark span { display:block; text-align:center; font-size:.62rem; line-height:1; }
      .mx-export-filemark-inner { display:flex; flex-direction:column; align-items:center; gap:1px; }
      .mx-export-format-title { color:var(--text-color,#20243A); font-weight:700; font-size:.9rem; line-height:1.25; }
      .mx-export-format-note { color:var(--muted,#65708A); font-size:.73rem; line-height:1.35; margin-top:.16rem; }
      @media (max-width:600px) {
        .mx-export-format { min-height:60px; padding:.6rem; gap:.6rem; }
        .mx-export-filemark { width:42px; height:46px; flex-basis:42px; }
      }
    </style>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        top, period_col = st.columns([2.5, 1.2], vertical_alignment="center")
        with top:
            st.markdown("**Export laporan**")
            st.caption("Unduh laporan performa Instagram Ads, funnel community, campaign, dan tren harian.")
        with period_col:
            st.caption("PERIODE LAPORAN")
            st.markdown(f"**{ctx.start:%d %b %Y} – {ctx.end:%d %b %Y}**")

        prepare = st.button(
            "Siapkan laporan", type="primary", icon=":material/description:",
            key="mx_prepare_full_report", width="stretch",
        )
        if prepare:
            try:
                with st.spinner("Menyusun laporan PDF dan Excel..."):
                    payload = _report_payload(ctx, raw_campaigns)
                    xlsx_bytes = _excel_report(payload)
                    pdf_bytes = _pdf_report(payload)
                    st.session_state["mx_full_report_files"] = {
                        "period": period_key, "xlsx": xlsx_bytes, "pdf": pdf_bytes,
                    }
                cached = st.session_state["mx_full_report_files"]
            except ImportError:
                st.error("Library export belum terpasang. Tambahkan reportlab dan XlsxWriter ke requirements.txt lalu deploy ulang.")
                return
            except Exception as exc:
                st.error(f"Laporan belum dapat dibuat: {exc}")
                return

        cached = st.session_state.get("mx_full_report_files")
        is_ready = bool(cached and cached.get("period") == period_key)
        base = f"SPN_Instagram_Ads_Report_{ctx.start:%Y%m%d}_{ctx.end:%Y%m%d}"
        pdf_col, excel_col = st.columns(2, gap="medium")
        with pdf_col:
            st.markdown("""
            <div class="mx-export-format">
              <div class="mx-export-filemark mx-export-filemark--pdf">
                <div class="mx-export-filemark-inner">
                  <svg viewBox="0 0 32 34" aria-hidden="true">
                    <path d="M7 1.8h12l7 7v22.4H7z" fill="currentColor" opacity=".18"/>
                    <path d="M19 1.8v7h7" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/>
                    <path d="M10 18.5h13M10 22h13" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
                  </svg>
                  <span>PDF</span>
                </div>
              </div>
              <div><div class="mx-export-format-title">PDF Report</div>
                <div class="mx-export-format-note">Dokumen siap dibagikan dan dicetak</div></div>
            </div>
            """, unsafe_allow_html=True)
            st.download_button(
                "Download PDF", cached["pdf"] if is_ready else b"",
                file_name=f"{base}.pdf", mime="application/pdf",
                width="stretch", disabled=not is_ready, key="mx_full_report_pdf",
                icon=":material/download:",
            )
        with excel_col:
            st.markdown("""
            <div class="mx-export-format">
              <div class="mx-export-filemark mx-export-filemark--xlsx">
                <div class="mx-export-filemark-inner">
                  <svg viewBox="0 0 32 34" aria-hidden="true">
                    <path d="M7 1.8h12l7 7v22.4H7z" fill="currentColor" opacity=".18"/>
                    <path d="M19 1.8v7h7" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/>
                    <path d="m10 17 5 7m0-7-5 7m7-3.5h5" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                  </svg>
                  <span>XLSX</span>
                </div>
              </div>
              <div><div class="mx-export-format-title">Excel Workbook</div>
                <div class="mx-export-format-note">Data lengkap dalam beberapa sheet</div></div>
            </div>
            """, unsafe_allow_html=True)
            st.download_button(
                "Download Excel", cached["xlsx"] if is_ready else b"",
                file_name=f"{base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch", disabled=not is_ready, key="mx_full_report_xlsx",
                icon=":material/download:",
            )
        if is_ready:
            st.caption("Laporan siap diunduh. Jika periode diubah, siapkan laporan kembali.")
        else:
            st.caption("Klik **Siapkan laporan** satu kali, lalu pilih format unduhan.")

# ---------------- titik masuk ----------------
META_CSS = """
<style>
.mx-kpis { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:.9rem; margin:.2rem 0 1.15rem; }
.mx-kpi { min-width:0; background:var(--panel,#F6F8FE); border:1px solid var(--line,#DCE4F5);
  border-top:3px solid var(--pink,#D91A7D); border-radius:14px; padding:.9rem 1rem .85rem;
  box-shadow:0 2px 10px rgba(26,20,100,.045); }
.mx-kpi-label { color:var(--muted,#55608A); font-size:.76rem; font-weight:600; letter-spacing:.015em; }
.mx-kpi-value { color:var(--navy,#1A1464); font-size:1.42rem; font-weight:800; line-height:1.3;
  letter-spacing:-.02em; margin-top:.14rem; font-variant-numeric:tabular-nums; overflow-wrap:anywhere; }
.mx-kpi-note { color:var(--muted,#55608A); font-size:.76rem; line-height:1.35; margin-top:.2rem; min-height:1.05rem; }
@media (max-width:1100px) { .mx-kpis { grid-template-columns:repeat(2,minmax(0,1fr)); } }
@media (max-width:600px) { .mx-kpis { grid-template-columns:repeat(2,minmax(0,1fr)); gap:.55rem; }
  .mx-kpi { padding:.72rem .75rem; } .mx-kpi-value { font-size:1.15rem; } }
</style>
"""


def render():
    """Tampilan Meta Ads ringkas dengan fokus penuh pada performa iklan Instagram."""
    # Sidebar utama SPN berisi filter pendaftar; tidak relevan untuk analytics iklan.
    st.markdown(META_CSS + """
    <style>
      [data-testid='stSidebar'], [data-testid='stSidebarCollapsedControl'],
      [data-testid='stExpandSidebarButton'] { display:none !important; }
    </style>
    """, unsafe_allow_html=True)
    if not _password_ok():
        return

    token, account = cfg("META_TOKEN"), cfg("AD_ACCOUNT_ID", "")
    version = cfg("API_VERSION", "v23.0")
    if not token or not account:
        st.error("Konfigurasi META_TOKEN dan AD_ACCOUNT_ID belum diisi pada Secrets Streamlit.")
        return
    account = str(account) if str(account).startswith("act_") else f"act_{account}"

    if st.session_state.get("mx_client_key") != (token, version):
        st.session_state.mx_client = MetaClient(token, version)
        st.session_state.mx_client_key = (token, version)
        st.session_state.pop("mx_health_v3", None)
        st.session_state.pop("mx_health_key_v3", None)
    client: MetaClient = st.session_state.mx_client

    menu = st.segmented_control(
        "Tampilan Meta Ads", ["Ringkasan", "Traffic", "Leads", "Admin"],
        default="Ringkasan", key="mx_section_v3", label_visibility="collapsed", width="stretch",
    ) or "Ringkasan"
    period = period_filter(account, version)
    if period is None:
        return
    start, end = period
    ctx = Ctx(client, account, version, None, start, end)
    st.caption(f"Instagram Ads · {start:%d %b %Y} – {end:%d %b %Y}")

    if menu == "Admin":
        render_data_check(ctx)
        with st.expander("Informasi akun dan log API", expanded=False):
            render_admin(ctx)
        return

    raw_campaigns = campaign_data(ctx)
    render_export_panel(ctx, raw_campaigns)
    if menu == "Ringkasan":
        render_overview(ctx, raw_campaigns)
    elif menu == "Traffic":
        render_stats(ctx, "Traffic", raw_campaigns)
    else:
        render_stats(ctx, "Leads", raw_campaigns)
