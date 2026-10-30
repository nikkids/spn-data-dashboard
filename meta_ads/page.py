"""Halaman Meta Ads untuk SPN Data Explorer.

Dipindahkan dari aplikasi Meta Ads yang berdiri sendiri (app.py) menjadi satu fungsi `render()`
yang dipanggil dari menu "Meta Ads" di app-publish.py. Perbedaan dengan app aslinya:

  * tidak memanggil st.set_page_config dan tidak menyuntik CSS global (tema SPN tetap dipakai)
  * tampilan ringkasan, Traffic, Leads, dan koneksi dibuat ringkas mengikuti tema SPN
  * laporan hanya menampilkan metrik iklan dan tidak mengambil data Facebook Pages
  * semua key session_state diawali "mx_" supaya tidak bentrok dengan filter dashboard SPN
  * tombol Refresh hanya menghapus cache milik Meta Ads, bukan seluruh cache aplikasi
  * kalau secrets belum diisi, hanya halaman ini yang menampilkan pesan error
"""
from __future__ import annotations

import hmac
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
    """Ambil jumlah follower aktual dan net follows/unfollows harian dari Instagram Insights."""
    ig_user_id = cfg("IG_USER_ID", cfg("INSTAGRAM_USER_ID", ""))
    if not ig_user_id:
        return None, pd.DataFrame(), "Isi IG_USER_ID pada Streamlit Secrets untuk menghubungkan akun Instagram."

    profile = safe_run(
        graph, ctx.client, ctx.version, str(ig_user_id),
        (("fields", "username,followers_count"),), False
    )
    follower_total = None if not profile else profile.get("followers_count")

    raw = safe_run(
        graph, ctx.client, ctx.version, f"{ig_user_id}/insights",
        (("metric", "follows_and_unfollows"),
         ("period", "day"),
         ("since", str(ctx.start)),
         ("until", str(ctx.end))), False
    )
    rows = []
    if isinstance(raw, dict):
        for metric in raw.get("data", []):
            for point in metric.get("values", []):
                value = point.get("value") or {}
                if isinstance(value, dict):
                    follows = float(value.get("follows", 0) or 0)
                    unfollows = float(value.get("unfollows", 0) or 0)
                    rows.append({
                        "Tanggal": pd.to_datetime(point.get("end_time"), errors="coerce").normalize(),
                        "Follows": follows,
                        "Unfollows": unfollows,
                        "Perubahan bersih": follows - unfollows,
                    })
    trend = pd.DataFrame(rows)
    if not trend.empty:
        trend = trend.dropna(subset=["Tanggal"]).set_index("Tanggal").sort_index()
        trend = trend[~trend.index.duplicated(keep="last")]
    return follower_total, trend, None


def render_stats(ctx: Ctx, category, raw_campaigns):
    """Traffic dan Leads memakai susunan KPI yang sama; tren menampilkan perubahan follower IG."""
    campaigns = category_frame(raw_campaigns, category) if raw_campaigns is not None else pd.DataFrame()
    if campaigns.empty:
        st.info(f"Belum ada campaign {category} pada periode ini.")

    campaigns = prepare_results(campaigns)
    spend = _series_sum(campaigns, "spend")
    ad_follows = _ad_follow_count(campaigns)
    cost_per_follow = div(spend, ad_follows)

    if category == "Traffic":
        result = _series_sum(campaigns, "profile_visits")
        result_label = "Kunjungan profil"
        cost_per_result = div(spend, result)
    else:
        result_action = lead_action_choice(campaigns)
        meta_lead_results = (
            float(primary_action_per_campaign(campaigns, result_action).sum())
            if result_action else 0.0
        )
        csv_path = str(cfg("SPN_CSV_PATH", "master_penting_3.csv"))
        registrations = _spn_registrations(str(ctx.start), str(ctx.end), csv_path)
        result = registrations
        result_label = "Pendaftar SPN unik"
        cost_per_result = div(spend, registrations or 0) if registrations is not None else None

    follower_total, follower_trend, follower_error = _instagram_profile_and_trend(ctx)

    headline_cards([
        ("Result", idn(result) if result is not None else "-", result_label),
        ("Cost per Result", rp(cost_per_result), "Spend dibagi result pada periode terpilih"),
        ("Follower By Ads", idn(ad_follows), "Follow yang dilaporkan Meta dari campaign ini"),
        ("Follower Asli IG", idn(follower_total) if follower_total is not None else "-",
         "Jumlah follower aktual akun Instagram"),
        ("Cost Per Follow", rp(cost_per_follow), "Spend dibagi Follower By Ads"),
        ("Total Spend", rp(spend), f"{category} · {ctx.start:%d %b %Y} – {ctx.end:%d %b %Y}"),
    ])

    if category == "Leads":
        if result is None:
            st.warning(
                "Data pendaftar SPN belum dapat dibaca. Pastikan master_penting_3.csv tersedia "
                "atau atur SPN_CSV_PATH pada Secrets."
            )
        else:
            st.caption(
                f"Cost per Result menggunakan {idn(result)} email unik yang mendaftar ke SPN "
                f"pada {ctx.start:%d %b %Y}–{ctx.end:%d %b %Y}. "
                "Hasil Leads dari Meta tidak dianggap sama dengan jumlah pendaftar SPN."
            )
            if meta_lead_results <= 0:
                st.caption("Meta belum mengembalikan action Leads untuk campaign pada periode ini.")

    with st.container(border=True):
        st.markdown('<div class="section-title">Tren follower Instagram</div>', unsafe_allow_html=True)
        if follower_error:
            st.info(follower_error)
        elif follower_trend.empty:
            st.info(
                "Data kenaikan/penurunan follower belum tersedia. Periksa izin Instagram Insights "
                "dan pastikan IG_USER_ID mengarah ke Instagram Professional Account."
            )
        else:
            full_index = pd.date_range(pd.Timestamp(ctx.start), pd.Timestamp(ctx.end), freq="D")
            plot = follower_trend.reindex(full_index, fill_value=0)
            plot.index.name = "Tanggal"
            st.caption("Perubahan bersih harian = follows baru − unfollows. Nilai positif berarti naik.")
            st.line_chart(plot[["Perubahan bersih"]], height=280)
            with st.expander("Rincian follows dan unfollows", expanded=False):
                detail = plot.rename_axis("Tanggal").reset_index()
                detail["Tanggal"] = detail["Tanggal"].dt.strftime("%d %b %Y")
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
        st.download_button(
            "Download data campaign (CSV)",
            display_raw.to_csv(index=False).encode("utf-8"),
            f"instagram_{category.lower()}_campaign_{ctx.start}_{ctx.end}.csv",
            "text/csv", key=f"mx_download_top_v4_{category}",
        )

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
                st.download_button(
                    "Download rincian (CSV)", shown.to_csv(index=False).encode("utf-8"),
                    f"instagram_{category.lower()}_{level}_{ctx.start}_{ctx.end}.csv",
                    "text/csv", key=f"mx_download_detail_v4_{category}",
                )
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
    if menu == "Ringkasan":
        render_overview(ctx, raw_campaigns)
    elif menu == "Traffic":
        render_stats(ctx, "Traffic", raw_campaigns)
    else:
        render_stats(ctx, "Leads", raw_campaigns)
