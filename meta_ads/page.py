"""Halaman Meta Ads untuk SPN Data Explorer.

Dipindahkan dari aplikasi Meta Ads yang berdiri sendiri (app.py) menjadi satu fungsi `render()`
yang dipanggil dari menu "Meta Ads" di app-publish.py. Perbedaan dengan app aslinya:

  * tidak memanggil st.set_page_config dan tidak menyuntik CSS global (tema SPN tetap dipakai)
  * navigasi (Traffic / Leads / Cek Data & Admin) dan filter periode ada di badan halaman,
    bukan di sidebar, karena sidebar SPN berisi filter data influencer yang tidak relevan di sini
  * semua key session_state diawali "mx_" supaya tidak bentrok dengan filter dashboard SPN
  * tombol Refresh hanya menghapus cache milik Meta Ads, bukan seluruh cache aplikasi
  * kalau secrets belum diisi, hanya halaman ini yang menampilkan pesan error
"""
from __future__ import annotations

import calendar
import hmac
import os
import re
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, timedelta

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
    for fn in (load_insights, load_campaigns, graph):
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


def headline_card(col, label, value):
    # Keep headline cards clean: no subtext/notes beneath the statistic.
    with col.container(border=True):
        st.caption(label)
        st.markdown(f"<div style='font-size:1.7rem;font-weight:700;line-height:1.22'>{value}</div>",
                    unsafe_allow_html=True)


def category_frame(df, category):
    if df is None or df.empty:
        return pd.DataFrame() if df is None else df.copy()
    return df[df["dashboard_category"] == category].copy()


# ---------------- halaman Traffic / Leads ----------------
def render_stats(ctx: Ctx, category, raw_campaigns):
    S, E = ctx.start, ctx.end
    campaigns = category_frame(raw_campaigns, category)
    if campaigns.empty:
        st.warning(f"Belum ada data campaign kategori {category} untuk periode ini.")
        other_count = 0 if raw_campaigns is None or raw_campaigns.empty else int((raw_campaigns["dashboard_category"] == "Lainnya").sum())
        if other_count:
            with st.expander("Campaign yang belum terklasifikasi"):
                st.dataframe(raw_campaigns.loc[raw_campaigns["dashboard_category"] == "Lainnya", ["campaign_name", "spend", "impressions"]], hide_index=True)
        return

    campaigns = prepare_results(campaigns)
    st.subheader("Ringkasan performa")
    spend = float(campaigns["spend"].sum())
    impressions = float(campaigns["impressions"].sum())
    clicks = float(campaigns["clicks"].sum())
    ctr = clicks / impressions * 100 if impressions else 0.0
    n_campaigns = int(campaigns["campaign_name"].nunique())

    ads = safe_run(load_insights, ctx.client, ctx.account, ctx.version, "ad", str(S), str(E), False, (), False, AD_FIELDS)
    if ads is None or ads.empty:
        n_ads = 0
    else:
        ads = tag_campaigns(ads)
        ads = category_frame(ads, category)
        if not ads.empty:
            ads = ads[pd.to_numeric(ads["impressions"], errors="coerce").fillna(0) > 0]
        n_ads = int(ads["ad_name"].nunique()) if not ads.empty and "ad_name" in ads else 0

    if category == "Traffic":
        visits = float(campaigns["profile_visits"].sum())
        avg_cost = div(spend, visits)
        cols = st.columns(4)
        headline_card(cols[0], "Total Spend", rp(spend))
        headline_card(cols[1], "IG Profile Visit", idn(visits))
        headline_card(cols[2], "CPV (Traffic)", rp(avg_cost))
        headline_card(cols[3], "Impresi", idn(impressions))
        cols = st.columns(4)
        headline_card(cols[0], "Reach", idn(campaigns["reach"].sum()))
        headline_card(cols[1], "Klik", idn(clicks))
        headline_card(cols[2], "CTR", f"{idn(ctr, 2)}%")
        headline_card(cols[3], "Content unik", idn(n_ads))
    else:
        result_action = lead_action_choice(campaigns)
        results = float(primary_action_per_campaign(campaigns, result_action).sum()) if result_action else 0.0
        cpl = div(spend, results)
        cols = st.columns(4)
        headline_card(cols[0], "Total Spend", rp(spend))
        headline_card(cols[1], "Hasil Leads", idn(results) if result_action else "-")
        headline_card(cols[2], "Cost per Lead", rp(cpl))
        headline_card(cols[3], "Impresi", idn(impressions))
        cols = st.columns(4)
        headline_card(cols[0], "Reach", idn(campaigns["reach"].sum()))
        headline_card(cols[1], "Klik", idn(clicks))
        headline_card(cols[2], "CTR", f"{idn(ctr, 2)}%")
        headline_card(cols[3], "Campaign", idn(n_campaigns))
        if not result_action:
            st.info("Meta tidak mengembalikan action Lead atau percakapan untuk periode ini. Spend dan metrik iklan tetap ditampilkan.")

    # Trend lines scoped to the active category and date filter.
    daily = safe_run(load_insights, ctx.client, ctx.account, ctx.version, "campaign", str(S), str(E), True)
    if daily is not None and not daily.empty and "date_start" in daily.columns:
        daily = tag_campaigns(daily)
        daily = category_frame(daily, category)
        if not daily.empty:
            daily["date_start"] = pd.to_datetime(daily["date_start"], errors="coerce").dt.normalize()
            trend = daily.groupby("date_start", as_index=True).agg(
                spend=("spend", "sum"), impressions=("impressions", "sum"), clicks=("clicks", "sum")
            ).sort_index()
            full_index = pd.date_range(pd.Timestamp(S), pd.Timestamp(E), freq="D")
            trend = trend.reindex(full_index, fill_value=0)
            trend.index.name = "Tanggal"
            st.subheader("Tren harian")
            left, right = st.columns(2)
            left.line_chart(trend["spend"])
            right.line_chart(trend["impressions"])
            left.bar_chart(trend["clicks"])

    st.subheader("Performa campaign")
    level_choice = st.selectbox(
        "Lihat rincian",
        ["Campaign", "Ad set", "Iklan"],
        key=f"mx_level_{category}",
    )
    level_map = {"Campaign": "campaign", "Ad set": "adset", "Iklan": "ad"}
    level = level_map[level_choice]
    if level == "campaign":
        detail = campaigns.copy()
    else:
        detail_raw = safe_run(load_insights, ctx.client, ctx.account, ctx.version, level, str(S), str(E))
        detail = tag_campaigns(detail_raw) if detail_raw is not None and not detail_raw.empty else pd.DataFrame()
        detail = category_frame(detail, category) if not detail.empty else detail
    if detail is not None and not detail.empty:
        chart_metric = st.selectbox(
            "Metrik grafik", ["spend", "impressions", "clicks", "reach", "ctr", "cpc", "cpm"],
            key=f"mx_chartmetric_{category}",
        )
        top_n = st.slider("Jumlah teratas", 5, 30, 10, key=f"mx_topn_{category}")
        label_col = {"campaign": "campaign_name", "adset": "adset_name", "ad": "ad_name"}[level]
        if label_col in detail.columns and chart_metric in detail.columns:
            chart_frame = detail.nlargest(top_n, chart_metric).set_index(label_col)[chart_metric]
            st.bar_chart(chart_frame, horizontal=True)
        shown_names = [col for col in ("campaign_name", "adset_name", "ad_name") if col in detail.columns]
        shown_metrics = [col for col in METRICS if col in detail.columns]
        show = detail.sort_values("spend", ascending=False)[shown_names + shown_metrics]
        st.dataframe(show, hide_index=True, width="stretch")
        st.download_button(
            "Download CSV", show.to_csv(index=False).encode("utf-8"),
            f"meta_{category.lower()}_{level}_{S}_{E}.csv", "text/csv", key=f"mx_download_{category}",
        )

    with st.expander("Breakdown performa"):
        breakdowns = {
            "Platform": ("publisher_platform",),
            "Placement": ("publisher_platform", "platform_position"),
            "Umur": ("age",),
            "Gender": ("gender",),
            "Umur × Gender": ("age", "gender"),
            "Device": ("impression_device",),
            "Negara": ("country",),
            "Region": ("region",),
        }
        b1, b2 = st.columns(2)
        selected_breakdown = b1.selectbox("Dimensi", list(breakdowns), key=f"mx_bd_{category}")
        selected_metric = b2.selectbox("Metrik breakdown", ["spend", "impressions", "clicks", "reach", "ctr", "cpc"], key=f"mx_bdmetric_{category}")
        dims = breakdowns[selected_breakdown]
        bd_raw = safe_run(load_insights, ctx.client, ctx.account, ctx.version, "campaign", str(S), str(E), False, dims)
        bd = tag_campaigns(bd_raw) if bd_raw is not None and not bd_raw.empty else pd.DataFrame()
        if not bd.empty:
            bd = category_frame(bd, category)
            if not bd.empty and all(dim in bd.columns for dim in dims):
                if len(dims) == 2:
                    st.bar_chart(bd, x=dims[0], y=selected_metric, color=dims[1])
                else:
                    st.bar_chart(bd.set_index(dims[0])[selected_metric])
                columns = [c for c in ["campaign_name", *dims, *METRICS] if c in bd.columns]
                st.dataframe(bd[columns], hide_index=True, width="stretch")

    st.subheader("Actions dari Meta")
    actions = action_rows(campaigns)
    if actions.empty:
        st.info("Belum ada action yang dikembalikan Meta untuk kategori dan periode ini.")
    else:
        action_summary = actions.groupby("action_type", as_index=False)["hasil"].sum().sort_values("hasil", ascending=False)
        st.bar_chart(action_summary.head(15).set_index("action_type")["hasil"], horizontal=True)
        st.dataframe(actions.sort_values("hasil", ascending=False), hide_index=True, width="stretch")

    other_count = 0 if raw_campaigns is None or raw_campaigns.empty else int((raw_campaigns["dashboard_category"] == "Lainnya").sum())
    if other_count:
        with st.expander(f"{other_count} campaign belum terklasifikasi sebagai Traffic atau Leads"):
            st.dataframe(raw_campaigns.loc[raw_campaigns["dashboard_category"] == "Lainnya", ["campaign_name", "spend", "impressions"]], hide_index=True)


# ---------------- Cek Data & Admin ----------------
def render_data_check(ctx: Ctx):
    S, E = ctx.start, ctx.end
    st.subheader("Pemeriksaan alur data")
    st.caption("Validasi token, akses ad account, dan konsistensi data Insights.")
    health_key = (S, E, ctx.account, ctx.business_id)
    if st.button("Jalankan ulang pengecekan", key="mx_run_health_check") or st.session_state.get("mx_health_key") != health_key:
        with st.spinner("Memeriksa token, akses akun, dan Insights..."):
            checks, hctx = run_health_checks(ctx.client, ctx.account, S, E, ctx.business_id)
        st.session_state.mx_health = dict(checks=checks, ctx=hctx, rng=(S, E), at=datetime.now())
        st.session_state.mx_health_key = health_key
    H = st.session_state.get("mx_health", {"checks": [], "ctx": {}, "rng": (S, E), "at": datetime.now()})
    counts = Counter(check.status for check in H["checks"])
    if counts["fail"]:
        st.error(f"{counts['fail']} pemeriksaan gagal.")
    elif counts["warn"]:
        st.warning(f"Ada {counts['warn']} peringatan.")
    else:
        st.success("Pemeriksaan selesai.")
    cols = st.columns(4)
    for col, (key, label) in zip(cols, [("pass", "Lolos"), ("warn", "Peringatan"), ("fail", "Gagal"), ("info", "Info")]):
        col.metric(label, counts[key])
    st.caption(f"Pengecekan terakhir: {H['at']:%H:%M:%S} · {H['rng'][0]} – {H['rng'][1]}")
    st.dataframe(pd.DataFrame([{
        "Status": check.status.upper(), "Grup": check.group,
        "Pengecekan": check.name, "Detail": check.detail,
    } for check in H["checks"]]), hide_index=True, width="stretch")
    daily = H["ctx"].get("daily")
    if daily is not None and not daily.empty:
        st.markdown("#### Rekonsiliasi ringkas akun")
        cols = st.columns(3)
        cols[0].metric("Spend akun", money(daily["spend"].sum()))
        cols[1].metric("Impresi akun", idn(daily["impressions"].sum()))
        cols[2].metric("Klik akun", idn(daily["clicks"].sum()))
        st.dataframe(
            daily.sort_values("date_start", ascending=False).head(15),
            hide_index=True, width="stretch"
        )


def render_admin(ctx: Ctx):
    client, version, account = ctx.client, ctx.version, ctx.account
    st.subheader("Administrasi akun")
    st.caption("Informasi ad account, Business Portfolio, Facebook Pages, dan log API.")

    with st.expander("Ad account", expanded=True):
        acc = safe_run(
            graph, client, version, account,
            (("fields", "name,account_status,timezone_name"),), False
        )
        if acc:
            cols = st.columns(3)
            cols[0].metric("Nama akun", acc.get("name", "-"))
            cols[1].metric("Status", ACCOUNT_STATUS.get(acc.get("account_status"), "-"))
            cols[2].metric("Timezone", acc.get("timezone_name", "-"))

    with st.expander("Business Portfolio & Pages", expanded=False):
        biz = ctx.business_id
        if not biz:
            try:
                biz = (graph(client, version, account, (("fields", "business"),), False).get("business") or {}).get("id")
            except MetaAPIError:
                biz = None
        if not biz:
            st.info("Business ID tidak ditemukan. Isi BUSINESS_ID di secrets jika fitur ini diperlukan.")
        else:
            st.markdown("##### Ad account dalam portfolio")
            owned = safe_run(
                graph, client, version, f"{biz}/owned_ad_accounts",
                (("fields", "id,name,account_status"), ("limit", "100"))
            )
            if owned:
                owned_df = pd.DataFrame(owned)
                if "account_status" in owned_df:
                    owned_df["account_status"] = owned_df["account_status"].map(ACCOUNT_STATUS)
                st.dataframe(owned_df, hide_index=True, width="stretch")

            st.markdown("##### Facebook Pages")
            pages = safe_run(
                graph, client, version, f"{biz}/owned_pages",
                (("fields", "id,name"), ("limit", "100"))
            )
            if pages:
                page_name = st.selectbox("Pilih Page", [p["name"] for p in pages], key="mx_admin_page_select")
                page_id = next(p["id"] for p in pages if p["name"] == page_name)
                page_token = safe_run(graph, client, version, page_id, (("fields", "access_token"),), False)
                page_token = (page_token or {}).get("access_token")
                info = safe_run(
                    graph, client, version, page_id,
                    (("fields", "name,fan_count,followers_count"),), False, page_token
                )
                if info:
                    cols = st.columns(2)
                    cols[0].metric("Fans", idn(info.get("fan_count", 0)))
                    cols[1].metric("Followers", idn(info.get("followers_count", 0)))
                post_fields = (
                    "id,created_time,message,permalink_url,shares,"
                    "reactions.summary(true).limit(0),comments.summary(true).limit(0)"
                )
                posts = safe_run(
                    graph, client, version, f"{page_id}/posts",
                    (("fields", post_fields), ("limit", "25")), False, page_token
                )
                if posts and posts.get("data"):
                    post_df = pd.DataFrame([{
                        "Tanggal": post.get("created_time", "")[:10],
                        "Pesan": (post.get("message") or "")[:100],
                        "Reaksi": post.get("reactions", {}).get("summary", {}).get("total_count", 0),
                        "Komentar": post.get("comments", {}).get("summary", {}).get("total_count", 0),
                        "Share": post.get("shares", {}).get("count", 0),
                        "Link": post.get("permalink_url"),
                    } for post in posts["data"]])
                    st.dataframe(post_df, hide_index=True, width="stretch")
                else:
                    st.info("Belum ada post yang dapat dibaca pada Page ini.")
            else:
                st.info("Belum ada Facebook Page yang terbaca dari Business Portfolio ini.")

    with st.expander("API Log", expanded=False):
        logs = pd.DataFrame(client.log[::-1])
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
    """Kembalikan (start, end), atau None kalau rentang tanggal belum lengkap / tidak valid."""
    today = date.today()
    with st.container(border=True):
        c_mode, c_pick, c_btn = st.columns([1.3, 3, 1], vertical_alignment="bottom")
        mode = c_mode.radio("Filter periode", ["Bulanan", "Rentang tanggal"], key="mx_date_mode", horizontal=True)
        if mode == "Bulanan":
            years = list(range(today.year - 4, today.year + 1))
            c_year, c_month = c_pick.columns(2)
            year = c_year.selectbox("Tahun", years, index=years.index(today.year), key="mx_filter_year")
            month = c_month.selectbox(
                "Bulan", list(range(1, 13)), index=today.month - 1,
                format_func=lambda m: MONTHS_ID[m - 1], key="mx_filter_month",
            )
            start = date(year, month, 1)
            last_day = calendar.monthrange(year, month)[1]
            # Bulan berjalan: month-to-date; bulan lampau: satu bulan penuh.
            end = date(year, month, min(last_day, today.day)) if (year, month) == (today.year, today.month) else date(year, month, last_day)
        else:
            values = c_pick.date_input("Rentang tanggal", (today - timedelta(days=29), today), key="mx_custom_date_range")
            if not isinstance(values, (tuple, list)) or len(values) != 2:
                st.info("Pilih tanggal mulai dan selesai.")
                return None
            start, end = values
            if start > end:
                st.error("Tanggal mulai harus sebelum tanggal selesai.")
                return None
        if c_btn.button("Refresh data", width="stretch", key="mx_refresh_data"):
            clear_meta_cache()
            st.rerun()
        st.caption(f"Ad account: {account} · API version: {version}")
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
def render():
    """Isi halaman Meta Ads. Dipanggil dari app-publish.py."""
    # Sidebar SPN berisi filter data influencer; tidak berlaku untuk data iklan, jadi disembunyikan di halaman ini.
    st.markdown(
        "<style>[data-testid='stSidebar'],[data-testid='stSidebarCollapsedControl'],"
        "[data-testid='stExpandSidebarButton']{display:none !important;}</style>",
        unsafe_allow_html=True,
    )
    if not _password_ok():
        return

    token, account = cfg("META_TOKEN"), cfg("AD_ACCOUNT_ID", "")
    version, business_id = cfg("API_VERSION", "v23.0"), cfg("BUSINESS_ID")
    if not token or not account:
        st.error("META_TOKEN / AD_ACCOUNT_ID belum diisi di .streamlit/secrets.toml (atau Secrets di Streamlit Cloud).")
        return
    account = str(account) if str(account).startswith("act_") else f"act_{account}"

    if st.session_state.get("mx_client_key") != (token, version):
        st.session_state.mx_client = MetaClient(token, version)
        st.session_state.mx_client_key = (token, version)
        st.session_state.pop("mx_health", None)
        st.session_state.pop("mx_health_key", None)
    client: MetaClient = st.session_state.mx_client

    menu = st.radio("Bagian", ["Traffic", "Leads", "Cek Data & Admin"], horizontal=True,
                    key="mx_menu", label_visibility="collapsed")
    period = period_filter(account, version)
    if period is None:
        return
    start, end = period
    ctx = Ctx(client, account, version, business_id, start, end)
    period_label = f"{start:%d %b %Y} – {end:%d %b %Y}"

    if menu == "Traffic":
        st.subheader("Traffic")
        st.caption(f"Performa campaign yang berfokus pada kunjungan profil · {period_label}")
        render_stats(ctx, "Traffic", campaign_data(ctx))
    elif menu == "Leads":
        st.subheader("Leads")
        st.caption(f"Performa campaign yang berfokus pada lead dan percakapan · {period_label}")
        render_stats(ctx, "Leads", campaign_data(ctx))
    else:
        st.subheader("Cek Data & Admin")
        st.caption(f"Monitoring alur data dan administrasi akun · {period_label}")
        check_tab, admin_tab = st.tabs(["Cek Data", "Admin"])
        with check_tab:
            render_data_check(ctx)
        with admin_tab:
            render_admin(ctx)
