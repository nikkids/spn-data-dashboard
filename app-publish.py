"""
SPN Data Explorer
-----------------
Alat penjelajah data berbasis pencarian untuk master_penting.csv, dibagi menjadi beberapa halaman
(menu halaman ada di bagian atas halaman utama):

    Data                 tabel, pencarian, filter, detail pendaftar, dan ekspor Excel
    Ringkasan            angka utama, tren pendaftaran (interaktif), perbandingan kelompok
    Audiens              usia, kota, hijab, keluarga
    Jangkauan sosmed     pengikut dan tier Instagram/TikTok, scatter interaktif 2D/3D
    Ekonomi dan profil   pendapatan, pengeluaran, pekerjaan, dan konten

Filter di sidebar berlaku di semua halaman.

Jalankan:  pip install -U streamlit pandas altair xlsxwriter plotly
           streamlit run app.py

Letakkan master_penting.csv dan folder .streamlit (berisi config.toml) di samping file ini.
Plotly opsional: tanpa plotly, grafik interaktif/3D otomatis diganti grafik Altair biasa.
"""
import base64
import functools
import html
import io
import re
import traceback
from contextlib import contextmanager
from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

try:
    import xlsxwriter  # noqa: F401

    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

try:
    import plotly.graph_objects as go

    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

alt.data_transformers.disable_max_rows()

# Logo SPN, letakkan di folder yang sama dengan file aplikasi:
#   logo-spn.jpeg            logo persegi / ikon tab
#   logo-spn-landscape.png   logo horizontal untuk sidebar dan Excel
LOGO_PATH = Path(__file__).parent / "logo-spn.jpeg"
LOGO_WIDE_PATH = Path(__file__).parent / "logo-spn-landscape.png"


def _insert_excel_logo(ws, cell: str = "B2") -> bool:
    """Sematkan logo SPN berukuran menonjol ke workbook Excel hasil ekspor."""
    logo_file = LOGO_WIDE_PATH if LOGO_WIDE_PATH.is_file() else LOGO_PATH
    if not logo_file.is_file():
        return False

    # Logo berfungsi sebagai elemen brand/subtitle yang jelas, bukan watermark kecil.
    # Batas 250 x 72 px; rasio gambar selalu dipertahankan.
    options = {
        "object_position": 1,
        "description": "Logo Smart Parents Network",
        "x_offset": 4,
        "y_offset": 2,
    }
    try:
        from PIL import Image

        with Image.open(logo_file) as image_file:
            image_width, image_height = image_file.size
        if image_width > 0 and image_height > 0:
            scale = min(250 / image_width, 72 / image_height)
            options["x_scale"] = scale
            options["y_scale"] = scale
        else:
            options.update({"x_scale": 0.45, "y_scale": 0.45})
    except Exception:
        # Fallback aman jika Pillow tidak tersedia atau metadata gambar tidak terbaca.
        options.update({"x_scale": 0.45, "y_scale": 0.45})

    ws.insert_image(cell, str(logo_file), options)
    return True


def _load_logo():
    try:
        from PIL import Image

        return Image.open(LOGO_PATH)
    except Exception:  # file belum ada atau tidak terbaca: pakai ikon bawaan
        return None


st.set_page_config(page_title="SPN Data Explorer", page_icon=_load_logo() or ":material/monitoring:", layout="wide")
if LOGO_WIDE_PATH.exists():
    st.logo(str(LOGO_WIDE_PATH), size="large")

MIN_STREAMLIT = (1, 65)  # versi tempat aplikasi ini dibuat dan diuji


def _version(v: str):
    return tuple(int(x) for x in re.findall(r"\d+", v)[:2])


if _version(st.__version__) < MIN_STREAMLIT:
    st.error(
        f"Aplikasi ini membutuhkan Streamlit ≥ {MIN_STREAMLIT[0]}.{MIN_STREAMLIT[1]} "
        f"(versi Anda {st.__version__}). Perbarui dengan:  `pip install -U streamlit`"
    )
    st.stop()

CSV_PATH = "master_penting_3.csv"

BLUE = "#4A7CD9"
BLUE_MUTED = "#2F5FC0"

TIER_ORDER = ["<1K", "Nano", "Mikro", "Makro"]
CHILD_ORDER = [
    "Belum Memiliki anak", "New Born (0-3 bulan)", "New Born (0-6 bulan)",
    "Baby (4-11 bulan)", "Infant (7-11 bulan)", "Toddler (1-5 tahun)",
    "Kids (6-13 tahun)", "Dewasa >13 tahun", "Unknown",
]
SEARCH_COLUMN = "_b_all"  # satu kolom pencarian yang mencakup semua kolom data
FOLLOWER_BINS = ["<1K", "1-3K", "3-10K", "10-30K", "30-100K", "100K-1M", "≥1M"]
MONTHS_ID = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]


# ============================================================ TAMPILAN (CSS)
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&display=swap');
:root {
  --bg: #FFFFFF; --panel: #F6F8FE; --panel-2: #FFFFFF; --line: #DCE4F5; --line-2: #BFCCEC;
  --text: #1A1464; --muted: #55608A; --faint: #69759F;
  --blue: #4A7CD9; --blue-hi: #2F5FC0; --blue-soft: rgba(74, 124, 217, 0.12);
  --pink: #D91A7D; --pink-hi: #B8136A; --navy: #1A1464;
}

/* elemen CSS tidak boleh menyisakan ruang kosong */
[data-testid="stElementContainer"]:has(> [data-testid="stMarkdown"] style) { display: none; }

/* header bawaan Streamlit (tempat menu halaman di bagian atas) */
[data-testid="stAppDeployButton"] { display: none; }
[data-testid="stHeader"] { background: var(--blue); border-bottom: 0; }
[data-testid="stHeader"], [data-testid="stHeader"] * { color: #fff; }
[data-testid="stHeader"] a, [data-testid="stHeader"] button {
  border-radius: 999px; font-weight: 600; white-space: nowrap;
}
/* Navigasi desktop lebih rapat agar seluruh menu utama muat tanpa masuk ke "1 more". */
[data-testid="stHeader"] nav { gap: 0.2rem !important; }
[data-testid="stHeader"] nav a {
  padding: 0.32rem 0.58rem !important; font-size: 0.88rem !important; gap: 0.32rem !important;
}
[data-testid="stHeader"] nav button { padding: 0.32rem 0.5rem !important; }
[data-testid="stHeader"] [data-testid*="Navigation"] { gap: 0.2rem !important; }
/* logo di samping judul halaman */
.page-head { display: flex; align-items: center; gap: 1rem; }
.page-logo { height: 3.2rem; width: auto; border-radius: 12px; display: block; flex: none; }
[data-testid="stHeader"] a:hover, [data-testid="stHeader"] button:hover { background: rgba(255, 255, 255, 0.16); }
[data-testid="stHeader"] [aria-current="page"] { background: #fff; }
[data-testid="stHeader"] [aria-current="page"], [data-testid="stHeader"] [aria-current="page"] * { color: var(--navy); }

/* kanvas utama */
[data-testid="stMainBlockContainer"] { padding: 4.6rem 2.6rem 3rem; max-width: 1600px; }
[data-testid="stMain"] [data-testid="stVerticalBlock"] { gap: 1.1rem; }

/* tipografi */
.count { font-size: 1.9rem; font-weight: 650; letter-spacing: -0.02em; line-height: 1.1; font-variant-numeric: tabular-nums; }
.count small { font-size: 0.95rem; font-weight: 400; color: var(--muted); letter-spacing: 0; margin-left: 0.35rem; }
.section-title { font-size: 1.05rem; font-weight: 600; letter-spacing: -0.005em; margin: 0; }
.subtle { color: var(--muted); font-size: 0.85rem; }
[data-testid="stCaptionContainer"] { color: var(--muted); font-size: 0.82rem; line-height: 1.45; }

/* label widget: kecil dan senyap */
[data-testid="stWidgetLabel"] p, .field-label {
  font-size: 0.76rem; font-weight: 600; color: var(--muted); letter-spacing: 0.02em;
}
.field-label { margin-bottom: 0; }

/* tombol */
[data-testid="stBaseButton-secondary"], [data-testid="stBaseButton-primary"], [data-testid="stBaseButton-tertiary"] {
  font-size: 0.86rem; font-weight: 550; min-height: 2.35rem; box-shadow: none; transition: border-color .12s, background .12s;
}
[data-testid="stBaseButton-secondary"] { background: transparent; border: 1.5px solid var(--navy); color: var(--navy); border-radius: 999px; }
[data-testid="stBaseButton-secondary"]:hover:not(:disabled) { border-color: var(--blue); background: var(--blue-soft); color: var(--text); }
[data-testid="stBaseButton-secondary"]:disabled { opacity: 0.4; }
[data-testid="stBaseButton-primary"] { background: var(--pink); border: 1px solid var(--pink); color: #fff; border-radius: 999px; }
[data-testid="stBaseButton-primary"]:hover:not(:disabled) { background: var(--pink-hi); border-color: var(--pink-hi); }
[data-testid="stBaseButton-tertiary"] { color: var(--muted); }
[data-testid="stBaseButton-tertiary"]:hover:not(:disabled) { color: var(--text); background: var(--blue-soft); }
[data-testid="stBaseButton-tertiary"]:disabled { opacity: 0.35; }

/* chip filter aktif */
.st-key-chips { gap: 0.4rem; }
.st-key-chips [data-testid="stBaseButton-secondary"] {
  min-height: 1.85rem; padding: 0.1rem 0.65rem 0.1rem 0.55rem; font-size: 0.8rem; font-weight: 500;
  background: var(--blue-soft); border: 1px solid rgba(47, 95, 192, 0.45); color: var(--blue-hi);
}
.st-key-chips [data-testid="stBaseButton-secondary"]:hover { background: rgba(47, 95, 192, 0.28); border-color: var(--blue); color: var(--text); }

/* kolom pencarian */
[data-testid="stTextInput"] input { font-size: 0.95rem; }

/* sidebar */
[data-testid="stSidebar"] { border-right: 1px solid var(--line); }
[data-testid="stSidebarUserContent"] { padding: 1.3rem 1.1rem 2rem; }
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.8rem; }
.sb-title { font-size: 1.05rem; font-weight: 650; letter-spacing: -0.005em; }
.sb-count { font-size: 0.86rem; color: var(--muted); font-variant-numeric: tabular-nums; margin-top: 0; }
.sb-count b { color: var(--text); font-weight: 650; }

/* kotak pencarian: satu blok bersama pilihan "cari di", dibedakan dari filter di bawahnya */
[data-testid="stSidebar"] .st-key-searchbox {
  background: var(--panel-2); border: 1px solid var(--blue); border-radius: 10px;
  padding: 1rem 0.9rem 1.05rem; gap: 0.6rem; box-shadow: 0 0 0 3px var(--blue-soft); margin: 0.5rem 0 0.4rem;
}
.st-key-searchbox [data-testid="stWidgetLabel"] { min-height: 0; margin-bottom: 0.1rem; }
.st-key-searchbox [data-testid="stWidgetLabel"] p {
  font-size: 0.72rem; font-weight: 700; letter-spacing: 0.09em; text-transform: uppercase; color: var(--blue-hi);
}

[data-testid="stSidebar"] [data-testid="stExpander"] { border: none; background: transparent; border-radius: 0; border-top: 1px solid var(--line); }
[data-testid="stSidebar"] [data-testid="stExpander"] details { border: none; }
[data-testid="stSidebar"] [data-testid="stExpander"] summary { padding: 0.85rem 0.1rem; }
[data-testid="stSidebar"] [data-testid="stExpander"] summary p {
  font-size: 0.72rem; font-weight: 700; letter-spacing: 0.09em; text-transform: uppercase; color: var(--blue-hi);
}
[data-testid="stSidebar"] [data-testid="stExpanderDetails"] { padding: 0.1rem 0.1rem 1rem; }

/* input angka: sembunyikan stepper dan petunjuk */
[data-testid="stNumberInput"] button { display: none; }
[data-testid="stNumberInput"] [data-testid="InputInstructions"] { display: none; }

/* kontrol segmen: boleh turun baris agar tidak saling tabrak di layar sempit */
[data-testid="stButtonGroup"], [data-testid="stButtonGroup"] > div { flex-wrap: wrap; row-gap: 0.3rem; }

/* logo Instagram dan TikTok: Font Awesome Free Brands (SIL OFL 1.1), hanya 2 glyph */
@font-face { font-family: "SPNBrands"; src: url(data:font/woff;base64,d09GRgABAAAAAAQAAAoAAAAABNwDBgEAAAAAAAAAAAAAAAAAAAAAAAAAAABPUy8yAAABZAAAAE4AAABgP9JP5GNtYXAAAAHAAAAAMgAAADwCcc/mZ2x5ZgAAAfwAAAHnAAAC2k+JYkVoZWFkAAAA9AAAADYAAAA2LZAksWhoZWEAAAEsAAAAHAAAACQDigF4aG10eAAAAbQAAAAKAAAACgNA//xsb2NhAAAB9AAAAAgAAAAIAJsBvm1heHAAAAFIAAAAGgAAACAACgCXbmFtZQAAA+QAAAAGAAAABgAGAABwb3N0AAAD7AAAABMAAAAg/94AGQABAAADBgEAGDU/y18PPPUACwIAAAAAAOK68HcAAAAA5vTyxP/8/7kBwAHAAAAACAACAAAAAAAAeJxjYGRgYDzzfyuQPPD/D4hkAIqgACYAm0sGA3icY2BkYGBgZpjKwMoAAoxgzMgABwAPuQCiAAB4nGNgYXzCOIGBlYGB0YcxjYGBwR1Kf2WQZGhhYGBiYGVmgAMBBJPBMTzYl+HAg+qPuYxn/m8F6j/D4A0UZgTJMToxbgBSCgwMAHtYDjoAAAGAAAABwP/8AAAAAHicY2BgYGJgYGAGYhEgyQimWRg0gDQbkGZkYHpQ/TH3/38GBggt38Y3FawKCADUTAsiAAAAAABRAJsBbXicfVBLjtNAEK0yGrdHiT1k3O1ORCbEduxmBIPSbsctsWLBkgOQICSOMisuMBxgzkBWbFCkuUjYsGTFEiVUBQKzomzpVb/+vPcKTgD2G7zGDYQQg4RHsADwXnhjB9qVzljhtdEDgtKV2vqq89Z2PreZzogORShyW5vaDF6uqMx6vV4tr5dUjtrlavcN40TiVCYxcrfbcvfiA9XuB8GGGuwzPPx75P2xw1sA3P/cf8YNOWwAqsIoK0kzNKxZ+853WtmaGSXJUcbMYmCdldoePOLmlb8Yz+18XubhyXD07MoNs+wsSc50NtRJ8iboPX08+24+tZeX6aDfN6axz690Fq5EJKQcapWKKIrSdITvej13IRGAp7YlT7dwDhX5egsfAdJQWJoHa5IHz/5UbSwNh5dM87B8KGrT+s41ROjW1GVdFmUhChFqikAhQgqR6cbRHYrSLdpFazifCPlgWfxObmreWLSc9zCFxjWHNw+foIdURne2dpoj5lN7xKlSiEpNj3jzYDgKItHvRQJfV6aT1WxWpd2TCoMAK+NTXsvOVIjzeCKLYnwqxOm4yNUkFkgl4onKj2whmQ2QfoHuvi7jl/vCjHejIBiRsoh6Y88SfwzM5GHFRRbkjC18/a8e0//MFRJ+Ae3zbSwAAAAAAAAGAAB4nGNgZgCD/7cZJBmwAAAuHwH3AA==) format("woff"); }
.brand-lg { display: flex; flex-wrap: wrap; gap: 0.4rem 1.4rem; font-size: 0.82rem; color: var(--muted); }
.brand-lg span { white-space: nowrap; }
.brand-lg b { font-family: "SPNBrands", sans-serif; font-weight: 400; font-size: 1.1rem; margin-right: 0.4rem; vertical-align: -0.12em; }

/* tabel */
[data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 16px; overflow: hidden; }

/* panel detail */
.st-key-inspector { background: var(--panel); border: 1px solid var(--line); border-radius: 16px; padding: 1.1rem 1.15rem; }
.st-key-inspector [data-testid="stVerticalBlock"] { gap: 0.55rem; }
.rec-name { font-size: 1.2rem; font-weight: 650; letter-spacing: -0.012em; line-height: 1.25; overflow-wrap: anywhere; }
.rec-meta { font-size: 0.78rem; color: var(--muted); margin-top: 0.25rem; font-variant-numeric: tabular-nums; }
.sec {
  font-size: 0.7rem; font-weight: 700; letter-spacing: 0.09em; text-transform: uppercase; color: var(--blue-hi);
  margin: 0.9rem 0 0.15rem; padding-bottom: 0.4rem; border-bottom: 1px solid var(--line);
}
.kv { display: grid; grid-template-columns: 6.3rem 1fr; gap: 0.75rem; padding: 0.38rem 0; font-size: 0.88rem; line-height: 1.4; }
.kv .k { color: var(--muted); }
.kv .v { color: var(--text); overflow-wrap: anywhere; }
.kv .sub { display: block; color: var(--muted); font-size: 0.8rem; font-variant-numeric: tabular-nums; }
.kv a, .row-link a { color: var(--blue-hi); text-decoration: none; }
.kv a:hover, .row-link a:hover { text-decoration: underline; }
.row-link { display: flex; justify-content: space-between; align-items: baseline; font-size: 0.88rem; margin-top: 0.2rem; }
.row-link .k { color: var(--muted); }
.dim { color: var(--faint); }
.tags { display: flex; flex-wrap: wrap; gap: 0.35rem; padding-top: 0.45rem; }
.tag { font-size: 0.76rem; padding: 0.15rem 0.55rem; border-radius: 999px; background: var(--blue-soft); color: var(--blue-hi); }
.dup { display: grid; grid-template-columns: 3.2rem 1fr auto; gap: 0.6rem; font-size: 0.8rem; color: var(--muted); padding: 0.25rem 0; }
.dup span:last-child { font-variant-numeric: tabular-nums; }
.st-key-inspector [data-testid="stCode"] pre { font-size: 0.85rem; padding: 0.45rem 0.7rem; }
.empty-title { font-size: 0.95rem; font-weight: 600; }

/* tombol "Show data" bawaan pada grafik disembunyikan: tampilan tabelnya menutupi toolbar sehingga tidak bisa kembali ke grafik */
[data-testid="stElementToolbarButton"]:has(button[aria-label="Show data"]) { display: none; }

/* donut */
.donut { display: flex; align-items: center; gap: 1.5rem; flex-wrap: wrap; padding: 0.5rem 0 0.3rem; }
.donut .ring { position: relative; width: 136px; height: 136px; border-radius: 50%; flex: none; }
.donut .ring::after { content: ""; position: absolute; inset: 24px; border-radius: 50%; background: var(--bg); }
.donut .ring span { position: absolute; inset: 0; z-index: 1; display: flex; flex-direction: column; align-items: center;
  justify-content: center; font-size: 1.1rem; font-weight: 650; font-variant-numeric: tabular-nums; line-height: 1.1; }
.donut .ring small { font-size: 0.68rem; font-weight: 500; color: var(--muted); }
.dlegend { flex: 1 1 11rem; min-width: 0; display: grid; gap: 0.5rem; }
.dl { display: grid; grid-template-columns: 0.7rem minmax(0, 1fr) auto; gap: 0.6rem; align-items: baseline; font-size: 0.86rem; }
.dl i { width: 0.65rem; height: 0.65rem; border-radius: 50%; align-self: center; }
.dl .dn { overflow-wrap: anywhere; }
.dl .dv { color: var(--muted); font-variant-numeric: tabular-nums; white-space: nowrap; }
.dl .dv b { color: var(--text); font-weight: 650; margin-left: 0.25rem; }

/* panel kesalahan */
.fb { border: 1px solid var(--line-2); border-left: 3px solid #D9694F; border-radius: 6px; background: var(--panel); padding: 0.9rem 1.1rem; }
.fb-title { font-weight: 650; margin-bottom: 0.2rem; }
.fb-text { color: var(--muted); font-size: 0.88rem; }

/* judul halaman */
.page-title { font-size: 1.45rem; font-weight: 650; letter-spacing: -0.02em; line-height: 1.2; margin: 0; }

/* kartu angka ringkasan */
.kpis { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 1rem; margin-bottom: 1.1rem; }
.kpi { background: var(--panel); border: 1px solid var(--line); border-top: 3px solid var(--blue); border-radius: 16px; padding: 0.95rem 1.1rem 1rem; min-width: 0; }
.kpi .l { font-size: 0.74rem; font-weight: 600; color: var(--muted); letter-spacing: 0.02em; }
.kpi .n { font-size: 1.6rem; font-weight: 650; letter-spacing: -0.02em; line-height: 1.2; margin-top: 0.15rem; font-variant-numeric: tabular-nums; }
.kpi .s { font-size: 0.78rem; color: var(--muted); margin-top: 0.1rem; }

/* gaya SPN: Poppins, judul tebal, kartu angka berwarna */
[data-testid="stApp"] { font-family: "Poppins", system-ui, sans-serif; }
.page-title { font-size: 1.7rem; font-weight: 800; letter-spacing: -0.02em; color: var(--navy); }
.section-title { font-weight: 700; color: var(--navy); }
.kpi { border-top-color: var(--pink); box-shadow: 0 2px 10px rgba(26, 20, 100, 0.05); }
.kpi .n { color: var(--navy); font-weight: 800; }
[data-testid="stSidebar"] { box-shadow: 2px 0 12px rgba(26, 20, 100, 0.04); }

@media (max-width: 1250px) { .kpis { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
@media (max-width: 900px) {
  [data-testid="stMainBlockContainer"] { padding: 4.2rem 1rem 2rem; }
  .count { font-size: 1.5rem; }
}
</style>
"""


# ============================================================ UTILITAS
def fmt_int(n) -> str:
    """12345 -> '12.345' (pemisah ribuan format Indonesia)."""
    return f"{int(n):,}".replace(",", ".")


def fmt_pct(x: float) -> str:
    return f"{x:.1f}".replace(".", ",")


def fmt_compact(n) -> str:
    """12345678 -> '12,3 jt', 45600 -> '45,6 rb'."""
    n = float(n)
    if n >= 1_000_000:
        return f"{n / 1_000_000:.1f}".replace(".", ",") + " jt"
    if n >= 1_000:
        return f"{n / 1_000:.1f}".replace(".", ",") + " rb"
    return fmt_int(n)


def fmt_date(ts) -> str:
    if ts is None or pd.isna(ts):
        return "-"
    return f"{ts.day} {MONTHS_ID[ts.month - 1]} {ts.year}"


def e(x) -> str:
    """Escape teks sebelum disisipkan ke HTML."""
    return html.escape(str(x), quote=True)


def has(x) -> bool:
    """True untuk string yang benar-benar terisi (nilai kosong datang sebagai NaN yang dianggap truthy)."""
    return isinstance(x, str) and x != ""


def S(key, default=None):
    return st.session_state.get(key, default)


def _defaults() -> dict:
    """Nilai awal setiap widget filter. Widget hanya benar-benar kembali kosong jika nilainya DIISI ulang
    (menghapus key dari session_state tidak cukup, karena browser masih menyimpan pilihan lama)."""
    d = {"f_q": "", "f_age": (AGE_MIN, AGE_MAX), "f_date": (DATE_MIN, DATE_MAX)}
    d.update({k: [] for k in MULTI})
    d.update({k: [] for k in MULTIVALUE})
    d.update({k: None for k in BINARY})
    for k in ("f_ig_min", "f_ig_max", "f_tt_min", "f_tt_max"):
        d[k] = None
    return d


def clear_keys(keys):
    defaults = _defaults()
    for k in keys:
        if k in defaults:
            st.session_state[k] = defaults[k]
        else:
            st.session_state.pop(k, None)


def reset_all():
    clear_keys(list(_defaults()))


def reset_everything():
    """Dipakai tombol 'Kembali ke dashboard': kosongkan seluruh state sesi."""
    reset_all()
    for k in list(st.session_state.keys()):
        if k not in _defaults():
            try:
                st.session_state.pop(k, None)
            except Exception:
                pass


# ---- fallback kesalahan
_fallback_count = [0]


def show_fallback(title: str, exc: Exception):
    _fallback_count[0] += 1
    n = _fallback_count[0]
    with st.container(key=f"fallback_{n}"):
        st.markdown(
            f'<div class="fb"><div class="fb-title">{e(title)}</div>'
            '<div class="fb-text">Data Anda tidak terpengaruh. Kembali ke dashboard untuk melanjutkan; '
            "semua filter akan diatur ulang.</div></div>",
            unsafe_allow_html=True,
        )
        st.button("Kembali ke dashboard", key=f"fb_back_{n}", on_click=reset_everything, type="primary")
        with st.expander("Detail teknis"):
            st.code("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)), language=None)


@contextmanager
def safe_section(title: str):
    """Bagian yang gagal tampil tidak menjatuhkan seluruh halaman.
    st.stop() dan st.rerun() adalah BaseException, jadi tetap lewat."""
    try:
        yield
    except Exception as exc:
        show_fallback(title, exc)


# ============================================================ DATA
def _clean_handle(s: pd.Series) -> pd.Series:
    s = s.fillna("").astype(str).str.strip().str.split("?").str[0]

    def last_segment(x: str) -> str:
        parts = [p for p in x.split("/") if p]
        return parts[-1] if "/" in x and parts else x

    s = s.apply(last_segment).str.lstrip("@").str.strip()
    return s.where(~s.str.lower().isin(["", "0", "-", "none", "nan", "n/a"]), "")


@st.cache_data(show_spinner="Memuat data…")
def load_data(path: str):
    raw = pd.read_csv(path, encoding="utf-8-sig")
    df = pd.DataFrame(index=raw.index)

    df["row_id"] = raw.index + 1
    df["period"] = raw["source_file"].str.replace(".csv", "", regex=False)
    df["registered_at"] = pd.to_datetime(raw["timestamp"], errors="coerce")
    df["_month"] = df["registered_at"].dt.to_period("M").dt.to_timestamp()
    df["name"] = raw["nama_lengkap"].fillna("").astype(str).str.strip()
    df["email"] = raw["email"].fillna("").astype(str).str.strip().str.lower()

    # --- telepon: normalisasi ke nomor nasional, buat tampilan + tautan wa.me
    digits = raw["nomor_whatsapp"].astype(str).str.replace(r"\D", "", regex=True)
    nsn = digits.str.replace(r"^(62|0)", "", regex=True)
    phone_ok = nsn.str.match(r"^8\d{8,11}$")
    df["phone"] = "0" + nsn
    df["wa_url"] = np.where(phone_ok, "https://wa.me/62" + nsn, None)
    df["_nsn"] = nsn

    df["age"] = raw["usia"]
    df["marital"] = raw["status_perkawinan"]
    df["pregnancy"] = raw["sedang_hamil"].map({True: "Hamil", False: "Tidak hamil"})
    df["kids"] = raw["jumlah_anak"]
    df["child_cat"] = raw["kategori_anak_terakhir"].fillna("Unknown")
    df["content"] = raw["kategori_content"].fillna("")
    df["income"] = raw["pendapatan_mom_perbulan"]
    df["spend"] = raw["pengeluaran_mom_perbulan"]
    df["phone_type"] = raw["type_hp"].map({"IOS (Iphone)": "iPhone", "Android (Selain Iphone)": "Android"})
    df["job"] = raw["pekerjaan"].str.replace(
        r"Fully Mom(my)?fluencer", "Full-time Momfluencer", regex=True
    )
    df["business"] = raw["memiliki_usaha"].map({"Ya": "Punya usaha", "Tidak": "Tanpa usaha"})
    df["school_type"] = raw["type_sekolah_anak"]
    df["hijab"] = raw["mengenakan_hijab"].map({"Berhijab": "Berhijab", "Non Hijab": "Tidak berhijab"})
    df["city"] = raw["Domisili"]

    df["ig_handle"] = _clean_handle(raw["nama_account_instagram"])
    df["tt_handle"] = _clean_handle(raw["nama_account_tiktok"])
    df["ig_url"] = np.where(df["ig_handle"] != "", "https://instagram.com/" + df["ig_handle"], None)
    df["tt_url"] = np.where(df["tt_handle"] != "", "https://www.tiktok.com/@" + df["tt_handle"], None)
    df["ig_followers"] = raw["jumlah_follower_instagram"].fillna(0).astype(int)
    df["tt_followers"] = raw["jumlah_follower_tiktok"].fillna(0).astype(int)
    df["tier_ig"] = raw["tier_instagram"].fillna("<1K")
    df["tier_tt"] = raw["tier_tiktok"].fillna("<1K")
    df["registrations"] = df.groupby("email")["email"].transform("size")

    # --- kategori konten (multi-nilai)
    cats = sorted({c.strip() for row in df["content"] for c in row.split(",") if c.strip()})
    content_dm = pd.DataFrame({c: df["content"].str.contains(re.escape(c)) for c in cats})

    # --- blob pencarian huruf kecil per cakupan
    def blob(*cols):
        out = df[cols[0]].astype("object").fillna("").astype(str)
        for c in cols[1:]:
            out = out + " | " + df[c].astype("object").fillna("").astype(str)
        return out.str.lower()

    df["_phone_variants"] = df["_nsn"] + " 0" + df["_nsn"] + " 62" + df["_nsn"]
    df["_b_all"] = blob(
        "name", "email", "_phone_variants", "ig_handle", "tt_handle", "city", "job", "marital", "hijab",
        "child_cat", "content", "business", "school_type", "phone_type", "income", "spend", "period",
        "pregnancy", "age", "kids", "tier_ig", "tier_tt",
    )
    df["_name_l"] = df["name"].str.lower()
    df["_ig_l"] = df["ig_handle"].str.lower()
    df["_tt_l"] = df["tt_handle"].str.lower()
    return df, content_dm, cats


st.markdown(CSS, unsafe_allow_html=True)

try:
    df, content_dm, CONTENT_CATS = load_data(CSV_PATH)
except FileNotFoundError:
    st.error(f"File `{CSV_PATH}` tidak ditemukan. Letakkan di folder yang sama dengan `app.py`, lalu muat ulang halaman.")
    st.stop()
except Exception as exc:
    show_fallback("Data tidak dapat dimuat", exc)
    st.stop()

N = len(df)
AGE_MIN, AGE_MAX = int(df["age"].min()), int(df["age"].max())
DATE_MIN, DATE_MAX = df["registered_at"].min().date(), df["registered_at"].max().date()
PERIOD_ORDER = list(dict.fromkeys(df["period"]))

# ============================================================ DEFINISI FILTER
# key -> (kolom, label, urutan opsi tetap atau None untuk "paling sering dulu")
MULTI = {
    "f_city": ("city", "Kota", None),
    "f_child": ("child_cat", "Anak terkecil", CHILD_ORDER),
    "f_tier_ig": ("tier_ig", "Tier Instagram", TIER_ORDER),
    "f_tier_tt": ("tier_tt", "Tier TikTok", TIER_ORDER),
    "f_job": ("job", "Pekerjaan", None),
    "f_marital": ("marital", "Status pernikahan", None),
    "f_period": ("period", "Periode pendaftaran", PERIOD_ORDER),
}
# key -> (kolom, label)   (kolom biner -> kontrol segmen satu klik)
BINARY = {
    "f_hijab": ("hijab", "Hijab"),
    "f_biz": ("business", "Usaha"),
    "f_preg": ("pregnancy", "Kehamilan"),
    "f_hp": ("phone_type", "Jenis HP"),
}
# key -> (matriks dummy, label)   (data cocok jika punya SALAH SATU nilai terpilih)
MULTIVALUE = {
    "f_content": (content_dm, "Kategori konten"),
}

OPTIONS = {
    key: ([o for o in order if o in set(df[col])] if order else df[col].value_counts().index.tolist())
    for key, (col, _, order) in MULTI.items()
}
BINARY_OPTIONS = {key: list(df[col].dropna().unique()) for key, (col, _) in BINARY.items()}
VALID = {
    **{k: set(v) for k, v in OPTIONS.items()},
    **{k: set(v) for k, v in BINARY_OPTIONS.items()},
    "f_content": set(CONTENT_CATS),
}

# Label opsi memuat jumlah data ("Jakarta · 245") dan berubah setiap filter lain berubah.
# Widget bisa menyimpan label lama itu sebagai nilai terpilih, sehingga perlu dipulihkan ke nilai aslinya.
_COUNT_TAIL = re.compile(r"\s*[·•]\s*[\d.,]+\s*$")


def _repair(value, valid):
    try:
        if value in valid:
            return value
    except TypeError:
        return None
    if isinstance(value, str):
        base = _COUNT_TAIL.sub("", value).strip()
        if base in valid:
            return base
    return None


def clean_multi(key: str):
    """Pilihan multi-nilai yang valid untuk `key`; label usang dipulihkan, nilai asing dibuang."""
    raw = S(key)
    if not raw:
        return []
    raw = [raw] if isinstance(raw, str) else list(raw)
    fixed = []
    for v in raw:
        r = _repair(v, VALID[key])
        if r is not None and r not in fixed:
            fixed.append(r)
    if fixed != raw:
        st.session_state[key] = fixed
    return fixed


def clean_single(key: str):
    val = S(key)
    if val is None:
        return None
    fixed = _repair(val, VALID[key])
    if fixed != val:
        if fixed is None:
            st.session_state.pop(key, None)
        else:
            st.session_state[key] = fixed
    return fixed


def _short(vals, n=2):
    vals = list(vals)
    text = ", ".join(str(v) for v in vals[:n])
    return text + (f" +{len(vals) - n}" if len(vals) > n else "")


def _tokens(q: str):
    q = q.lower().strip()
    q = re.sub(r"(?<=\d)[\s\-.](?=\d)", "", q)  # "0812 3456-789" -> "08123456789"
    toks = []
    for t in q.split():
        if "/" in t:  # tautan profil yang ditempel -> ambil handle-nya
            parts = [p for p in t.split("?")[0].split("/") if p]
            t = parts[-1] if parts else t
        t = t.lstrip("@+").strip(",;")
        if t:
            toks.append(t)
    return toks


def collect_active():
    """Semua filter yang sedang aktif sebagai {grup: {mask, chip, keys}}."""
    out = {}

    q = (S("f_q") or "").strip()
    toks = _tokens(q)
    if toks:
        m = np.ones(N, dtype=bool)
        for t in toks:
            m &= df[SEARCH_COLUMN].str.contains(t, regex=False).to_numpy()
        out["search"] = dict(mask=m, chip=f"Cari: {q}", keys=["f_q"])

    for key, (col, label, _) in MULTI.items():
        sel = clean_multi(key)
        if sel:
            out[key] = dict(mask=df[col].isin(sel).to_numpy(), chip=f"{label}: {_short(sel)}", keys=[key])

    for key, (col, label) in BINARY.items():
        val = clean_single(key)
        if val:
            out[key] = dict(mask=(df[col] == val).to_numpy(), chip=f"{label}: {val}", keys=[key])

    for key, (dm, label) in MULTIVALUE.items():
        sel = clean_multi(key)
        if sel:
            out[key] = dict(mask=dm[sel].any(axis=1).to_numpy(), chip=f"{label}: {_short(sel)}", keys=[key])

    age = S("f_age")
    if age and tuple(age) != (AGE_MIN, AGE_MAX):
        out["f_age"] = dict(mask=df["age"].between(*age).to_numpy(), chip=f"Usia {age[0]}-{age[1]}", keys=["f_age"])

    for key, col, label in [("f_ig", "ig_followers", "Pengikut IG"), ("f_tt", "tt_followers", "Pengikut TikTok")]:
        lo, hi = S(f"{key}_min"), S(f"{key}_max")
        if lo is not None or hi is not None:
            m = np.ones(N, dtype=bool)
            if lo is not None:
                m &= (df[col] >= lo).to_numpy()
            if hi is not None:
                m &= (df[col] <= hi).to_numpy()
            if lo is not None and hi is not None:
                rng = f"{fmt_int(lo)}-{fmt_int(hi)}"
            elif lo is not None:
                rng = f"min {fmt_int(lo)}"
            else:
                rng = f"maks {fmt_int(hi)}"
            out[key] = dict(mask=m, chip=f"{label} {rng}", keys=[f"{key}_min", f"{key}_max"])

    d = S("f_date")
    if d and len(d) == 2 and (d[0], d[1]) != (DATE_MIN, DATE_MAX):
        day = df["registered_at"].dt.normalize()
        m = ((day >= pd.Timestamp(d[0])) & (day <= pd.Timestamp(d[1]))).to_numpy()
        out["f_date"] = dict(mask=m, chip=f"Daftar {fmt_date(d[0])} s/d {fmt_date(d[1])}", keys=["f_date"])

    return out


try:
    ACTIVE = collect_active()
except Exception as exc:
    show_fallback("Filter tidak dapat diterapkan", exc)
    st.stop()


def others(group):
    """Mask semua filter aktif KECUALI `group` (untuk jumlah opsi yang selalu terbarui)."""
    m = np.ones(N, dtype=bool)
    for g, f in ACTIVE.items():
        if g != group:
            m &= f["mask"]
    return m


ALL_MASK = others(None)
n_match = int(ALL_MASK.sum())
view = df[ALL_MASK]


# ============================================================ SIDEBAR (filter)
def _count_label(cnt):
    return lambda v: f"{v} · {fmt_int(cnt.get(v, 0))}"


def multiselect_filter(key: str):
    col, label, _ = MULTI[key]
    cnt = df.loc[others(key), col].value_counts().to_dict()
    st.multiselect(label, OPTIONS[key], key=key, placeholder="Semua", format_func=_count_label(cnt))


def binary_filter(key: str):
    col, label = BINARY[key]
    cnt = df.loc[others(key), col].value_counts().to_dict()
    st.segmented_control(
        label, BINARY_OPTIONS[key], key=key, selection_mode="single", format_func=_count_label(cnt),
    )


def multivalue_filter(key: str, options):
    dm, label = MULTIVALUE[key]
    cnt = dm.loc[others(key)].sum().to_dict()
    st.multiselect(label, options, key=key, placeholder="Semua", format_func=_count_label(cnt))


def follower_range(key: str, label: str):
    st.markdown(f'<div class="field-label">{e(label)}</div>', unsafe_allow_html=True)
    a, b = st.columns(2)
    a.number_input(f"{label} min", min_value=0, value=None, step=1000, key=f"{key}_min",
                   placeholder="Min", label_visibility="collapsed")
    b.number_input(f"{label} max", min_value=0, value=None, step=1000, key=f"{key}_max",
                   placeholder="Maks", label_visibility="collapsed")


with st.sidebar:
    with safe_section("Panel filter tidak dapat ditampilkan"):
        st.markdown('<div class="sb-title">Filter</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="sb-count"><b>{fmt_int(n_match)}</b> dari {fmt_int(N)} data</div>',
            unsafe_allow_html=True,
        )

        # Satu kotak pencarian yang mencakup semua kolom data.
        with st.container(key="searchbox"):
            st.text_input(
                "Pencarian", key="f_q", type="search", live="400ms",
                placeholder="Cari di semua kolom", icon=":material/search:",
            )
        st.button("Atur ulang filter", on_click=reset_all, disabled=not ACTIVE, type="secondary",
                  width="stretch", key="reset_sidebar")

        with st.expander("Audiens", expanded=False):
            multiselect_filter("f_city")
            st.slider("Usia", AGE_MIN, AGE_MAX, (AGE_MIN, AGE_MAX), key="f_age")
            multiselect_filter("f_child")
            binary_filter("f_hijab")

        with st.expander("Jangkauan", expanded=False):
            multiselect_filter("f_tier_ig")
            multiselect_filter("f_tier_tt")
            follower_range("f_ig", "Pengikut Instagram")
            follower_range("f_tt", "Pengikut TikTok")

        with st.expander("Profil", expanded=False):
            multiselect_filter("f_job")
            multivalue_filter("f_content", CONTENT_CATS)
            binary_filter("f_biz")

        with st.expander("Lainnya", expanded=False):
            multiselect_filter("f_period")
            st.date_input("Tanggal daftar", value=(DATE_MIN, DATE_MAX), min_value=DATE_MIN,
                          max_value=DATE_MAX, key="f_date", format="DD/MM/YYYY")
            multiselect_filter("f_marital")
            binary_filter("f_preg")
            binary_filter("f_hp")


# ============================================================ EKSPOR: nama kolom dan berkas Excel
EXPORT_NAMES = {
    "row_id": "No. data", "period": "Periode", "registered_at": "Waktu daftar", "name": "Nama",
    "email": "Email", "phone": "WhatsApp", "age": "Usia", "marital": "Status pernikahan",
    "pregnancy": "Kehamilan", "kids": "Jumlah anak", "child_cat": "Anak terkecil",
    "content": "Kategori konten", "income": "Pendapatan per bulan", "spend": "Pengeluaran per bulan",
    "phone_type": "Jenis HP", "job": "Pekerjaan",
    "business": "Usaha", "school_type": "Jenis sekolah anak", "hijab": "Hijab", "city": "Kota",
    "ig_handle": "Instagram", "ig_url": "Link Instagram", "ig_followers": "Pengikut IG", "tier_ig": "Tier IG",
    "tt_handle": "TikTok", "tt_url": "Link TikTok", "tt_followers": "Pengikut TikTok", "tier_tt": "Tier TikTok",
}


def _dash_tables(raw: pd.DataFrame) -> list:
    """Agregasi sumber untuk grafik dashboard, dengan jenis visual yang sesuai dengan datanya."""
    def count_by(col, order=None, top=None):
        s = raw[col].astype("object").replace("", np.nan).dropna()
        vc = s.value_counts()
        if order:
            vc = vc.reindex([o for o in order if o in vc.index])
        elif top:
            vc = vc.head(top)
        return vc.rename_axis("Kelompok").reset_index(name="Jumlah data")

    def bins_by(col):
        values = pd.to_numeric(raw[col], errors="coerce").fillna(0).clip(lower=0)
        b = pd.cut(values, bins=[-1, 999, 2999, 9999, 29999, 99999, 999999, np.inf], labels=FOLLOWER_BINS)
        return b.value_counts().reindex(FOLLOWER_BINS, fill_value=0).rename_axis("Rentang pengikut").reset_index(name="Jumlah akun")

    def followers_by_city(top=10):
        t = raw.copy()
        t["city"] = t["city"].astype("object").replace("", np.nan)
        t = t.dropna(subset=["city"])
        for col in ("ig_followers", "tt_followers"):
            t[col] = pd.to_numeric(t[col], errors="coerce").fillna(0).clip(lower=0)
        g = t.groupby("city", dropna=True).agg(
            Instagram=("ig_followers", "sum"), TikTok=("tt_followers", "sum")
        )
        g["_total"] = g["Instagram"] + g["TikTok"]
        g = g.sort_values("_total", ascending=False).head(top).drop(columns="_total")
        return g.rename_axis("Kota").reset_index()

    return [
        ("Kota dengan pendaftar terbanyak", "bar", count_by("city", top=12)),
        ("Pekerjaan pendaftar terbanyak", "bar", count_by("job", top=12)),
        ("Komposisi tier Instagram", "doughnut", count_by("tier_ig", TIER_ORDER)),
        ("Komposisi tier TikTok", "doughnut", count_by("tier_tt", TIER_ORDER)),
        ("Kategori anak terkecil", "bar", count_by("child_cat", CHILD_ORDER)),
        ("Tren pendaftaran per periode", "line", count_by("period", PERIOD_ORDER)),
        ("Status hijab", "doughnut", count_by("hijab")),
        ("Distribusi pengikut Instagram", "column", bins_by("ig_followers")),
        ("Distribusi pengikut TikTok", "column", bins_by("tt_followers")),
        ("Perbandingan pengikut IG dan TikTok per kota", "clustered_column", followers_by_city(top=10)),
    ]


def _write_dashboard(book, ws, raw: pd.DataFrame, filters: list) -> None:
    """Dashboard Excel bergaya terang dengan KPI berbingkai, grafik mudah dipindai, dan tabel sumber."""
    # Palet terang: putih, abu-biru, dan aksen biru agar nyaman dibaca di layar maupun saat dicetak.
    BG = "#F4F7FB"
    WHITE = "#FFFFFF"
    PANEL = "#FFFFFF"
    PANEL_2 = "#EAF2FF"
    LINE = "#D8E1ED"
    LINE_2 = "#B7C8DD"
    TEXT = "#1E293B"
    MUTED = "#64748B"
    BLUE_HI = "#1D4ED8"
    TEAL = "#0F9D8A"

    ws.hide_gridlines(2)
    ws.set_tab_color(BLUE)
    ws.set_default_row(19)
    ws.set_zoom(85)

    # Kanvas dashboard. Tabel sumber grafik tetap terlihat di sisi kanan, di luar area cetak.
    base = book.add_format({
        "bg_color": BG, "font_name": "Aptos", "font_color": TEXT, "valign": "vcenter"
    })
    ws.set_column(0, 0, 2.2, base)
    ws.set_column(1, 20, 10.2, base)
    ws.set_column(21, 22, 2.2, base)
    ws.set_column(23, 23, 2.0, base)
    # Kolom sumber data sengaja tidak disembunyikan: pengguna bisa mengecek angka di balik grafik.
    ws.set_column(24, 24, 34, base)
    ws.set_column(25, 26, 18, base)

    title = book.add_format({
        "font_name": "Aptos Display", "bold": True, "font_size": 22,
        "font_color": "#17365D", "bg_color": BG, "valign": "vcenter",
        "bottom": 2, "bottom_color": BLUE,
    })
    subtitle = book.add_format({
        "font_name": "Aptos", "font_size": 9, "font_color": MUTED,
        "bg_color": BG, "valign": "vcenter", "text_wrap": True,
    })
    note_fmt = book.add_format({
        "font_name": "Aptos", "font_size": 8, "font_color": MUTED,
        "bg_color": BG, "valign": "vcenter", "text_wrap": True,
    })
    kpi_value = book.add_format({
        "font_name": "Aptos Display", "bold": True, "font_size": 21,
        "font_color": BLUE_HI, "bg_color": PANEL, "num_format": "#,##0",
        "align": "left", "valign": "vcenter", "left": 1, "right": 1,
        "top": 1, "bottom": 0, "border_color": LINE_2,
    })
    kpi_label = book.add_format({
        "font_name": "Aptos", "font_size": 9, "bold": True, "font_color": MUTED,
        "bg_color": PANEL_2, "valign": "vcenter", "text_wrap": True,
        "left": 1, "right": 1, "top": 0, "bottom": 1, "border_color": LINE_2,
    })
    section = book.add_format({
        "font_name": "Aptos", "bold": True, "font_size": 11,
        "font_color": "#17365D", "bg_color": "#E3EDFA", "bottom": 1,
        "top": 1, "left": 1, "right": 1,
        "top_color": LINE, "bottom_color": LINE_2, "left_color": LINE, "right_color": LINE,
        "valign": "vcenter",
    })
    support_title = book.add_format({
        "bold": True, "font_name": "Aptos", "font_size": 10,
        "font_color": "#17365D", "bg_color": PANEL_2, "valign": "vcenter",
        "border": 1, "border_color": LINE,
    })
    support_head = book.add_format({
        "bold": True, "font_name": "Aptos", "font_size": 9,
        "font_color": WHITE, "bg_color": BLUE, "border": 1,
        "border_color": LINE_2, "text_wrap": True, "valign": "vcenter",
    })
    support_cell = book.add_format({
        "font_name": "Aptos", "font_size": 9, "font_color": TEXT,
        "bg_color": WHITE, "bottom": 1, "bottom_color": LINE,
    })
    support_num = book.add_format({
        "font_name": "Aptos", "font_size": 9, "font_color": TEXT,
        "bg_color": WHITE, "bottom": 1, "bottom_color": LINE,
        "num_format": "#,##0",
    })

    # Header dibuat dua tingkat: judul utama di atas, lalu logo SPN berukuran besar
    # sebagai brand/subtitle di kiri. Informasi filter berada di kanan logo.
    ws.merge_range(0, 1, 0, 20, "SPN DATA EXPLORER  |  DASHBOARD", title)
    ws.set_row(0, 39)
    filter_text = "; ".join(filters) if filters else "Semua data"
    ws.merge_range(
        1, 5, 1, 20,
        f"Ringkasan hasil ekspor\nFilter aktif: {filter_text}",
        subtitle,
    )
    ws.set_row(1, 60)
    ws.set_row(2, 9)
    _insert_excel_logo(ws, "B2")

    if raw.empty:
        empty_fmt = book.add_format({
            "font_name": "Aptos", "font_size": 11, "font_color": TEXT,
            "bg_color": WHITE, "border": 1, "border_color": LINE,
            "text_wrap": True, "valign": "vcenter",
        })
        ws.merge_range(3, 1, 4, 20,
                       "Tidak ada data yang cocok dengan filter. Ubah filter di aplikasi lalu ekspor kembali.",
                       empty_fmt)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.print_area(0, 0, 5, 20)
        return

    ig = int(pd.to_numeric(raw["ig_followers"], errors="coerce").fillna(0).sum())
    tt = int(pd.to_numeric(raw["tt_followers"], errors="coerce").fillna(0).sum())
    kpis = [
        (1, len(raw), "JUMLAH DATA TERPILIH"),
        (8, ig, "TOTAL PENGIKUT INSTAGRAM"),
        (15, tt, "TOTAL PENGIKUT TIKTOK"),
    ]
    for col, val, label in kpis:
        ws.merge_range(3, col, 3, col + 5, val, kpi_value)
        ws.merge_range(4, col, 4, col + 5, label, kpi_label)
    ws.set_row(3, 36)
    ws.set_row(4, 25)
    ws.merge_range(
        5, 1, 5, 20,
        "Catatan: grafik dihitung dari data hasil filter. Total pengikut merupakan penjumlahan nilai yang tercatat; "
        "bukan jumlah akun unik.",
        note_fmt,
    )
    ws.set_row(5, 23)
    ws.merge_range(6, 1, 6, 20, "DISTRIBUSI DAN PERBANDINGAN", section)
    ws.set_row(6, 25)

    # Tabel sumber ditampilkan di sebelah kanan dashboard. Area ini tidak ikut area cetak utama.
    ws.merge_range(0, 24, 0, 26, "DATA SUMBER GRAFIK", title)
    ws.merge_range(1, 24, 1, 26, "Angka agregat yang digunakan oleh setiap plot", subtitle)
    support_table_title = book.add_format({
        "bold": True, "font_name": "Aptos", "font_size": 9,
        "font_color": "#17365D", "bg_color": "#E3EDFA", "valign": "vcenter",
        "border": 1, "border_color": LINE, "text_wrap": True,
    })
    support_head = book.add_format({
        "bold": True, "font_name": "Aptos", "font_size": 9,
        "font_color": WHITE, "bg_color": BLUE, "border": 1,
        "border_color": LINE_2, "text_wrap": True, "valign": "vcenter",
    })
    support_cell = book.add_format({
        "font_name": "Aptos", "font_size": 9, "font_color": TEXT,
        "bg_color": WHITE, "bottom": 1, "bottom_color": LINE,
        "left": 1, "left_color": LINE, "right": 1, "right_color": LINE,
    })
    support_num = book.add_format({
        "font_name": "Aptos", "font_size": 9, "font_color": TEXT,
        "bg_color": WHITE, "bottom": 1, "bottom_color": LINE,
        "left": 1, "left_color": LINE, "right": 1, "right_color": LINE,
        "num_format": "#,##0",
    })

    tables = _dash_tables(raw)
    r = 3
    spots = []
    for name, kind, table in tables:
        last_support_col = 24 + max(len(table.columns) - 1, 1)
        ws.merge_range(r, 24, r, min(last_support_col, 26), name, support_table_title)
        for j, col_name in enumerate(table.columns):
            ws.write_string(r + 1, 24 + j, str(col_name), support_head)
        for i, row in enumerate(table.itertuples(index=False), start=r + 2):
            for j, value in enumerate(row):
                col_idx = 24 + j
                if j == 0:
                    ws.write_string(i, col_idx, "" if pd.isna(value) else str(value), support_cell)
                elif pd.isna(value):
                    ws.write_blank(i, col_idx, None, support_num)
                else:
                    ws.write_number(i, col_idx, float(value), support_num)
        spots.append((name, kind, table, r + 2, r + 1 + len(table)))
        r += len(table) + 4

    chart_width, chart_height, row_step = 650, 300, 15
    chart_colors = [BLUE, TEAL, "#F59E0B", "#8B5CF6", "#EC6A8D", "#60A5FA", "#84CC16"]
    for k, (name, kind, table, first_data_row, last_data_row) in enumerate(spots):
        if table.empty:
            continue

        if kind == "doughnut":
            chart = book.add_chart({"type": "doughnut"})
            chart.add_series({
                "name": name,
                "categories": [ws.get_name(), first_data_row, 24, last_data_row, 24],
                "values": [ws.get_name(), first_data_row, 25, last_data_row, 25],
                "points": [{"fill": {"color": c}, "border": {"color": WHITE}} for c in chart_colors[:len(table)]],
                "data_labels": {
                    "category": True, "percentage": True, "leader_lines": True,
                    "position": "best_fit",
                    "font": {"name": "Aptos", "size": 9, "color": TEXT},
                },
            })
            chart.set_hole_size(58)
            chart.set_legend({"position": "bottom", "font": {"name": "Aptos", "size": 9, "color": MUTED}})
        elif kind == "line":
            chart = book.add_chart({"type": "line"})
            chart.add_series({
                "name": str(table.columns[-1]),
                "categories": [ws.get_name(), first_data_row, 24, last_data_row, 24],
                "values": [ws.get_name(), first_data_row, 25, last_data_row, 25],
                "line": {"color": BLUE, "width": 2.5},
                "marker": {"type": "circle", "size": 6, "border": {"color": BLUE}, "fill": {"color": WHITE}},
                "data_labels": {"value": True, "position": "above", "num_format": "#,##0",
                                "font": {"name": "Aptos", "size": 8, "color": TEXT}},
            })
            chart.set_legend({"none": True})
        elif kind == "clustered_column":
            chart = book.add_chart({"type": "column", "subtype": "clustered"})
            for j in range(1, len(table.columns)):
                chart.add_series({
                    "name": [ws.get_name(), first_data_row - 1, 24 + j],
                    "categories": [ws.get_name(), first_data_row, 24, last_data_row, 24],
                    "values": [ws.get_name(), first_data_row, 24 + j, last_data_row, 24 + j],
                    "fill": {"color": chart_colors[(j - 1) % len(chart_colors)]},
                    "border": {"none": True},
                    "gap": 70,
                })
            chart.set_legend({"position": "bottom", "font": {"name": "Aptos", "size": 9, "color": MUTED}})
        elif kind == "column":
            chart = book.add_chart({"type": "column"})
            chart.add_series({
                "name": str(table.columns[-1]),
                "categories": [ws.get_name(), first_data_row, 24, last_data_row, 24],
                "values": [ws.get_name(), first_data_row, 25, last_data_row, 25],
                "fill": {"color": BLUE}, "border": {"none": True}, "gap": 55,
                "data_labels": {"value": True, "position": "outside_end", "num_format": "#,##0",
                                "font": {"name": "Aptos", "size": 8, "color": TEXT}},
            })
            chart.set_legend({"none": True})
        else:  # horizontal bar cocok untuk kategori dengan label panjang
            chart = book.add_chart({"type": "bar"})
            chart.add_series({
                "name": str(table.columns[-1]),
                "categories": [ws.get_name(), first_data_row, 24, last_data_row, 24],
                "values": [ws.get_name(), first_data_row, 25, last_data_row, 25],
                "fill": {"color": BLUE}, "border": {"none": True}, "gap": 55,
                "data_labels": {"value": True, "position": "outside_end", "num_format": "#,##0",
                                "font": {"name": "Aptos", "size": 8, "color": TEXT}},
            })
            chart.set_legend({"none": True})
            chart.set_x_axis({
                "name": str(table.columns[-1]),
                "name_font": {"name": "Aptos", "size": 9, "color": MUTED},
                "num_format": "#,##0",
                "num_font": {"name": "Aptos", "size": 8, "color": MUTED},
                "line": {"color": LINE_2}, "major_tick_mark": "none",
                "major_gridlines": {"visible": True, "line": {"color": "#E8EEF5", "width": 0.75}},
            })
            chart.set_y_axis({
                "reverse": True, "num_font": {"name": "Aptos", "size": 8, "color": TEXT},
                "line": {"none": True}, "major_tick_mark": "none",
                "major_gridlines": {"visible": False},
            })

        # Axis styling untuk grafik non-horizontal.
        if kind in ("line", "column", "clustered_column"):
            chart.set_x_axis({
                "num_font": {"name": "Aptos", "size": 8, "color": MUTED},
                "line": {"color": LINE_2}, "major_tick_mark": "none",
                "major_gridlines": {"visible": False},
                "label_position": "low",
            })
            chart.set_y_axis({
                "num_format": "#,##0", "num_font": {"name": "Aptos", "size": 8, "color": MUTED},
                "line": {"none": True}, "major_tick_mark": "none",
                "major_gridlines": {"visible": True, "line": {"color": "#E8EEF5", "width": 0.75}},
                "min": 0,
            })
        chart.set_title({
            "name": name,
            "name_font": {"name": "Aptos Display", "size": 11, "bold": True, "color": "#17365D"},
        })
        chart.set_size({"width": chart_width, "height": chart_height})
        chart.set_chartarea({"border": {"color": LINE, "width": 1}, "fill": {"color": WHITE}})
        chart.set_plotarea({"border": {"color": "#E8EEF5", "width": 0.5}, "fill": {"color": WHITE}})
        chart.show_hidden_data()
        chart_row = 7 + (k // 2) * row_step
        chart_col = 1 if k % 2 == 0 else 11
        ws.insert_chart(chart_row, chart_col, chart, {"x_offset": 2, "y_offset": 5})

    last_chart_row = 7 + ((len(spots) - 1) // 2) * row_step + 18 if spots else 6
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(0.25, 0.25, 0.45, 0.45)
    ws.center_horizontally()
    ws.fit_to_pages(1, 0)  # Satu halaman lebar; tinggi bebas agar chart tidak diperkecil berlebihan.
    ws.print_area(0, 0, last_chart_row, 20)


def build_xlsx(frame: pd.DataFrame, raw=None, filters=None) -> bytes:
    """Workbook dengan Dashboard visual dan Data terfilter, tanpa sheet Review."""
    out = io.BytesIO()
    opts = {"strings_to_formulas": False, "strings_to_urls": False}
    with pd.ExcelWriter(out, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm",
                        engine_kwargs={"options": opts}) as xw:
        frame.to_excel(xw, sheet_name="Data", index=False)
        ws = xw.sheets["Data"]
        book = xw.book
        header = book.add_format({
            "bold": True, "font_name": "Aptos", "font_size": 10,
            "font_color": "#FFFFFF", "bg_color": BLUE,
            "bottom": 2, "bottom_color": "#1B5F7A",
            "valign": "vcenter", "text_wrap": True,
        })
        link_fmt = book.add_format({"font_color": "#0563C1", "underline": 1})
        number_fmt = book.add_format({"num_format": "#,##0", "font_color": "#334155"})
        date_fmt = book.add_format({"num_format": "dd mmm yyyy hh:mm", "font_color": "#334155"})

        ws.hide_gridlines(2)
        ws.set_tab_color(BLUE)
        ws.set_zoom(90)
        ws.set_row(0, 32, header)
        max_row = max(len(frame), 1)
        columns_for_table = [{"header": str(c)} for c in frame.columns]
        # Lebar dibatasi agar kolom teks panjang tidak membuat tabel melebar berlebihan.
        for i, name in enumerate(frame.columns):
            sample = frame.iloc[:300, i].astype(str).str.len().max() if len(frame) else 0
            width = min(max(len(str(name)), int(sample or 0)) + 2, 38)
            width = max(width, 12)
            fmt = None
            if name in ("No. data", "Usia", "Jumlah anak", "Pengikut IG", "Pengikut TikTok"):
                fmt = number_fmt
            elif name in ("Waktu daftar",):
                fmt = date_fmt
            ws.set_column(i, i, width, fmt)

        # Tabel Excel memberi filter dropdown, zebra rows, dan rentang data yang lebih mudah dikelola.
        ws.add_table(0, 0, max_row, len(frame.columns) - 1, {
            "name": "SPNExportData",
            "style": "Table Style Medium 2",
            "columns": columns_for_table,
            "autofilter": True,
        })
        ws.freeze_panes(1, 2)

        # Tautan profil hanya dibuat aktif pada kolom link yang memang dihasilkan aplikasi.
        for name in ("Link Instagram", "Link TikTok"):
            if name in frame.columns:
                col = frame.columns.get_loc(name)
                for row_num, value in enumerate(frame[name], start=1):
                    if isinstance(value, str) and value.startswith(("https://instagram.com/", "https://www.tiktok.com/@")):
                        ws.write_url(row_num, col, value, link_fmt, string=value)

        # Data bar halus membantu menemukan akun dengan follower tinggi tanpa mengganti nilainya.
        for name in ("Pengikut IG", "Pengikut TikTok"):
            if name in frame.columns and len(frame):
                col = frame.columns.get_loc(name)
                ws.conditional_format(1, col, len(frame), col, {
                    "type": "data_bar", "bar_color": "#93C5FD", "bar_solid": True,
                    "bar_border_color": "#60A5FA",
                })

        if raw is not None:
            dashboard_ws = xw.book.add_worksheet("Dashboard")
            _write_dashboard(xw.book, dashboard_ws, raw, filters or [])
            dashboard_ws.activate()
            dashboard_ws.set_first_sheet()
    return out.getvalue()

# ============================================================ EKSPOR: pilihan kolom
# Kolom yang tetap ikut pada ekspor "hanya kolom terfilter" (bisa diubah pengguna di menu Ekspor).
MANDATORY_DEFAULT = ["row_id", "name", "email", "phone", "ig_handle", "tt_handle"]
EXPORT_MODES = ["Semua kolom", "Hanya kolom terfilter + wajib"]

# grup filter aktif -> kolom data yang disaring oleh filter itu
FILTER_EXPORT = {
    **{k: [col] for k, (col, _, _) in MULTI.items()},
    **{k: [col] for k, (col, _) in BINARY.items()},
    "f_content": ["content"],
    "f_age": ["age"],
    "f_ig": ["ig_followers"],
    "f_tt": ["tt_followers"],
    "f_date": ["registered_at"],
}
# kolom yang dicari oleh kotak pencarian, menurut cakupannya
SEARCH_EXPORT = ["name", "email", "phone", "ig_handle", "tt_handle", "city"]


def filtered_columns() -> list:
    """Kolom yang sedang disaring filter aktif, dengan urutan kolom ekspor."""
    cols = set()
    for g in ACTIVE:
        if g == "search":
            cols.update(SEARCH_EXPORT)
        else:
            cols.update(FILTER_EXPORT.get(g, []))
    return [c for c in EXPORT_NAMES if c in cols]


def _export_frame(source: pd.DataFrame, filters: list, cols: list) -> bytes:
    return build_xlsx(source[cols].rename(columns=EXPORT_NAMES), source, filters)


def export_control():
    """Tombol Ekspor: pilih semua kolom atau hanya kolom terfilter + kolom wajib."""
    with st.popover("Ekspor ke Excel", icon=":material/download:", type="primary", width="stretch"):
        st.markdown('<div class="field-label">Kolom yang diekspor</div>', unsafe_allow_html=True)
        if ACTIVE:
            mode = st.radio(
                "Kolom yang diekspor", EXPORT_MODES, key="x_mode", label_visibility="collapsed",
                captions=["Semua kolom, untuk baris hasil filter.",
                          "Hanya kolom yang sedang difilter, ditambah kolom wajib."],
            )
        else:
            mode = EXPORT_MODES[0]
            st.caption("Belum ada filter aktif, jadi semua kolom diekspor. Setelah ada filter, "
                       "Anda bisa memilih hanya kolom yang difilter.")

        if mode == EXPORT_MODES[0]:
            cols = list(EXPORT_NAMES)
        else:
            mand = st.multiselect("Kolom wajib (selalu ikut)", list(EXPORT_NAMES), default=MANDATORY_DEFAULT,
                                  key="x_mand", format_func=EXPORT_NAMES.get)
            need = set(mand) | set(filtered_columns())
            cols = [c for c in EXPORT_NAMES if c in need] or list(EXPORT_NAMES)
        st.caption(f"{len(cols)} kolom: " + ", ".join(EXPORT_NAMES[c] for c in cols))

        all_cols = mode == EXPORT_MODES[0]
        # Berkas dibuat hanya saat tombol diklik (data berupa callable), jadi filter dan klik baris tetap instan.
        st.download_button(
            "Unduh .xlsx",
            functools.partial(_export_frame, view, [f["chip"] for f in ACTIVE.values()], cols) if HAS_XLSX else b"",
            file_name="spn_data_semua_kolom.xlsx" if all_cols else "spn_data_kolom_terfilter.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click="ignore", type="primary", width="stretch", disabled=not HAS_XLSX,
            icon=":material/download:",
            help=("Berisi Dashboard dengan beragam grafik dan tabel sumber di sisi kanan, serta Data hasil filter dengan dropdown filter dan format tabel."
                  if HAS_XLSX else "Pasang dulu dengan: pip install xlsxwriter"),
        )


# ============================================================ KOMPONEN HALAMAN
@functools.lru_cache(maxsize=1)
def _logo_b64() -> str:
    try:
        return base64.b64encode(LOGO_PATH.read_bytes()).decode()
    except OSError:  # logo belum ada: judul tampil tanpa logo
        return ""


def page_head(title: str, sub: str):
    b64 = _logo_b64()
    logo = f'<img class="page-logo" src="data:image/jpeg;base64,{b64}" alt="Logo SPN">' if b64 else ""
    st.markdown(f'<div class="page-head">{logo}<div><div class="page-title">{e(title)}</div>'
                f'<div class="subtle">{e(sub)}</div></div></div>', unsafe_allow_html=True)


def status_row(export: bool = False):
    """Jumlah data terpilih, tombol hapus filter, dan chip filter aktif. Sama di semua halaman."""
    cols = st.columns([5, 1.9, 1.3] if export else [7, 1.3], vertical_alignment="center")
    pct = n_match / N * 100
    cols[0].markdown(
        f'<div class="count">{fmt_int(n_match)}<small>dari {fmt_int(N)} data ({fmt_pct(pct)}%)</small></div>',
        unsafe_allow_html=True,
    )
    if export:
        with cols[1]:
            export_control()
    cols[-1].button("Hapus filter", on_click=reset_all, disabled=not ACTIVE, width="stretch",
                    key="reset_main", type="secondary")
    if ACTIVE:
        with st.container(horizontal=True, gap="small", key="chips"):
            for g, f in ACTIVE.items():
                st.button(f["chip"], key=f"chip_{g}", on_click=clear_keys, args=(f["keys"],),
                          icon=":material/close:", icon_position="right", help="Klik untuk menghapus filter ini")
    else:
        st.caption("Belum ada filter. Menampilkan semua data.")


@contextmanager
def card(title: str, note: str = None):
    """Kartu grafik: judul, isi, catatan kecil. Grafik yang gagal tidak menjatuhkan halaman."""
    with st.container(border=True):
        st.markdown(f'<div class="section-title">{e(title)}</div>', unsafe_allow_html=True)
        with safe_section(f"Grafik \u201c{title}\u201d tidak dapat ditampilkan"):
            yield
        if note:
            st.caption(note)


def show(chart, empty: str = "Belum ada data untuk grafik ini."):
    if chart is None:
        st.caption(empty)
    elif isinstance(chart, str):
        st.markdown(chart, unsafe_allow_html=True)
    else:
        st.altair_chart(chart, width="stretch")


def show_plotly(fig, height: int, empty: str = "Belum ada data untuk grafik ini."):
    """Tampilkan grafik Plotly dengan tema gelap aplikasi (interaktif: zoom, geser, sorot, putar 3D)."""
    if fig is None:
        st.caption(empty)
        return
    fig.update_layout(height=height)
    st.plotly_chart(fig, theme=None, width="stretch",
                    config={"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d"]})


def kpi_cards(items):
    """items: (label, nilai, keterangan, warna aksen)."""
    def big(v):
        return ' style="font-size:1.25rem;line-height:2.05rem"' if len(str(v)) > 9 else ""

    cells = "".join(
        f'<div class="kpi" style="border-top-color:{c}"><div class="l">{e(l)}</div>'
        f'<div class="n"{big(v)}>{e(v)}</div><div class="s">{e(s)}</div></div>'
        for l, v, s, c in items
    )
    st.markdown(f'<div class="kpis">{cells}</div>', unsafe_allow_html=True)


# ============================================================ TABEL: kolom dan panel detail
COL_LABEL = {
    "name": "Nama", "email": "Email", "wa_url": "WhatsApp", "ig_handle": "Instagram", "tt_handle": "TikTok",
    "ig_url": "Instagram", "tt_url": "TikTok", "ig_followers": "Pengikut IG", "tt_followers": "Pengikut TikTok",
    "tier_ig": "Tier IG", "tier_tt": "Tier TikTok", "age": "Usia", "city": "Kota", "marital": "Status",
    "pregnancy": "Kehamilan", "kids": "Anak", "child_cat": "Anak terkecil", "hijab": "Hijab",
    "income": "Pendapatan", "spend": "Pengeluaran", "job": "Pekerjaan",
    "business": "Usaha", "content": "Kategori konten", "phone_type": "Jenis HP", "period": "Periode",
    "registered_at": "Terdaftar",
}
PRESETS = {
    "Esensial": ["name", "ig_handle", "ig_followers", "tt_handle", "tt_followers", "city", "age"],
    "Profil": ["name", "age", "city", "marital", "child_cat", "hijab", "job", "business", "content"],
    "Kontak": ["name", "email", "wa_url", "ig_url", "tt_url", "city"],
    "Semua": ["name", "email", "wa_url", "age", "city", "marital", "pregnancy", "kids", "child_cat", "hijab",
              "income", "spend", "job", "business", "content", "phone_type",
              "ig_url", "ig_followers", "tier_ig", "tt_url", "tt_followers", "tier_tt",
              "period", "registered_at"],
}
COLUMN_CONFIG = {
    "name": st.column_config.TextColumn(COL_LABEL["name"], pinned=True),
    "wa_url": st.column_config.LinkColumn(COL_LABEL["wa_url"], display_text=r"https://wa\.me/(.*)"),
    "ig_url": st.column_config.LinkColumn(COL_LABEL["ig_url"], display_text=r"https://instagram\.com/(.*)"),
    "tt_url": st.column_config.LinkColumn(COL_LABEL["tt_url"], display_text=r"https://www\.tiktok\.com/@(.*)"),
    "ig_followers": st.column_config.NumberColumn(COL_LABEL["ig_followers"], format="localized"),
    "tt_followers": st.column_config.NumberColumn(COL_LABEL["tt_followers"], format="localized"),
    "age": st.column_config.NumberColumn(COL_LABEL["age"], format="%d"),
    "registered_at": st.column_config.DatetimeColumn(COL_LABEL["registered_at"], format="YYYY-MM-DD"),
    **{c: label for c, label in COL_LABEL.items()
       if c not in ("name", "wa_url", "ig_url", "tt_url", "ig_followers", "tt_followers", "age",
                    "registered_at")},
}


def render_record(r):
    name = r["name"] if has(r["name"]) else "(tanpa nama)"
    st.markdown(
        f'<div><div class="rec-name">{e(name)}</div>'
        f'<div class="rec-meta">No. {r["row_id"]} · {e(r["period"])} · daftar {fmt_date(r["registered_at"])}</div></div>',
        unsafe_allow_html=True,
    )

    # --- kontak (dengan tombol salin)
    wa = (f'<a href="{e(r["wa_url"])}" target="_blank" rel="noopener noreferrer">Buka chat</a>'
          if has(r["wa_url"]) else '<span class="warn">Nomor tidak valid</span>')
    st.markdown(f'<div><div class="sec">Kontak</div><div class="row-link"><span class="k">WhatsApp</span>{wa}</div></div>',
                unsafe_allow_html=True)
    st.code(r["phone"], language=None)
    st.markdown('<div class="row-link"><span class="k">Email</span></div>', unsafe_allow_html=True)
    st.code(r["email"], language=None)

    # --- bagian lain
    def row(label, value):
        shown = value if value not in ("", None) and not (isinstance(value, float) and pd.isna(value)) else None
        return (f'<div class="kv"><div class="k">{e(label)}</div>'
                f'<div class="v">{e(shown) if shown is not None else "<span class=dim>-</span>"}</div></div>')

    def social(label, handle, url, followers, tier):
        link = (f'<a href="{e(url)}" target="_blank" rel="noopener noreferrer">@{e(handle)}</a>'
                if has(url) else '<span class="dim">-</span>')
        return (f'<div class="kv"><div class="k">{label}</div><div class="v">{link}'
                f'<span class="sub">{fmt_int(followers)} pengikut · {e(tier)}</span></div></div>')

    age = "-" if pd.isna(r["age"]) else f"{int(r['age'])} tahun"
    st.markdown("".join([
        '<div><div class="sec">Media sosial</div>',
        social("Instagram", r["ig_handle"], r["ig_url"], r["ig_followers"], r["tier_ig"]),
        social("TikTok", r["tt_handle"], r["tt_url"], r["tt_followers"], r["tier_tt"]),
        "</div>",
    ]), unsafe_allow_html=True)
    parts = [
        '<div><div class="sec">Profil</div>',
        row("Kota", r["city"]), row("Usia", age), row("Status", r["marital"]),
        row("Hijab", r["hijab"]), row("Kehamilan", r["pregnancy"]), row("Jenis HP", r["phone_type"]),
        '<div class="sec">Keluarga</div>',
        row("Jumlah anak", r["kids"]), row("Anak terkecil", r["child_cat"]), row("Sekolah anak", r["school_type"]),
        '<div class="sec">Pekerjaan dan ekonomi</div>',
        row("Pekerjaan", r["job"]), row("Usaha", r["business"]),
        row("Pendapatan", r["income"]), row("Pengeluaran", r["spend"]),
    ]
    if r["content"]:
        tags = "".join(f'<span class="tag">{e(c.strip())}</span>' for c in r["content"].split(",") if c.strip())
        parts += ['<div class="sec">Kategori konten</div>', f'<div class="tags">{tags}</div>']

    same = df[((df["email"] == r["email"]) | (df["phone"] == r["phone"])) & (df["row_id"] != r["row_id"])]
    if len(same):
        parts.append(f'<div class="sec">Terdaftar juga di tempat lain ({len(same)})</div>')
        for _, o in same.sort_values("registered_at").iterrows():
            parts.append(f'<div class="dup"><span>No. {o["row_id"]}</span>'
                         f'<span>{e(o["period"])} · {e(o["name"])}</span><span>{fmt_date(o["registered_at"])}</span></div>')
    parts.append("</div>")
    st.markdown("".join(parts), unsafe_allow_html=True)


# ============================================================ GRAFIK: pembantu
# Setiap grafik memakai satu warna aksen dari palet ini (bukan satu biru untuk semuanya),
# dan batang diberi gradasi menurut nilai atau urutan supaya tetap sederhana tapi tidak monoton.
BG = "#FFFFFF"
TXT = "#1A1464"
PANEL = "#F6F8FE"
A = {
    "blue": "#4A7CD9", "teal": "#2BA89A", "amber": "#F2A33A", "violet": "#7B5CD6",
    "rose": "#E0457F", "green": "#3FA66B", "sky": "#4FA3E0",
}
CAT_PALETTE = [A["blue"], A["sky"], A["violet"], A["teal"], A["rose"], A["amber"], A["green"]]
PAL2 = [A["blue"], A["sky"]]
HEAT = ["#EAF0FC", "#B4C8F0", "#4A7CD9", "#1A1464"]
AGE_BINS = ["<20", "20-24", "25-29", "30-34", "35-39", "40-44", "45-49", "\u226550"]
PRESENCE_ORDER = ["Keduanya", "Hanya Instagram", "Hanya TikTok", "Tidak ada"]
LOG_TICKS = [10, 100, 1000, 10000, 100000, 1000000, 10000000]
LOG_TEXT = ["10", "100", "1K", "10K", "100K", "1M", "10M"]
MONTH_EXPR = ("['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des']"
              "[month(datum.value)] + ' ' + year(datum.value)")


def _mix(c1: str, c2: str, t: float) -> str:
    """Campur dua warna hex. t=0 -> c1, t=1 -> c2."""
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


def shade_value(data: pd.DataFrame, val: str, accent: str, floor: float = 0.55) -> pd.DataFrame:
    """Warna batang: nilai terbesar = aksen penuh, terkecil = aksen yang diredupkan."""
    v = data[val].astype(float)
    span = v.max() - v.min()
    t = (v - v.min()) / span if span > 0 else pd.Series(1.0, index=v.index)
    return data.assign(color=[_mix(accent, BG, floor * (1 - x)) for x in t])


def shade_order(data: pd.DataFrame, accent: str, floor: float = 0.55) -> pd.DataFrame:
    """Warna batang menurut urutan kategori: makin tinggi tingkat, makin terang."""
    n = len(data)
    return data.assign(color=[_mix(accent, BG, floor * (1 - (i / (n - 1) if n > 1 else 1))) for i in range(n)])


# ---- agregasi
def clean_cat(s: pd.Series) -> pd.Series:
    s = s.astype("object")
    return s[s.notna() & (s.astype(str).str.strip() != "")]


def counts(s: pd.Series, order=None, top=None) -> pd.DataFrame:
    """Jumlah per kategori -> kolom group, n, share. `order` = urutan tetap (nilai lain ditaruh di belakang)."""
    vc = clean_cat(s).value_counts()
    if order:
        keep = [o for o in order if o in vc.index]
        vc = vc.reindex(keep + [i for i in vc.index if i not in set(order)])
    elif top:
        vc = vc.head(top)
    out = vc.rename_axis("group").reset_index(name="n")
    total = int(out["n"].sum())
    out["share"] = out["n"] / total if total else 0.0
    out["group"] = out["group"].astype(str)
    return out


_NUM = re.compile(r"\d[\d.,]*")


def _to_num(txt: str) -> float:
    txt = re.sub(r"\.(?=\d{3}\b)", "", txt).replace(",", ".")
    try:
        return float(txt.rstrip("."))
    except ValueError:
        return 0.0


def bucket_order(values) -> list:
    """Urutkan label rentang ('1-5 jt', '6-10 jt', '>20 jt') menurut angka pertama di dalamnya."""
    def key(v):
        m = _NUM.search(str(v))
        if not m:
            return (1, 0.0, str(v))
        n = _to_num(m.group())
        t = str(v).strip().lower()
        if t.startswith(("<", "kurang", "di bawah")):
            n -= 0.5
        elif t.startswith((">", "\u2265", "lebih", "di atas", "diatas")):
            n += 0.5
        return (0, n, str(v))
    return sorted(values, key=key)


def age_groups(s: pd.Series) -> pd.Series:
    g = pd.cut(pd.to_numeric(s, errors="coerce"), bins=[-1, 19, 24, 29, 34, 39, 44, 49, np.inf], labels=AGE_BINS)
    return g.astype("object")


def kids_series(s: pd.Series):
    """Jumlah anak -> label '0'..'4+' bila kolomnya angka; kalau bukan angka, pakai teks aslinya."""
    num = pd.to_numeric(s, errors="coerce")
    if num.notna().sum() >= max(1, 0.5 * s.notna().sum()):
        out = num.dropna().astype(int).clip(upper=4).astype(str).replace({"4": "4+"})
        return out, ["0", "1", "2", "3", "4+"]
    return clean_cat(s).astype(str), None


def tip_counts(label: str):
    return [alt.Tooltip("group:N", title=label), alt.Tooltip("n:Q", title="Jumlah data", format=",.0f"),
            alt.Tooltip("share:Q", title="Porsi", format=".1%")]


# ---- jenis grafik
def _axis_fmt(data, val) -> str:
    return "~s" if float(data[val].max()) >= 10000 else ",.0f"


def hbar(data, *, cat, val, sort, tooltip, fmt=",.0f", step=30):
    """Batang horizontal: label kategori selalu terbaca di lebar layar berapa pun."""
    base = alt.Chart(data).encode(
        y=alt.Y(f"{cat}:N", sort=sort, title=None,
                axis=alt.Axis(ticks=False, domain=False, labelLimit=240, labelPadding=10)),
        x=alt.X(f"{val}:Q", title=None,
                axis=alt.Axis(ticks=False, domain=False, tickCount=4, format=_axis_fmt(data, val), grid=True)),
    )
    bars = base.mark_bar(cornerRadiusEnd=3, height=alt.RelativeBandSize(0.62)).encode(
        color=alt.Color("color:N", scale=None, legend=None), tooltip=tooltip)
    labels = base.mark_text(align="left", dx=6, fontSize=12, color=TXT).encode(text=alt.Text(f"{val}:Q", format=fmt), tooltip=tooltip)
    return (bars + labels).properties(height=alt.Step(step)).configure_view(stroke=None)


def vbar(data, *, cat, val, sort, tooltip, fmt=",.0f", height=250):
    """Batang vertikal untuk kategori pendek (kelompok usia, jumlah anak)."""
    base = alt.Chart(data).encode(
        x=alt.X(f"{cat}:N", sort=sort, title=None,
                axis=alt.Axis(ticks=False, domain=False, labelAngle=0, labelPadding=8, labelLimit=110)),
        y=alt.Y(f"{val}:Q", title=None,
                axis=alt.Axis(ticks=False, domain=False, tickCount=4, format=_axis_fmt(data, val), grid=True)),
    )
    bars = base.mark_bar(cornerRadiusEnd=3, width=alt.RelativeBandSize(0.7)).encode(
        color=alt.Color("color:N", scale=None, legend=None), tooltip=tooltip)
    labels = base.mark_text(dy=-9, fontSize=11, color=TXT).encode(text=alt.Text(f"{val}:Q", format=fmt), tooltip=tooltip)
    return (bars + labels).properties(height=height).configure_view(stroke=None)


def chart_counts(s, *, label, accent, order=None, top=None, shade="value", colors=None,
                 vertical=False, height=250, step=30):
    """Grafik jumlah per kategori. shade='value' (gradasi menurut nilai) atau 'order' (menurut urutan)."""
    d = counts(s, order=order, top=top)
    if d.empty:
        return None
    if colors:
        d = d.assign(color=[colors[i % len(colors)] for i in range(len(d))])
    elif shade == "order":
        d = shade_order(d, accent)
    else:
        d = shade_value(d, "n", accent)
    sort = d["group"].tolist()
    if vertical:
        return vbar(d, cat="group", val="n", sort=sort, tooltip=tip_counts(label), height=height)
    return hbar(d, cat="group", val="n", sort=sort, tooltip=tip_counts(label), step=step)


def chart_donut(s, *, label, order=None, colors=None):
    """Donut berbasis HTML/CSS: cincin di kiri, legenda lengkap (jumlah dan porsi) di kanan.
    Tidak bergantung pada ukuran kanvas grafik, jadi tidak ada bagian yang terpotong."""
    d = counts(s, order=order)
    if d.empty:
        return None
    pal = colors or CAT_PALETTE
    total = int(d["n"].sum())
    stops, acc, rows = [], 0.0, []
    for i, (g, n, sh) in enumerate(zip(d["group"], d["n"], d["share"])):
        c = pal[i % len(pal)]
        stops.append(f"{c} {acc * 100:.3f}% {(acc + sh) * 100:.3f}%")
        acc += sh
        tip = f"{label}: {g} \u00b7 {fmt_int(n)} data ({fmt_pct(sh * 100)}%)"
        rows.append(
            f'<div class="dl" title="{e(tip)}"><i style="background:{c}"></i><span class="dn">{e(g)}</span>'
            f'<span class="dv">{fmt_int(n)} <b>{sh * 100:.0f}%</b></span></div>'
        )
    ring = f'<div class="ring" style="background:conic-gradient({", ".join(stops)})"><span>{fmt_int(total)}<small>data</small></span></div>'
    return f'<div class="donut">{ring}<div class="dlegend">{"".join(rows)}</div></div>'


def chart_hist(s, *, label, accent, maxbins=16):
    num = pd.to_numeric(s, errors="coerce").dropna()
    if num.empty:
        return None
    d = pd.DataFrame({"v": num.to_numpy()})
    bins = alt.Bin(maxbins=maxbins)
    return alt.Chart(d).mark_bar(color=accent, cornerRadiusTopLeft=3, cornerRadiusTopRight=3, binSpacing=1).encode(
        x=alt.X("v:Q", bin=bins, title=label, axis=alt.Axis(grid=False, ticks=False, domain=False)),
        y=alt.Y("count():Q", title=None, axis=alt.Axis(ticks=False, domain=False, tickCount=4, format="~s")),
        tooltip=[alt.Tooltip("v:Q", bin=bins, title=label), alt.Tooltip("count():Q", title="Jumlah data")],
    ).properties(height=250).configure_view(stroke=None)


def chart_trend(frame):
    """Pendaftar per bulan (area bergradasi). Dipakai sebagai cadangan jika plotly tidak terpasang."""
    m = frame.dropna(subset=["_month"]).groupby("_month").size().rename("n")
    if m.empty:
        return None
    full = pd.date_range(m.index.min(), m.index.max(), freq="MS")
    m = m.reindex(full, fill_value=0).rename_axis("month").reset_index()
    m["label"] = [f"{MONTHS_ID[t.month - 1]} {t.year}" for t in m["month"]]
    grad = alt.Gradient(
        gradient="linear", x1=1, x2=1, y1=1, y2=0,
        stops=[alt.GradientStop(color="rgba(37,150,190,0.0)", offset=0),
               alt.GradientStop(color="rgba(37,150,190,0.55)", offset=1)],
    )
    base = alt.Chart(m).encode(
        x=alt.X("month:T", title=None, axis=alt.Axis(grid=False, ticks=False, domain=False, labelExpr=MONTH_EXPR,
                                                    tickCount={"interval": "month", "step": 3}, labelPadding=8)),
        y=alt.Y("n:Q", title=None, axis=alt.Axis(ticks=False, domain=False, tickCount=4, format="~s")),
    )
    area = base.mark_area(line={"color": A["blue"], "strokeWidth": 2.5}, color=grad, interpolate="monotone")
    dots = base.mark_circle(color=A["blue"], size=46).encode(
        tooltip=[alt.Tooltip("label:N", title="Bulan"), alt.Tooltip("n:Q", title="Pendaftar", format=",.0f")])
    return (area + dots).properties(height=260).configure_view(stroke=None)


def heatmap(long, *, x, y, val, x_sort, y_sort, x_title, y_title, step=34):
    """Peta panas kategori x kategori dengan angka di dalam sel."""
    mx = float(long[val].max()) if len(long) else 0.0
    long = long.assign(_t=[f"{int(v):,}".replace(",", ".") if v > 0 else "" for v in long[val]])
    base = alt.Chart(long).encode(
        x=alt.X(f"{x}:N", sort=x_sort, title=None,
                axis=alt.Axis(ticks=False, domain=False, labelAngle=0, labelLimit=120, orient="top", labelPadding=6)),
        y=alt.Y(f"{y}:N", sort=y_sort, title=None,
                axis=alt.Axis(ticks=False, domain=False, labelLimit=190, labelPadding=8)),
    )
    tips = [alt.Tooltip(f"{y}:N", title=y_title), alt.Tooltip(f"{x}:N", title=x_title),
            alt.Tooltip(f"{val}:Q", title="Jumlah data", format=",.0f")]
    rect = base.mark_rect(cornerRadius=3, stroke=PANEL, strokeWidth=2).encode(
        color=alt.Color(f"{val}:Q", scale=alt.Scale(range=HEAT), legend=None), tooltip=tips)
    txt = base.mark_text(fontSize=11).encode(
        text="_t:N", tooltip=tips,
        color=alt.condition(f"datum['{val}'] > {mx * 0.45}", alt.value("#FFFFFF"), alt.value("#1A1464")))
    return (rect + txt).properties(height=alt.Step(step)).configure_view(stroke=None)


def chart_cross(frame, xcol, ycol, *, x_order, y_order, x_title, y_title, step=34):
    """Peta panas dari dua kolom kategori. x_order/y_order None -> urut menurut jumlah."""
    d = pd.DataFrame({"x": frame[xcol], "y": frame[ycol]}).dropna()
    d = d[(d["x"].astype(str).str.strip() != "") & (d["y"].astype(str).str.strip() != "")]
    if d.empty:
        return None
    xs = x_order(d["x"].unique()) if callable(x_order) else [o for o in (x_order or []) if o in set(d["x"])] or d["x"].value_counts().index.tolist()
    ys = y_order(d["y"].unique()) if callable(y_order) else [o for o in (y_order or []) if o in set(d["y"])] or d["y"].value_counts().index.tolist()
    ct = d.groupby(["x", "y"]).size().rename("n")
    ct = ct.reindex(pd.MultiIndex.from_product([xs, ys], names=["x", "y"]), fill_value=0).reset_index()
    return heatmap(ct, x="x", y="y", val="n", x_sort=xs, y_sort=ys, x_title=x_title, y_title=y_title, step=step)


def chart_median_pair(frame, col, *, order=None, top=10):
    """Median pengikut IG dan TikTok per kelompok (batang berpasangan). Kelompok < 5 data disembunyikan."""
    g = frame.groupby(col, observed=True).agg(
        n=("row_id", "size"),
        ig=("ig_followers", lambda s: s[s > 0].median()),
        tt=("tt_followers", lambda s: s[s > 0].median()),
    )
    g = g[g["n"] >= 5]
    idx = [o for o in order if o in g.index] if order else g.sort_values("n", ascending=False).head(top).index.tolist()
    g = g.loc[idx]
    long = g.reset_index().melt(id_vars=[col, "n"], value_vars=["ig", "tt"], var_name="platform", value_name="median")
    long = long.dropna(subset=["median"])
    if long.empty:
        return None
    long["platform"] = long["platform"].map({"ig": "Instagram", "tt": "TikTok"})
    long = long.rename(columns={col: "group"})
    long["group"] = long["group"].astype(str)
    dom = ["Instagram", "TikTok"]
    base = alt.Chart(long).encode(
        y=alt.Y("group:N", sort=[str(i) for i in idx], title=None,
                axis=alt.Axis(ticks=False, domain=False, labelLimit=240, labelPadding=10)),
        yOffset=alt.YOffset("platform:N", sort=dom),
        x=alt.X("median:Q", title=None, axis=alt.Axis(ticks=False, domain=False, tickCount=4, format="~s")),
        color=alt.Color("platform:N", scale=alt.Scale(domain=dom, range=[A["rose"], A["sky"]]),
                        legend=alt.Legend(title=None, orient="top", symbolType="circle", labelFontSize=12)),
    )
    tips = [alt.Tooltip("group:N", title="Kelompok"), alt.Tooltip("platform:N", title="Platform"),
            alt.Tooltip("median:Q", title="Median pengikut", format=",.0f"), alt.Tooltip("n:Q", title="Jumlah data", format=",.0f")]
    bars = base.mark_bar(cornerRadiusEnd=3).encode(tooltip=tips)
    labels = base.mark_text(align="left", dx=5, fontSize=11, color=TXT).encode(text=alt.Text("median:Q", format=",.0f"), tooltip=tips)
    return (bars + labels).properties(height=alt.Step(38)).configure_view(stroke=None)


def chart_total_followers(frame, col, *, top=10):
    """Total pengikut (IG + TikTok) per kelompok, urut dari terbesar."""
    t = frame.assign(_t=frame["ig_followers"] + frame["tt_followers"])
    t = t[t[col].astype("object").notna() & (t[col].astype(str).str.strip() != "")]
    g = t.groupby(col, observed=True).agg(tot=("_t", "sum"), n=("row_id", "size"))
    g = g[g["tot"] > 0].sort_values("tot", ascending=False).head(top)
    if g.empty:
        return None
    d = g.reset_index().rename(columns={col: "group"})
    d["group"] = d["group"].astype(str)
    d = shade_value(d, "tot", A["rose"])
    return hbar(d, cat="group", val="tot", sort=d["group"].tolist(), tooltip=[
        alt.Tooltip("group:N", title="Kelompok"), alt.Tooltip("tot:Q", title="Total pengikut", format=",.0f"),
        alt.Tooltip("n:Q", title="Jumlah data", format=",.0f")], step=28)


def _scatter_groups(d: pd.DataFrame, color_col, color_order):
    """Beri kolom `_c` (kelompok warna) pada data scatter. Mengembalikan (data, kategori, palet)."""
    d = d.copy()
    if not color_col:
        d["_c"] = "Akun"
        return d, ["Akun"], [A["blue"]]
    d["_c"] = d[color_col].astype("object")
    cats = color_order or d["_c"].dropna().value_counts().index.tolist()
    cats = [c for c in cats if c in set(d["_c"])][:6]
    d["_c"] = d["_c"].where(d["_c"].isin(cats), "Lainnya").fillna("Lainnya")
    if (d["_c"] == "Lainnya").any():
        cats = cats + ["Lainnya"]
    pal = [CAT_PALETTE[i % len(CAT_PALETTE)] for i in range(len(cats))]
    if "Lainnya" in cats:
        pal[-1] = "#69759F"
    return d, cats, pal


def chart_scatter(frame, color_col, color_order, color_label):
    """IG vs TikTok (skala log), versi Altair. Dipakai jika plotly tidak terpasang.
    Mengembalikan (grafik, jumlah titik ditampilkan, total titik)."""
    d = frame[(frame["ig_followers"] > 0) & (frame["tt_followers"] > 0)]
    total = len(d)
    if total == 0:
        return None, 0, 0
    if total > 3000:
        d = d.sample(3000, random_state=0)
    d = d[["ig_followers", "tt_followers", "ig_handle", "tt_handle", "city"] + ([color_col] if color_col else [])].copy()
    enc = dict(
        x=alt.X("ig_followers:Q", title="Pengikut Instagram", scale=alt.Scale(type="log"),
                axis=alt.Axis(grid=True, ticks=False, domain=False, format="~s", values=LOG_TICKS)),
        y=alt.Y("tt_followers:Q", title="Pengikut TikTok", scale=alt.Scale(type="log"),
                axis=alt.Axis(grid=True, ticks=False, domain=False, format="~s", values=LOG_TICKS)),
    )
    tips = [alt.Tooltip("ig_handle:N", title="Instagram"), alt.Tooltip("ig_followers:Q", title="Pengikut IG", format=",.0f"),
            alt.Tooltip("tt_handle:N", title="TikTok"), alt.Tooltip("tt_followers:Q", title="Pengikut TikTok", format=",.0f"),
            alt.Tooltip("city:N", title="Kota")]
    if color_col:
        d, cats, pal = _scatter_groups(d, color_col, color_order)
        enc["color"] = alt.Color("_c:N", scale=alt.Scale(domain=cats, range=pal),
                                 legend=alt.Legend(title=color_label, orient="top", symbolType="circle", labelFontSize=12))
        tips.append(alt.Tooltip("_c:N", title=color_label))
        mark = dict(size=42, opacity=0.6)
    else:
        mark = dict(size=42, opacity=0.5, color=A["blue"])
    ch = alt.Chart(d).mark_circle(**mark).encode(tooltip=tips, **enc).properties(height=380).configure_view(stroke=None)
    return ch, len(d), total


# ---- grafik interaktif (Plotly)
def _plotly_base(**extra) -> dict:
    return dict(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#1A1464", size=12),
        margin=dict(l=0, r=0, t=36, b=0),
        hoverlabel=dict(bgcolor="#FFFFFF", font_color="#1A1464", bordercolor="#BFCCEC"),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0),
        **extra,
    )


def _log_axis(title: str) -> dict:
    return dict(type="log", title=title, tickvals=LOG_TICKS, ticktext=LOG_TEXT, gridcolor="#DCE4F5",
                zeroline=False, showline=False)


def chart_trend_plotly(frame):
    """Pendaftar per bulan: bisa di-zoom, digeser, dan dipersempit lewat tombol rentang atau slider di bawah."""
    m = frame.dropna(subset=["_month"]).groupby("_month").size().rename("n")
    if m.empty:
        return None
    full = pd.date_range(m.index.min(), m.index.max(), freq="MS")
    m = m.reindex(full, fill_value=0).rename_axis("month").reset_index()
    m["label"] = [f"{MONTHS_ID[t.month - 1]} {t.year}" for t in m["month"]]
    fig = go.Figure(go.Scatter(
        x=m["month"], y=m["n"], mode="lines+markers", customdata=m["label"],
        line=dict(color=BLUE, width=3, shape="spline"), marker=dict(size=7, color=BLUE),
        fill="tozeroy", fillcolor="rgba(37,150,190,0.22)",
        hovertemplate="<b>%{customdata}</b><br>%{y:,.0f} pendaftar<extra></extra>",
    ))
    fig.update_layout(**_plotly_base(showlegend=False, hovermode="x unified"))
    fig.update_xaxes(
        showgrid=False, tickformat="%b %Y",
        rangeslider=dict(visible=True, thickness=0.09, bgcolor="#F6F8FE", bordercolor="#DCE4F5", borderwidth=1),
        rangeselector=dict(
            bgcolor="#FFFFFF", activecolor=BLUE, font=dict(color="#1A1464"), bordercolor="#BFCCEC", borderwidth=1,
            buttons=[dict(count=6, label="6 bln", step="month", stepmode="backward"),
                     dict(count=1, label="1 thn", step="year", stepmode="backward"),
                     dict(step="all", label="Semua")],
        ),
    )
    fig.update_yaxes(gridcolor="#DCE4F5", zeroline=False, rangemode="tozero")
    return fig


# ---- scatter pengikut IG vs TikTok (Plotly)
EXT_HI = "#1E8A4F"   # hijau: pengikut tertinggi
EXT_LO = "#D6336C"   # merah muda: pengikut terendah
BRAND_FONT = "SPNBrands, sans-serif"  # font logo (didefinisikan di CSS)


def _scatter_custom(g: pd.DataFrame):
    """Kolom pendamping untuk hover: handle IG/TikTok, kota, usia, total pengikut."""
    age = g["age"].apply(lambda x: "-" if pd.isna(x) else str(int(x)))
    return np.column_stack([
        g["ig_handle"].replace("", "-").astype(str),
        g["tt_handle"].replace("", "-").astype(str),
        g["city"].fillna("-").astype(str),
        age,
        g["_tot"].apply(lambda x: fmt_int(x)),
    ])


def chart_scatter_plotly(frame, color_col, color_order, color_label, three_d=False,
                         zcol="age", zlabel="Usia", n_ext=5):
    """Pengikut IG (sumbu X) vs TikTok (sumbu Y), skala log.
    - Data ekstrem (n_ext tertinggi dan terendah di IG, TikTok, dan total) SELALU tampil, diberi
      penanda besar berwarna; yang tertinggi diberi label nama akun.
    - Titik biasa disampel bila > 3000, tapi ekstrem tidak pernah ikut terbuang.
    - Garis putus-putus = median, supaya terlihat di mana kerumunan berada.
    - Mode 3D hanya menambah sumbu ketiga (usia / jumlah anak); sumbu X dan Y tetap pengikut.
    Mengembalikan (figure, jumlah titik ditampilkan, total titik)."""
    d = frame[(frame["ig_followers"] > 0) & (frame["tt_followers"] > 0)].copy()
    if three_d:
        d["_z"] = pd.to_numeric(d[zcol], errors="coerce")
        d = d.dropna(subset=["_z"])
    total = len(d)
    if total == 0:
        return None, 0, 0
    d["_tot"] = d["ig_followers"] + d["tt_followers"]

    # tiap titik ekstrem hanya masuk SATU kelompok, jadi alasannya selalu jelas
    ext_groups, ext_idx = [], set()
    for how in ("Tertinggi", "Terendah"):
        for col, plat, sym, hcol in (("ig_followers", "IG", "\uf16d", "ig_handle"),
                                     ("tt_followers", "TikTok", "\ue07b", "tt_handle")):
            ranked = d[col].nlargest(n_ext) if how == "Tertinggi" else d[col].nsmallest(n_ext)
            fresh = [i for i in ranked.index if i not in ext_idx]
            ext_idx.update(fresh)
            if fresh:
                ext_groups.append((f"{plat} {how.lower()}", how, sym, hcol, d.loc[fresh]))

    rest = d.drop(index=list(ext_idx))
    if len(rest) > 3000:
        rest = rest.sample(3000, random_state=0)
    rest, cats, pal = _scatter_groups(rest, color_col, color_order)

    hover_core = ("IG <b>%{customdata[0]}</b>: %{x:,.0f}<br>TikTok <b>%{customdata[1]}</b>: %{y:,.0f}<br>"
                  "Total: %{customdata[4]}<br>Usia: %{customdata[3]} · %{customdata[2]}")
    fig = go.Figure()

    # --- titik biasa (dibuat pudar agar ekstrem menonjol)
    for c, color in zip(cats, pal):
        g = rest[rest["_c"] == c]
        if g.empty:
            continue
        common = dict(name=str(c), x=g["ig_followers"], y=g["tt_followers"], customdata=_scatter_custom(g),
                      hovertemplate=hover_core + ("<extra>%{fullData.name}</extra>" if color_col else "<extra></extra>"),
                      showlegend=bool(color_col), mode="markers")
        if three_d:
            fig.add_trace(go.Scatter3d(z=g["_z"], marker=dict(size=3, color=color, opacity=0.5), **common))
        else:
            fig.add_trace(go.Scatter(marker=dict(size=7, color=color, opacity=0.45, line=dict(width=0)), **common))

    # --- ekstrem: logo platform (Instagram / TikTok), warna = tertinggi (hijau) atau terendah (merah muda)
    for name, how, glyph, hcol, g in ext_groups:
        color = EXT_HI if how == "Tertinggi" else EXT_LO
        common = dict(name=name, x=g["ig_followers"], y=g["tt_followers"], customdata=_scatter_custom(g),
                      hovertemplate=hover_core + f"<br><b>Sorotan: {name}</b><extra></extra>",
                      mode="markers+text", text=[glyph] * len(g), textposition="middle center",
                      textfont=dict(size=16 if three_d else 20, color=color, family=BRAND_FONT),
                      showlegend=False)
        hit = dict(size=24, color="rgba(0,0,0,0)", line=dict(width=0))  # area sentuh tak terlihat untuk hover
        if three_d:
            fig.add_trace(go.Scatter3d(z=g["_z"], marker=hit, **common))
        else:
            fig.add_trace(go.Scatter(marker=hit, **common))

    med_ig, med_tt = float(d["ig_followers"].median()), float(d["tt_followers"].median())
    if three_d:
        axis3d = dict(showbackground=False, gridcolor="#DCE4F5", zeroline=False)
        fig.update_layout(**_plotly_base(scene=dict(
            xaxis={**_log_axis("Pengikut Instagram"), **axis3d},
            yaxis={**_log_axis("Pengikut TikTok"), **axis3d},
            zaxis={**axis3d, "title": f"{zlabel} (sumbu ke-3)"},
            camera=dict(eye=dict(x=1.7, y=-1.7, z=0.8)),
        )))
    else:
        fig.update_layout(**_plotly_base(dragmode="zoom"))
        fig.update_xaxes(**_log_axis("Pengikut Instagram"))
        fig.update_yaxes(**_log_axis("Pengikut TikTok"))
        for kw, txt in ((dict(x=med_ig), f"median IG {fmt_compact(med_ig)}"),
                        (dict(y=med_tt), f"median TikTok {fmt_compact(med_tt)}")):
            (fig.add_vline if "x" in kw else fig.add_hline)(
                line=dict(color="#69759F", width=1, dash="dot"), annotation_text=txt,
                annotation_font=dict(size=10, color="#55608A"), **kw)
    return fig, len(rest) + len(ext_idx), total


# ============================================================ HALAMAN 1: DATA
def page_data():
    page_head("Data", "Cari, saring, dan ekspor data pendaftar. Pilih satu baris untuk melihat detailnya.")
    status_row(export=True)

    toks = _tokens(S("f_q") or "")
    v = view.copy()
    if toks:
        score = np.zeros(len(v))
        for t in toks:
            score += 3 * ((v["_ig_l"] == t) | (v["_tt_l"] == t)).to_numpy()
            score += 2 * v["_name_l"].str.startswith(t).to_numpy()
            score += 1 * v["_name_l"].str.contains(r"\b" + re.escape(t), regex=True).to_numpy()
        v = v.assign(_score=score).sort_values(["_score", "registered_at"], ascending=[False, False])
        sort_note = "Diurutkan dari yang paling cocok, lalu pendaftar terbaru. Klik judul kolom untuk mengurutkan ulang."
    else:
        v = v.sort_values("registered_at", ascending=False)
        sort_note = "Diurutkan dari pendaftar terbaru. Klik judul kolom untuk mengurutkan ulang."
    v = v.reset_index(drop=True)

    left, right = st.columns([3.1, 1.7], gap="large")
    with left:
        with safe_section("Tabel tidak dapat ditampilkan"):
            preset = st.segmented_control("Kolom", list(PRESETS), default="Esensial", key="v_cols",
                                          required=True, label_visibility="collapsed")
            if preset not in PRESETS:
                preset = "Esensial"
            cols = PRESETS[preset]
            st.caption("Kolom: " + (", ".join(dict.fromkeys(COL_LABEL[c] for c in cols))
                                    if preset != "Semua" else "semua kolom yang tersedia"))
            event = st.dataframe(
                v[cols], hide_index=True, height=520, key="tbl", on_select="rerun",
                selection_mode="single-row", column_config=COLUMN_CONFIG,
            )
            st.caption(sort_note)

    with right:
        with safe_section("Panel detail tidak dapat ditampilkan"):
            with st.container(key="inspector", height=596):
                picked = event.selection.rows if event and event.selection else []
                if picked and picked[0] < len(v):
                    render_record(v.iloc[picked[0]])
                else:
                    st.markdown('<div class="empty-title">Detail data</div>', unsafe_allow_html=True)
                    st.caption("Pilih satu baris pada tabel lewat kotak di sisi kirinya.")


# ============================================================ HALAMAN 2: RINGKASAN
DIMS = {
    "Kota (12 teratas)": ("city", None),
    "Tier Instagram": ("tier_ig", TIER_ORDER),
    "Tier TikTok": ("tier_tt", TIER_ORDER),
    "Pekerjaan": ("job", None),
    "Anak terkecil": ("child_cat", CHILD_ORDER),
    "Hijab": ("hijab", None),
    "Status pernikahan": ("marital", None),
    "Usaha": ("business", None),
    "Jenis HP": ("phone_type", None),
    "Periode": ("period", PERIOD_ORDER),
}
METRICS = {"Jumlah data": "n", "Total pengikut": "tot", "Median IG": "ig", "Median TikTok": "tt"}
METRIC_TITLE = {"n": "Jumlah data", "tot": "Total pengikut", "ig": "Median pengikut IG", "tt": "Median pengikut TikTok"}
METRIC_COLOR = {"n": A["blue"], "tot": A["amber"], "ig": A["rose"], "tt": A["sky"]}


def _med(s: pd.Series, positive: bool = False):
    s = pd.to_numeric(s, errors="coerce").dropna()
    if positive:
        s = s[s > 0]
    return s.median() if len(s) else np.nan


def _num_txt(x, suffix=""):
    return "-" if pd.isna(x) else f"{fmt_int(round(float(x)))}{suffix}"


def page_overview():
    page_head("Ringkasan", "Gambaran umum data terpilih dan perbandingan antar kelompok.")
    status_row()

    with safe_section("Angka ringkasan tidak dapat ditampilkan"):
        n_v = len(view)

        def share(k):
            return f"{fmt_pct(k / n_v * 100)}%" if n_v else "-"

        uniq = view[view["email"] != ""].drop_duplicates("email")
        repeat = int((uniq["registrations"] > 1).sum())
        ig_has = (view["ig_handle"] != "") | (view["ig_followers"] > 0)
        tt_has = (view["tt_handle"] != "") | (view["tt_followers"] > 0)
        both = int((ig_has & tt_has).sum())
        tot_ig, tot_tt = int(view["ig_followers"].sum()), int(view["tt_followers"].sum())
        mikro = int((view["tier_ig"].isin(["Mikro", "Makro"]) | view["tier_tt"].isin(["Mikro", "Makro"])).sum())
        cities = clean_cat(view["city"]).value_counts()
        biz = view["business"].dropna()
        hij = view["hijab"].dropna()
        age = pd.to_numeric(view["age"], errors="coerce").dropna()
        age_txt = f" · median usia {_num_txt(_med(age))} th" if len(age) else ""
        kpi_cards([
            ("Data terpilih", fmt_int(n_match), f"{fmt_pct(n_match / N * 100)}% dari {fmt_int(N)} data{age_txt}", A["blue"]),
            ("Pendaftar unik", fmt_int(len(uniq)), f"{fmt_int(repeat)} mendaftar lebih dari sekali", A["teal"]),
            ("Total pengikut", fmt_compact(tot_ig + tot_tt),
             f"IG {fmt_compact(tot_ig)} · TikTok {fmt_compact(tot_tt)}", A["rose"]),
            ("Mikro ke atas", share(mikro), f"{fmt_int(mikro)} pendaftar (tier IG atau TikTok)", A["sky"]),
            ("Punya kedua akun", share(both), f"{fmt_int(both)} pendaftar punya IG dan TikTok", A["amber"]),
            ("Kota terbanyak", str(cities.index[0]) if len(cities) else "-",
             f"{fmt_int(cities.iloc[0])} data ({share(cities.iloc[0])})" if len(cities) else "-", A["green"]),
            ("Berhijab", f"{fmt_pct((hij == 'Berhijab').mean() * 100)}%" if len(hij) else "-",
             "dari yang mengisi", A["violet"]),
            ("Punya usaha", f"{fmt_pct((biz == 'Punya usaha').mean() * 100)}%" if len(biz) else "-",
             "dari yang mengisi", A["teal"]),
        ])

    c1, c2 = st.columns([2, 1], gap="large")
    with c1:
        with card("Pendaftar per bulan",
                  note="Geser slider di bawah grafik atau pakai tombol rentang untuk memperbesar periode tertentu."
                  if HAS_PLOTLY else None):
            if HAS_PLOTLY:
                show_plotly(chart_trend_plotly(view), height=360)
            else:
                show(chart_trend(view))
    with c2:
        with card("Periode pendaftaran"):
            show(chart_counts(view["period"], label="Periode", accent=A["violet"], order=PERIOD_ORDER, shade="value", step=28))

    with card("Bandingkan kelompok",
              note="Total pengikut = Instagram + TikTok. Median dihitung dari akun dengan pengikut minimal 1; kelompok di bawah 5 data disembunyikan untuk median."):
        a, b = st.columns([1, 1.8], vertical_alignment="bottom")
        dim = a.selectbox("Kelompokkan menurut", list(DIMS), key="v_dim", label_visibility="collapsed")
        metric = b.segmented_control("Ukuran", list(METRICS), default="Jumlah data", key="v_metric",
                                     required=True, label_visibility="collapsed")
        if dim not in DIMS:
            dim = next(iter(DIMS))
        if metric not in METRICS:
            metric = "Jumlah data"
        col, order = DIMS[dim]
        agg = view.assign(_tot=view["ig_followers"] + view["tt_followers"]).groupby(col, observed=True).agg(
            n=("row_id", "size"),
            tot=("_tot", "sum"),
            ig=("ig_followers", lambda s: s[s > 0].median()),
            tt=("tt_followers", lambda s: s[s > 0].median()),
        )
        mcol = METRICS[metric]
        if mcol in ("ig", "tt"):
            agg = agg[agg["n"] >= 5].dropna(subset=[mcol])
        agg = agg.reset_index().rename(columns={col: "group"})
        agg["group"] = agg["group"].astype(str)
        if order:
            agg = agg[agg["group"].isin(order)]
            sort = [o for o in order if o in set(agg["group"])]
            agg = agg.set_index("group").loc[sort].reset_index()
        else:
            agg = agg.sort_values(mcol, ascending=False).head(12)
            sort = agg["group"].tolist()
        if agg.empty:
            st.caption("Belum cukup data per kelompok (minimal 5) untuk menghitung median.")
        else:
            agg = shade_value(agg, mcol, METRIC_COLOR[mcol])
            st.altair_chart(
                hbar(agg, cat="group", val=mcol, sort=sort, fmt=",.0f", tooltip=[
                    alt.Tooltip("group:N", title=dim), alt.Tooltip("n:Q", title="Jumlah data", format=",.0f"),
                    alt.Tooltip("tot:Q", title="Total pengikut", format=",.0f"),
                    alt.Tooltip("ig:Q", title="Median pengikut IG", format=",.0f"),
                    alt.Tooltip("tt:Q", title="Median pengikut TikTok", format=",.0f")]),
                width="stretch",
            )

    d1, d2 = st.columns(2, gap="large")
    with d1:
        with card("Berapa kali mendaftar", note="Per email unik, dihitung dari seluruh data (bukan hanya hasil filter)."):
            u = view[view["email"] != ""].drop_duplicates("email")
            reg = u["registrations"].clip(upper=4).astype(int).astype(str).replace({"4": "4+"})
            show(chart_counts(reg, label="Jumlah pendaftaran", accent=A["amber"], order=["1", "2", "3", "4+"],
                              vertical=True, shade="order", height=230))
    with d2:
        with card("Kategori konten (15 teratas)", note="Satu pendaftar bisa memilih lebih dari satu kategori."):
            cnt = content_dm.loc[ALL_MASK].sum().sort_values(ascending=False).head(15)
            cnt = cnt[cnt > 0]
            if cnt.empty:
                st.caption("Belum ada data untuk grafik ini.")
            else:
                d = cnt.rename_axis("group").reset_index(name="n")
                d["share"] = d["n"] / max(n_match, 1)
                d = shade_value(d, "n", A["teal"])
                st.altair_chart(hbar(d, cat="group", val="n", sort=d["group"].tolist(),
                                     tooltip=[alt.Tooltip("group:N", title="Kategori"),
                                              alt.Tooltip("n:Q", title="Jumlah data", format=",.0f"),
                                              alt.Tooltip("share:Q", title="Porsi pendaftar", format=".1%")], step=26),
                                width="stretch")


# ============================================================ HALAMAN 3: AUDIENS
def page_audience():
    page_head("Audiens", "Siapa saja yang mendaftar: usia, kota, keluarga, dan latar belakang.")
    status_row()

    a, b = st.columns(2, gap="large")
    with a:
        with card("Kelompok usia"):
            show(chart_counts(age_groups(view["age"]), label="Usia", accent=A["blue"], order=AGE_BINS,
                              vertical=True, shade="value"))
    with b:
        with card("Kota (12 teratas)"):
            show(chart_counts(view["city"], label="Kota", accent=A["teal"], top=12, shade="value", step=24))

    d1, d2, d3 = st.columns(3, gap="large")
    with d1:
        with card("Hijab"):
            show(chart_donut(view["hijab"], label="Hijab", colors=PAL2))
    with d2:
        with card("Status pernikahan"):
            show(chart_donut(view["marital"], label="Status", colors=[A["violet"], A["rose"], A["sky"], A["green"]]))
    with d3:
        with card("Kehamilan"):
            show(chart_donut(view["pregnancy"], label="Kehamilan", colors=[A["rose"], A["teal"]]))

    e1, e2 = st.columns(2, gap="large")
    with e1:
        with card("Jumlah anak"):
            ks, korder = kids_series(view["kids"])
            show(chart_counts(ks, label="Jumlah anak", accent=A["amber"], order=korder, vertical=True, shade="value"))
    with e2:
        with card("Anak terkecil"):
            show(chart_counts(view["child_cat"], label="Anak terkecil", accent=A["violet"], order=CHILD_ORDER,
                              shade="order", step=26))

    f1, f2 = st.columns(2, gap="large")
    with f1:
        with card("Jenis sekolah anak"):
            show(chart_counts(view["school_type"], label="Sekolah anak", accent=A["green"], top=10, shade="value"))
    with f2:
        with card("Usia dan anak terkecil", note="Angka di dalam kotak adalah jumlah data."):
            show(chart_cross(view.assign(_age=age_groups(view["age"])), "_age", "child_cat",
                             x_order=AGE_BINS, y_order=CHILD_ORDER,
                             x_title="Usia", y_title="Anak terkecil", step=28))


# ============================================================ HALAMAN 4: JANGKAUAN
def _followers_bins(s: pd.Series) -> pd.Series:
    b = pd.cut(pd.to_numeric(s, errors="coerce").fillna(0), bins=[-1, 999, 2999, 9999, 29999, 99999, 999999, np.inf],
               labels=FOLLOWER_BINS)
    return b.astype("object")


SCATTER_COLORS = {
    "Tanpa warna": (None, None),
    "Hijab": ("hijab", None),
    "Usaha": ("business", None),
    "Jenis HP": ("phone_type", None),
    "Pekerjaan": ("job", None),
}
SCATTER_Z = {"Usia": "age", "Jumlah anak": "kids"}


def extreme_table(frame: pd.DataFrame, n: int):
    """Daftar akun dengan pengikut tertinggi dan terendah (menurut total IG + TikTok)."""
    d = frame[(frame["ig_followers"] > 0) | (frame["tt_followers"] > 0)].copy()
    d["Total"] = d["ig_followers"] + d["tt_followers"]
    cols = {"name": "Nama", "ig_handle": "Instagram", "ig_followers": "Pengikut IG",
            "tt_handle": "TikTok", "tt_followers": "Pengikut TikTok", "Total": "Total", "city": "Kota"}
    cfg = {c: st.column_config.NumberColumn(format="localized") for c in ("Pengikut IG", "Pengikut TikTok", "Total")}
    a, b = st.columns(2, gap="large")
    a.markdown('<div class="field-label">Pengikut tertinggi</div>', unsafe_allow_html=True)
    a.dataframe(d.nlargest(n, "Total")[list(cols)].rename(columns=cols), hide_index=True,
                column_config=cfg, width="stretch")
    b.markdown('<div class="field-label">Pengikut terendah (minimal punya 1 akun berpengikut)</div>', unsafe_allow_html=True)
    b.dataframe(d.nsmallest(n, "Total")[list(cols)].rename(columns=cols), hide_index=True,
                column_config=cfg, width="stretch")


def brand_legend():
    """Legenda logo untuk akun yang disorot."""
    items = [(EXT_HI, "f16d", "Instagram tertinggi"), (EXT_HI, "e07b", "TikTok tertinggi"),
             (EXT_LO, "f16d", "Instagram terendah"), (EXT_LO, "e07b", "TikTok terendah")]
    cells = "".join(f'<span><b style="color:{c}">&#x{cp};</b>{e(t)}</span>' for c, cp, t in items)
    st.markdown(f'<div class="brand-lg">{cells}</div>', unsafe_allow_html=True)


def scatter_card():
    with card("Pengikut Instagram dan TikTok per akun"):
        # baris 1: warna (5 pilihan, butuh lebar penuh)
        pick = st.segmented_control("Warnai menurut", list(SCATTER_COLORS), default="Tanpa warna",
                                    key="v_scatter", required=True)
        if pick not in SCATTER_COLORS:
            pick = "Tanpa warna"

        # baris 2: tampilan, sumbu ketiga, jumlah ekstrem
        c1, c2, c3 = st.columns([1, 1.6, 2], gap="medium", vertical_alignment="bottom")
        mode = c1.segmented_control("Tampilan", ["2D", "3D"], default="3D", key="v_scatter_mode3",
                                    required=True) if HAS_PLOTLY else "2D"
        three_d = mode == "3D"
        zlabel = c2.segmented_control("Sumbu ketiga", list(SCATTER_Z), default="Usia", key="v_scatter_z",
                                      required=True, disabled=not three_d)
        if zlabel not in SCATTER_Z:
            zlabel = "Usia"
        n_ext = c3.slider("Jumlah akun disorot per kategori", 3, 20, 5, key="v_scatter_ext")

        data = view
        ccol, corder = SCATTER_COLORS[pick]
        if HAS_PLOTLY:
            brand_legend()
            fig, shown, total = chart_scatter_plotly(data, ccol, corder, pick, three_d=three_d,
                                                     zcol=SCATTER_Z[zlabel], zlabel=zlabel, n_ext=n_ext)
            show_plotly(fig, height=580 if three_d else 480, empty="Belum ada akun dengan pengikut di kedua platform.")
        else:
            chart, shown, total = chart_scatter(data, ccol, corder, pick)
            show(chart, "Belum ada akun dengan pengikut di kedua platform.")
            st.caption("Pasang plotly (pip install plotly) untuk tampilan interaktif dan 3D.")

        with st.expander("Peringkat akun menurut pengikut"):
            extreme_table(data, n_ext)


def page_reach():
    page_head("Jangkauan sosmed", "Ukuran akun Instagram dan TikTok pendaftar.")
    status_row()

    # --- grafik utama: scatter interaktif, bisa 2D atau 3D
    scatter_card()

    a, b = st.columns(2, gap="large")
    with a:
        with card("Sebaran pengikut Instagram"):
            show(chart_counts(_followers_bins(view["ig_followers"]), label="Pengikut IG", accent=A["rose"],
                              order=FOLLOWER_BINS, vertical=True, shade="value"))
    with b:
        with card("Sebaran pengikut TikTok"):
            show(chart_counts(_followers_bins(view["tt_followers"]), label="Pengikut TikTok", accent=A["sky"],
                              order=FOLLOWER_BINS, vertical=True, shade="value"))

    c, d = st.columns(2, gap="large")
    with c:
        with card("Tier Instagram"):
            show(chart_counts(view["tier_ig"], label="Tier IG", accent=A["rose"], order=TIER_ORDER, shade="order", step=34))
    with d:
        with card("Tier TikTok"):
            show(chart_counts(view["tier_tt"], label="Tier TikTok", accent=A["sky"], order=TIER_ORDER, shade="order", step=34))

    e1, e2 = st.columns(2, gap="large")
    with e1:
        with card("Punya akun di platform mana"):
            ig_has = (view["ig_handle"] != "") | (view["ig_followers"] > 0)
            tt_has = (view["tt_handle"] != "") | (view["tt_followers"] > 0)
            presence = pd.Series(np.select([ig_has & tt_has, ig_has, tt_has], PRESENCE_ORDER[:3], PRESENCE_ORDER[3]),
                                 index=view.index)
            show(chart_donut(presence, label="Akun", order=PRESENCE_ORDER,
                             colors=[A["teal"], A["rose"], A["sky"], "#69759F"]))
    with e2:
        with card("Tier Instagram dan tier TikTok", note="Angka di dalam kotak adalah jumlah data."):
            show(chart_cross(view, "tier_tt", "tier_ig", x_order=TIER_ORDER, y_order=TIER_ORDER,
                             x_title="Tier TikTok", y_title="Tier Instagram", step=36))

    g1, g2 = st.columns(2, gap="large")
    with g1:
        with card("Total pengikut menurut kota (10 teratas)",
                  note="Jumlah pengikut Instagram dan TikTok dari semua pendaftar di kota itu."):
            show(chart_total_followers(view, "city", top=10))
    with g2:
        with card("Median pengikut menurut pekerjaan", note="10 pekerjaan terbanyak, kelompok di bawah 5 data disembunyikan."):
            show(chart_median_pair(view, "job", top=10), "Belum cukup data per kelompok (minimal 5).")


# ============================================================ HALAMAN 5: EKONOMI & PROFIL
def _money_chart(s: pd.Series, label: str, accent: str):
    """Kolom pendapatan/pengeluaran: rentang berlabel -> batang berurutan; angka murni -> histogram."""
    if pd.api.types.is_numeric_dtype(s) and s.nunique() > 12:
        return chart_hist(s, label=label, accent=accent)
    return chart_counts(s, label=label, accent=accent, order=bucket_order(clean_cat(s).unique()), shade="value",
                        vertical=True, height=250)


def page_economy():
    page_head("Ekonomi dan profil", "Pendapatan, pengeluaran, pekerjaan, usaha, dan perangkat.")
    status_row()

    c, d = st.columns(2, gap="large")
    with c:
        with card("Pendapatan per bulan"):
            show(_money_chart(view["income"], "Pendapatan", A["green"]))
    with d:
        with card("Pengeluaran per bulan"):
            show(_money_chart(view["spend"], "Pengeluaran", A["amber"]))

    numeric_money = pd.api.types.is_numeric_dtype(view["income"]) or pd.api.types.is_numeric_dtype(view["spend"])
    if not numeric_money:
        with card("Pendapatan dan pengeluaran", note="Angka di dalam kotak adalah jumlah data."):
            show(chart_cross(view, "spend", "income", x_order=bucket_order, y_order=bucket_order,
                             x_title="Pengeluaran", y_title="Pendapatan", step=34))

    e1, e2 = st.columns(2, gap="large")
    with e1:
        with card("Pekerjaan (10 teratas)"):
            show(chart_counts(view["job"], label="Pekerjaan", accent=A["rose"], top=10, shade="value", step=28))
    with e2:
        with card("Punya usaha"):
            show(chart_donut(view["business"], label="Usaha", order=["Punya usaha", "Tanpa usaha"],
                             colors=[A["green"], "#69759F"]))

    with card("Jenis HP"):
        show(chart_donut(view["phone_type"], label="Jenis HP", colors=[A["teal"], A["violet"]]))


# ============================================================ META ADS (halaman tambahan)
# Kode ada di folder meta_ads/. Kalau modulnya bermasalah, hanya halaman ini yang menampilkan pesan error.
try:
    from meta_ads.page import render as _render_meta_ads
    _META_ADS_ERR = None
except Exception as _exc:
    _render_meta_ads, _META_ADS_ERR = None, _exc


def page_meta_ads():
    page_head("Meta Ads", "Ringkasan performa iklan Instagram, Traffic, dan Leads.")
    if _render_meta_ads is None:
        st.error(f"Modul Meta Ads tidak dapat dimuat: {_META_ADS_ERR}")
        return
    with safe_section("Halaman Meta Ads tidak dapat ditampilkan"):
        _render_meta_ads()


# ============================================================ NAVIGASI
PAGES = [
    st.Page(page_data, title="Data", icon=":material/table_view:", default=True),
    st.Page(page_overview, title="Ringkasan", icon=":material/dashboard:", url_path="ringkasan"),
    st.Page(page_audience, title="Audiens", icon=":material/groups:", url_path="audiens"),
    # Label navigasi dibuat singkat agar seluruh menu muat pada desktop.
    # Judul halaman di dalam konten tetap memakai nama lengkap.
    st.Page(page_reach, title="Jangkauan", icon=":material/campaign:", url_path="jangkauan"),
    st.Page(page_economy, title="Ekonomi", icon=":material/payments:", url_path="ekonomi"),
    st.Page(page_meta_ads, title="Meta Ads", icon=":material/ads_click:", url_path="meta-ads"),
]
# Menu halaman di bagian atas halaman utama (bukan di sidebar); sidebar khusus untuk pencarian dan filter.
pg = st.navigation(PAGES, position="top")

# ---- tampilan kosong: tunjukkan filter mana yang bisa dilonggarkan (berlaku di semua halaman)
if n_match == 0 and pg.url_path != "meta-ads":  # halaman Meta Ads tidak memakai filter influencer
    status_row()
    st.warning("Tidak ada data yang cocok dengan kombinasi ini.")
    rescue = sorted(((int(others(g).sum()), g) for g in ACTIVE), reverse=True)
    rescue = [(n, g) for n, g in rescue if n > 0][:4]
    if rescue:
        st.markdown('<div class="field-label">Longgarkan salah satu filter</div>', unsafe_allow_html=True)
        for n, g in rescue:
            st.button(f"Hapus {ACTIVE[g]['chip']}  \u2192  {fmt_int(n)} data", key=f"rescue_{g}",
                      on_click=clear_keys, args=(ACTIVE[g]["keys"],))
    st.stop()

pg.run()
