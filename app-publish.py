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


# ============================================================ NAVIGASI
PAGES = [
    st.Page(page_data, title="Data", icon=":material/table_view:", default=True),
    st.Page(page_overview, title="Ringkasan", icon=":material/dashboard:", url_path="ringkasan"),
    st.Page(page_audience, title="Audiens", icon=":material/groups:", url_path="audiens"),
    st.Page(page_reach, title="Jangkauan sosmed", icon=":material/campaign:", url_path="jangkauan"),
    st.Page(page_economy, title="Ekonomi dan profil", icon=":material/payments:", url_path="ekonomi"),
]
# Menu halaman di bagian atas halaman utama (bukan di sidebar); sidebar khusus untuk pencarian dan filter.
pg = st.navigation(PAGES, position="top")

# ---- tampilan kosong: tunjukkan filter mana yang bisa dilonggarkan (berlaku di semua halaman)
if n_match == 0:
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
