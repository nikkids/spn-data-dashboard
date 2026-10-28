# Meta Ads (menu tambahan di SPN Data Explorer)

Menu **Meta Ads** di `app-publish.py` memanggil `meta_ads/page.py`. Isinya: Traffic, Leads, Cek Data & Admin.

- `page.py`: halaman Streamlit (fungsi `render()`)
- `meta_client.py`: client Graph API (log, pagination, error handling)
- `checks.py`: pengecekan alur data / health check
- `kpi.py`: helper klasifikasi dan actions
- `check_flow.py`: versi terminal untuk cron: `python -m meta_ads.check_flow --days 7` (exit code 1 kalau ada FAIL)

Secrets: lihat `.streamlit/secrets.toml.example`. Tanpa secrets, hanya halaman Meta Ads yang menampilkan pesan, halaman SPN lain tidak terpengaruh.
