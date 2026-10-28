"""KPI ala dashboard Get Follow / Traffic dari insights level campaign + actions."""
from __future__ import annotations

import re

import pandas as pd


def norm(s) -> str:
    return re.sub(r"[\s_\-\.]+", " ", str(s).lower()).strip()


def parse_keywords(text) -> list[str]:
    return [norm(k) for k in str(text).split(",") if k.strip()]


def classify(name, follow_kw, traffic_kw, follow_count=0) -> str:
    """Classify from campaign naming first, then infer Get Follow from observed results."""
    n = norm(name)
    # Explicit Get Follow labels win. Generic 'follow' keyword is checked only after
    # traffic markers so a Traffic campaign is not misclassified by incidental wording.
    explicit_follow = [k for k in follow_kw if k not in {"follow", "ig", "instagram"}]
    if any(k in n for k in explicit_follow):
        return "Get Follow"
    if any(k in n for k in traffic_kw):
        return "Traffic"
    if any(k in n for k in follow_kw):
        return "Get Follow"
    # Backend inference: an unlabeled campaign with observed follow results is treated
    # as Get Follow unless its name explicitly marks it as Traffic.
    try:
        if float(follow_count or 0) > 0:
            return "Get Follow"
    except (TypeError, ValueError):
        pass
    return "Lainnya"


def action_map(actions) -> dict:
    out: dict = {}
    if not isinstance(actions, list):
        return out
    for a in actions:
        try:
            out[a["action_type"]] = out.get(a["action_type"], 0.0) + float(a["value"])
        except (KeyError, TypeError, ValueError):
            pass
    return out


def action_totals(df: pd.DataFrame) -> pd.Series:
    tot: dict = {}
    if "actions" in df:
        for acts in df["actions"]:
            for k, v in action_map(acts).items():
                tot[k] = tot.get(k, 0.0) + v
    return pd.Series(tot, dtype=float).sort_values(ascending=False)


def detect(totals: pd.Series, kind: str) -> list[str]:
    """Tebak action_type follow / profile visit (urut dari total terbesar)."""
    low = {t: t.lower() for t in totals.index}
    if kind == "follow":
        return [t for t, l in low.items() if "follow" in l and "unfollow" not in l]
    return [t for t, l in low.items() if "profile" in l and "visit" in l]


def enrich(
    df, follow_types, visit_types, follow_kw, traffic_kw,
    visit_source="actions", follow_source="actions",
) -> pd.DataFrame:
    d = df.copy()
    maps = d["actions"].map(action_map) if "actions" in d else pd.Series([{}] * len(d), index=d.index)
    if follow_source == "instagram_profile_follow" and "instagram_profile_follow" in d:
        d["follows"] = pd.to_numeric(d["instagram_profile_follow"], errors="coerce").fillna(0.0)
    else:
        d["follows"] = maps.map(lambda m: sum(m.get(t, 0.0) for t in follow_types)).astype(float)
    if visit_source == "instagram_profile_visits" and "instagram_profile_visits" in d:
        d["visits"] = pd.to_numeric(d["instagram_profile_visits"], errors="coerce").fillna(0.0)
    else:
        d["visits"] = maps.map(lambda m: sum(m.get(t, 0.0) for t in visit_types)).astype(float)
    d["kategori"] = d.apply(
        lambda row: classify(row.get("campaign_name", ""), follow_kw, traffic_kw, row.get("follows", 0)),
        axis=1,
    )
    return d


def div(a, b):
    return a / b if b else None


def summarize(d: pd.DataFrame) -> dict:
    def tot(col, cat=None):
        x = d if cat is None else d[d["kategori"] == cat]
        return float(x[col].sum())

    return dict(
        spend=tot("spend"), impressions=tot("impressions"),
        spend_gf=tot("spend", "Get Follow"), spend_tr=tot("spend", "Traffic"),
        follows_all=tot("follows"), follows_gf=tot("follows", "Get Follow"),
        visits_all=tot("visits"), visits_gf=tot("visits", "Get Follow"), visits_tr=tot("visits", "Traffic"),
        n_traffic=int((d["kategori"] == "Traffic").sum()),
        max_freq=float(d["frequency"].max()) if len(d) else 0.0,
    )


def by_category(d: pd.DataFrame) -> pd.DataFrame:
    g = (d.groupby("kategori")
         .agg(campaign=("campaign_name", "nunique"), spend=("spend", "sum"),
              impressions=("impressions", "sum"), follows=("follows", "sum"), visits=("visits", "sum"))
         .reset_index())
    order = {"Get Follow": 0, "Traffic": 1, "Lainnya": 2}
    g = g.sort_values("kategori", key=lambda s: s.map(order)).reset_index(drop=True)
    total = g["spend"].sum()
    g["% spend"] = g["spend"] / total * 100 if total else 0.0
    g["CPF"] = g["spend"] / g["follows"].where(g["follows"] > 0)
    g["CPV"] = g["spend"] / g["visits"].where(g["visits"] > 0)
    g["CR follow÷visit (%)"] = g["follows"] / g["visits"].where(g["visits"] > 0) * 100
    return g
