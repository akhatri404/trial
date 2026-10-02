"""Turns the tidy table into KPIs, growth bridges, SKU-vs-sales verdicts, rankings, direction buckets and checks.

Everything is deterministic and every threshold lives in `Thresholds`, so the same file always produces the same
report and the rules can be shown (and debated) in the appendix.

Scope: all rows from the loader (sub-category sheet plus any category that only exists in the category sheet), so the
totals reconcile with the source's own department totals. `reconcile` proves that on every run.
"""
from __future__ import annotations

from dataclasses import dataclass, field, replace

import numpy as np
import pandas as pd

from .loader import WorkbookData

RECAPTURE, REVIEW, SCALE, REPLICATE, TEST, DEFEND, DEPRIORITIZE, MONITOR = (
    "シェア奪回", "要点検", "拡大", "横展開", "小規模テスト", "防衛", "優先度低", "経過観察",
)
ORDER = [SCALE, REPLICATE, RECAPTURE, TEST, REVIEW, DEFEND, DEPRIORITIZE, MONITOR]

V_EFFICIENT, V_SKULED, V_SKUDOWNSALES = "売上がSKU増に追随", "SKU増が売上増を上回る", "SKU増・売上減"
V_PRUNED, V_SHRINK = "SKU削減・売上増", "SKU削減・売上減"
V_STABLE_UP, V_STABLE_DOWN = "SKU横ばい・売上増", "SKU横ばい・売上減"
V_NEW, V_EXIT, V_SMALL = "PB新規投入", "PB撤退", "小規模(判定外)"
T_UP, T_FLAT, T_DOWN = "上昇", "横ばい", "低下"      # PB share trend: newest (part-year) period vs the last full period

# Internal keys stay Japanese; `tr` translates them for display only.
_EN = {
    SCALE: "Scale", REPLICATE: "Replicate", RECAPTURE: "Recapture", TEST: "Test small", REVIEW: "Review",
    DEFEND: "Defend", DEPRIORITIZE: "Deprioritize", MONITOR: "Monitor",
    V_EFFICIENT: "Sales keep pace with SKUs", V_SKULED: "SKUs outpace sales", V_SKUDOWNSALES: "SKUs up, sales down",
    V_PRUNED: "Pruned, sales up", V_SHRINK: "Shrinking", V_STABLE_UP: "Stable range, sales up",
    V_STABLE_DOWN: "Stable range, sales down", V_NEW: "New PB range", V_EXIT: "PB range exited", V_SMALL: "Small base",
    T_UP: "Rising", T_FLAT: "Flat", T_DOWN: "Falling",
}


def tr(key, lang="ja"):
    """Display name of a direction or verdict in the requested language."""
    return key if lang == "ja" else _EN.get(key, key)


@dataclass
class Thresholds:
    min_headroom: float = 200e6          # non-PB sales (JPY) needed to call a pool "big"
    scale_share_gain_pt: float = 2.0     # PB share gain (pt) vs previous period to count as momentum
    scale_min_pb_sales: float = 50e6
    scale_min_market_yoy: float = -0.03  # market must not be shrinking faster than this
    white_space_max_share: float = 0.05  # PB share at or below this = "no PB position"
    test_min_market_yoy: float = -0.03   # white-space pools shrinking faster than this are not Replicate/Test candidates
    recapture_drop_pt: float = 5.0       # PB share loss (pt) that counts as a collapse
    recapture_min_prev_pb: float = 20e6
    review_drop_pt: float = 2.0          # smaller share loss that deserves a look
    review_min_prev_pb: float = 20e6
    defend_min_share: float = 0.50       # PB share at/above this = saturated
    defend_min_pb_sales: float = 100e6
    deprioritize_max_sales: float = 200e6
    deprioritize_market_yoy: float = -0.30
    deprioritize_max_growth: float = 0.20  # a small no-PB pool growing faster than this stays on watch
    sku_growth_flag: float = 0.50        # PB SKU count growth that triggers a dilution check
    sku_productivity_drop_flag: float = 0.25
    sku_stable_band: float = 0.10        # SKU change within +/- this is "stable range"
    min_base_pb_sales: float = 50e6      # minimum PB sales base for growth-% rankings and SKU verdicts
    watch_min_pb_sales: float = 100e6    # minimum PB sales for the two-step share-decline watchlist
    watch_min_step_pt: float = 0.5       # each step must fall by at least this many pt
    upside_scale_cap: float = 0.70       # max PB share assumed for Scale items
    upside_scale_momentum_years: float = 2.0
    upside_replicate_share: float = 0.10  # PB share assumed for Replicate items
    recon_sales_tolerance: float = 0.002   # reconciliation: allowed relative sales difference
    recon_ratio_tolerance_pt: float = 0.10  # reconciliation: allowed PB-ratio difference in pt
    asp_flag: float = 0.30               # price/mix change (avg yen per unit) that deserves a data check
    momentum_pt: float = 2.0             # PB share move (pt), newest part-year period vs last full period, that counts as rising / falling


@dataclass
class Result:
    periods_full: tuple
    period_latest: str | None
    partial_months: int | None
    kpi: dict
    subs: pd.DataFrame
    cats: pd.DataFrame
    directions: dict
    sku_flags: pd.DataFrame
    upside: pd.DataFrame
    upside_total: dict
    caveats: list
    thresholds: Thresholds
    findings: list = field(default_factory=list)
    bridge: dict = field(default_factory=dict)
    rankings: dict = field(default_factory=dict)
    watchlist: pd.DataFrame = None
    concentration: dict = field(default_factory=dict)
    recon: pd.DataFrame = None
    dq: list = field(default_factory=list)
    borderline: pd.DataFrame = None
    scope_note: str = ""
    lang: str = "ja"
    depts: pd.DataFrame = None                       # one row per department (部門), largest sales first
    dept_findings: dict = field(default_factory=dict)  # {department: [plain-language takeaways]}


# ------------------------------------------------------------------------------------------------------------
def _jpy(x) -> str:
    """Japanese money units: 12.34億円 / 5,300万円 / 8,000円."""
    if x is None or x != x:
        return "－"
    sign, a = ("-" if x < 0 else ""), abs(x)
    if a >= 1e10:
        return f"{sign}{a / 1e8:,.1f}億円"
    if a >= 1e8:
        return f"{sign}{a / 1e8:,.2f}億円"
    if a >= 1e4:
        return f"{sign}{a / 1e4:,.0f}万円"
    return f"{sign}{a:,.0f}円"


def _usd(x) -> str:
    """English money format: ¥1.78B / ¥53M / ¥8,000."""
    if x is None or x != x:
        return "–"
    sign, a = ("-" if x < 0 else ""), abs(x)
    if a >= 1e9:
        return f"{sign}¥{a / 1e9:,.2f}B"
    if a >= 1e8:
        return f"{sign}¥{a / 1e6:,.0f}M"
    if a >= 1e6:
        return f"{sign}¥{a / 1e6:,.1f}M"
    return f"{sign}¥{a:,.0f}"


def money(lang="ja"):
    return _jpy if lang == "ja" else _usd


def _div(a, b):
    a, b = np.asarray(a, dtype=float), np.asarray(b, dtype=float)
    with np.errstate(divide="ignore", invalid="ignore"):
        return np.where(b != 0, a / b, np.nan)


def _verdict(sku_prev, sku_cur, pb_prev, pb_cur, t: Thresholds) -> str:
    """Compares the change in PB SKU count with the change in PB sales."""
    if sku_prev <= 0 and sku_cur > 0:
        return V_NEW
    if sku_cur <= 0 and sku_prev > 0:
        return V_EXIT
    if sku_prev <= 0 or pb_prev < t.min_base_pb_sales:
        return V_SMALL
    sg, pg = sku_cur / sku_prev - 1, (pb_cur / pb_prev - 1) if pb_prev else np.nan
    if sg > t.sku_stable_band:
        return V_EFFICIENT if pg >= sg else (V_SKULED if pg >= 0 else V_SKUDOWNSALES)
    if sg < -t.sku_stable_band:
        return V_PRUNED if pg >= 0 else V_SHRINK
    return V_STABLE_UP if pg >= 0 else V_STABLE_DOWN


def _classify_all(d: pd.DataFrame, t: Thresholds) -> pd.Series:
    scale_cats = set(d.loc[(d.d_share_pt >= t.scale_share_gain_pt) & (d.pb_cur >= t.scale_min_pb_sales)
                           & (d.nb_cur >= t.min_headroom) & (d.yoy > t.scale_min_market_yoy), "cat"])

    def one(r):
        if r.share_prev > 0 and r.d_share_pt <= -t.recapture_drop_pt and r.pb_prev >= t.recapture_min_prev_pb \
                and r.sales_prev >= t.min_headroom:
            return RECAPTURE
        if r.d_share_pt <= -t.review_drop_pt and r.pb_prev >= t.review_min_prev_pb:
            return REVIEW
        if r.d_share_pt >= t.scale_share_gain_pt and r.pb_cur >= t.scale_min_pb_sales \
                and r.nb_cur >= t.min_headroom and r.yoy > t.scale_min_market_yoy:
            return SCALE
        if r.share_cur >= t.defend_min_share and r.pb_cur >= t.defend_min_pb_sales:
            return DEFEND
        if r.share_cur <= t.white_space_max_share and r.nb_cur >= t.min_headroom \
                and (pd.isna(r.yoy) or r.yoy >= t.test_min_market_yoy):
            return REPLICATE if r["cat"] in scale_cats else TEST
        latest_ok = pd.isna(r.share_latest) or r.share_latest <= t.white_space_max_share
        small_dead = (r.sales < t.deprioritize_max_sales and r.share_cur <= t.white_space_max_share and latest_ok
                      and not (pd.notna(r.yoy) and r.yoy > t.deprioritize_max_growth))
        shrinking = pd.notna(r.yoy) and r.yoy <= t.deprioritize_market_yoy and r.share_cur <= t.white_space_max_share
        if small_dead or shrinking:
            return DEPRIORITIZE
        return MONITOR

    return d.apply(one, axis=1)


# ------------------------------------------------------------------------------------------------------------
def analyze(data: WorkbookData, t: Thresholds | None = None, lang: str = "ja") -> Result:
    t = t or Thresholds()
    T = lambda ja, en: ja if lang == "ja" else en     # pick the display text
    J = money(lang)
    SEP = T("、", ", ")
    df, periods = data.df.copy(), data.periods

    # ---- which periods are full years, which one is partial ------------------------------------------------
    tot = {p: df[f"sales_{p}"].sum() for p in periods}
    latest, full = None, list(periods)
    if len(periods) >= 3:
        last, before = periods[-1], periods[-2]
        ratio = tot[last] / tot[before] if tot[before] else 1.0
        if data.partial_months or ratio < 0.6:
            latest, full = last, periods[:-1]
    prev, cur = full[-2], full[-1]

    for p in periods:
        df[f"pb_sales_{p}"] = df[f"sales_{p}"] * df[f"pb_ratio_sales_{p}"]
        df[f"pb_units_{p}"] = df[f"units_{p}"] * df[f"pb_ratio_units_{p}"]

    d = pd.DataFrame({"dept": df["dept"], "cat": df["cat"], "sub": df["sub"], "source": df["source"]})
    d["name"] = np.where(d["sub"] == d["cat"], d["cat"], d["cat"] + " / " + d["sub"])
    d["sales"], d["sales_prev"] = df[f"sales_{cur}"], df[f"sales_{prev}"]
    d["yoy"] = _div(d.sales - d.sales_prev, d.sales_prev)
    d["units"], d["units_prev"] = df[f"units_{cur}"], df[f"units_{prev}"]
    d["units_yoy"] = _div(d.units - d.units_prev, d.units_prev)
    d["asp"], d["asp_prev"] = _div(d.sales, d.units), _div(d.sales_prev, d.units_prev)
    d["asp_chg"] = _div(d.asp - d.asp_prev, d.asp_prev)
    d["pb_prev"], d["pb_cur"] = df[f"pb_sales_{prev}"], df[f"pb_sales_{cur}"]
    d["nb_prev"], d["nb_cur"] = d.sales_prev - d.pb_prev, d.sales - d.pb_cur
    d["share_prev"], d["share_cur"] = df[f"pb_ratio_sales_{prev}"], df[f"pb_ratio_sales_{cur}"]
    d["share_latest"] = df[f"pb_ratio_sales_{latest}"] if latest else np.nan
    if latest:
        d.loc[df[f"sales_{latest}"] <= 0, "share_latest"] = np.nan
    d["d_share_pt"] = (d.share_cur - d.share_prev) * 100
    d["d_pb"] = d.pb_cur - d.pb_prev
    d["pb_growth"] = _div(d.d_pb, d.pb_prev)
    d["pb_units"], d["pb_units_prev"] = df[f"pb_units_{cur}"], df[f"pb_units_{prev}"]
    nb_units = d.units - d.pb_units
    d["price_idx"] = _div(_div(d.pb_cur, d.pb_units), _div(d.nb_cur, nb_units))   # PB avg yen/unit vs non-PB
    d["sku_prev"], d["sku_cur"] = df[f"pb_skus_{prev}"], df[f"pb_skus_{cur}"]
    d["sku_latest"] = df[f"pb_skus_{latest}"] if latest else np.nan
    d["sales_latest"] = df[f"sales_{latest}"] if latest else np.nan
    d["pb_latest"] = df[f"pb_sales_{latest}"] if latest else np.nan
    d["mom_pt"] = ((d.share_latest - d.share_cur) * 100).round(1)         # PB share move in points (rounded as shown), newest vs last full period
    sized = d.pb_cur >= t.min_base_pb_sales                               # tiny bases are not tagged
    d["trend"] = np.where(~sized | d.mom_pt.isna(), "",
                          np.where(d.mom_pt >= t.momentum_pt, T_UP, np.where(d.mom_pt <= -t.momentum_pt, T_DOWN, T_FLAT)))
    d["sku_growth"] = _div(d.sku_cur - d.sku_prev, d.sku_prev)
    d["per_sku_prev"], d["per_sku_cur"] = _div(d.pb_prev, d.sku_prev), _div(d.pb_cur, d.sku_cur)
    d["per_sku_chg"] = _div(d.per_sku_cur - d.per_sku_prev, d.per_sku_prev)
    d_sku = d.sku_cur - d.sku_prev
    d["marg_per_sku"] = np.where(d_sku > 0, d.d_pb / d_sku.replace(0, np.nan), np.nan)
    d["verdict"] = [_verdict(a, b, c, e, t) for a, b, c, e in zip(d.sku_prev, d.sku_cur, d.pb_prev, d.pb_cur)]

    # ---- growth bridge: PB sales change = market effect + share effect (exact, row by row) ---------------
    d["market_eff"] = (d.sales - d.sales_prev) * d.share_prev
    d["share_eff"] = d.sales * (d.share_cur - d.share_prev)
    S0, S1, PB0, PB1 = d.sales_prev.sum(), d.sales.sum(), d.pb_prev.sum(), d.pb_cur.sum()
    R0, R1 = PB0 / S0, PB1 / S1
    w0, w1 = d.sales_prev / S0, d.sales / S1
    d["rate_pt"] = ((w0 + w1) / 2 * (d.share_cur - d.share_prev)) * 100        # within-sub-category share change
    d["mix_pt"] = ((d.share_prev + d.share_cur) / 2 * (w1 - w0)) * 100         # shift of sales towards higher/lower PB areas
    Rbar = (R0 + R1) / 2
    d["mix_c_pt"] = (((d.share_prev + d.share_cur) / 2 - Rbar) * (w1 - w0)) * 100   # sums to the same total as mix_pt
    bridge = dict(
        pb_prev=PB0, pb_cur=PB1, d_pb=PB1 - PB0,
        market=float(d.market_eff.sum()), share=float(d.share_eff.sum()),
        share_gain=float(d.loc[d.share_eff > 0, "share_eff"].sum()), share_loss=float(d.loc[d.share_eff < 0, "share_eff"].sum()),
        ratio_prev=R0, ratio_cur=R1, d_ratio_pt=(R1 - R0) * 100,
        rate_pt=float(d.rate_pt.sum()), mix_pt=float(d.mix_pt.sum()),
    )
    assert abs(bridge["market"] + bridge["share"] - bridge["d_pb"]) < 1.0, "bridge does not add up"
    assert abs(bridge["rate_pt"] + bridge["mix_pt"] - bridge["d_ratio_pt"]) < 1e-6, "ratio bridge does not add up"

    d["direction"] = _classify_all(d, t)

    # ---- category level --------------------------------------------------------------------------------
    agg = dict(sales=("sales", "sum"), sales_prev=("sales_prev", "sum"), pb_cur=("pb_cur", "sum"), pb_prev=("pb_prev", "sum"),
               sku_prev=("sku_prev", "sum"), sku_cur=("sku_cur", "sum"), market_eff=("market_eff", "sum"),
               share_eff=("share_eff", "sum"), rate_pt=("rate_pt", "sum"), mix_pt=("mix_pt", "sum"),
               units=("units", "sum"), units_prev=("units_prev", "sum"))
    agg.update(sku_latest=("sku_latest", "sum"), sales_latest=("sales_latest", "sum"), pb_latest=("pb_latest", "sum"))
    c = d.groupby(["dept", "cat"], as_index=False).agg(**agg)
    c["share_prev"], c["share_cur"] = _div(c.pb_prev, c.sales_prev), _div(c.pb_cur, c.sales)
    c["share_latest"] = _div(c.pb_latest, c.sales_latest)
    c["yoy"] = _div(c.sales - c.sales_prev, c.sales_prev)
    c["units_yoy"] = _div(c.units - c.units_prev, c.units_prev)
    c["pb_growth"] = _div(c.pb_cur - c.pb_prev, c.pb_prev)
    c["sku_growth"] = _div(c.sku_cur - c.sku_prev, c.sku_prev)
    c["per_sku_prev"], c["per_sku_cur"] = _div(c.pb_prev, c.sku_prev), _div(c.pb_cur, c.sku_cur)
    c["per_sku_chg"] = _div(c.per_sku_cur - c.per_sku_prev, c.per_sku_prev)
    c["d_pb"] = c.pb_cur - c.pb_prev
    c["d_share_pt"] = (c.share_cur - c.share_prev) * 100
    dsk = c.sku_cur - c.sku_prev
    c["marg_per_sku"] = np.where(dsk > 0, c.d_pb / dsk.replace(0, np.nan), np.nan)
    c["verdict"] = [_verdict(a, b, p0, p1, t) for a, b, p0, p1 in zip(c.sku_prev, c.sku_cur, c.pb_prev, c.pb_cur)]
    c["rank_pb"] = c.pb_cur.rank(ascending=False, method="min").astype(int)
    c["rank_share"] = c.share_cur.rank(ascending=False, method="min").astype(int)
    c["rank_growth"] = c.pb_growth.rank(ascending=False, method="min")
    c["rank_per_sku"] = c.per_sku_cur.rank(ascending=False, method="min")
    c["rank_dept"] = c.groupby("dept").pb_cur.rank(ascending=False, method="min").astype(int)   # rank inside the department
    c = c.sort_values("pb_cur", ascending=False).reset_index(drop=True)
    flags = c[(c.sku_growth >= t.sku_growth_flag) & (c.per_sku_chg <= -t.sku_productivity_drop_flag)] \
        .sort_values("sku_growth", ascending=False).reset_index(drop=True)

    # ---- department level (部門) ----------------------------------------------------------------------------
    dp = d.groupby("dept", as_index=False, sort=False).agg(
        cats=("cat", "nunique"), subs=("sub", "count"), sales=("sales", "sum"), sales_prev=("sales_prev", "sum"),
        pb_cur=("pb_cur", "sum"), pb_prev=("pb_prev", "sum"), nb_cur=("nb_cur", "sum"),
        sku_prev=("sku_prev", "sum"), sku_cur=("sku_cur", "sum"), market_eff=("market_eff", "sum"),
        share_eff=("share_eff", "sum"), rate_pt=("rate_pt", "sum"), mix_pt=("mix_c_pt", "sum"),
        sku_latest=("sku_latest", "sum"), sales_latest=("sales_latest", "sum"), pb_latest=("pb_latest", "sum"))
    dp["share_prev"], dp["share_cur"] = _div(dp.pb_prev, dp.sales_prev), _div(dp.pb_cur, dp.sales)
    dp["yoy"] = _div(dp.sales - dp.sales_prev, dp.sales_prev)
    dp["d_pb"] = dp.pb_cur - dp.pb_prev
    dp["pb_growth"] = _div(dp.d_pb, dp.pb_prev)
    dp["d_share_pt"] = (dp.share_cur - dp.share_prev) * 100
    dp["share_latest"] = _div(dp.pb_latest, dp.sales_latest)
    dp["per_sku_latest"] = _div(dp.pb_latest, dp.sku_latest)
    dp["sku_growth"] = _div(dp.sku_cur - dp.sku_prev, dp.sku_prev)
    dp["per_sku_prev"], dp["per_sku_cur"] = _div(dp.pb_prev, dp.sku_prev), _div(dp.pb_cur, dp.sku_cur)
    dp["per_sku_chg"] = _div(dp.per_sku_cur - dp.per_sku_prev, dp.per_sku_prev)
    dp["verdict"] = [_verdict(a, b_, p0, p1, t) for a, b_, p0, p1 in zip(dp.sku_prev, dp.sku_cur, dp.pb_prev, dp.pb_cur)]
    dp = dp.sort_values("sales", ascending=False).reset_index(drop=True)

    # ---- KPIs ----------------------------------------------------------------------------------------------
    sk0, sk1 = d.sku_prev.sum(), d.sku_cur.sum()
    U0, U1 = d.units_prev.sum(), d.units.sum()
    pbu0, pbu1 = d.pb_units_prev.sum(), d.pb_units.sum()
    nbu0, nbu1 = U0 - pbu0, U1 - pbu1
    nb0, nb1 = S0 - PB0, S1 - PB1
    kpi = dict(
        prev=prev, cur=cur, latest=latest,
        sales_prev=S0, sales_cur=S1, sales_yoy=S1 / S0 - 1,
        units_yoy=U1 / U0 - 1 if U0 else np.nan, asp_chg=(S1 / U1) / (S0 / U0) - 1 if U0 and U1 else np.nan,
        pb_prev=PB0, pb_cur=PB1, pb_yoy=PB1 / PB0 - 1 if PB0 else np.nan, d_pb=PB1 - PB0,
        nb_yoy=nb1 / nb0 - 1 if nb0 else np.nan,
        share_prev=R0, share_cur=R1,
        pb_units_yoy=pbu1 / pbu0 - 1 if pbu0 else np.nan, nb_units_yoy=nbu1 / nbu0 - 1 if nbu0 else np.nan,
        sku_prev=sk0, sku_cur=sk1, sku_growth=sk1 / sk0 - 1 if sk0 else np.nan,
        per_sku_prev=PB0 / sk0 if sk0 else np.nan, per_sku_cur=PB1 / sk1 if sk1 else np.nan,
        marg_per_sku=(PB1 - PB0) / (sk1 - sk0) if sk1 > sk0 else np.nan,
    )
    kpi["per_sku_chg"] = kpi["per_sku_cur"] / kpi["per_sku_prev"] - 1 if sk0 and sk1 else np.nan
    if latest:
        sl = df[f"sales_{latest}"].sum()
        kpi["share_latest"] = df[f"pb_sales_{latest}"].sum() / sl if sl else np.nan
        kpi["sales_latest"] = sl
        kpi["sku_latest"] = d.sku_latest.sum()
        kpi["pb_latest"] = float(df[f"pb_sales_{latest}"].sum())
        kpi["nb_latest"] = float(sl - kpi["pb_latest"])
        kpi["per_sku_latest"] = kpi["pb_latest"] / kpi["sku_latest"] if kpi["sku_latest"] else np.nan

    # ---- concentration -----------------------------------------------------------------------------------
    pbs = d.sort_values("pb_cur", ascending=False)
    shares = pbs.pb_cur / PB1 if PB1 else pbs.pb_cur * np.nan
    conc = dict(top3=float(shares.head(3).sum()), top5=float(shares.head(5).sum()), hhi=float((shares ** 2).sum()),
                top3_names=pbs["name"].head(3).tolist(), n_for_80=int((shares.cumsum() < 0.8).sum() + 1),
                n_active=int((pbs.pb_cur > 0).sum()))
    pbs_prev = d.sort_values("pb_prev", ascending=False)
    conc["top3_prev"] = float((pbs_prev.pb_prev / PB0).head(3).sum()) if PB0 else np.nan

    # ---- rankings ------------------------------------------------------------------------------------------
    base = d[d.pb_prev >= t.min_base_pb_sales]
    sk2 = d[(d.sku_cur >= 2) & (d.pb_cur >= t.min_base_pb_sales)]
    rankings = dict(
        sub_pb=d.sort_values("pb_cur", ascending=False).head(10).reset_index(drop=True),
        sub_share=d[d.pb_cur >= t.min_base_pb_sales].sort_values("share_cur", ascending=False).head(10).reset_index(drop=True),
        growth_top=base.sort_values("pb_growth", ascending=False).head(5).reset_index(drop=True),
        growth_bottom=base.sort_values("pb_growth").head(5).reset_index(drop=True),
        per_sku_top=sk2.sort_values("per_sku_cur", ascending=False).head(5).reset_index(drop=True),
        per_sku_bottom=sk2.sort_values("per_sku_cur").head(5).reset_index(drop=True),
    )

    # ---- momentum watchlist: PB share falling two steps in a row (needs the latest period) ----------------
    watch = pd.DataFrame()
    if latest:
        step1, step2 = d.share_cur - d.share_prev, d.share_latest - d.share_cur
        m = (step1 * 100 <= -t.watch_min_step_pt) & (step2 * 100 <= -t.watch_min_step_pt) & (d.pb_cur >= t.watch_min_pb_sales)
        watch = d[m].sort_values("pb_cur", ascending=False).reset_index(drop=True)

    # ---- upside -------------------------------------------------------------------------------------------
    rows = []
    for _, r in d.iterrows():
        tgt = None
        if r.direction == SCALE:
            tgt = min(t.upside_scale_cap, r.share_cur + t.upside_scale_momentum_years * (r.share_cur - r.share_prev))
            basis = T(f"勢い{t.upside_scale_momentum_years:g}年継続(上限{t.upside_scale_cap:.0%})",
                      f"momentum x{t.upside_scale_momentum_years:g} yrs (cap {t.upside_scale_cap:.0%})")
        elif r.direction == RECAPTURE:
            tgt, basis = r.share_prev, T("前期シェアに回復", "restore previous-period share")
        elif r.direction == REPLICATE:
            tgt, basis = t.upside_replicate_share, T(f"シェア{t.upside_replicate_share:.0%}", f"{t.upside_replicate_share:.0%} share")
        if tgt is not None and tgt > r.share_cur:
            rows.append(dict(name=r["name"], dept=r["dept"], direction=r.direction, sales=r.sales, share_cur=r.share_cur,
                             target=tgt, add=(tgt - r.share_cur) * r.sales, basis=basis))
    up = pd.DataFrame(rows, columns=["name", "dept", "direction", "sales", "share_cur", "target", "add", "basis"])
    up = up.sort_values("add", ascending=False).reset_index(drop=True)
    up_total = dict(add=float(up["add"].sum()), pct_of_pb=float(up["add"].sum() / PB1) if PB1 else np.nan,
                    share_after=float((PB1 + up["add"].sum()) / S1) if S1 else np.nan)

    # ---- reconciliation against the source's own totals ---------------------------------------------------
    recon_rows = []
    pcheck = [prev, cur] + ([latest] if latest else [])
    scopes = [(nm, df[df["dept"] == nm]) for nm in dict.fromkeys(df["dept"])]
    if data.total_name:
        scopes = [(data.total_name, df)] + scopes
    for nm, sub in scopes:
        rep = data.reported.get(nm)
        if not rep:
            continue
        for p in pcheck:
            rp = rep.get(p)
            if not rp:
                continue
            cs = float(sub[f"sales_{p}"].sum())
            cr = float(sub[f"pb_sales_{p}"].sum() / cs) if cs else np.nan
            rs, rr = rp.get("sales", np.nan), rp.get("pb_ratio_sales", np.nan)
            sd = (cs - rs) / rs if rs else np.nan
            rd = (cr - rr) * 100 if rr == rr else np.nan
            ok = (abs(sd) <= t.recon_sales_tolerance if sd == sd else True) and (abs(rd) <= t.recon_ratio_tolerance_pt if rd == rd else True)
            recon_rows.append(dict(scope=nm, period=p, rep_sales=rs, calc_sales=cs, sales_diff=sd, rep_ratio=rr,
                                   calc_ratio=cr, ratio_diff_pt=rd, ok=bool(ok)))
    recon = pd.DataFrame(recon_rows, columns=["scope", "period", "rep_sales", "calc_sales", "sales_diff", "rep_ratio",
                                              "calc_ratio", "ratio_diff_pt", "ok"])

    # ---- data-quality checks ------------------------------------------------------------------------------
    dq = []
    for w in data.warnings:
        dq.append(("warn", w))
    if data.category_only:
        covered = float(d.loc[d.source == "カテゴリー", "sales"].sum())
        cats_ = SEP.join(data.category_only)
        dq.append(("info", T(f"サブカテ行のないカテゴリーは「カテゴリー」シートから取り込み、全ての合計に含めました: "
                             f"{cats_}({cur}期売上 {J(covered)})。これらはカテゴリー単位の明細のみです。",
                             f"Categories with no sub-category rows were taken from the category sheet and included in all totals: "
                             f"{cats_} ({J(covered)} of {cur}期 sales). Their detail is category-level only.")))
    if len(recon):
        bad = recon[~recon.ok]
        if bad.empty:
            dq.append(("info", T(f"照合: 算出した売上とPB比率は、元資料の部門合計と全{recon.scope.nunique()}区分・"
                                 f"{recon.period.nunique()}期で一致しました。",
                                 f"Reconciliation: computed sales and PB ratios match the source's department totals for all "
                                 f"{recon.scope.nunique()} scopes and {recon.period.nunique()} periods.")))
        else:
            for _, r in bad.head(4).iterrows():
                dq.append(("warn", T(f"照合差異 {r.scope}({r.period}期): 元資料合計に対し売上 {r.sales_diff:+.2%}、PB比率 {r.ratio_diff_pt:+.2f}pt。",
                                     f"Reconciliation gap in {r.scope} ({r.period}期): sales {r.sales_diff:+.2%}, PB ratio {r.ratio_diff_pt:+.2f} pt vs the source total.")))
    elif not data.reported:
        dq.append(("info", T("「方向性」シートの部門合計が見つからないため、元資料との照合はできませんでした。",
                             "No 「方向性」 department totals were found, so the totals could not be reconciled with the source.")))
    bad_ratio = int(((df[[f"pb_ratio_sales_{p}" for p in periods]] > 1.0001) | (df[[f"pb_ratio_sales_{p}" for p in periods]] < 0)).sum().sum()
                    + ((df[[f"pb_ratio_units_{p}" for p in periods]] > 1.0001) | (df[[f"pb_ratio_units_{p}" for p in periods]] < 0)).sum().sum())
    neg = int((df[[f"sales_{p}" for p in periods] + [f"units_{p}" for p in periods]] < 0).sum().sum())
    if bad_ratio or neg:
        dq.append(("warn", T(f"不正な値: 0〜100%の範囲外のPB比率が{bad_ratio}件、売上・売れ数のマイナス値が{neg}件あります。",
                             f"Invalid values: {bad_ratio} PB ratios outside 0-100% and {neg} negative sales/units cells.")))
    dup = int(d.duplicated(["cat", "sub"]).sum())
    if dup:
        dq.append(("warn", T(f"カテゴリー/サブカテの重複行が{dup}件あります。元資料を確認してください。",
                             f"{dup} duplicated category/sub-category rows were found; check the source.")))
    ghost = d[(d.sku_cur > 0) & (d.pb_cur <= 0)]
    if len(ghost):
        ex = SEP.join(ghost["name"].head(4))
        dq.append(("info", T(f"PB SKUはあるが{cur}期のPB売上がないサブカテが{len(ghost)}件あります: {ex}。",
                             f"{len(ghost)} sub-categories list PB SKUs but have no PB sales in {cur}期: {ex}.")))
    nosku = d[(d.sku_cur <= 0) & (d.pb_cur > 0)]
    if len(nosku):
        ex = SEP.join(nosku["name"].head(4))
        dq.append(("warn", T(f"PB売上はあるがPB SKU数が0のサブカテが{len(nosku)}件あります: {ex}。",
                             f"{len(nosku)} sub-categories have PB sales but a PB SKU count of 0: {ex}.")))
    zero_units = d[((d.sales > 0) & (d.units <= 0)) | ((d.sales <= 0) & (d.units > 0))]
    if len(zero_units):
        ex = SEP.join(zero_units["name"].head(4))
        dq.append(("warn", T(f"売上があるのに売れ数がない(またはその逆の)行が{len(zero_units)}件あります: {ex}。",
                             f"{len(zero_units)} rows have sales without units (or units without sales): {ex}.")))
    asp_out = d[(d.asp_chg.abs() > t.asp_flag) & (d.sales >= t.min_headroom / 2)].sort_values("sales", ascending=False)
    if len(asp_out):
        examples = SEP.join(f"{r['name']} {r.asp_chg:+.0%}" for _, r in asp_out.head(3).iterrows())
        dq.append(("info", T(f"規模の大きいサブカテ{len(asp_out)}件で、平均単価が{t.asp_flag:.0%}超変動しています({examples})。"
                             "価格・容量・商品構成の変化によるものです。数量の伸びを読む前に売れ数の定義を確認してください。",
                             f"Average price per unit moved by more than {t.asp_flag:.0%} in {len(asp_out)} larger sub-categories "
                             f"({examples}). This is price or pack-size/mix change; check the unit definition before reading volume growth.")))
    gone = d[(d.yoy <= -0.5) & (d.sales_prev >= 5e6)]
    if len(gone):
        examples = SEP.join(f"{r['name']} {r.yoy:+.0%}" for _, r in gone.sort_values('sales_prev', ascending=False).head(3).iterrows())
        dq.append(("info", T(f"売上が前年比で半分以下になったサブカテが{len(gone)}件あります(例: {examples})。"
                             "小規模または終売の可能性があり、成長率は参考になりません。",
                             f"{len(gone)} sub-categories lost more than half of their sales year on year (e.g. {examples}): "
                             "small bases or discontinued ranges. Growth percentages there are not meaningful.")))

    # ---- sensitivity: which classifications flip when thresholds move? -------------------------------------
    base_lab = d.direction
    variants = [replace(t, min_headroom=t.min_headroom * 0.75), replace(t, min_headroom=t.min_headroom * 1.25),
                replace(t, scale_share_gain_pt=max(0.5, t.scale_share_gain_pt - 1)), replace(t, scale_share_gain_pt=t.scale_share_gain_pt + 1),
                replace(t, white_space_max_share=t.white_space_max_share * 0.6), replace(t, white_space_max_share=t.white_space_max_share * 1.6),
                replace(t, defend_min_share=t.defend_min_share - 0.1), replace(t, defend_min_share=t.defend_min_share + 0.1)]
    alts = [set() for _ in range(len(d))]
    for v in variants:
        lab = _classify_all(d, v)
        for i, (a, b) in enumerate(zip(base_lab, lab)):
            if a != b:
                alts[i].add(b)
    flip = [i for i, a in enumerate(alts) if a]
    bl = pd.DataFrame({"name": d["name"].iloc[flip].values, "base": base_lab.iloc[flip].values,
                       "alternatives": ["・".join(sorted(alts[i])) for i in flip],
                       "sales": d["sales"].iloc[flip].values})
    bl = bl[bl.base.isin([SCALE, REPLICATE, RECAPTURE, TEST, REVIEW, DEFEND])].sort_values("sales", ascending=False).reset_index(drop=True)

    # ---- caveats -------------------------------------------------------------------------------------------
    cav = []
    if latest:
        sl = kpi["sales_latest"]
        actual = sl / S1 if S1 else float("nan")
        cav.append(T(f"{latest}期は他の期より対象期間が短く、売上は{cur}期の{actual:.0%}です。売上・PB売上は他の期と比べられないため、"
                     f"増減額・成長率・増減要因・方向性の分析には使わず、PB比率・SKU数・シェア動向を{cur}期と並べて表示しています。",
                     f"The {latest} period covers a shorter span than the others (its sales are {actual:.0%} of {cur}期). Sales and PB sales cannot be compared "
                     f"with the other periods, so it is not used for changes, growth, drivers or directions; PB ratio, SKU counts and share trend are shown next to {cur}期."))
        cav.append(T(f"対象期間の異なる期のPBシェアは季節性の影響を受けるため、{latest}期と{cur}期のシェア比較は参考値です。",
                     f"PB share over a different span is affected by seasonality, so {latest} vs {cur}期 share changes are indicative, not like-for-like."))
    cav += [
        T("粗利・原価のデータはありません。本資料はすべて売上シェアに基づくため、PBシェアが高くても利益が高いとは限りません。",
          "There is no margin or cost data. Everything here is sales share, so a high PB share is not necessarily high profit."),
        T("平均単価には価格変更・容量変更・商品構成の変化が混在しています。数量と価格に関する記述は参考値です。",
          "Average price per unit mixes price changes, pack-size changes and product mix. Volume vs price statements are indicative only."),
        T("期中に追加したSKUは通年の売上実績がないため、SKU当たり売上は定常時の生産性を過小評価しています。"
          "「PB SKU数」が期中に販売実績のあるSKUのみを数えている場合、SKU希薄化の指摘もやや過大になります。",
          "SKUs added during the year have not had a full year of sales, so sales per SKU understates their steady-state productivity. "
          "If 「PB SKU数」 counts only SKUs that sold in the period, the dilution finding is also somewhat overstated."),
    ]

    directions = {k: d[d.direction == k].sort_values("nb_cur", ascending=False).reset_index(drop=True) for k in ORDER}

    # ---- headline findings (numbers computed above, wording stays neutral) ------------------------------------
    f = []
    top = c.sort_values("d_pb", ascending=False).iloc[0]
    sn = top.d_pb / bridge["d_pb"] if bridge["d_pb"] > 0 else np.nan
    has_sn = pd.notna(sn) and sn > 0
    f.append(T(f"PB売上は{J(PB0)}→{J(PB1)}({J(bridge['d_pb'])}、{kpi['pb_yoy']:+.1%})、PB比率は"
               f"{R0:.1%}→{R1:.1%}({(R1 - R0) * 100:+.2f}pt)。{top['cat']}の寄与は{J(top.d_pb)}"
               + (f"で、純増の約{sn:.0%}を占めます。" if has_sn else "です。"),
               f"PB sales moved {J(PB0)} → {J(PB1)} ({J(bridge['d_pb'])}, {kpi['pb_yoy']:+.1%}) and the PB ratio "
               f"{R0:.1%} → {R1:.1%} ({(R1 - R0) * 100:+.2f} pt). {top['cat']} contributed {J(top.d_pb)}"
               + (f", about {sn:.0%} of the net change." if has_sn else ".")))
    if pd.notna(kpi["per_sku_chg"]):
        s_ja = (f"PB SKU数は{sk0:.0f}→{sk1:.0f}({kpi['sku_growth']:+.0%})、PB売上は{kpi['pb_yoy']:+.0%}。PB SKU当たり売上は"
                f"{J(kpi['per_sku_prev'])}→{J(kpi['per_sku_cur'])}({kpi['per_sku_chg']:+.0%})")
        s_en = (f"PB SKUs went {sk0:.0f} → {sk1:.0f} ({kpi['sku_growth']:+.0%}) while PB sales moved {kpi['pb_yoy']:+.0%}; sales per PB SKU "
                f"{J(kpi['per_sku_prev'])} → {J(kpi['per_sku_cur'])} ({kpi['per_sku_chg']:+.0%})")
        if pd.notna(kpi["marg_per_sku"]):
            s_ja += f"。追加SKU1つ当たりのPB売上増は平均{J(kpi['marg_per_sku'])}"
            s_en += f"; each added SKU brought {J(kpi['marg_per_sku'])} of extra PB sales on average"
        if latest and pd.notna(kpi.get("sku_latest")):
            s_ja += f"。{latest}期のPB SKU数は{kpi['sku_latest']:.0f}"
            s_en += f". The PB SKU count in {latest} is {kpi['sku_latest']:.0f}"
        f.append(T(s_ja + "です。", s_en + "."))
    f.append(T(f"PB売上は集中しています。上位3サブカテ({'、'.join(conc['top3_names'])})でPB売上の{conc['top3']:.0%}、"
               f"{conc['n_for_80']}サブカテで80%を占めます。",
               f"PB sales are concentrated: the top 3 sub-categories ({', '.join(conc['top3_names'])}) are {conc['top3']:.0%} of PB sales "
               f"and {conc['n_for_80']} sub-categories make up 80%."))

    # ---- per-department takeaways ---------------------------------------------------------------------------
    PC = lambda x: f"{x:.1%}" if pd.notna(x) else T("－", "–")
    dept_findings = {}
    for _, dr in dp.iterrows():
        sub_d = d[d["dept"] == dr["dept"]]
        fl = []
        mv_ = sub_d.sort_values("d_pb", ascending=False)
        g_, l_ = mv_.iloc[0], mv_.iloc[-1]
        pj, pe = [], []
        if g_.d_pb > 0:
            pj.append(f"最大の増加は{g_['name']}({J(g_.d_pb)})")
            pe.append(f"largest gain {g_['name']} ({J(g_.d_pb)})")
        if l_.d_pb < 0:
            pj.append(f"最大の減少は{l_['name']}({J(l_.d_pb)})")
            pe.append(f"largest decline {l_['name']} ({J(l_.d_pb)})")
        if pj:
            fl.append(T("PB売上の動き: " + "、".join(pj) + "です。", "PB sales by sub-category: " + "; ".join(pe) + "."))
        if latest and pd.notna(dr.share_latest):
            dpt = (dr.share_latest - dr.share_cur) * 100
            fl.append(T(f"{latest}期: PB売上 {J(dr.pb_latest)}、PB比率 {PC(dr.share_latest)}({cur}期比 {dpt:+.2f}pt)。"
                        f"{latest}期は他の期より対象期間が短いため、売上の増減は比較していません。",
                        f"{latest}期: PB sales {J(dr.pb_latest)}, PB ratio {PC(dr.share_latest)} ({dpt:+.2f} pt vs {cur}期). "
                        f"It covers a shorter span than the other periods, so changes in sales are not compared."))
        tg = sub_d[sub_d["trend"] != ""] if latest else sub_d.iloc[0:0]
        if len(tg):
            ups = tg[tg.trend == T_UP].sort_values("mom_pt", ascending=False).head(2)
            downs = tg[tg.trend == T_DOWN].sort_values("mom_pt").head(2)
            mj, me = [], []
            if len(ups):
                mj.append("上昇 " + "、".join(f"{r['name']}({r.mom_pt:+.1f}pt)" for _, r in ups.iterrows()))
                me.append("rising: " + ", ".join(f"{r['name']} ({r.mom_pt:+.1f} pt)" for _, r in ups.iterrows()))
            if len(downs):
                mj.append("低下 " + "、".join(f"{r['name']}({r.mom_pt:+.1f}pt)" for _, r in downs.iterrows()))
                me.append("falling: " + ", ".join(f"{r['name']} ({r.mom_pt:+.1f} pt)" for _, r in downs.iterrows()))
            if mj:
                fl.append(T(f"PBシェアの最新の動き({latest}期、{cur}期比): " + " / ".join(mj) + "。",
                            f"PB share momentum ({latest}期 vs {cur}期): " + "; ".join(me) + "."))
            warn = tg[(tg.direction == SCALE) & (tg.trend == T_DOWN)]
            if len(warn):
                fl.append(T(f"要注意: 「{SCALE}」の対象ですが、直近のPBシェアが低下しています: " + "、".join(warn["name"]) + "。",
                            f"Watch: classified as {tr(SCALE, 'en')}, but PB share has slipped recently: " + ", ".join(warn["name"]) + "."))
        pool_ = sub_d.sort_values("nb_cur", ascending=False).iloc[0]
        if pool_.nb_cur > 0:
            fl.append(T(f"非PB売上が最も大きいのは{pool_['name']}({J(pool_.nb_cur)}、PB比率{PC(pool_.share_cur)})です。",
                        f"Largest non-PB pool: {pool_['name']} ({J(pool_.nb_cur)}, PB ratio {PC(pool_.share_cur)})."))
        ud = float(up.loc[up["dept"] == dr["dept"], "add"].sum()) if len(up) else 0.0
        if ud > 0:
            fl.append(T(f"伸びしろの試算(仮定に基づく): {SCALE}・{RECAPTURE}・{REPLICATE}の合計で約{J(ud)}。",
                        f"Illustrative upside (assumption-based): about {J(ud)} across Scale, Recapture and Replicate items."))
        if len(watch):
            wn = watch.loc[watch["dept"] == dr["dept"], "name"].tolist()
            if wn:
                fl.append(T("要注視(PBシェアが2期連続で低下): " + "、".join(wn) + "。",
                            "Watch (PB share down two periods in a row): " + ", ".join(wn) + "."))
        dept_findings[dr["dept"]] = fl

    cat_only = SEP.join(data.category_only)
    scope = T(f"対象: 全{len(d)}サブカテ" + (f"(カテゴリー単位のみの{cat_only}を含む)" if data.category_only else "") + "。",
              f"Scope: all {len(d)} sub-categories" + (f" (incl. category-level only: {cat_only})" if data.category_only else "") + ".")
    return Result((prev, cur), latest, data.partial_months, kpi, d, c, directions, flags, up, up_total, cav, t, f,
                  bridge, rankings, watch, conc, recon, dq, bl, scope, lang, depts=dp, dept_findings=dept_findings)


def key_notes(res: Result) -> list:
    """The few assumptions a reader needs to know. Everything else (data checks, rules) stays off the report."""
    ja = res.lang == "ja"
    notes = []
    if res.period_latest:
        cf = res.periods_full[1]
        sh_ = res.kpi["sales_latest"] / res.kpi["sales_cur"] if res.kpi["sales_cur"] else float("nan")
        notes.append(f"{res.period_latest}期は他の期より対象期間が短く(売上は{cf}期の{sh_:.0%})、売上・PB売上は他の期と比べられません。"
                     f"PB比率・SKU数・シェア動向のみ{cf}期と並べて表示しており、成長率や方向性の判定には使っていません。" if ja else
                     f"{res.period_latest}期 covers a shorter span than the others (its sales are {sh_:.0%} of {cf}期), so sales and PB sales cannot be compared with the other periods. "
                     f"Only PB ratio, SKU counts and share trend are shown next to {cf}期; it is not used for growth or directions.")
    notes.append("粗利・原価のデータは含まれていません。PBシェアが高くても、利益が高いとは限りません。" if ja else
                 "No margin or cost data is included. A high PB share does not necessarily mean high profit.")
    notes.append("平均単価の変化には、価格・容量・商品構成の変化が含まれます。数量に関する記述は参考値です。" if ja else
                 "Average price per unit reflects price, pack-size and product-mix changes, so statements about volume are indicative.")
    return notes
