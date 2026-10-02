"""Turns the tidy table into KPIs, PB-ratio-vs-SKU verdicts, direction buckets and checks.

Everything is deterministic and every threshold lives in `Thresholds`, so the same file always produces the same
report and the rules can be shown (and debated) in the appendix.

The workbook has no PB sales column, so PB sales are NOT derived (sales x PB ratio) anywhere in the report. Every
figure is the workbook's own value at its own level: sub-categories from the 「サブカテ」 sheet, categories from
「カテゴリー」, departments and the company from 「方向性」. Only when a level has no figure of its own in the workbook
(e.g. a file without those sheets) is the PB ratio taken as the sales-weighted average of the rows below it, and the
data checks say so. `recon` compares the sub-category rows with those figures on every run.
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

V_EFFICIENT, V_SKULED, V_SKUDOWNSALES = "PB比率がSKU増に追随", "SKU増がPB比率上昇を上回る", "SKU増・PB比率低下"
V_PRUNED, V_SHRINK = "SKU削減・PB比率上昇", "SKU削減・PB比率低下"
V_STABLE_UP, V_STABLE_DOWN = "SKU横ばい・PB比率上昇", "SKU横ばい・PB比率低下"
V_NEW, V_EXIT, V_SMALL = "PB新規投入", "PB撤退", "小規模(判定外)"
T_UP, T_FLAT, T_DOWN = "上昇", "横ばい", "低下"      # PB share trend: newest (part-year) period vs the last full period

# Internal keys stay Japanese; `tr` translates them for display only.
_EN = {
    SCALE: "Scale", REPLICATE: "Replicate", RECAPTURE: "Recapture", TEST: "Test small", REVIEW: "Review",
    DEFEND: "Defend", DEPRIORITIZE: "Deprioritize", MONITOR: "Monitor",
    V_EFFICIENT: "PB ratio keeps pace with SKUs", V_SKULED: "SKUs outpace PB ratio gain", V_SKUDOWNSALES: "SKUs up, PB ratio down",
    V_PRUNED: "Pruned, PB ratio up", V_SHRINK: "Pruned, PB ratio down", V_STABLE_UP: "Stable range, PB ratio up",
    V_STABLE_DOWN: "Stable range, PB ratio down", V_NEW: "New PB range", V_EXIT: "PB range exited", V_SMALL: "Small base",
    T_UP: "Rising", T_FLAT: "Flat", T_DOWN: "Falling",
}


def tr(key, lang="ja"):
    """Display name of a direction or verdict in the requested language."""
    return key if lang == "ja" else _EN.get(key, key)


@dataclass
class Thresholds:
    min_headroom: float = 200e6          # total sales (JPY) needed to call a market "big"
    scale_share_gain_pt: float = 2.0     # PB ratio gain (pt) vs previous period to count as momentum
    scale_max_share: float = 0.70        # PB ratio must still be below this for a market to have room left
    scale_min_market_yoy: float = -0.03  # market must not be shrinking faster than this
    white_space_max_share: float = 0.05  # PB ratio at or below this = "no PB position"
    test_min_market_yoy: float = -0.03   # white-space markets shrinking faster than this are not Replicate/Test candidates
    recapture_drop_pt: float = 5.0       # PB ratio loss (pt) that counts as a collapse
    review_drop_pt: float = 2.0          # smaller PB ratio loss that deserves a look
    review_min_sales: float = 100e6      # previous-period sales needed for a "Review" flag
    defend_min_share: float = 0.50       # PB ratio at/above this = saturated
    deprioritize_max_sales: float = 200e6
    deprioritize_market_yoy: float = -0.30
    deprioritize_max_growth: float = 0.20  # a small no-PB market growing faster than this stays on watch
    sku_growth_flag: float = 0.50        # PB SKU count growth that triggers a dilution check
    sku_stable_band: float = 0.10        # SKU change within +/- this is "stable range"
    min_share: float = 0.02              # minimum PB ratio (base period) for SKU verdicts / (current) for share-trend tags
    min_sales: float = 100e6             # minimum total sales for SKU verdicts, share-trend tags and "largest mover" lists
    watch_min_sales: float = 200e6       # minimum total sales for the two-step PB-ratio-decline watchlist
    watch_min_step_pt: float = 0.5       # each step must fall by at least this many pt
    recon_sales_tolerance: float = 0.002   # reconciliation: allowed relative sales difference
    recon_ratio_tolerance_pt: float = 0.10  # reconciliation: allowed PB-ratio difference in pt
    asp_flag: float = 0.30               # price/mix change (avg yen per unit) that deserves a data check
    momentum_pt: float = 2.0             # PB ratio move (pt), newest part-year period vs last full period, that counts as rising / falling


@dataclass
class Result:
    periods_full: tuple
    period_latest: str | None
    partial_months: int | None
    kpi: dict
    subs: pd.DataFrame
    cats: pd.DataFrame
    directions: dict
    caveats: list
    thresholds: Thresholds
    findings: list = field(default_factory=list)
    watchlist: pd.DataFrame = None
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


def _verdict(sku_prev, sku_cur, share_prev, share_cur, sales_prev, t: Thresholds) -> str:
    """Compares the change in PB SKU count with the change in PB ratio (both are the workbook's own figures)."""
    if sku_prev <= 0 and sku_cur > 0:
        return V_NEW
    if sku_cur <= 0 and sku_prev > 0:
        return V_EXIT
    if sku_prev <= 0 or share_prev < t.min_share or sales_prev < t.min_sales:
        return V_SMALL
    sg, pg = sku_cur / sku_prev - 1, share_cur / share_prev - 1
    if sg > t.sku_stable_band:
        return V_EFFICIENT if pg >= sg else (V_SKULED if pg >= 0 else V_SKUDOWNSALES)
    if sg < -t.sku_stable_band:
        return V_PRUNED if pg >= 0 else V_SHRINK
    return V_STABLE_UP if pg >= 0 else V_STABLE_DOWN


def _classify_all(d: pd.DataFrame, t: Thresholds) -> pd.Series:
    scale_cats = set(d.loc[(d.d_share_pt >= t.scale_share_gain_pt) & (d.sales >= t.min_headroom)
                           & (d.share_cur < t.scale_max_share) & (d.yoy > t.scale_min_market_yoy), "cat"])

    def one(r):
        if r.share_prev > 0 and r.d_share_pt <= -t.recapture_drop_pt and r.sales_prev >= t.min_headroom:
            return RECAPTURE
        if r.d_share_pt <= -t.review_drop_pt and r.sales_prev >= t.review_min_sales:
            return REVIEW
        if r.d_share_pt >= t.scale_share_gain_pt and r.sales >= t.min_headroom \
                and r.share_cur < t.scale_max_share and r.yoy > t.scale_min_market_yoy:
            return SCALE
        if r.share_cur >= t.defend_min_share and r.sales >= t.min_headroom:
            return DEFEND
        if r.share_cur <= t.white_space_max_share and r.sales >= t.min_headroom \
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
def _wavg(ratio, weight) -> float:
    """Sales-weighted average of a ratio. Fallback only, for a level the workbook gives no figure for."""
    w = float(weight.sum())
    return float((ratio * weight).sum() / w) if w else 0.0


def _own(e: dict, key: str):
    v = e.get(key)
    return None if v is None or v != v else float(v)


def _level(sub: pd.DataFrame, ext: dict | None, periods: list):
    """Sales, units, PB ratios and PB SKU count per period for one department / category / the company.

    Uses the workbook's own figure for the level (`ext`) wherever it has one; otherwise adds up the rows below it
    (sales, units, SKUs) or takes the sales-weighted average (PB ratios). Returns (metrics, used_fallback_for_ratio).
    """
    out, fallback = {}, False
    for p in periods:
        e = (ext or {}).get(p) or {}
        s, u, k = _own(e, "sales"), _own(e, "units"), _own(e, "pb_skus")
        rs, ru = _own(e, "pb_ratio_sales"), _own(e, "pb_ratio_units")
        if s is None:
            s = float(sub[f"sales_{p}"].sum())
        if u is None:
            u = float(sub[f"units_{p}"].sum())
        if k is None:
            k = float(sub[f"pb_skus_{p}"].sum())
        if rs is None:
            rs, fallback = _wavg(sub[f"pb_ratio_sales_{p}"], sub[f"sales_{p}"]), True
        if ru is None:
            ru, fallback = _wavg(sub[f"pb_ratio_units_{p}"], sub[f"units_{p}"]), True
        out[p] = dict(sales=s, units=u, rs=rs, ru=ru, sku=k)
    return out, fallback


def _level_row(m: dict, prev: str, cur: str, latest: str | None, extra: dict) -> dict:
    a, b = m[prev], m[cur]
    r = dict(extra)
    r.update(sales_prev=a["sales"], sales=b["sales"], units_prev=a["units"], units=b["units"],
             share_prev=a["rs"], share_cur=b["rs"], share_units_prev=a["ru"], share_units_cur=b["ru"],
             sku_prev=a["sku"], sku_cur=b["sku"])
    if latest:
        c = m[latest]
        r.update(sales_latest=c["sales"], share_latest=c["rs"] if c["sales"] > 0 else np.nan, sku_latest=c["sku"])
    else:
        r.update(sales_latest=np.nan, share_latest=np.nan, sku_latest=np.nan)
    return r


def _derive(x: pd.DataFrame, t: Thresholds) -> pd.DataFrame:
    """Growth and PB-ratio changes. Only the workbook's sales, ratios and SKU counts go in; no PB sales anywhere."""
    x["yoy"] = _div(x.sales - x.sales_prev, x.sales_prev)
    x["units_yoy"] = _div(x.units - x.units_prev, x.units_prev)
    x["asp"], x["asp_prev"] = _div(x.sales, x.units), _div(x.sales_prev, x.units_prev)
    x["asp_chg"] = _div(x.asp - x.asp_prev, x.asp_prev)
    x["d_share_pt"] = (x.share_cur - x.share_prev) * 100
    x["share_growth"] = _div(x.share_cur - x.share_prev, x.share_prev)
    x["sku_growth"] = _div(x.sku_cur - x.sku_prev, x.sku_prev)
    x["verdict"] = [_verdict(a, b, c, e, f, t) for a, b, c, e, f in
                    zip(x.sku_prev, x.sku_cur, x.share_prev, x.share_cur, x.sales_prev)]
    return x


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

    # ---- sub-category level: the workbook's own columns, as they are ----------------------------------------
    d = pd.DataFrame({"dept": df["dept"], "cat": df["cat"], "sub": df["sub"], "source": df["source"]})
    d["name"] = np.where(d["sub"] == d["cat"], d["cat"], d["cat"] + " / " + d["sub"])
    d["sales"], d["sales_prev"] = df[f"sales_{cur}"], df[f"sales_{prev}"]
    d["units"], d["units_prev"] = df[f"units_{cur}"], df[f"units_{prev}"]
    d["share_prev"], d["share_cur"] = df[f"pb_ratio_sales_{prev}"], df[f"pb_ratio_sales_{cur}"]
    d["share_units_prev"], d["share_units_cur"] = df[f"pb_ratio_units_{prev}"], df[f"pb_ratio_units_{cur}"]
    d["sku_prev"], d["sku_cur"] = df[f"pb_skus_{prev}"], df[f"pb_skus_{cur}"]
    d["sales_latest"] = df[f"sales_{latest}"] if latest else np.nan
    d["share_latest"] = df[f"pb_ratio_sales_{latest}"] if latest else np.nan
    d["sku_latest"] = df[f"pb_skus_{latest}"] if latest else np.nan
    if latest:
        d.loc[df[f"sales_{latest}"] <= 0, "share_latest"] = np.nan
    d = _derive(d, t)
    d["mom_pt"] = ((d.share_latest - d.share_cur) * 100).round(1)         # PB ratio move in points (rounded as shown), newest vs last full period
    sized = (d.sales >= t.min_sales) & (d.share_cur >= t.min_share)       # tiny bases are not tagged
    d["trend"] = np.where(~sized | d.mom_pt.isna(), "",
                          np.where(d.mom_pt >= t.momentum_pt, T_UP, np.where(d.mom_pt <= -t.momentum_pt, T_DOWN, T_FLAT)))
    d["direction"] = _classify_all(d, t)

    # ---- category / department / company level: the workbook's own figures --------------------------------
    fell_back = []
    keys = list(dict.fromkeys(zip(df["dept"], df["cat"])))
    crows = []
    for dp_, ct_ in keys:
        m, fb = _level(df[(df["dept"] == dp_) & (df["cat"] == ct_)], data.cat_reported.get((dp_, ct_)), periods)
        if fb:
            fell_back.append(ct_)
        crows.append(_level_row(m, prev, cur, latest, dict(dept=dp_, cat=ct_)))
    c = _derive(pd.DataFrame(crows), t)
    c["rank_dept"] = c.groupby("dept").sales.rank(ascending=False, method="min").astype(int)   # rank inside the department
    c = c.sort_values("sales", ascending=False).reset_index(drop=True)

    drows = []
    for dp_ in dict.fromkeys(df["dept"]):
        sub = df[df["dept"] == dp_]
        m, fb = _level(sub, data.reported.get(dp_), periods)
        if fb:
            fell_back.append(dp_)
        drows.append(_level_row(m, prev, cur, latest, dict(dept=dp_, cats=sub["cat"].nunique(), subs=len(sub))))
    dp = _derive(pd.DataFrame(drows), t).sort_values("sales", ascending=False).reset_index(drop=True)

    tm, fb = _level(df, data.reported.get(data.total_name) if data.total_name else None, periods)
    if fb:
        fell_back.append(T("ペット計", "company total"))
    tr_ = _derive(pd.DataFrame([_level_row(tm, prev, cur, latest, {})]), t).iloc[0]
    S0, S1, R0, R1 = tr_.sales_prev, tr_.sales, tr_.share_prev, tr_.share_cur
    sk0, sk1 = tr_.sku_prev, tr_.sku_cur
    kpi = dict(
        prev=prev, cur=cur, latest=latest,
        sales_prev=S0, sales_cur=S1, sales_yoy=tr_.yoy, units_yoy=tr_.units_yoy, asp_chg=tr_.asp_chg,
        share_prev=R0, share_cur=R1, share_units_prev=tr_.share_units_prev, share_units_cur=tr_.share_units_cur,
        sku_prev=sk0, sku_cur=sk1, sku_growth=tr_.sku_growth, share_growth=tr_.share_growth,
    )
    if latest:
        kpi.update(share_latest=tr_.share_latest, sales_latest=tr_.sales_latest, sku_latest=tr_.sku_latest,
                   share_units_latest=tm[latest]["ru"])

    # ---- momentum watchlist: PB ratio falling two steps in a row (needs the latest period) ----------------
    watch = pd.DataFrame()
    if latest:
        step1, step2 = d.share_cur - d.share_prev, d.share_latest - d.share_cur
        m = (step1 * 100 <= -t.watch_min_step_pt) & (step2 * 100 <= -t.watch_min_step_pt) & (d.sales >= t.watch_min_sales)
        watch = d[m].sort_values("sales", ascending=False).reset_index(drop=True)

    # ---- reconciliation: the sub-category rows against the workbook's own category / department / total figures
    recon_rows = []
    pcheck = [prev, cur] + ([latest] if latest else [])
    scopes = [(nm, df[df["dept"] == nm], data.reported.get(nm), False) for nm in dict.fromkeys(df["dept"])]
    scopes += [(ct_, df[(df["dept"] == dp_) & (df["cat"] == ct_) & (df["source"] == "サブカテ")], data.cat_reported.get((dp_, ct_)), True)
               for dp_, ct_ in keys]
    if data.total_name:
        scopes = [(data.total_name, df, data.reported.get(data.total_name), False)] + scopes
    for nm, sub, rep, with_sku in scopes:
        if not rep or sub.empty:
            continue
        for p in pcheck:
            rp = rep.get(p)
            if not rp:
                continue
            cs = float(sub[f"sales_{p}"].sum())
            cr = _wavg(sub[f"pb_ratio_sales_{p}"], sub[f"sales_{p}"]) if cs else np.nan
            rs, rr = rp.get("sales", np.nan), rp.get("pb_ratio_sales", np.nan)
            sd = (cs - rs) / rs if rs else np.nan
            rd = (cr - rr) * 100 if rr == rr else np.nan
            rk = rp.get("pb_skus", np.nan) if with_sku else np.nan
            ck = float(sub[f"pb_skus_{p}"].sum()) if with_sku else np.nan
            ok = ((abs(sd) <= t.recon_sales_tolerance if sd == sd else True) and (abs(rd) <= t.recon_ratio_tolerance_pt if rd == rd else True)
                  and (rk == ck if rk == rk else True))
            recon_rows.append(dict(scope=nm, period=p, rep_sales=rs, calc_sales=cs, sales_diff=sd, rep_ratio=rr,
                                   calc_ratio=cr, ratio_diff_pt=rd, rep_sku=rk, calc_sku=ck, ok=bool(ok)))
    recon = pd.DataFrame(recon_rows, columns=["scope", "period", "rep_sales", "calc_sales", "sales_diff", "rep_ratio",
                                              "calc_ratio", "ratio_diff_pt", "rep_sku", "calc_sku", "ok"])

    # ---- data-quality checks ------------------------------------------------------------------------------
    dq = []
    for w in data.warnings:
        dq.append(("warn", w))
    dq.append(("info", T("元資料にPB売上額の列がないため、PB売上額は算出・表示していません。PB比率とPB SKU数は元資料の値をそのまま使っています。",
                         "The workbook has no PB sales column, so PB sales are not derived or shown. PB ratios and PB SKU counts are the workbook's own values.")))
    if fell_back:
        dq.append(("warn", T("元資料に値のない区分は、配下のサブカテの売上加重平均でPB比率を算出しました: " + SEP.join(fell_back) + "。",
                             "These levels have no figure of their own in the workbook, so their PB ratio is the sales-weighted average of the rows below: "
                             + SEP.join(fell_back) + ".")))
    if data.category_only:
        covered = float(d.loc[d.source == "カテゴリー", "sales"].sum())
        cats_ = SEP.join(data.category_only)
        dq.append(("info", T(f"サブカテ行のないカテゴリーは「カテゴリー」シートから取り込みました: "
                             f"{cats_}({cur}期売上 {J(covered)})。これらはカテゴリー単位の明細のみです。",
                             f"Categories with no sub-category rows were taken from the category sheet: "
                             f"{cats_} ({J(covered)} of {cur}期 sales). Their detail is category-level only.")))
    if len(recon):
        bad = recon[~recon.ok]
        if bad.empty:
            dq.append(("info", T(f"照合: サブカテ行の売上・PB比率・PB SKU数は、元資料の部門・カテゴリー・全体の値と全{recon.scope.nunique()}区分・"
                                 f"{recon.period.nunique()}期で一致しました。",
                                 f"Reconciliation: the sub-category rows match the workbook's own department, category and total figures for all "
                                 f"{recon.scope.nunique()} scopes and {recon.period.nunique()} periods.")))
        else:
            for _, r in bad.head(4).iterrows():
                dq.append(("warn", T(f"照合差異 {r.scope}({r.period}期): 元資料の値に対し売上 {r.sales_diff:+.2%}、PB比率 {r.ratio_diff_pt:+.2f}pt。",
                                     f"Reconciliation gap in {r.scope} ({r.period}期): sales {r.sales_diff:+.2%}, PB ratio {r.ratio_diff_pt:+.2f} pt vs the workbook's figure.")))
    elif not data.reported and not data.cat_reported:
        dq.append(("info", T("「方向性」「カテゴリー」シートの値が見つからないため、元資料との照合はできませんでした。",
                             "No 「方向性」 or 「カテゴリー」 figures were found, so the rows could not be reconciled.")))
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
    ghost = d[(d.sku_cur > 0) & (d.share_cur <= 0)]
    if len(ghost):
        ex = SEP.join(ghost["name"].head(4))
        dq.append(("info", T(f"PB SKUはあるが{cur}期のPB比率が0のサブカテが{len(ghost)}件あります: {ex}。",
                             f"{len(ghost)} sub-categories list PB SKUs but have a PB ratio of 0 in {cur}期: {ex}.")))
    nosku = d[(d.sku_cur <= 0) & (d.share_cur > 0)]
    if len(nosku):
        ex = SEP.join(nosku["name"].head(4))
        dq.append(("warn", T(f"PB比率はあるがPB SKU数が0のサブカテが{len(nosku)}件あります: {ex}。",
                             f"{len(nosku)} sub-categories have a PB ratio above 0 but a PB SKU count of 0: {ex}.")))
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
        cav.append(T(f"{latest}期は他の期より対象期間が短く、売上は{cur}期の{actual:.0%}です。売上は他の期と比べられないため、"
                     f"増減額・成長率・方向性の分析には使わず、PB比率・SKU数・シェア動向を{cur}期と並べて表示しています。",
                     f"The {latest} period covers a shorter span than the others (its sales are {actual:.0%} of {cur}期). Sales cannot be compared "
                     f"with the other periods, so it is not used for changes, growth or directions; PB ratio, SKU counts and share trend are shown next to {cur}期."))
        cav.append(T(f"対象期間の異なる期のPB比率は季節性の影響を受けるため、{latest}期と{cur}期のPB比率の比較は参考値です。",
                     f"PB ratio over a different span is affected by seasonality, so {latest} vs {cur}期 PB ratio changes are indicative, not like-for-like."))
    cav += [
        T("PB売上額は元資料にないため、PB比率(売上に占めるPBの割合)をそのまま使っています。PB売上額の増減や、SKU当たりのPB売上は算出していません。",
          "The workbook has no PB sales amounts, so the PB ratio (PB share of sales) is used as given. Changes in PB sales and PB sales per SKU are not calculated."),
        T("粗利・原価のデータはありません。本資料はすべて売上シェアに基づくため、PB比率が高くても利益が高いとは限りません。",
          "There is no margin or cost data. Everything here is sales share, so a high PB ratio is not necessarily high profit."),
        T("平均単価には価格変更・容量変更・商品構成の変化が混在しています。数量と価格に関する記述は参考値です。",
          "Average price per unit mixes price changes, pack-size changes and product mix. Volume vs price statements are indicative only."),
    ]

    directions = {k: d[d.direction == k].sort_values("sales", ascending=False).reset_index(drop=True) for k in ORDER}

    # ---- headline findings (numbers are the workbook's own; wording stays neutral) -------------------------
    f = []
    cs_ = c[c.sales_prev >= t.min_sales]
    top = cs_.sort_values("d_share_pt", ascending=False).iloc[0] if len(cs_) else None
    f.append(T(f"ペット全体の売上は{J(S0)}→{J(S1)}({tr_.yoy:+.1%})、PB比率は{R0:.1%}→{R1:.1%}({(R1 - R0) * 100:+.2f}pt)。"
               + (f"PB比率の上昇が最も大きいカテゴリーは{top['cat']}({top.share_prev:.1%}→{top.share_cur:.1%}、{top.d_share_pt:+.1f}pt)です。"
                  if top is not None and top.d_share_pt > 0 else ""),
               f"Total sales moved {J(S0)} → {J(S1)} ({tr_.yoy:+.1%}) and the PB ratio {R0:.1%} → {R1:.1%} ({(R1 - R0) * 100:+.2f} pt)."
               + (f" The largest PB ratio gain by category is {top['cat']} ({top.share_prev:.1%} → {top.share_cur:.1%}, {top.d_share_pt:+.1f} pt)."
                  if top is not None and top.d_share_pt > 0 else "")))
    if pd.notna(tr_.sku_growth):
        s_ja = f"PB SKU数は{sk0:.0f}→{sk1:.0f}({tr_.sku_growth:+.0%})、PB比率は{R0:.1%}→{R1:.1%}({(R1 - R0) * 100:+.2f}pt)"
        s_en = f"PB SKUs went {sk0:.0f} → {sk1:.0f} ({tr_.sku_growth:+.0%}) while the PB ratio moved {R0:.1%} → {R1:.1%} ({(R1 - R0) * 100:+.2f} pt)"
        if latest and pd.notna(kpi.get("sku_latest")):
            s_ja += f"。{latest}期のPB SKU数は{kpi['sku_latest']:.0f}"
            s_en += f". The PB SKU count in {latest} is {kpi['sku_latest']:.0f}"
        f.append(T(s_ja + "です。", s_en + "."))
    hi, lo = dp.sort_values("share_cur").iloc[-1], dp.sort_values("share_cur").iloc[0]
    f.append(T(f"部門によってPB比率は大きく異なります。最も高いのは{hi['dept']}({hi.share_cur:.1%})、最も低いのは{lo['dept']}({lo.share_cur:.1%})です。",
               f"PB ratios differ widely by department: highest {hi['dept']} ({hi.share_cur:.1%}), lowest {lo['dept']} ({lo.share_cur:.1%})."))

    # ---- per-department takeaways ---------------------------------------------------------------------------
    PC = lambda x: f"{x:.1%}" if pd.notna(x) else T("－", "–")
    dept_findings = {}
    for _, dr in dp.iterrows():
        sub_d = d[d["dept"] == dr["dept"]]
        fl = []
        mv_ = sub_d[sub_d.sales >= t.min_sales].sort_values("d_share_pt", ascending=False)
        pj, pe = [], []
        if len(mv_):
            g_, l_ = mv_.iloc[0], mv_.iloc[-1]
            if g_.d_share_pt >= 0.5:
                pj.append(f"PB比率の上昇が最大: {g_['name']}({PC(g_.share_prev)}→{PC(g_.share_cur)}、{g_.d_share_pt:+.1f}pt)")
                pe.append(f"largest PB ratio gain {g_['name']} ({PC(g_.share_prev)} → {PC(g_.share_cur)}, {g_.d_share_pt:+.1f} pt)")
            if l_.d_share_pt <= -0.5:
                pj.append(f"PB比率の低下が最大: {l_['name']}({PC(l_.share_prev)}→{PC(l_.share_cur)}、{l_.d_share_pt:+.1f}pt)")
                pe.append(f"largest PB ratio drop {l_['name']} ({PC(l_.share_prev)} → {PC(l_.share_cur)}, {l_.d_share_pt:+.1f} pt)")
        if pj:
            fl.append(T("PB比率の動き: " + "、".join(pj) + "です。", "PB ratio by sub-category: " + "; ".join(pe) + "."))
        if latest and pd.notna(dr.share_latest):
            dpt = (dr.share_latest - dr.share_cur) * 100
            fl.append(T(f"{latest}期: PB比率 {PC(dr.share_latest)}({cur}期比 {dpt:+.2f}pt)。"
                        f"{latest}期は他の期より対象期間が短いため、売上の増減は比較していません。",
                        f"{latest}期: PB ratio {PC(dr.share_latest)} ({dpt:+.2f} pt vs {cur}期). "
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
                fl.append(T(f"PB比率の最新の動き({latest}期、{cur}期比): " + " / ".join(mj) + "。",
                            f"PB ratio momentum ({latest}期 vs {cur}期): " + "; ".join(me) + "."))
            warn = tg[(tg.direction == SCALE) & (tg.trend == T_DOWN)]
            if len(warn):
                fl.append(T(f"要注意: 「{SCALE}」の対象ですが、直近のPB比率が低下しています: " + "、".join(warn["name"]) + "。",
                            f"Watch: classified as {tr(SCALE, 'en')}, but PB ratio has slipped recently: " + ", ".join(warn["name"]) + "."))
        pools_ = sub_d[sub_d.share_cur <= t.white_space_max_share].sort_values("sales", ascending=False)
        if len(pools_) and pools_.iloc[0].sales >= t.min_sales:
            pool_ = pools_.iloc[0]
            fl.append(T(f"PB比率が{t.white_space_max_share:.0%}以下で最も売上が大きいのは{pool_['name']}({J(pool_.sales)}、PB比率{PC(pool_.share_cur)})です。",
                        f"Largest market with PB ratio at or below {t.white_space_max_share:.0%}: {pool_['name']} ({J(pool_.sales)}, PB ratio {PC(pool_.share_cur)})."))
        if len(watch):
            wn = watch.loc[watch["dept"] == dr["dept"], "name"].tolist()
            if wn:
                fl.append(T("要注視(PB比率が2期連続で低下): " + "、".join(wn) + "。",
                            "Watch (PB ratio down two periods in a row): " + ", ".join(wn) + "."))
        dept_findings[dr["dept"]] = fl

    cat_only = SEP.join(data.category_only)
    scope = T(f"対象: 全{len(d)}サブカテ" + (f"(カテゴリー単位のみの{cat_only}を含む)" if data.category_only else "") + "。",
              f"Scope: all {len(d)} sub-categories" + (f" (incl. category-level only: {cat_only})" if data.category_only else "") + ".")
    return Result(periods_full=(prev, cur), period_latest=latest, partial_months=data.partial_months, kpi=kpi, subs=d, cats=c,
                  directions=directions, caveats=cav, thresholds=t, findings=f, watchlist=watch, recon=recon, dq=dq,
                  borderline=bl, scope_note=scope, lang=lang, depts=dp, dept_findings=dept_findings)


def key_notes(res: Result) -> list:
    """The few assumptions a reader needs to know. Everything else (data checks, rules) stays off the report."""
    ja = res.lang == "ja"
    notes = []
    if res.period_latest:
        cf = res.periods_full[1]
        sh_ = res.kpi["sales_latest"] / res.kpi["sales_cur"] if res.kpi["sales_cur"] else float("nan")
        notes.append(f"{res.period_latest}期は他の期より対象期間が短く(売上は{cf}期の{sh_:.0%})、売上は他の期と比べられません。"
                     f"PB比率・SKU数・シェア動向のみ{cf}期と並べて表示しており、成長率や方向性の判定には使っていません。" if ja else
                     f"{res.period_latest}期 covers a shorter span than the others (its sales are {sh_:.0%} of {cf}期), so sales cannot be compared with the other periods. "
                     f"Only PB ratio, SKU counts and share trend are shown next to {cf}期; it is not used for growth or directions.")
    notes.append("元資料にPB売上額はないため、PB比率(売上に占めるPBの割合)と売上は元資料の値をそのまま表示しています。" if ja else
                 "The workbook has no PB sales amounts, so sales and the PB ratio are shown exactly as given in the workbook.")
    notes.append("粗利・原価のデータは含まれていません。PB比率が高くても、利益が高いとは限りません。" if ja else
                 "No margin or cost data is included. A high PB ratio does not necessarily mean high profit.")
    notes.append("平均単価の変化には、価格・容量・商品構成の変化が含まれます。数量に関する記述は参考値です。" if ja else
                 "Average price per unit reflects price, pack-size and product-mix changes, so statements about volume are indicative.")
    return notes
