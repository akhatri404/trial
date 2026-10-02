"""Reads the workbook and normalises it into one tidy table.

Expected layout (same as 開発方向性_コミュニケーション資料):
  * a sheet whose name contains 「サブカテ」: one row per sub-category, with 部門 / カテゴリ / サブカテ and, for each
    period, 全体売上 / 全体売れ数 / PB比率(売上) / PB比率(売れ数) / PB SKU数
  * sheet 「カテゴリー」 (optional): same metrics at category level. Categories that have NO rows in the sub-category
    sheet (e.g. 観賞魚, 鳥, 昆虫) are added as their own rows so that totals cover the whole business.
  * sheet 「方向性」 (optional): department totals as reported by the source, plus a 「※Nヶ月分」 note. Used only to
    reconcile our totals against the source, never as an input to the analysis.

Everything else (トライアル direction markers, 過去資料) is deliberately ignored.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd


class WorkbookError(ValueError):
    """Raised with a message that is safe to show to end users."""


_METRICS = {
    "全体売上": "sales",
    "全体売れ数": "units",
    "PB比率(売上)": "pb_ratio_sales",
    "PB比率(売れ数)": "pb_ratio_units",
    "PBSKU数": "pb_skus",
}
_LABELS = {v: k for k, v in _METRICS.items()}
_COL_RE = re.compile(r"^(\d+/\d+)期(.+)$")


def _norm(x) -> str:
    s = "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x)
    s = s.replace("（", "(").replace("）", ")")
    return re.sub(r"\s+", "", s)


def _strip_code(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"^\d+\s*", "", str(s)).strip()


@dataclass
class WorkbookData:
    df: pd.DataFrame                       # one row per sub-category (+ category-only rows); wide columns like sales_25/6
    periods: list[str]                     # in file order
    partial_months: int | None = None      # from the 「※Nヶ月分」 note, if present
    reported: dict = field(default_factory=dict)   # {dept_name: {period: {sales, units, pb_ratio_sales, pb_ratio_units}}}
    total_name: str | None = None          # key in `reported` of the grand-total row (ペット計)
    category_only: list[str] = field(default_factory=list)   # categories added from the カテゴリー sheet
    warnings: list[str] = field(default_factory=list)


def _find_sheet(xl: pd.ExcelFile, keyword: str):
    for n in xl.sheet_names:
        if keyword in n:
            return n
    return None


def _header_row(raw: pd.DataFrame, must_have: tuple, limit: int = 15):
    for i in range(min(limit, len(raw))):
        vals = [_norm(v) for v in raw.iloc[i].tolist()]
        if all(m in vals for m in must_have):
            return i, vals
    return None, None


def _period_columns(headers: list, metrics: dict):
    periods: list = []
    found: dict = {}
    for j, h in enumerate(headers):
        m = _COL_RE.match(h)
        if m and m.group(2) in metrics:
            p = m.group(1)
            if p not in periods:
                periods.append(p)
            found[(p, metrics[m.group(2)])] = j
    return periods, found


def _num(series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def load_workbook_data(xlsx) -> WorkbookData:
    try:
        xl = pd.ExcelFile(xlsx)
    except Exception as e:
        raise WorkbookError(f"Excelファイルとして開けませんでした({e})。") from e

    sub_name = _find_sheet(xl, "サブカテ")
    if sub_name is None:
        raise WorkbookError("「サブカテ」を含むシートが見つかりません。このファイルのシート: " + "、".join(xl.sheet_names))
    raw = xl.parse(sub_name, header=None)
    hdr_idx, headers = _header_row(raw, ("部門", "サブカテ"))
    if hdr_idx is None:
        raise WorkbookError(f"シート「{sub_name}」に見出し行(部門 / カテゴリ / サブカテ)が見つかりません。")
    body = raw.iloc[hdr_idx + 1:]

    def col(name):
        if name not in headers:
            raise WorkbookError(f"シート「{sub_name}」に列「{name}」がありません。")
        return headers.index(name)

    out = pd.DataFrame({
        "dept": body.iloc[:, col("部門")].map(_strip_code),
        "cat": body.iloc[:, col("カテゴリ")].map(_strip_code),
        "sub": body.iloc[:, col("サブカテ")].map(_strip_code),
    })
    periods, found = _period_columns(headers, _METRICS)
    if len(periods) < 2:
        raise WorkbookError("2期以上(例: 25/6期と26/6期)のデータが必要です。見つかった期: " + ("、".join(periods) or "なし"))
    for p in periods:
        for k in _METRICS.values():
            if (p, k) not in found:
                raise WorkbookError(f"{p}期に「{_LABELS[k]}」の列がありません。")
            out[f"{k}_{p}"] = _num(body.iloc[:, found[(p, k)]])
    out = out[(out["sub"] != "") & (out["sub"].str.lower() != "nan")].reset_index(drop=True)
    if out.empty:
        raise WorkbookError("サブカテのシートにデータ行がありません。")
    out["source"] = "サブカテ"

    warnings: list = []
    category_only: list = []

    # --- category sheet: add categories that have no sub-category rows ------------------------------
    cat_name = _find_sheet(xl, "カテゴリー")
    if cat_name:
        try:
            c = xl.parse(cat_name, header=None)
            ci, ch = _header_row(c, ("部門", "カテゴリ"))
            cperiods, cfound = _period_columns(ch, _METRICS)
            if ci is None or not all((p, k) in cfound for p in periods for k in _METRICS.values()):
                raise ValueError("category sheet layout differs")
            rows = c.iloc[ci + 1:]
            cc = pd.DataFrame({"dept": rows.iloc[:, ch.index("部門")].map(_strip_code),
                               "cat": rows.iloc[:, ch.index("カテゴリ")].map(_strip_code)})
            for p in periods:
                for k in _METRICS.values():
                    cc[f"{k}_{p}"] = _num(rows.iloc[:, cfound[(p, k)]])
            cc = cc[(cc["cat"] != "") & (cc["cat"].str.lower() != "nan")]
            have = set(zip(out["dept"], out["cat"]))
            extra = cc[[(d, k) not in have for d, k in zip(cc["dept"], cc["cat"])]].copy()
            if len(extra):
                extra["sub"] = extra["cat"]
                extra["source"] = "カテゴリー"
                category_only = extra["cat"].tolist()
                out = pd.concat([out, extra[out.columns]], ignore_index=True)
        except Exception:
            warnings.append("「カテゴリー」シートを読み込めなかったため、サブカテ行のないカテゴリー(ある場合)は"
                            "合計に含まれていません。")

    # --- reported department totals from 方向性 (reconciliation only) --------------------------------
    partial_months, reported, total_name = None, {}, None
    dir_name = _find_sheet(xl, "方向性")
    if dir_name:
        d = xl.parse(dir_name, header=None)
        for i in range(min(8, len(d))):
            for v in d.iloc[i].tolist():
                m = re.search(r"※\s*(\d+)\s*[ヶか箇ケ]月分", str(v))
                if m:
                    partial_months = int(m.group(1))
        hi, hh = _header_row(d, ("部門",), limit=8)
        if hi is not None:
            m4 = {k: v for k, v in _METRICS.items() if v != "pb_skus"}
            dper, dfound = _period_columns(hh, m4)
            name_col = hh.index("部門")
            for i in range(hi + 1, len(d)):
                nm = _strip_code(d.iloc[i, name_col])
                if not nm or nm.lower() == "nan":
                    continue
                reported[nm] = {p: {k: float(pd.to_numeric(d.iloc[i, dfound[(p, k)]], errors="coerce") or 0.0)
                                    for k in m4.values() if (p, k) in dfound} for p in dper}
                if "計" in nm and total_name is None:
                    total_name = nm

    return WorkbookData(out, periods, partial_months, reported, total_name, category_only, warnings)
