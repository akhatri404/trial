"""Renders a Result into a PDF with ReportLab (pure Python, no browser or Office needed).

Japanese text: an embedded TrueType font is used so the PDF looks the same on every machine.
Font lookup order: $PBREPORT_FONT, ./fonts/*.ttf (bundled IPAex Gothic), common system TrueType fonts,
and finally ReportLab's built-in (non-embedded) HeiseiKakuGo CID font.
"""
from __future__ import annotations

import datetime as dt
import io
import os
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.graphics.shapes import Drawing, Line, Rect, String
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle)

from .analysis import (DEFEND, DEPRIORITIZE, MONITOR, ORDER, RECAPTURE, REPLICATE, REVIEW, SCALE, TEST, V_SMALL, Result,
                       key_notes, money, tr)

NAVY = colors.HexColor("#1F3864")
GREY = colors.HexColor("#595959")
LIGHT = colors.HexColor("#EEF2F8")
GREEN = colors.HexColor("#2E7D32")
RED = colors.HexColor("#C62828")
BAR_BLUE = colors.HexColor("#3B6FB6")

DIR_COLOR = {SCALE: "#2E7D32", REPLICATE: "#3B6FB6", RECAPTURE: "#C77700", TEST: "#7B4EA3",
             REVIEW: "#B71C1C", DEFEND: "#455A64", DEPRIORITIZE: "#8D8D8D", MONITOR: "#8D8D8D"}

DIR_BLURB = {
    SCALE: "PBシェアが上昇しており、PB売上の規模があり、非PBの伸びしろも残っています。SKU開発はまずここに注力します。",
    REPLICATE: "PBはまだほぼ未展開ですが、同じカテゴリー内に「拡大」の実績があるサブカテがあります。その成功パターンを横展開します。",
    RECAPTURE: "PBシェアが前期から大きく低下しました。投資の前に原因(取扱終了、競合、価格)を特定します。",
    TEST: "PB未展開の大きな市場で、同カテゴリー内に実績のあるサブカテもありません。実現性はこのデータでは判断できないため、1〜2SKUでテストします。",
    REVIEW: "PBシェアが低下傾向です。商品・価格・売場のどこに問題があるかを確認します。",
    DEFEND: "PBシェアはすでに高い水準です。シェアを守り、生産性を下げるだけのSKU追加は避けます。",
    DEPRIORITIZE: "PB未展開で、規模が小さいか縮小している市場です。優先度は低くします。",
    MONITOR: "現在のルールでは明確なシグナルはありません。",
}

DIR_BLURB_EN = {
    SCALE: "PB share has been rising, PB sales are material and non-PB headroom remains. Put SKU effort here first.",
    REPLICATE: "No real PB position yet, but a sibling sub-category in the same category is a proven Scale item. Reuse that playbook.",
    RECAPTURE: "PB share collapsed versus the previous period. Find the root cause (delisting, competitor, pricing) before investing.",
    TEST: "Large pools with no PB position and no proven sibling. Feasibility is not visible in this data, so test with 1-2 SKUs.",
    REVIEW: "PB share is slipping. Check whether this is a product, price or placement issue.",
    DEFEND: "PB share is already high. Protect it and avoid adding SKUs that only dilute productivity.",
    DEPRIORITIZE: "Small or shrinking pools with no PB position. Low priority.",
    MONITOR: "No clear signal under the current rules.",
}

_FONT = "JP"
_registered = False


def _register_font() -> str:
    global _registered, _FONT
    if _registered:
        return _FONT
    here = Path(__file__).resolve().parent.parent
    cands = []
    if os.environ.get("PBREPORT_FONT"):
        cands.append(os.environ["PBREPORT_FONT"])
    cands += sorted(str(p) for p in (here / "fonts").glob("*.ttf"))
    cands += [
        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
        "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
        "/usr/share/fonts/truetype/ipaexfont-gothic/ipaexg.ttf",
        "C:/Windows/Fonts/msgothic.ttc",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    ]
    for c in cands:
        if c and Path(c).exists():
            try:
                pdfmetrics.registerFont(TTFont("JP", c, subfontIndex=0) if c.lower().endswith(".ttc") else TTFont("JP", c))
                _FONT, _registered = "JP", True
                return _FONT
            except Exception:
                continue
    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))   # last resort: not embedded
    _FONT, _registered = "HeiseiKakuGo-W5", True
    return _FONT


def _styles(lang="ja"):
    f = _register_font()
    # wordWrap="CJK": Japanese has no spaces, so lines must be allowed to break between any characters
    P = lambda name, **kw: ParagraphStyle(name, fontName=f, wordWrap="CJK" if lang == "ja" else None, **kw)
    return {
        "title": P("title", fontSize=20, leading=25, textColor=NAVY, spaceAfter=2),
        "sub": P("sub", fontSize=9, leading=12, textColor=GREY),
        "h1": P("h1", fontSize=13.5, leading=17, textColor=NAVY, spaceBefore=12, spaceAfter=5),
        "h2": P("h2", fontSize=10.5, leading=14, textColor=NAVY, spaceBefore=8, spaceAfter=3),
        "body": P("body", fontSize=9, leading=13, spaceAfter=3),
        "small": P("small", fontSize=7.5, leading=10, textColor=GREY),
        "bullet": P("bullet", fontSize=9, leading=13, leftIndent=11, bulletIndent=1, spaceAfter=3),
        "cell": P("cell", fontSize=7.5, leading=9.5),
        "cellh": P("cellh", fontSize=7.5, leading=9.5, textColor=colors.white),
    }


def _p(text, st):
    return Paragraph(escape(str(text)), st)


def _pct0(x, d=1, na="－"):
    return na if x is None or x != x else f"{x * 100:.{d}f}%"


def _signed_pct0(x, d=1, na="－"):
    return na if x is None or x != x else f"{x * 100:+.{d}f}%"


def _table(rows, widths, st, header=True, align_right_from=1, zebra=True):
    data = []
    for i, r in enumerate(rows):
        data.append([_p(c, st["cellh"] if (header and i == 0) else st["cell"]) if not hasattr(c, "wrap") else c for c in r])
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    style = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#C9CED6")),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), NAVY))
    if zebra:
        for i in range(1 if header else 0, len(rows)):
            if i % 2 == 0:
                style.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
    t.setStyle(TableStyle(style))
    return t


def _fit(text, font, size, width):
    if pdfmetrics.stringWidth(text, font, size) <= width:
        return text
    while text and pdfmetrics.stringWidth(text + "…", font, size) > width:
        text = text[:-1]
    return text + "…"


def _bar_chart(items, width_mm=180, label_w_mm=62, row_h=13, title=None, signed=False):
    """items: list of (label, value, color, value_text)."""
    f = _register_font()
    W, rh = width_mm * mm, row_h
    top = 16 if title else 4
    H = top + rh * len(items) + 4
    dr = Drawing(W, H)
    if title:
        dr.add(String(0, H - 11, title, fontName=f, fontSize=8.5, fillColor=NAVY))
    lw = label_w_mm * mm
    area = W - lw - 62
    vmax = max([abs(v) for _, v, _, _ in items] + [1e-9])
    vmin = min([v for _, v, _, _ in items] + [0]) if signed else 0
    span = (max(v for _, v, _, _ in items) - vmin) or 1e-9
    zero_x = lw + (-vmin / span) * area
    for i, (label, v, col, txt) in enumerate(items):
        y = H - top - (i + 1) * rh + 3
        dr.add(String(0, y + 2, _fit(label, f, 7, lw - 6), fontName=f, fontSize=7, fillColor=colors.black))
        w = abs(v) / span * area
        x = zero_x if v >= 0 else zero_x - w
        dr.add(Rect(x, y, max(w, 0.6), rh - 5, fillColor=col, strokeColor=None))
        tx = x + w + 3 if v >= 0 else x - 3
        dr.add(String(tx, y + 2, txt, fontName=f, fontSize=7, fillColor=GREY, textAnchor="start" if v >= 0 else "end"))
    if signed:
        dr.add(Line(zero_x, 2, zero_x, H - top + 2, strokeColor=GREY, strokeWidth=0.5))
    return dr


def _n(x, d=0):
    return "－" if x is None or x != x else f"{x:,.{d}f}"


def build_pdf(res: Result, source_name: str = "") -> bytes:
    lang = res.lang
    T = lambda ja, en: ja if lang == "ja" else en
    J = money(lang)
    D = lambda key: tr(key, lang)                     # direction / verdict display name
    NA = T("－", "–")
    _pct = lambda x, d=1: _pct0(x, d, NA)
    _signed_pct = lambda x, d=1: _signed_pct0(x, d, NA)
    BLURB = DIR_BLURB if lang == "ja" else DIR_BLURB_EN
    SEP = T("、", ", ")
    ARROW = "→"

    st = _styles(lang)
    f = _register_font()
    prev, cur = res.periods_full
    k, b, t = res.kpi, res.bridge, res.thresholds
    buf = io.BytesIO()
    title = T("ペットカテゴリー PB戦略レポート", "Pet category PB strategy report")
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=14 * mm,
                            bottomMargin=16 * mm, title=title, author=T("PB戦略レポート", "PB strategy report"))
    FW = A4[0] - 30 * mm
    S = []
    H1 = lambda txt: S.append(_p(txt, st["h1"]))
    H2 = lambda txt: S.append(_p(txt, st["h2"]))
    NOTE = lambda txt: S.append(_p(txt, st["small"]))
    BODY = lambda txt: S.append(_p(txt, st["body"]))

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont(f, 7)
        canvas.setFillColor(GREY)
        canvas.drawString(15 * mm, 8 * mm, T("PB戦略レポート", "PB strategy report"))
        canvas.drawRightString(A4[0] - 15 * mm, 8 * mm, T(f"{doc_.page}ページ", f"Page {doc_.page}"))
        canvas.restoreState()

    # ---- title ----
    S.append(_p(title, st["title"]))
    if res.period_latest:
        tail = T(f"(PB比率とSKU数は{res.period_latest}期" + (f"・{res.partial_months}ヶ月分" if res.partial_months else "・途中期") + "も表示)",
                 f" (PB ratio and SKU trend also show {res.period_latest}期" + (f", {res.partial_months} months" if res.partial_months else ", partial") + ")")
    else:
        tail = ""
    S.append(_p(T(f"比較期間: {prev}期 vs {cur}期", f"Periods compared: {prev}期 vs {cur}期") + tail, st["sub"]))
    S.append(_p(T(f"作成日時 {dt.datetime.now():%Y-%m-%d %H:%M}", f"Generated {dt.datetime.now():%Y-%m-%d %H:%M}")
                + (T(f"  |  元データ: {source_name}", f"  |  Source: {source_name}") if source_name else ""), st["sub"]))
    S.append(_p(res.scope_note, st["sub"]))
    S.append(Spacer(1, 4))

    # ---- 1. takeaways ----
    H1(T("1. 要点", "1. Key takeaways"))
    for line in res.findings:
        S.append(Paragraph(escape(line), st["bullet"], bulletText="•"))
    rows = [[T("指標", "Metric"), f"{prev}期", f"{cur}期", T("増減", "Change")],
            [T("全体売上", "Total sales"), J(k["sales_prev"]), J(k["sales_cur"]), _signed_pct(k["sales_yoy"])],
            [T("全体売れ数", "Total units"), "", "", _signed_pct(k["units_yoy"])],
            [T("平均単価(価格・構成)", "Average yen per unit (price/mix)"), "", "", _signed_pct(k["asp_chg"])],
            [T("PB売上", "PB sales"), J(k["pb_prev"]), J(k["pb_cur"]), f"{J(k['d_pb'])} ({_signed_pct(k['pb_yoy'])})"],
            [T("非PB売上", "Non-PB sales"), J(k["sales_prev"] - k["pb_prev"]), J(k["sales_cur"] - k["pb_cur"]), _signed_pct(k["nb_yoy"])],
            [T("PB売れ数 vs 非PB売れ数", "PB units vs non-PB units"), "", "", f"{_signed_pct(k['pb_units_yoy'])} vs {_signed_pct(k['nb_units_yoy'])}"],
            [T("PB比率(売上)", "PB ratio (sales)"), _pct(k["share_prev"], 2), _pct(k["share_cur"], 2), f"{(k['share_cur'] - k['share_prev']) * 100:+.2f} pt"],
            [T("PB SKU数", "PB SKUs"), f"{k['sku_prev']:.0f}", f"{k['sku_cur']:.0f}", _signed_pct(k["sku_growth"], 0)],
            [T("PB SKU当たりPB売上", "PB sales per PB SKU"), J(k["per_sku_prev"]), J(k["per_sku_cur"]), _signed_pct(k["per_sku_chg"], 0)]]
    if res.period_latest:
        rows.append([T(f"{res.period_latest}期(直近・途中期): PB比率 / PB SKU数", f"{res.period_latest}期 (latest, partial): PB ratio / PB SKUs"),
                     "", f"{_pct(k.get('share_latest'), 2)} / {k.get('sku_latest', float('nan')):.0f}",
                     T(f"{cur}期比 {(k['share_latest'] - k['share_cur']) * 100:+.2f}pt", f"{(k['share_latest'] - k['share_cur']) * 100:+.2f} pt vs {cur}期")])
    S.append(Spacer(1, 4))
    S.append(_table(rows, [FW * .38, FW * .17, FW * .17, FW * .28], st))

    # ---- 2. growth bridge ----
    H1(T("2. PB成長の要因", "2. What drove PB growth"))
    BODY(T(f"PB売上は{J(b['d_pb'])}変化しました。各サブカテの{prev}期PBシェアを固定した場合、市場成長だけで"
           f"{J(b['market'])}増加し、残りの{J(b['share'])}はシェアの純変化によるものです。両者の合計は増減額と完全に一致します。",
           f"PB sales changed by {J(b['d_pb'])}. Holding each sub-category's {prev}期 PB share fixed, market growth alone would have added "
           f"{J(b['market'])}; the rest, {J(b['share'])}, is net share change. The two parts add up exactly."))
    items = [(T("前期PBシェアでの市場成長", "Market growth at prior PB shares"), b["market"], BAR_BLUE, J(b["market"])),
             (T("シェア上昇(上昇したサブカテ)", "Share gains (sub-categories that gained)"), b["share_gain"], GREEN, J(b["share_gain"])),
             (T("シェア低下(低下したサブカテ)", "Share losses (sub-categories that lost)"), b["share_loss"], RED, J(b["share_loss"])),
             (T("PB売上の純増減", "Net change in PB sales"), b["d_pb"], NAVY, J(b["d_pb"]))]
    S.append(_bar_chart(items, label_w_mm=70, title=T(f"PB売上の増減要因 {prev}期→{cur}期", f"PB sales bridge, {prev}期 → {cur}期"), signed=True))
    S.append(Spacer(1, 4))
    BODY(T(f"PB比率 {b['ratio_prev']:.2%}→{b['ratio_cur']:.2%}({b['d_ratio_pt']:+.2f}pt)= サブカテ内のシェア変化 {b['rate_pt']:+.2f}pt "
           f"+ 構成変化 {b['mix_pt']:+.2f}pt(PB比率の高い/低いサブカテ間での売上の移動)。",
           f"PB ratio {b['ratio_prev']:.2%} → {b['ratio_cur']:.2%} ({b['d_ratio_pt']:+.2f} pt) = {b['rate_pt']:+.2f} pt from share change inside sub-categories "
           f"and {b['mix_pt']:+.2f} pt from mix (sales shifting between higher-PB and lower-PB sub-categories)."))
    cb = res.cats.sort_values("d_pb", ascending=False)
    rows = [[T("カテゴリー", "Category"), T("PB売上の増減", "Change in PB sales"), T("市場要因", "Market effect"), T("シェア要因", "Share effect"),
             T("PB比率pt: シェア", "PB ratio pt: share"), T("PB比率pt: 構成", "PB ratio pt: mix")]]
    for _, r in cb.iterrows():
        if abs(r.d_pb) < 1e5 and abs(r.market_eff) < 1e5:
            continue
        rows.append([r["cat"], J(r.d_pb), J(r.market_eff), J(r.share_eff), f"{r.rate_pt:+.2f}", f"{r.mix_pt:+.2f}"])
    rows.append([T("合計", "Total"), J(b["d_pb"]), J(b["market"]), J(b["share"]), f"{b['rate_pt']:+.2f}", f"{b['mix_pt']:+.2f}"])
    S.append(_table(rows, [FW * .22, FW * .18, FW * .16, FW * .16, FW * .14, FW * .14], st))
    cn = res.concentration
    S.append(Spacer(1, 3))
    NOTE(T(f"集中度: 上位3サブカテ = PB売上の{cn['top3']:.0%}({prev}期は{cn['top3_prev']:.0%})、上位5 = {cn['top5']:.0%}。"
           f"PB展開中の{cn['n_active']}サブカテのうち{cn['n_for_80']}サブカテで80%を占めます。",
           f"Concentration: top 3 sub-categories = {cn['top3']:.0%} of PB sales ({cn['top3_prev']:.0%} in {prev}期); top 5 = {cn['top5']:.0%}; "
           f"{cn['n_for_80']} of {cn['n_active']} PB-active sub-categories make up 80%."))

    # ---- 3. movers + pools ----
    H1(T("3. PB売上の増減と伸びしろ", "3. Where PB sales moved, and where the headroom is"))
    sub_df = res.subs
    mv = sub_df.sort_values("d_pb", ascending=False)
    gains, losses = mv[mv.d_pb > 0].head(8), mv[mv.d_pb < 0].tail(5)
    items = [(r["name"], r.d_pb, GREEN, J(r.d_pb)) for _, r in gains.iterrows()] + \
            [(r["name"], r.d_pb, RED, J(r.d_pb)) for _, r in losses.sort_values("d_pb").iterrows()]
    if items:
        S.append(_bar_chart(items, title=T(f"PB売上の増減 {prev}期→{cur}期(増減の大きいサブカテ)",
                                           f"Change in PB sales, {prev}期 → {cur}期 (largest movers)"), signed=True))
    S.append(Spacer(1, 6))
    pools = sub_df.sort_values("nb_cur", ascending=False).head(10)
    items = [(r["name"], r.nb_cur, BAR_BLUE, f"{J(r.nb_cur)}  (PB {r.share_cur * 100:.1f}%)") for _, r in pools.iterrows()]
    S.append(_bar_chart(items, label_w_mm=62, title=T(f"非PB売上の大きいサブカテ {cur}期", f"Largest non-PB sales pools, {cur}期")))

    # ---- 4. category scorecard + SKU vs sales ----
    S.append(PageBreak())
    H1(T("4. カテゴリー別スコアカード: 成長・順位・SKUと売上", "4. Category scorecard: growth, ranking and SKU vs sales"))
    BODY(T("順位はPB売上順(1 = 最大)。判定はPB SKU数の変化とPB売上の変化を比較したものです"
           f"(SKU変化が±{t.sku_stable_band:.0%}以内は横ばい、PB売上{J(t.min_base_pb_sales)}未満の小規模なものは判定対象外)。",
           "Rank is by PB sales (1 = largest). Verdict compares the change in PB SKU count with the change in PB sales "
           f"(SKU change within ±{t.sku_stable_band:.0%} counts as a stable range; small bases under {J(t.min_base_pb_sales)} PB sales are not judged)."))
    rows = [[T("カテゴリー", "Category"), T("PB売上(順位)", "PB sales (rank)"), T("PBシェア(増減pt)", "PB share (Δ pt)"), T("PB成長率", "PB growth"),
             T("市場成長率", "Market growth"), T("PB SKU数", "PB SKUs"), T("SKU当たりPB売上", "PB sales / SKU"), T("判定", "Verdict")]]
    for _, r in res.cats.iterrows():
        rows.append([r["cat"], T(f"{J(r.pb_cur)}({r.rank_pb}位)", f"{J(r.pb_cur)} (#{r.rank_pb})"), f"{_pct(r.share_cur)} ({r.d_share_pt:+.1f})",
                     _signed_pct(r.pb_growth, 0), _signed_pct(r.yoy),
                     f"{r.sku_prev:.0f}{ARROW}{r.sku_cur:.0f} ({_signed_pct(r.sku_growth, 0)})",
                     f"{J(r.per_sku_cur)} ({_signed_pct(r.per_sku_chg, 0)})", D(r.verdict)])
    S.append(_table(rows, [FW * .12, FW * .13, FW * .13, FW * .08, FW * .09, FW * .15, FW * .15, FW * .15], st))
    S.append(Spacer(1, 3))
    marg = J(k["marg_per_sku"]) if k["marg_per_sku"] == k["marg_per_sku"] else NA
    NOTE(T("PB成長率 = PB売上の前期比。市場成長率 = 全体(PB+非PB)売上の前期比。"
           f"追加SKU1つ当たりのPB売上増: 全体で{marg}(参考: {prev}期の既存SKU1つ当たり{J(k['per_sku_prev'])})。",
           "PB growth = change in PB sales vs the previous period. Market growth = change in total (PB + non-PB) sales. "
           f"Marginal PB sales per added SKU: {marg} overall vs {J(k['per_sku_prev'])} per existing SKU in {prev}期."))

    H2(T("サブカテ別 PB SKU数の変化とPB売上の変化(一定規模以上)", "PB SKU change vs PB sales change, by sub-category (larger bases)"))
    sv = sub_df[(sub_df.verdict != V_SMALL)].sort_values("pb_cur", ascending=False).head(16)
    rows = [[T("サブカテ", "Sub-category"), T("PB SKU数", "PB SKUs"), T("PB売上", "PB sales"), T("PB売上成長率", "PB sales growth"),
             T("追加SKU当たりPB売上増", "Extra PB sales per added SKU"), T("判定", "Verdict")]]
    for _, r in sv.iterrows():
        rows.append([r["name"], f"{r.sku_prev:.0f}{ARROW}{r.sku_cur:.0f}", f"{J(r.pb_prev)}{ARROW}{J(r.pb_cur)}", _signed_pct(r.pb_growth, 0),
                     J(r.marg_per_sku), D(r.verdict)])
    S.append(_table(rows, [FW * .27, FW * .1, FW * .2, FW * .11, FW * .14, FW * .18], st))

    # ---- 5. rankings ----
    H1(T("5. ランキング(サブカテ)", "5. Rankings (sub-category)"))
    rk = res.rankings

    def rank_table(title_, df_, cols, widths, hdr_):
        H2(title_)
        rows_ = [hdr_]
        for _, r in df_.iterrows():
            rows_.append([r["name"]] + [fn(r) for fn in cols])
        S.append(_table(rows_, widths, st))
    sub_h = T("サブカテ", "Sub-category")
    base_txt = J(t.min_base_pb_sales)
    rank_table(T("PB売上 上位10", "By PB sales (top 10)"), rk["sub_pb"],
               [lambda r: J(r.pb_cur), lambda r: _pct(r.share_cur), lambda r: _signed_pct(r.pb_growth, 0),
                lambda r: f"{r.sku_cur:.0f}", lambda r: J(r.per_sku_cur)],
               [FW * .33, FW * .14, FW * .13, FW * .13, FW * .1, FW * .17],
               [sub_h, T("PB売上", "PB sales"), T("PBシェア", "PB share"), T("PB成長率", "PB growth"), T("SKU数", "SKUs"), T("SKU当たりPB売上", "PB sales / SKU")])
    growth_cols = [lambda r: _signed_pct(r.pb_growth, 0), lambda r: J(r.d_pb), lambda r: f"{_pct(r.share_prev)}{ARROW}{_pct(r.share_cur)}"]
    growth_hdr = [sub_h, T("PB成長率", "PB growth"), T("増減額", "Change"), T("PBシェア", "PB share")]
    rank_table(T(f"PB売上成長率 上位(前期PB売上{base_txt}以上)", f"Fastest PB sales growth (PB sales base of at least {base_txt})"),
               rk["growth_top"], growth_cols, [FW * .36, FW * .16, FW * .16, FW * .32], growth_hdr)
    rank_table(T(f"PB売上成長率 下位(前期PB売上{base_txt}以上)", f"Slowest PB sales growth (PB sales base of at least {base_txt})"),
               rk["growth_bottom"], growth_cols, [FW * .36, FW * .16, FW * .16, FW * .32], growth_hdr)
    sku_cols = [lambda r: J(r.per_sku_cur), lambda r: f"{r.sku_cur:.0f}", lambda r: _signed_pct(r.per_sku_chg, 0)]
    sku_hdr = [sub_h, T("SKU当たりPB売上", "PB sales / SKU"), T("PB SKU数", "PB SKUs"), T("増減", "Change")]
    rank_table(T("SKU当たりPB売上 上位(2SKU以上)", "Highest PB sales per SKU (at least 2 SKUs)"),
               rk["per_sku_top"], sku_cols, [FW * .36, FW * .2, FW * .2, FW * .24], sku_hdr)
    rank_table(T("SKU当たりPB売上 下位(2SKU以上)", "Lowest PB sales per SKU (at least 2 SKUs)"),
               rk["per_sku_bottom"], sku_cols, [FW * .36, FW * .2, FW * .2, FW * .24], sku_hdr)
    if res.period_latest and res.watchlist is not None and len(res.watchlist):
        H2(T(f"要注視リスト: PBシェアが2期連続で低下({prev}→{cur}→{res.period_latest})",
             f"Momentum watchlist: PB share falling two steps in a row ({prev} → {cur} → {res.period_latest})"))
        rows = [[sub_h, T("PB売上", "PB sales"), T("PBシェアの推移", "PB share path")]]
        for _, r in res.watchlist.iterrows():
            rows.append([r["name"], J(r.pb_cur), f"{_pct(r.share_prev)} {ARROW} {_pct(r.share_cur)} {ARROW} {_pct(r.share_latest)}"])
        S.append(_table(rows, [FW * .4, FW * .2, FW * .4], st))
        NOTE(T(f"{res.period_latest}期は途中期のため、シェアは参考値です(季節性)。",
               f"The {res.period_latest} period is partial, so its share is indicative (seasonality)."))

    # ---- 6. directions ----
    S.append(PageBreak())
    H1(T("6. 方向性の提案(データに基づく。トライアルの記載は不使用)", "6. Proposed direction (data-driven; トライアル markers not used)"))
    BODY(T("各サブカテは、PBシェアの水準と推移、市場の規模と成長に基づいて1つの方向性に分類しています。非PB = 直近の通期における全体売上 − PB売上。",
           "Each sub-category is placed in one direction based on its PB share level and trend, and on market size and growth. Non-PB = total sales minus PB sales in the latest full period."))
    latest_suffix = f"{ARROW}{res.period_latest}" if res.period_latest else ""
    hdr = [sub_h, T(f"売上 {cur}", f"Sales {cur}"), T("非PB", "Non-PB"), T(f"PBシェア {prev}→{cur}", f"PB share {prev}→{cur}") + latest_suffix,
           T("市場前年比", "Mkt YoY")]
    wd = [FW * .34, FW * .13, FW * .13, FW * .28, FW * .12]
    for dname in ORDER:
        x = res.directions[dname]
        if x.empty or dname == MONITOR:
            continue
        block = [_p(T(f"{D(dname)}({len(x)}件)", f"{D(dname)}  ({len(x)})"), st["h2"]), _p(BLURB[dname], st["small"]), Spacer(1, 2)]
        if dname == DEPRIORITIZE:
            block.append(_p(SEP.join(x["name"].tolist()), st["body"]))
            S.append(KeepTogether(block))
            continue
        show = x.head(10)
        rows = [hdr]
        for _, r in show.iterrows():
            lat = f"{ARROW}{r.share_latest * 100:.1f}%" if r.share_latest == r.share_latest else ""
            rows.append([r["name"], J(r.sales), J(r.nb_cur), f"{r.share_prev * 100:.1f}%{ARROW}{r.share_cur * 100:.1f}%{lat}", _signed_pct(r.yoy)])
        block.append(_table(rows, wd, st))
        if len(x) > len(show):
            block.append(_p(T(f"ほか{len(x) - len(show)}件(付録参照)", f"+ {len(x) - len(show)} more (see appendix)"), st["small"]))
        S.append(KeepTogether(block))

    # ---- 7. upside ----
    H1(T("7. 伸びしろの試算(仮定に基づく)", "7. Illustrative upside (assumption-based)"))
    S.append(_p(T(f"{D(SCALE)}: PBシェアが現在のシェアに直近の上昇幅{t.upside_scale_momentum_years:g}年分を加えた水準に到達"
                  f"(上限{t.upside_scale_cap:.0%})。{D(RECAPTURE)}: 前期シェアに回復。{D(REPLICATE)}: シェア"
                  f"{t.upside_replicate_share:.0%}。{D(TEST)}は対象外。これは計画用の仮定であり予測ではありません。"
                  "売上はNB(ナショナルブランド)から獲得し、全体売上は横ばいと仮定しています。",
                  f"{D(SCALE)} items: PB share reaches current share plus {t.upside_scale_momentum_years:g} years of the last observed gain "
                  f"(capped at {t.upside_scale_cap:.0%}). {D(RECAPTURE)} items: previous-period share is restored. {D(REPLICATE)} items: "
                  f"{t.upside_replicate_share:.0%} share. {D(TEST)} items are excluded. These are planning assumptions, not forecasts, "
                  "and assume the sales come from national brands with total sales flat."), st["body"]))
    if not res.upside.empty:
        rows = [[sub_h, T("方向性", "Direction"), T(f"売上 {cur}", f"Sales {cur}"), T("PBシェア 現在→目標", "PB share now → target"),
                 T("PB売上の増加額", "Added PB sales")]]
        for _, r in res.upside.head(12).iterrows():
            rows.append([r["name"], D(r.direction), J(r.sales), f"{r.share_cur * 100:.1f}% {ARROW} {r.target * 100:.1f}%", J(r["add"])])
        u = res.upside_total
        rows.append([T("合計(表示外の項目を含む)", "Total (all items incl. any not shown)"), "", "",
                     T(f"PB比率 {_pct(k['share_cur'])}→{_pct(u['share_after'])}", f"PB ratio {_pct(k['share_cur'])} → {_pct(u['share_after'])}"),
                     T(f"{J(u['add'])}(PB売上の{_signed_pct(u['pct_of_pb'], 0)})", f"{J(u['add'])} ({_signed_pct(u['pct_of_pb'], 0)} of PB sales)")])
        S.append(_table(rows, [FW * .34, FW * .13, FW * .13, FW * .2, FW * .2], st))

    # ---- 8. basis of this report ----
    H1(T("8. 前提・注意事項", "8. Basis and notes"))
    for line in key_notes(res):
        S.append(Paragraph(escape(line), st["bullet"], bulletText="•"))

    # ---- appendix: every sub-category ----
    S.append(PageBreak())
    S.append(_p(T("付録. 全サブカテ一覧", "Appendix. All sub-categories"), st["h1"]))
    rows = [[T("方向性", "Direction"), sub_h, T(f"売上 {cur}", f"Sales {cur}"), T("非PB", "Non-PB"), T(f"PBシェア {prev}→{cur}", f"PB share {prev}→{cur}"),
             T("市場前年比", "Mkt YoY"), T("PB SKU数", "PB SKUs")]]
    for dname in ORDER:
        for _, r in res.directions[dname].iterrows():
            rows.append([D(dname), r["name"], J(r.sales), J(r.nb_cur), f"{r.share_prev * 100:.1f}%{ARROW}{r.share_cur * 100:.1f}%",
                         _signed_pct(r.yoy), f"{r.sku_prev:.0f}{ARROW}{r.sku_cur:.0f}"])
    S.append(_table(rows, [FW * .12, FW * .32, FW * .11, FW * .11, FW * .15, FW * .09, FW * .1], st))

    doc.build(S, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()
