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
                       T_DOWN, T_FLAT, T_UP, key_notes, money, tr)

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
        "h1": P("h1", fontSize=13.5, leading=17, textColor=NAVY, spaceBefore=12, spaceAfter=5, keepWithNext=1),
        "h2": P("h2", fontSize=10.5, leading=14, textColor=NAVY, spaceBefore=8, spaceAfter=3, keepWithNext=1),
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
    LAT = res.period_latest                            # the newest period, when it covers less than the others: shown alongside, not compared in yen
    LATH = f"{LAT}期*" if LAT else ""

    st = _styles(lang)
    f = _register_font()
    prev, cur = res.periods_full
    k, b, t = res.kpi, res.bridge, res.thresholds
    LATNOTE = (T(f"* {LAT}期は他の期より対象期間が短く(売上は{cur}期の{k['sales_latest'] / k['sales_cur']:.0%})、売上・PB売上は他の期と比べられないため、増減は表示していません。"
                 "PB比率とPB SKU数は並べて比較できます。",
                 f"* {LAT}期 covers a shorter span than the others (its sales are {k['sales_latest'] / k['sales_cur']:.0%} of {cur}期), so changes in sales and PB sales are not shown. "
                 "PB ratio and PB SKU count can be compared directly.")) if LAT else ""
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
    BULLET = lambda txt: S.append(Paragraph(escape(txt), st["bullet"], bulletText="•"))

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont(f, 7)
        canvas.setFillColor(GREY)
        canvas.drawString(15 * mm, 8 * mm, T("PB戦略レポート", "PB strategy report"))
        canvas.drawRightString(A4[0] - 15 * mm, 8 * mm, T(f"{doc_.page}ページ", f"Page {doc_.page}"))
        canvas.restoreState()

    # ---- title ----
    S.append(_p(title, st["title"]))
    S.append(_p(T(f"比較期間: {prev}期 vs {cur}期", f"Periods compared: {prev}期 vs {cur}期")
                + (T(f"(あわせて{LAT}期を並べて表示)", f" (with {LAT}期 shown alongside)") if LAT else ""), st["sub"]))
    S.append(_p(T(f"作成日時 {dt.datetime.now():%Y-%m-%d %H:%M}", f"Generated {dt.datetime.now():%Y-%m-%d %H:%M}")
                + (T(f"  |  元データ: {source_name}", f"  |  Source: {source_name}") if source_name else ""), st["sub"]))
    S.append(_p(res.scope_note, st["sub"]))
    S.append(Spacer(1, 4))

    # ---- 1. takeaways ----
    H1(T("1. 要点", "1. Key takeaways"))
    for line in res.findings:
        BULLET(line)
    ch = lambda a, b: _signed_pct(b / a - 1) if a else NA          # change from one period to the next
    L = lambda fn: fn() if LAT else []                              # the extra cells for the part-year period
    nb_cur = k["sales_cur"] - k["pb_cur"]
    rows = [[T("指標", "Metric"), f"{prev}期", f"{cur}期", T("増減", "Change")] + L(lambda: [LATH, f"{cur}→{LAT}"]),
            [T("全体売上", "Total sales"), J(k["sales_prev"]), J(k["sales_cur"]), _signed_pct(k["sales_yoy"])]
            + L(lambda: [J(k["sales_latest"]), NA]),
            [T("全体売れ数", "Total units"), "", "", _signed_pct(k["units_yoy"])] + L(lambda: ["", ""]),
            [T("平均単価(価格・構成)", "Average yen per unit (price/mix)"), "", "", _signed_pct(k["asp_chg"])] + L(lambda: ["", ""]),
            [T("PB売上", "PB sales"), J(k["pb_prev"]), J(k["pb_cur"]), f"{J(k['d_pb'])} ({_signed_pct(k['pb_yoy'])})"]
            + L(lambda: [J(k["pb_latest"]), NA]),
            [T("非PB売上", "Non-PB sales"), J(k["sales_prev"] - k["pb_prev"]), J(nb_cur), _signed_pct(k["nb_yoy"])]
            + L(lambda: [J(k["nb_latest"]), NA]),
            [T("PB売れ数 vs 非PB売れ数", "PB units vs non-PB units"), "", "", f"{_signed_pct(k['pb_units_yoy'])} vs {_signed_pct(k['nb_units_yoy'])}"]
            + L(lambda: ["", ""]),
            [T("PB比率(売上)", "PB ratio (sales)"), _pct(k["share_prev"], 2), _pct(k["share_cur"], 2), f"{(k['share_cur'] - k['share_prev']) * 100:+.2f} pt"]
            + L(lambda: [_pct(k["share_latest"], 2), f"{(k['share_latest'] - k['share_cur']) * 100:+.2f} pt"]),
            [T("PB SKU数", "PB SKUs"), f"{k['sku_prev']:.0f}", f"{k['sku_cur']:.0f}", _signed_pct(k["sku_growth"], 0)]
            + L(lambda: [f"{k['sku_latest']:.0f}", ch(k["sku_cur"], k["sku_latest"])]),
            [T("PB SKU当たりPB売上", "PB sales per PB SKU"), J(k["per_sku_prev"]), J(k["per_sku_cur"]), _signed_pct(k["per_sku_chg"], 0)]
            + L(lambda: [NA, NA])]
    S.append(Spacer(1, 4))
    S.append(_table(rows, [FW * .28, FW * .13, FW * .13, FW * .17, FW * .15, FW * .14] if LAT else [FW * .38, FW * .17, FW * .17, FW * .28], st))
    if LAT:
        NOTE(LATNOTE)


    # ---- 2. growth bridge (whole pet business) ----
    H1(T("2. PB成長の要因(ペット全体)", "2. What drove PB growth (whole pet business)"))
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

    # ---- 4. department comparison ----
    dp = res.depts
    H1(T("4. 部門別比較", "4. Department comparison"))
    BODY(T("ペット全体の数字は、PB比率の大きく異なる部門の平均です。部門ごとに見ると、次のとおりです。",
           "The pet total is an average of departments whose PB ratios are very different. Department by department:"))
    items = [(r["dept"], r.share_cur, BAR_BLUE, f"{_pct(r.share_cur)} ({r.d_share_pt:+.1f}pt)") for _, r in dp.iterrows()]
    S.append(_bar_chart(items, label_w_mm=45, title=T(f"部門別のPB比率 {cur}期(括弧内は{prev}期からの増減)",
                                                      f"PB ratio by department, {cur}期 (change vs {prev}期 in brackets)")))
    S.append(Spacer(1, 6))
    rows = [[T("部門", "Department"), T("売上", "Sales"), T("市場前年比", "Market YoY"), T("PB売上", "PB sales"), T("PB売上の増減", "Change in PB sales"),
             T(f"PB比率 {prev}→{cur}", f"PB ratio {prev}→{cur}"), T("PB SKU数", "PB SKUs"), T("SKU当たりPB売上", "PB sales / SKU")]]
    for _, r in dp.iterrows():
        rows.append([r["dept"], J(r.sales), _signed_pct(r.yoy), J(r.pb_cur), f"{J(r.d_pb)} ({_signed_pct(r.pb_growth, 0)})",
                     f"{_pct(r.share_prev)}{ARROW}{_pct(r.share_cur)}", f"{r.sku_prev:.0f}{ARROW}{r.sku_cur:.0f}", J(r.per_sku_cur)])
    rows.append([T("ペット計", "Pet total"), J(k["sales_cur"]), _signed_pct(k["sales_yoy"]), J(k["pb_cur"]), f"{J(k['d_pb'])} ({_signed_pct(k['pb_yoy'], 0)})",
                 f"{_pct(k['share_prev'])}{ARROW}{_pct(k['share_cur'])}", f"{k['sku_prev']:.0f}{ARROW}{k['sku_cur']:.0f}", J(k["per_sku_cur"])])
    S.append(_table(rows, [FW * .15, FW * .11, FW * .1, FW * .11, FW * .16, FW * .17, FW * .1, FW * .1], st))
    if LAT:
        H2(T("3期の推移(部門別)", "Three periods by department"))
        pc = lambda a, b: _signed_pct(b / a - 1) if a else NA
        sa, pa = T("売上", "Sales"), T("PB売上", "PB sales")
        step = f"{cur}→{LAT}"
        rows = [[T("部門", "Department"), f"{sa} {prev}", f"{sa} {cur}", f"{sa} {LATH}", f"{pa} {prev}", f"{pa} {cur}", f"{pa} {LATH}"]]
        for _, r in dp.iterrows():
            rows.append([r["dept"], J(r.sales_prev), J(r.sales), J(r.sales_latest), J(r.pb_prev), J(r.pb_cur), J(r.pb_latest)])
        rows.append([T("ペット計", "Pet total"), J(k["sales_prev"]), J(k["sales_cur"]), J(k["sales_latest"]),
                     J(k["pb_prev"]), J(k["pb_cur"]), J(k["pb_latest"])])
        S.append(_table(rows, [FW * .2, FW * .13, FW * .13, FW * .14, FW * .13, FW * .13, FW * .14], st))
        w3 = [FW * .19, FW * .105, FW * .105, FW * .105, FW * .09, FW * .105, FW * .105, FW * .105, FW * .09]
        S.append(Spacer(1, 4))
        ra, ka = T("PB比率", "PB ratio"), T("PB SKU数", "PB SKUs")
        rows = [[T("部門", "Department"), f"{ra} {prev}", f"{ra} {cur}", f"{ra} {LATH}", f"{step} (pt)", f"{ka} {prev}", f"{ka} {cur}", f"{ka} {LATH}", step]]
        for _, r in dp.iterrows():
            rows.append([r["dept"], _pct(r.share_prev), _pct(r.share_cur), _pct(r.share_latest), f"{(r.share_latest - r.share_cur) * 100:+.1f}",
                         f"{r.sku_prev:.0f}", f"{r.sku_cur:.0f}", f"{r.sku_latest:.0f}", pc(r.sku_cur, r.sku_latest)])
        rows.append([T("ペット計", "Pet total"), _pct(k["share_prev"]), _pct(k["share_cur"]), _pct(k["share_latest"]),
                     f"{(k['share_latest'] - k['share_cur']) * 100:+.1f}", f"{k['sku_prev']:.0f}", f"{k['sku_cur']:.0f}", f"{k['sku_latest']:.0f}",
                     pc(k["sku_cur"], k["sku_latest"])])
        S.append(_table(rows, w3, st))
        NOTE(LATNOTE)
    H2(T("PB売上とPB比率の増減要因(部門別)", "What moved PB sales and the PB ratio, by department"))
    rows = [[T("部門", "Department"), T("PB売上の増減", "Change in PB sales"), T("うち市場成長", "of which market growth"),
             T("うちシェア変化", "of which share change"), T("PB比率への寄与: 部門内のシェア変化", "PB ratio pt: share change inside"),
             T("PB比率への寄与: 売上構成の変化", "PB ratio pt: sales mix")]]
    for _, r in dp.iterrows():
        rows.append([r["dept"], J(r.d_pb), J(r.market_eff), J(r.share_eff), f"{r.rate_pt:+.2f}pt", f"{r.mix_pt:+.2f}pt"])
    rows.append([T("ペット計", "Pet total"), J(b["d_pb"]), J(b["market"]), J(b["share"]), f"{b['rate_pt']:+.2f}pt", f"{b['mix_pt']:+.2f}pt"])
    S.append(_table(rows, [FW * .18, FW * .15, FW * .15, FW * .15, FW * .185, FW * .185], st))
    S.append(Spacer(1, 3))
    NOTE(T("市場成長 = 部門の売上増減 × 前期のPB比率。シェア変化 = 当期売上 × PB比率の増減。"
           "売上構成の変化は、PB比率の高い部門へ売上が移ると+、低い部門へ移ると−になります。",
           "Market growth = change in department sales x prior PB ratio. Share change = current sales x change in PB ratio. "
           "Sales mix is positive when sales move towards departments with a higher PB ratio, negative when they move to lower ones."))
    H2(T("方向性の見方(次ページ以降の部門別分析で使用)", "How to read the directions (used in the department pages that follow)"))
    for dname in ORDER:
        S.append(Paragraph(f'<font color="{DIR_COLOR[dname]}">{escape(D(dname))}</font>: {escape(BLURB[dname])}', st["small"]))

    NOTE(T("カテゴリー表の「判定」は、PB SKU数の変化とPB売上の変化の比較です"
           f"(SKU数の変化が±{t.sku_stable_band:.0%}以内は横ばい、PB売上{J(t.min_base_pb_sales)}未満の小規模なものは判定対象外)。",
           "'Verdict' in the category tables compares the change in PB SKU count with the change in PB sales "
           f"(SKU change within ±{t.sku_stable_band:.0%} counts as a stable range; PB sales under {J(t.min_base_pb_sales)} are not judged)."))

    # ---- 5. one page per department ----
    rank_dir = {key: i for i, key in enumerate(ORDER)}
    TREND_STYLE = {T_UP: ("↑", "#2E7D32"), T_FLAT: ("→", "#595959"), T_DOWN: ("↓", "#C62828")}

    def trend_cell(r):
        if not r.trend:
            return NA
        arrow, col = TREND_STYLE[r.trend]
        return Paragraph(f'<font color="{col}">{arrow} {escape(D(r.trend))} {r.mom_pt:+.1f}pt</font>', st["cell"])
    for i, (_, dr) in enumerate(dp.iterrows(), start=1):
        dept = dr["dept"]
        S.append(PageBreak())
        H1(T(f"5-{i}. 部門別分析: {dept}", f"5-{i}. Department analysis: {dept}"))
        share3 = f"{_pct(dr.share_prev)}{ARROW}{_pct(dr.share_cur)}" + (f"{ARROW}{_pct(dr.share_latest)}*" if LAT else "")
        sku3 = f"{dr.sku_prev:.0f}{ARROW}{dr.sku_cur:.0f}" + (f"{ARROW}{dr.sku_latest:.0f}*" if LAT else "")
        kp = [[T("売上", "Sales"), T("市場前年比", "Market YoY"), T("PB売上", "PB sales"), T("PB売上の増減", "Change in PB sales"),
               T("PB比率", "PB ratio"), T("PB SKU数", "PB SKUs"), T("SKU当たりPB売上", "PB sales / SKU")],
              [J(dr.sales), _signed_pct(dr.yoy), J(dr.pb_cur), f"{J(dr.d_pb)} ({_signed_pct(dr.pb_growth, 0)})",
               f"{share3} ({dr.d_share_pt:+.1f}pt)", sku3, f"{J(dr.per_sku_cur)} ({_signed_pct(dr.per_sku_chg, 0)})"]]
        S.append(_table(kp, [FW * .12, FW * .1, FW * .12, FW * .17, FW * .2, FW * .12, FW * .17], st, zebra=False))
        if LAT:
            NOTE(LATNOTE)
        S.append(Spacer(1, 4))
        for line in res.dept_findings.get(dept, []):
            BULLET(line)

        H2(T("カテゴリー別", "By category"))
        rows = [[T("カテゴリー", "Category"), T("PB売上(部門内順位)", "PB sales (rank in dept.)"), T("PBシェア(増減pt)", "PB share (Δ pt)"), T("PB成長率", "PB growth"),
                 T("市場成長率", "Market growth"), T("PB SKU数", "PB SKUs"), T("SKU当たりPB売上", "PB sales / SKU"), T("判定", "Verdict")]
                + ([T(f"{LATH} PBシェア / SKU数", f"{LATH} PB share / SKUs")] if LAT else [])]
        for _, r in res.cats[res.cats["dept"] == dept].sort_values("pb_cur", ascending=False).iterrows():
            rows.append([r["cat"], T(f"{J(r.pb_cur)}({r.rank_dept}位)", f"{J(r.pb_cur)} (#{r.rank_dept})"), f"{_pct(r.share_cur)} ({r.d_share_pt:+.1f})",
                         _signed_pct(r.pb_growth, 0), _signed_pct(r.yoy), f"{r.sku_prev:.0f}{ARROW}{r.sku_cur:.0f} ({_signed_pct(r.sku_growth, 0)})",
                         f"{J(r.per_sku_cur)} ({_signed_pct(r.per_sku_chg, 0)})", D(r.verdict)]
                        + ([f"{_pct(r.share_latest)} / {r.sku_latest:.0f}"] if LAT else []))
        cw = ([FW * .11, FW * .13, FW * .11, FW * .07, FW * .08, FW * .13, FW * .14, FW * .12, FW * .11] if LAT
              else [FW * .12, FW * .14, FW * .12, FW * .08, FW * .09, FW * .15, FW * .15, FW * .15])
        S.append(_table(rows, cw, st))

        H2(T("サブカテ別と方向性", "By sub-category, with direction"))
        rows = [[T("サブカテ", "Sub-category"), T("市場前年比", "Mkt YoY"), T("非PB売上", "Non-PB sales"),
                 T(f"PBシェア {prev}→{cur}", f"PB share {prev}→{cur}") + (f"{ARROW}{LATH}" if LAT else ""), T("PB売上", "PB sales")]
                + [T("PB成長率", "PB growth"), T("PB SKU数", "PB SKUs")] + ([T("シェア動向", "Share trend")] if LAT else [])
                + [T("方向性", "Direction")]]
        sd = res.subs[res.subs["dept"] == dept].copy()
        sd["_o"] = sd["direction"].map(rank_dir)
        for _, r in sd.sort_values(["_o", "nb_cur"], ascending=[True, False]).iterrows():
            lat = f"{ARROW}{r.share_latest * 100:.1f}%" if LAT and r.share_latest == r.share_latest else ""
            sku3 = f"{r.sku_prev:.0f}{ARROW}{r.sku_cur:.0f}" + (f"{ARROW}{r.sku_latest:.0f}" if LAT else "")
            rows.append([r["name"], _signed_pct(r.yoy), J(r.nb_cur), f"{r.share_prev * 100:.1f}%{ARROW}{r.share_cur * 100:.1f}%{lat}", J(r.pb_cur)]
                        + [_signed_pct(r.pb_growth, 0), sku3] + ([trend_cell(r)] if LAT else [])
                        + [Paragraph(f'<font color="{DIR_COLOR[r.direction]}">{escape(D(r.direction))}</font>', st["cell"])])
        sw = ([FW * .17, FW * .07, FW * .10, FW * .18, FW * .10, FW * .07, FW * .10, FW * .10, FW * .11] if LAT
              else [FW * .24, FW * .09, FW * .1, FW * .19, FW * .09, FW * .08, FW * .08, FW * .13])
        S.append(_table(rows, sw, st))
        if LAT:
            NOTE(T(f"「シェア動向」は、PBシェアの{LAT}期と{cur}期の差です(±{t.momentum_pt:g}pt以上で上昇・低下。{cur}期のPB売上が{J(t.min_base_pb_sales)}未満は表示しません)。",
                   f"'Share trend' is the change in PB share between {LAT}期 and {cur}期 (±{t.momentum_pt:g} pt or more counts as rising / falling; "
                   f"not shown where {cur}期 PB sales are under {J(t.min_base_pb_sales)})."))

    # ---- 6. upside ----
    S.append(PageBreak())
    H1(T("6. 伸びしろの試算(仮定に基づく)", "6. Illustrative upside (assumption-based)"))
    S.append(_p(T(f"{D(SCALE)}: PBシェアが現在のシェアに直近の上昇幅{t.upside_scale_momentum_years:g}年分を加えた水準に到達"
                  f"(上限{t.upside_scale_cap:.0%})。{D(RECAPTURE)}: 前期シェアに回復。{D(REPLICATE)}: シェア"
                  f"{t.upside_replicate_share:.0%}。{D(TEST)}は対象外。これは計画用の仮定であり予測ではありません。"
                  "売上はNB(ナショナルブランド)から獲得し、全体売上は横ばいと仮定しています。",
                  f"{D(SCALE)} items: PB share reaches current share plus {t.upside_scale_momentum_years:g} years of the last observed gain "
                  f"(capped at {t.upside_scale_cap:.0%}). {D(RECAPTURE)} items: previous-period share is restored. {D(REPLICATE)} items: "
                  f"{t.upside_replicate_share:.0%} share. {D(TEST)} items are excluded. These are planning assumptions, not forecasts, "
                  "and assume the sales come from national brands with total sales flat."), st["body"]))
    if not res.upside.empty:
        rows = [[T("サブカテ", "Sub-category"), T("部門", "Department"), T("方向性", "Direction"), T(f"売上 {cur}", f"Sales {cur}"),
                 T("PBシェア 現在→目標", "PB share now → target"), T("PB売上の増加額", "Added PB sales")]]
        for _, r in res.upside.head(12).iterrows():
            rows.append([r["name"], r["dept"], D(r.direction), J(r.sales), f"{r.share_cur * 100:.1f}% {ARROW} {r.target * 100:.1f}%", J(r["add"])])
        u = res.upside_total
        rows.append([T("合計(表示外の項目を含む)", "Total (all items incl. any not shown)"), "", "", "",
                     T(f"PB比率 {_pct(k['share_cur'])}→{_pct(u['share_after'])}", f"PB ratio {_pct(k['share_cur'])} → {_pct(u['share_after'])}"),
                     T(f"{J(u['add'])}(PB売上の{_signed_pct(u['pct_of_pb'], 0)})", f"{J(u['add'])} ({_signed_pct(u['pct_of_pb'], 0)} of PB sales)")])
        S.append(_table(rows, [FW * .27, FW * .13, FW * .1, FW * .12, FW * .18, FW * .2], st))

    # ---- 7. basis of this report ----
    H1(T("7. 前提・注意事項", "7. Basis and notes"))
    for line in key_notes(res):
        BULLET(line)

    doc.build(S, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()
