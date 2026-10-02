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
                       _jpy, thresholds_table)

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


def _styles():
    f = _register_font()
    # wordWrap="CJK": Japanese has no spaces, so lines must be allowed to break between any characters
    P = lambda name, **kw: ParagraphStyle(name, fontName=f, wordWrap="CJK", **kw)
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


def _pct(x, d=1):
    return "－" if x is None or x != x else f"{x * 100:.{d}f}%"


def _signed_pct(x, d=1):
    return "－" if x is None or x != x else f"{x * 100:+.{d}f}%"


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


def _footer(canvas, doc):
    f = _register_font()
    canvas.saveState()
    canvas.setFont(f, 7)
    canvas.setFillColor(GREY)
    canvas.drawString(15 * mm, 8 * mm, "PB戦略レポート(自動生成)")
    canvas.drawRightString(A4[0] - 15 * mm, 8 * mm, f"{doc.page}ページ")
    canvas.restoreState()


def _n(x, d=0):
    return "－" if x is None or x != x else f"{x:,.{d}f}"


def build_pdf(res: Result, source_name: str = "") -> bytes:
    st = _styles()
    f = _register_font()
    prev, cur = res.periods_full
    k, b, t = res.kpi, res.bridge, res.thresholds
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=14 * mm,
                            bottomMargin=16 * mm, title="ペットカテゴリー PB戦略レポート", author="PBレポート自動生成")
    FW = A4[0] - 30 * mm
    S = []
    H1 = lambda txt: S.append(_p(txt, st["h1"]))
    H2 = lambda txt: S.append(_p(txt, st["h2"]))
    NOTE = lambda txt: S.append(_p(txt, st["small"]))
    BODY = lambda txt: S.append(_p(txt, st["body"]))

    # ---- title ----
    S.append(_p("ペットカテゴリー PB戦略レポート", st["title"]))
    sub = f"比較期間: {prev}期 vs {cur}期" + (f"(PB比率とSKU数は{res.period_latest}期" +
          (f"・{res.partial_months}ヶ月分" if res.partial_months else "・途中期") + "も表示)" if res.period_latest else "")
    S.append(_p(sub, st["sub"]))
    S.append(_p(f"作成日時 {dt.datetime.now():%Y-%m-%d %H:%M}" + (f"  |  元データ: {source_name}" if source_name else ""), st["sub"]))
    S.append(_p(res.scope_note, st["sub"]))
    S.append(Spacer(1, 4))

    # ---- 1. takeaways ----
    H1("1. 要点")
    for line in res.findings:
        S.append(Paragraph(escape(line), st["bullet"], bulletText="•"))
    rows = [["指標", f"{prev}期", f"{cur}期", "増減"],
            ["全体売上", _jpy(k["sales_prev"]), _jpy(k["sales_cur"]), _signed_pct(k["sales_yoy"])],
            ["全体売れ数", "", "", _signed_pct(k["units_yoy"])],
            ["平均単価(価格・構成)", "", "", _signed_pct(k["asp_chg"])],
            ["PB売上", _jpy(k["pb_prev"]), _jpy(k["pb_cur"]), f"{_jpy(k['d_pb'])}({_signed_pct(k['pb_yoy'])})"],
            ["非PB売上", _jpy(k["sales_prev"] - k["pb_prev"]), _jpy(k["sales_cur"] - k["pb_cur"]), _signed_pct(k["nb_yoy"])],
            ["PB売れ数 vs 非PB売れ数", "", "", f"{_signed_pct(k['pb_units_yoy'])} vs {_signed_pct(k['nb_units_yoy'])}"],
            ["PB比率(売上)", _pct(k["share_prev"], 2), _pct(k["share_cur"], 2), f"{(k['share_cur'] - k['share_prev']) * 100:+.2f}pt"],
            ["PB SKU数", f"{k['sku_prev']:.0f}", f"{k['sku_cur']:.0f}", _signed_pct(k["sku_growth"], 0)],
            ["PB SKU当たりPB売上", _jpy(k["per_sku_prev"]), _jpy(k["per_sku_cur"]), _signed_pct(k["per_sku_chg"], 0)]]
    if res.period_latest:
        rows.append([f"{res.period_latest}期(直近・途中期): PB比率 / PB SKU数", "", f"{_pct(k.get('share_latest'), 2)} / {k.get('sku_latest', float('nan')):.0f}",
                     f"{cur}期比 {(k['share_latest'] - k['share_cur']) * 100:+.2f}pt"])
    S.append(Spacer(1, 4))
    S.append(_table(rows, [FW * .38, FW * .17, FW * .17, FW * .28], st))

    # ---- 2. growth bridge ----
    H1("2. PB成長の要因")
    BODY(f"PB売上は{_jpy(b['d_pb'])}変化しました。各サブカテの{prev}期PBシェアを固定した場合、市場成長だけで"
         f"{_jpy(b['market'])}増加し、残りの{_jpy(b['share'])}はシェアの純変化によるものです。両者の合計は増減額と完全に一致します。")
    items = [("前期PBシェアでの市場成長", b["market"], BAR_BLUE, _jpy(b["market"])),
             ("シェア上昇(上昇したサブカテ)", b["share_gain"], GREEN, _jpy(b["share_gain"])),
             ("シェア低下(低下したサブカテ)", b["share_loss"], RED, _jpy(b["share_loss"])),
             ("PB売上の純増減", b["d_pb"], NAVY, _jpy(b["d_pb"]))]
    S.append(_bar_chart(items, label_w_mm=70, title=f"PB売上の増減要因 {prev}期→{cur}期", signed=True))
    S.append(Spacer(1, 4))
    BODY(f"PB比率 {b['ratio_prev']:.2%}→{b['ratio_cur']:.2%}({b['d_ratio_pt']:+.2f}pt)= サブカテ内のシェア変化 {b['rate_pt']:+.2f}pt "
         f"+ 構成変化 {b['mix_pt']:+.2f}pt(PB比率の高い/低いサブカテ間での売上の移動)。")
    cb = res.cats.sort_values("d_pb", ascending=False)
    rows = [["カテゴリー", "PB売上の増減", "市場要因", "シェア要因", "PB比率pt: シェア", "PB比率pt: 構成"]]
    for _, r in cb.iterrows():
        if abs(r.d_pb) < 1e5 and abs(r.market_eff) < 1e5:
            continue
        rows.append([r["cat"], _jpy(r.d_pb), _jpy(r.market_eff), _jpy(r.share_eff), f"{r.rate_pt:+.2f}", f"{r.mix_pt:+.2f}"])
    rows.append(["合計", _jpy(b["d_pb"]), _jpy(b["market"]), _jpy(b["share"]), f"{b['rate_pt']:+.2f}", f"{b['mix_pt']:+.2f}"])
    S.append(_table(rows, [FW * .22, FW * .18, FW * .16, FW * .16, FW * .14, FW * .14], st))
    cn = res.concentration
    S.append(Spacer(1, 3))
    NOTE(f"集中度: 上位3サブカテ = PB売上の{cn['top3']:.0%}({prev}期は{cn['top3_prev']:.0%})、上位5 = {cn['top5']:.0%}。"
         f"PB展開中の{cn['n_active']}サブカテのうち{cn['n_for_80']}サブカテで80%を占めます(HHI {cn['hhi']:.3f})。")

    # ---- 3. movers + pools ----
    H1("3. PB売上の増減と伸びしろ")
    sub_df = res.subs
    mv = sub_df.sort_values("d_pb", ascending=False)
    gains, losses = mv[mv.d_pb > 0].head(8), mv[mv.d_pb < 0].tail(5)
    items = [(r["name"], r.d_pb, GREEN, _jpy(r.d_pb)) for _, r in gains.iterrows()] + \
            [(r["name"], r.d_pb, RED, _jpy(r.d_pb)) for _, r in losses.sort_values("d_pb").iterrows()]
    if items:
        S.append(_bar_chart(items, title=f"PB売上の増減 {prev}期→{cur}期(増減の大きいサブカテ)", signed=True))
    S.append(Spacer(1, 6))
    pools = sub_df.sort_values("nb_cur", ascending=False).head(10)
    items = [(r["name"], r.nb_cur, BAR_BLUE, f"{_jpy(r.nb_cur)}  (PB {r.share_cur * 100:.1f}%)") for _, r in pools.iterrows()]
    S.append(_bar_chart(items, label_w_mm=62, title=f"非PB売上の大きいサブカテ {cur}期"))

    # ---- 4. category scorecard + SKU vs sales ----
    S.append(PageBreak())
    H1("4. カテゴリー別スコアカード: 成長・順位・SKUと売上")
    BODY("順位はPB売上順(1 = 最大)。判定はPB SKU数の変化とPB売上の変化を比較したものです"
         f"(SKU変化が±{t.sku_stable_band:.0%}以内は横ばい、PB売上{_jpy(t.min_base_pb_sales)}未満の小規模なものは判定対象外)。")
    rows = [["カテゴリー", "PB売上(順位)", "PBシェア(増減pt)", "PB成長率", "市場成長率", "PB SKU数", "SKU当たりPB売上", "判定"]]
    for _, r in res.cats.iterrows():
        rows.append([r["cat"], f"{_jpy(r.pb_cur)}({r.rank_pb}位)", f"{_pct(r.share_cur)}({r.d_share_pt:+.1f})", _signed_pct(r.pb_growth, 0), _signed_pct(r.yoy),
                     f"{r.sku_prev:.0f}→{r.sku_cur:.0f}({_signed_pct(r.sku_growth, 0)})",
                     f"{_jpy(r.per_sku_cur)}({_signed_pct(r.per_sku_chg, 0)})", r.verdict])
    S.append(_table(rows, [FW * .12, FW * .13, FW * .13, FW * .08, FW * .09, FW * .15, FW * .15, FW * .15], st))
    S.append(Spacer(1, 3))
    NOTE("PB成長率 = PB売上の前期比。市場成長率 = 全体(PB+非PB)売上の前期比。"
         "追加SKU1つ当たりのPB売上増: 全体で" + (_jpy(k["marg_per_sku"]) if k["marg_per_sku"] == k["marg_per_sku"] else "－") +
         f"(参考: {prev}期の既存SKU1つ当たり{_jpy(k['per_sku_prev'])})。")

    H2("サブカテ別 PB SKU数の変化とPB売上の変化(一定規模以上)")
    sv = sub_df[(sub_df.verdict != V_SMALL)].sort_values("pb_cur", ascending=False).head(16)
    rows = [["サブカテ", "PB SKU数", "PB売上", "PB売上成長率", "追加SKU当たりPB売上増", "判定"]]
    for _, r in sv.iterrows():
        rows.append([r["name"], f"{r.sku_prev:.0f}→{r.sku_cur:.0f}", f"{_jpy(r.pb_prev)}→{_jpy(r.pb_cur)}", _signed_pct(r.pb_growth, 0),
                     _jpy(r.marg_per_sku), r.verdict])
    S.append(_table(rows, [FW * .27, FW * .1, FW * .2, FW * .11, FW * .14, FW * .18], st))

    # ---- 5. rankings ----
    H1("5. ランキング(サブカテ)")
    rk = res.rankings
    def rank_table(title, df_, cols, widths, hdr_):
        H2(title)
        rows_ = [hdr_]
        for _, r in df_.iterrows():
            rows_.append([r["name"]] + [fn(r) for fn in cols])
        S.append(_table(rows_, widths, st))
    rank_table("PB売上 上位10", rk["sub_pb"],
               [lambda r: _jpy(r.pb_cur), lambda r: _pct(r.share_cur), lambda r: _signed_pct(r.pb_growth, 0),
                lambda r: f"{r.sku_cur:.0f}", lambda r: _jpy(r.per_sku_cur)],
               [FW * .33, FW * .14, FW * .13, FW * .13, FW * .1, FW * .17], ["サブカテ", "PB売上", "PBシェア", "PB成長率", "SKU数", "SKU当たりPB売上"])
    rank_table(f"PB売上成長率 上位(前期PB売上{_jpy(t.min_base_pb_sales)}以上)", rk["growth_top"],
               [lambda r: _signed_pct(r.pb_growth, 0), lambda r: _jpy(r.d_pb), lambda r: f"{_pct(r.share_prev)}→{_pct(r.share_cur)}"],
               [FW * .36, FW * .16, FW * .16, FW * .32], ["サブカテ", "PB成長率", "増減額", "PBシェア"])
    rank_table(f"PB売上成長率 下位(前期PB売上{_jpy(t.min_base_pb_sales)}以上)", rk["growth_bottom"],
               [lambda r: _signed_pct(r.pb_growth, 0), lambda r: _jpy(r.d_pb), lambda r: f"{_pct(r.share_prev)}→{_pct(r.share_cur)}"],
               [FW * .36, FW * .16, FW * .16, FW * .32], ["サブカテ", "PB成長率", "増減額", "PBシェア"])
    rank_table("SKU当たりPB売上 上位(2SKU以上)", rk["per_sku_top"],
               [lambda r: _jpy(r.per_sku_cur), lambda r: f"{r.sku_cur:.0f}", lambda r: _signed_pct(r.per_sku_chg, 0)],
               [FW * .36, FW * .2, FW * .2, FW * .24], ["サブカテ", "SKU当たりPB売上", "PB SKU数", "増減"])
    rank_table("SKU当たりPB売上 下位(2SKU以上)", rk["per_sku_bottom"],
               [lambda r: _jpy(r.per_sku_cur), lambda r: f"{r.sku_cur:.0f}", lambda r: _signed_pct(r.per_sku_chg, 0)],
               [FW * .36, FW * .2, FW * .2, FW * .24], ["サブカテ", "SKU当たりPB売上", "PB SKU数", "増減"])
    if res.period_latest and res.watchlist is not None and len(res.watchlist):
        H2(f"要注視リスト: PBシェアが2期連続で低下({prev}→{cur}→{res.period_latest})")
        rows = [["サブカテ", "PB売上", "PBシェアの推移"]]
        for _, r in res.watchlist.iterrows():
            rows.append([r["name"], _jpy(r.pb_cur), f"{_pct(r.share_prev)}→{_pct(r.share_cur)}→{_pct(r.share_latest)}"])
        S.append(_table(rows, [FW * .4, FW * .2, FW * .4], st))
        NOTE(f"{res.period_latest}期は途中期のため、シェアは参考値です(季節性)。")

    # ---- 6. directions ----
    S.append(PageBreak())
    H1("6. 方向性の提案(データに基づく。Trial / トーエーの記載は不使用)")
    BODY("各サブカテは明示的なルール(付録参照)により1つの方向性に分類しています。非PB = 直近の通期における全体売上 − PB売上。")
    hdr = ["サブカテ", f"売上 {cur}", "非PB", f"PBシェア {prev}→{cur}" + (f"→{res.period_latest}" if res.period_latest else ""), "市場前年比"]
    wd = [FW * .34, FW * .13, FW * .13, FW * .28, FW * .12]
    for dname in ORDER:
        x = res.directions[dname]
        if x.empty or dname == MONITOR:
            continue
        block = [_p(f"{dname}({len(x)}件)", st["h2"]), _p(DIR_BLURB[dname], st["small"]), Spacer(1, 2)]
        if dname == DEPRIORITIZE:
            block.append(_p("、".join(x["name"].tolist()), st["body"]))
            S.append(KeepTogether(block))
            continue
        show = x.head(10)
        rows = [hdr]
        for _, r in show.iterrows():
            lat = f"→{r.share_latest * 100:.1f}%" if r.share_latest == r.share_latest else ""
            rows.append([r["name"], _jpy(r.sales), _jpy(r.nb_cur), f"{r.share_prev * 100:.1f}%→{r.share_cur * 100:.1f}%{lat}", _signed_pct(r.yoy)])
        block.append(_table(rows, wd, st))
        if len(x) > len(show):
            block.append(_p(f"ほか{len(x) - len(show)}件(付録参照)", st["small"]))
        S.append(KeepTogether(block))
    if res.borderline is not None and len(res.borderline):
        H2("境界線上の分類")
        NOTE("以下は、基準値を概ね±25%(伸びしろ)、±1pt(シェア上昇)、±60%(PB未展開シェア)、±10pt(防衛シェア)動かすと方向性が変わります。"
             "判断が必要な項目として扱ってください: " + "/".join(f"{r['name']}({r.base}→{r.alternatives})" for _, r in res.borderline.iterrows()) + "。")

    # ---- 7. upside ----
    H1("7. 伸びしろの試算(仮定に基づく)")
    S.append(_p(f"拡大: PBシェアが現在のシェアに直近の上昇幅{t.upside_scale_momentum_years:g}年分を加えた水準に到達"
                f"(上限{t.upside_scale_cap:.0%})。シェア奪回: 前期シェアに回復。横展開: シェア"
                f"{t.upside_replicate_share:.0%}。小規模テストは対象外。これは計画用の仮定であり予測ではありません。"
                "売上はNB(ナショナルブランド)から獲得し、全体売上は横ばいと仮定しています。", st["body"]))
    if not res.upside.empty:
        rows = [["サブカテ", "方向性", f"売上 {cur}", "PBシェア 現在→目標", "PB売上の増加額"]]
        for _, r in res.upside.head(12).iterrows():
            rows.append([r["name"], r.direction, _jpy(r.sales), f"{r.share_cur * 100:.1f}%→{r.target * 100:.1f}%", _jpy(r["add"])])
        u = res.upside_total
        rows.append(["合計(表示外の項目を含む)", "", "", f"PB比率 {_pct(k['share_cur'])}→{_pct(u['share_after'])}",
                     f"{_jpy(u['add'])}(PB売上の{_signed_pct(u['pct_of_pb'], 0)})"])
        S.append(_table(rows, [FW * .34, FW * .13, FW * .13, FW * .2, FW * .2], st))

    # ---- 8. caveats + data checks ----
    H1("8. 会議前に確認すべき注意点")
    for c in res.caveats:
        S.append(Paragraph(escape(c), st["bullet"], bulletText="•"))
    H1("9. データチェックと照合")
    for sev, msg in res.dq:
        S.append(Paragraph(escape(("要確認: " if sev == "warn" else "") + msg), st["bullet"], bulletText="•" if sev == "info" else "!"))
    if res.recon is not None and len(res.recon):
        rc = res.recon[res.recon.period == cur]
        rows = [["区分", f"元資料の売上 {cur}", "算出した売上", "元資料のPB比率", "算出したPB比率", "一致"]]
        for _, r in rc.iterrows():
            rows.append([r.scope, _jpy(r.rep_sales), _jpy(r.calc_sales), _pct(r.rep_ratio, 2), _pct(r.calc_ratio, 2), "○" if r.ok else "×"])
        S.append(Spacer(1, 3))
        S.append(_table(rows, [FW * .24, FW * .19, FW * .19, FW * .13, FW * .15, FW * .1], st))

    # ---- appendix ----
    S.append(PageBreak())
    S.append(_p("付録A. 使用したルールと基準値", st["h1"]))
    S.append(_p(f"ルールは次の順に適用し、最初に該当したものを採用します: {RECAPTURE}→{REVIEW}→{SCALE}→{DEFEND}→{REPLICATE} / {TEST}→"
                f"{DEPRIORITIZE}→{MONITOR}。基準値はアプリのサイドバーまたはコマンドラインで変更できます。", st["body"]))
    tt = thresholds_table(res.thresholds)
    half = (len(tt) + 1) // 2
    rows = [["項目", "値", "項目", "値"]]
    for i in range(half):
        a = tt[i]
        b2 = tt[i + half] if i + half < len(tt) else ("", "")
        rows.append([a[0], a[1], b2[0], b2[1]])
    S.append(_table(rows, [FW * .32, FW * .18, FW * .32, FW * .18], st))
    S.append(_p("算出方法", st["h2"]))
    for line in [
        "PB売上 = 全体売上 × PB比率(売上ベース)。非PB売上 = 全体 − PB。PB売れ数 = 全体売れ数 × PB比率(売れ数ベース)。",
        "増減要因: PB売上の増減 =(売上の増減 × 前期PBシェア)+(当期売上 × PBシェアの増減)をサブカテ全体で合計。誤差なしで完全に一致します。",
        "PB比率の要因分解: シェア要因 = 平均売上構成比 × シェア変化。構成要因 = 平均PBシェア × 売上構成比の変化。誤差なしで完全に一致します。",
        "追加SKU当たりPB売上増 = PB売上の増減 ÷ PB SKU数の増減(SKUが増えた場合のみ)。",
    ]:
        S.append(Paragraph(escape(line), st["bullet"], bulletText="•"))

    S.append(_p("付録B. 全サブカテ一覧", st["h1"]))
    rows = [["方向性", "サブカテ", f"売上 {cur}", "非PB", f"PBシェア {prev}→{cur}", "市場前年比", "PB SKU数"]]
    for dname in ORDER:
        for _, r in res.directions[dname].iterrows():
            rows.append([dname, r["name"], _jpy(r.sales), _jpy(r.nb_cur), f"{r.share_prev * 100:.1f}%→{r.share_cur * 100:.1f}%",
                         _signed_pct(r.yoy), f"{r.sku_prev:.0f}→{r.sku_cur:.0f}"])
    S.append(_table(rows, [FW * .12, FW * .32, FW * .11, FW * .11, FW * .15, FW * .09, FW * .1], st))

    doc.build(S, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()
