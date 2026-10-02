"""Renders a Result into a native, editable PowerPoint deck (python-pptx).

* real slide titles (title placeholders), native charts and native tables - nothing is a picture,
  so the team can restyle and edit the deck in PowerPoint
* 16:9, Calibri for Latin text and Yu Gothic as the East Asian font for the Japanese category names
* speaker notes on every slide
"""
from __future__ import annotations

import datetime as dt
import io

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .analysis import (DEFEND, DEPRIORITIZE, MONITOR, RECAPTURE, REPLICATE, REVIEW, SCALE, TEST, V_SMALL, Result, _jpy,
                       key_notes)
from .pdf import DIR_BLURB

# palette: deep teal dominant, warm amber as the single sharp accent
DARK, TEAL, ACCENT = "0E3B43", "1C7C7D", "F2A33A"
LIGHT, CARD, TEXT, MUTED = "F3F6F6", "E6EEEE", "1F2933", "5F6C72"
GREEN, RED = "2E7D32", "C62828"
DIR_COLOR = {SCALE: "2E7D32", REPLICATE: "1C7C7D", RECAPTURE: "C77700", TEST: "6A4C93", REVIEW: "B71C1C", DEFEND: "455A64"}
DIR_TITLE = {
    SCALE: "拡大: PBが伸びており、まだ伸びしろがある",
    REPLICATE: "横展開: 実績のある近接サブカテの成功パターンを活用",
    RECAPTURE: "シェア奪回: PBシェアが前期から大きく低下",
    TEST: "小規模テスト: PB未展開の大きな市場",
    REVIEW: "要点検: PBシェアが低下傾向",
    DEFEND: "防衛: PBシェアが高く、SKUの希薄化を避ける",
}
LATIN, EA = "Calibri", "Yu Gothic"
SW, SH = 13.333, 7.5
MX = 0.6  # side margin (inches)


def _rgb(h):
    return RGBColor.from_string(h)


def _style(run, size, color=TEXT, bold=False, italic=False):
    f = run.font
    f.size, f.bold, f.italic = Pt(size), bold, italic
    f.color.rgb = _rgb(color)
    f.name = LATIN
    rPr = run._r.get_or_add_rPr()
    for old in rPr.findall(qn("a:ea")):
        rPr.remove(old)
    rPr.append(rPr.makeelement(qn("a:ea"), {"typeface": EA}))


def _bullet(paragraph, char="•", indent=0.22):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(indent))))
    pPr.set("indent", str(-int(Inches(indent))))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for e in pPr.findall(qn(tag)):
            pPr.remove(e)
    bu = pPr.makeelement(qn("a:buChar"), {"char": char})
    pPr.append(bu)


def _text(slide, x, y, w, h, text, size=14, color=TEXT, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
          name=None, italic=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap, tf.auto_size, tf.vertical_anchor = True, MSO_AUTO_SIZE.NONE, anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    _style(r, size, color, bold, italic)
    return tb


def _bullets(slide, x, y, w, h, items, size=15, color=TEXT, gap=8, name=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap, tf.auto_size = True, MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        _bullet(p)
        r = p.add_run()
        r.text = it
        _style(r, size, color)
    return tb


def _box(slide, x, y, w, h, fill=CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE, name=None, radius=0.06):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    s.fill.solid()
    s.fill.fore_color.rgb = _rgb(fill)
    s.line.fill.background()
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    return s


def _title(slide, text, color=DARK):
    t = slide.shapes.title
    t.left, t.top, t.width, t.height = Inches(MX), Inches(0.38), Inches(SW - 2 * MX), Inches(0.95)
    tf = t.text_frame
    tf.word_wrap, tf.auto_size, tf.vertical_anchor = True, MSO_AUTO_SIZE.NONE, MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    r = p.add_run()
    r.text = text
    _style(r, 28, color, bold=True)
    return t


def _notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def _table(slide, rows, widths, x, y, row_h=0.42, size=11, right_cols=(), header_fill=DARK, name="Table"):
    n_r, n_c = len(rows), len(rows[0])
    gf = slide.shapes.add_table(n_r, n_c, Inches(x), Inches(y), Inches(sum(widths)), Inches(row_h * n_r))
    gf.name = name
    tbl = gf.table
    for j, w in enumerate(widths):
        tbl.columns[j].width = Inches(w)
    for i in range(n_r):
        tbl.rows[i].height = Inches(row_h)
        for j in range(n_c):
            cell = tbl.cell(i, j)
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(header_fill if i == 0 else ("FFFFFF" if i % 2 else LIGHT))
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT if j in right_cols else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = str(rows[i][j])
            _style(r, size, "FFFFFF" if i == 0 else TEXT, bold=(i == 0 or (j == 0 and False)))
    return gf


def _bar_chart(slide, cats, vals, colors, x, y, w, h, name, size=11, label_fmt='#,##0'):
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("値", vals)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    gf.name = name
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    ch.font.size, ch.font.name = Pt(size), LATIN
    ch.font.color.rgb = _rgb(TEXT)
    plot = ch.plots[0]
    plot.gap_width = 45
    plot.vary_by_categories = False
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format, dl.number_format_is_linked = label_fmt, False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size, dl.font.color.rgb = Pt(size), _rgb(MUTED)
    ser = plot.series[0]
    ser.invert_if_negative = False   # otherwise negative bars render white
    for i, c in enumerate(colors):
        pt = ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = _rgb(c)
        dpt = ser._element.get_or_add_dPt_for_point(i)   # per-point fill needs its own flag, too
        if dpt.find(qn("c:invertIfNegative")) is None:
            dpt.idx.addnext(dpt.makeelement(qn("c:invertIfNegative"), {"val": "0"}))
    ca, va = ch.category_axis, ch.value_axis
    ca.reverse_order = True
    ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    ca.format.line.fill.background()
    ca.has_major_gridlines = False
    va.visible = False
    va.has_major_gridlines = False
    return gf


def _pct(x, d=1):
    return "－" if x is None or x != x else f"{x * 100:.{d}f}%"


def _spct(x, d=1):
    return "－" if x is None or x != x else f"{x * 100:+.{d}f}%"


def build_pptx(res: Result, source_name: str = "") -> bytes:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
    L_TITLE, L_TONLY = prs.slide_layouts[0], prs.slide_layouts[5]
    prev, cur = res.periods_full
    k = res.kpi
    d = res.subs
    prs.core_properties.title = "ペットカテゴリー PB戦略レポート"
    prs.core_properties.author = "PB戦略レポート"

    # ---------------------------------------------------------------- 1. title
    s = prs.slides.add_slide(L_TITLE)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = _rgb(DARK)
    t, sub = s.shapes.title, s.placeholders[1]
    t.left, t.top, t.width, t.height = Inches(0.9), Inches(1.7), Inches(11.5), Inches(1.7)
    t.text_frame.word_wrap = True
    t.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    p = t.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = "ペットカテゴリー PB戦略レポート"
    _style(r, 44, "FFFFFF", bold=True)
    sub.left, sub.top, sub.width, sub.height = Inches(0.9), Inches(3.5), Inches(11.5), Inches(0.9)
    sub.text_frame.word_wrap = True
    sp = sub.text_frame.paragraphs[0]
    sp.alignment = PP_ALIGN.LEFT
    r = sp.add_run()
    extra = f"  |  PB比率の推移は{res.period_latest}期も表示" if res.period_latest else ""
    r.text = f"{prev}期 vs {cur}期{extra}"
    _style(r, 20, "BFD9DA")
    sp2 = sub.text_frame.add_paragraph()
    sp2.alignment = PP_ALIGN.LEFT
    r = sp2.add_run()
    r.text = f"作成日 {dt.datetime.now():%Y-%m-%d}" + (f"  |  元データ: {source_name}" if source_name else "")
    _style(r, 13, "8FB5B7")
    chips = [("PB売上", _jpy(k["pb_cur"]), _spct(k["pb_yoy"])),
             ("PB比率", _pct(k["share_cur"]), f"{(k['share_cur'] - k['share_prev']) * 100:+.1f}pt"),
             ("PB SKU数", f"{k['sku_cur']:.0f}", _spct(k["sku_growth"], 0))]
    for i, (lab, val, chg) in enumerate(chips):
        x = 0.9 + i * 3.0
        _box(s, x, 5.0, 2.7, 1.35, fill=TEAL, name=f"KPI chip {i + 1}")
        _text(s, x + 0.2, 5.12, 2.3, 0.3, lab, 12, "D6EBEC")
        _text(s, x + 0.2, 5.42, 2.3, 0.55, val, 28, "FFFFFF", bold=True)
        _text(s, x + 0.2, 5.98, 2.3, 0.3, f"{prev}期比 {chg}", 12, ACCENT, bold=True)
    _notes(s, f"PB戦略資料。比較期間: {prev}期 vs {cur}期。トライアルの方向性の記載は使用していません。")

    # ---------------------------------------------------------------- 2. takeaways
    s = prs.slides.add_slide(L_TONLY)
    top_cat = res.cats.sort_values("d_pb", ascending=False).iloc[0]
    share_net = top_cat.d_pb / k["d_pb"] if k["d_pb"] > 0 else None
    _title(s, f"PB成長の{share_net:.0%}を{top_cat['cat']}が牽引" if share_net and share_net > 0.3
           else "PB実績の概要")
    cards = [("PB売上", _jpy(k["pb_cur"]), f"{prev}期比 {_spct(k['pb_yoy'])}({_jpy(k['d_pb'])})"),
             ("PB比率(売上)", _pct(k["share_cur"]), f"{prev}期比 {(k['share_cur'] - k['share_prev']) * 100:+.2f}pt"),
             ("PB SKU当たりPB売上", _jpy(k["per_sku_cur"]), f"{prev}期比 {_spct(k['per_sku_chg'], 0)}(SKU数 {_spct(k['sku_growth'], 0)})")]
    cw = (SW - 2 * MX - 0.5) / 3
    for i, (lab, val, chg) in enumerate(cards):
        x = MX + i * (cw + 0.25)
        _box(s, x, 1.55, cw, 1.75, name=f"Stat card {i + 1}")
        _text(s, x + 0.25, 1.72, cw - 0.5, 0.3, lab, 13, MUTED)
        _text(s, x + 0.25, 2.05, cw - 0.5, 0.7, val, 36, DARK, bold=True)
        bad = i == 2 and (k["per_sku_chg"] or 0) < 0
        _text(s, x + 0.25, 2.8, cw - 0.5, 0.35, chg, 12, RED if bad else TEAL, bold=True)
    _bullets(s, MX, 3.6, SW - 2 * MX, 3.6, res.findings[:6], size=13, gap=6, name="Findings")
    _notes(s, "\n".join(res.findings))

    # ---------------------------------------------------------------- 2b. growth bridge
    s = prs.slides.add_slide(L_TONLY)
    b = res.bridge
    if b["d_pb"] > 0:
        _title(s, f"PB成長の{b['market'] / b['d_pb']:.0%}は市場成長、残りはシェア上昇による")
    else:
        _title(s, "PB売上の増減要因")
    _text(s, MX, 1.28, 8, 0.25, f"PB売上の増減要因 {prev}期→{cur}期(百万円)", 11, MUTED)
    _bar_chart(s, ["前期PBシェアでの市場成長", "シェア上昇(上昇したサブカテ)",
                   "シェア低下(低下したサブカテ)", "PB売上の純増減"],
               [round(b["market"] / 1e6, 1), round(b["share_gain"] / 1e6, 1), round(b["share_loss"] / 1e6, 1), round(b["d_pb"] / 1e6, 1)],
               [TEAL, GREEN, RED, DARK], MX, 1.55, 7.9, 3.6, "PB sales bridge chart", size=12)
    _box(s, 8.95, 1.55, SW - 8.95 - MX, 4.0, name="Bridge notes card")
    bn = [f"PB比率 {b['ratio_prev']:.2%}→{b['ratio_cur']:.2%}({b['d_ratio_pt']:+.2f}pt)。",
          f"うちサブカテ内のシェア変化が{b['rate_pt']:+.2f}pt、サブカテ間の売上構成の変化が{b['mix_pt']:+.2f}pt。",
          f"売れ数: PB {_spct(k['pb_units_yoy'])}、非PB {_spct(k['nb_units_yoy'])}、全体 {_spct(k['units_yoy'])}。全体売上は{_spct(k['sales_yoy'])}で、市場成長は価格・構成要因。"]
    _text(s, 9.2, 1.8, SW - 8.95 - MX - 0.5, 0.4, "グラフの見方", 16, DARK, bold=True)
    _bullets(s, 9.2, 2.35, SW - 8.95 - MX - 0.5, 3.1, bn, size=12, name="Bridge notes")
    cn = res.concentration
    _text(s, MX, 5.45, 7.9, 1.2, f"集中度: 上位3サブカテでPB売上の{cn['top3']:.0%}({prev}期は{cn['top3_prev']:.0%})。"
          f"PB展開中の{cn['n_active']}サブカテのうち{cn['n_for_80']}サブカテで80%。市場要因 = 売上の増減 × 前期PBシェア、シェア要因 = 当期売上 × PBシェアの増減。"
          "両者の合計はPB売上の増減と完全に一致します。", 12, MUTED)
    _notes(s, " ".join(bn))

    # ---------------------------------------------------------------- 3. movers
    s = prs.slides.add_slide(L_TONLY)
    mv = d.sort_values("d_pb", ascending=False)
    gains, losses = mv[mv.d_pb > 0].head(8), mv[mv.d_pb < 0].tail(5).sort_values("d_pb")
    sel = list(gains.iterrows()) + list(losses.iterrows())
    _title(s, f"PB売上の増減 {prev}期→{cur}期")
    if sel:
        _bar_chart(s, [r["name"] for _, r in sel], [round(r.d_pb / 1e6, 1) for _, r in sel],
                   [GREEN if r.d_pb > 0 else RED for _, r in sel], MX, 1.5, 8.4, 5.4, "PB sales change chart")
    _text(s, MX, 1.28, 8, 0.25, "PB売上の増減(百万円、増減の大きいサブカテ)", 11, MUTED)
    _box(s, 9.3, 1.55, SW - 9.3 - MX, 4.0, name="Mover notes card")
    notes = []
    if len(gains):
        g = gains.iloc[0]
        notes.append(f"最大の増加: {g['name']}({_jpy(g.d_pb)})、PBシェア {_pct(g.share_prev)}→{_pct(g.share_cur)}。")
    if len(losses):
        l = losses.iloc[0]
        notes.append(f"最大の減少: {l['name']}({_jpy(l.d_pb)})、PBシェア {_pct(l.share_prev)}→{_pct(l.share_cur)}。")
    gross_loss = float(d.loc[d.d_pb < 0, "d_pb"].sum())
    gross_gain = float(d.loc[d.d_pb > 0, "d_pb"].sum())
    notes.append(f"増加の総額 {_jpy(gross_gain)}、減少の総額 {_jpy(gross_loss)}、純増減 {_jpy(k['d_pb'])}。")
    _text(s, 9.55, 1.8, SW - 9.3 - MX - 0.5, 0.4, "グラフの見方", 16, DARK, bold=True)
    _bullets(s, 9.55, 2.35, SW - 9.3 - MX - 0.5, 3.1, notes, size=12, name="Mover notes")
    _notes(s, "棒グラフは2つの通期間のPB売上(シェア × 全体売上)の増減です。" + " ".join(notes))

    # ---------------------------------------------------------------- 4. non-PB pools
    s = prs.slides.add_slide(L_TONLY)
    pools = d.sort_values("nb_cur", ascending=False).head(10)
    _title(s, "非PB売上の大きい市場にPBの伸びしろがある")
    _text(s, MX, 1.28, 9, 0.25, f"{cur}期の非PB売上(百万円、括弧内はPBシェア)", 11, MUTED)
    _bar_chart(s, [f"{r['name']}  (PB {r.share_cur * 100:.1f}%)" for _, r in pools.iterrows()],
               [round(r.nb_cur / 1e6, 1) for _, r in pools.iterrows()], [TEAL] * len(pools),
               MX, 1.5, 8.6, 5.4, "Non-PB pools chart")
    _box(s, 9.6, 1.55, SW - 9.6 - MX, 4.0, name="Pool notes card")
    n0 = (pools.share_cur <= res.thresholds.white_space_max_share).sum()
    pn = [f"上位10市場のうち{n0}市場はPBシェアが{res.thresholds.white_space_max_share:.0%}以下です。",
          "非PB売上 = 全体売上 − PB売上。つまりPBがこれから獲得しうる売上です。",
          "規模だけでは推奨になりません。実現性と粗利はこのデータに含まれていません。"]
    _text(s, 9.85, 1.8, SW - 9.6 - MX - 0.5, 0.4, "示唆", 16, DARK, bold=True)
    _bullets(s, 9.85, 2.35, SW - 9.6 - MX - 0.5, 3.1, pn, size=12, name="Pool notes")
    _notes(s, " ".join(pn))

    # ---------------------------------------------------------------- category scorecard
    s = prs.slides.add_slide(L_TONLY)
    _title(s, "カテゴリー別スコアカード: 成長・順位・SKUと売上")
    sc = res.cats[res.cats.pb_cur > 0].head(12)
    rows = [["カテゴリー", "PB売上(順位)", "PBシェア(増減pt)", "PB成長率", "市場成長率", "PB SKU数", "SKU当たりPB売上", "判定"]]
    for _, r_ in sc.iterrows():
        rows.append([r_["cat"], f"{_jpy(r_.pb_cur)}({r_.rank_pb}位)", f"{_pct(r_.share_cur)}({r_.d_share_pt:+.1f})", _spct(r_.pb_growth, 0), _spct(r_.yoy),
                     f"{r_.sku_prev:.0f}→{r_.sku_cur:.0f}({_spct(r_.sku_growth, 0)})", f"{_jpy(r_.per_sku_cur)}({_spct(r_.per_sku_chg, 0)})", r_.verdict])
    _table(s, rows, [1.5, 1.55, 1.55, 0.95, 1.05, 1.75, 1.7, 2.1], MX, 1.4, row_h=0.4, size=10.5, right_cols=(3, 4), name="Category scorecard")
    _text(s, MX, 1.4 + 0.4 * len(rows) + 0.12, SW - 2 * MX, 0.5,
          "判定はPB SKU数の変化とPB売上の変化の比較です。SKU変化が±"
          f"{res.thresholds.sku_stable_band:.0%}以内は横ばい、PB売上{_jpy(res.thresholds.min_base_pb_sales)}未満は判定対象外です。", 11, MUTED)
    _notes(s, "順位はPB売上順(1 = 最大)。PB成長率はPB売上の前期比、市場成長率は全体売上の前期比です。")

    # ---------------------------------------------------------------- SKU vs sales
    s = prs.slides.add_slide(L_TONLY)
    cs = res.cats[(res.cats.pb_prev >= res.thresholds.min_base_pb_sales) & (res.cats.sku_prev > 0)].sort_values("sku_growth", ascending=False)
    if (k["sku_growth"] or 0) > (k["pb_yoy"] or 0):
        _title(s, f"PB SKU数{_spct(k['sku_growth'], 0)}に対しPB売上{_spct(k['pb_yoy'], 0)}: SKU増が売上増を上回る")
    else:
        _title(s, f"PB SKU数{_spct(k['sku_growth'], 0)}、PB売上{_spct(k['pb_yoy'], 0)}: 売上がSKU増に追随")
    if len(cs):
        cd = CategoryChartData()
        cd.categories = cs["cat"].tolist()
        cd.add_series("PB SKU数の増加率", [round(float(v), 4) for v in cs.sku_growth])
        cd.add_series("PB売上の成長率", [round(float(v), 4) for v in cs.pb_growth])
        gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(MX), Inches(1.75), Inches(7.7), Inches(4.2), cd)
        gf.name = "SKU vs sales growth chart"
        ch = gf.chart
        ch.has_title = False
        ch.has_legend = True
        ch.legend.position, ch.legend.include_in_layout = XL_LEGEND_POSITION.TOP, False
        ch.legend.font.size = Pt(12)
        ch.font.size, ch.font.name = Pt(12), LATIN
        pl = ch.plots[0]
        pl.gap_width, pl.overlap = 60, -5
        pl.has_data_labels = True
        pl.data_labels.number_format, pl.data_labels.number_format_is_linked = '0%', False
        pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        pl.data_labels.font.size = Pt(11)
        for ser, col in zip(pl.series, (ACCENT, DARK)):
            ser.invert_if_negative = False
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = _rgb(col)
        ch.value_axis.visible = False
        ch.value_axis.has_major_gridlines = False
        ch.category_axis.format.line.fill.background()
        ch.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    _text(s, MX, 1.4, 8, 0.28, f"カテゴリー別の伸び {prev}期→{cur}期(前期PB売上{_jpy(res.thresholds.min_base_pb_sales)}以上)", 11, MUTED)
    rows = [["カテゴリー", "SKU当たり", "判定"]]
    for _, r_ in cs.iterrows():
        rows.append([r_["cat"], _spct(r_.per_sku_chg, 0), r_.verdict])
    _table(s, rows, [1.45, 0.9, 2.15], 8.55, 1.75, row_h=0.38, size=10.5, right_cols=(1,), name="SKU verdict table")
    marg = _jpy(k["marg_per_sku"]) if k["marg_per_sku"] == k["marg_per_sku"] else "－"
    _box(s, MX, 6.05, SW - 2 * MX, 0.9, name="SKU takeaway")
    lat = f"{res.period_latest}期のPB SKU数は{k['sku_latest']:.0f}です。" if res.period_latest and k.get("sku_latest") == k.get("sku_latest") else ""
    _text(s, MX + 0.25, 6.15, SW - 2 * MX - 0.5, 0.7,
          f"追加SKU1つ当たりのPB売上増は平均{marg}で、{prev}期の既存SKU1つ当たり{_jpy(k['per_sku_prev'])}を下回ります"
          f"(PB SKU当たり売上 {_spct(k['per_sku_chg'], 0)})。新規SKUはまだ通年の実績がありません。{lat}"
          if k["marg_per_sku"] == k["marg_per_sku"] and k["marg_per_sku"] < k["per_sku_prev"] else
          f"追加SKU1つ当たりのPB売上増は平均{marg}({prev}期の既存SKU1つ当たり{_jpy(k['per_sku_prev'])}、"
          f"PB SKU当たり売上 {_spct(k['per_sku_chg'], 0)})。新規SKUはまだ通年の実績がありません。{lat}", 13, DARK, anchor=MSO_ANCHOR.MIDDLE)
    _notes(s, "カテゴリーごとに2本の棒を比較してください。SKUの棒が売上の棒より大きく高い場合、追加したSKUに見合う売上が出ていません。")

    # ---------------------------------------------------------------- rankings
    s = prs.slides.add_slide(L_TONLY)
    rk = res.rankings
    _title(s, "サブカテ別ランキング: 規模・成長・シェア")
    _text(s, MX, 1.3, 6, 0.3, "PB売上 上位8", 14, DARK, bold=True)
    rows = [["サブカテ", "PB売上", "PBシェア", "PB成長率"]]
    for _, r_ in rk["sub_pb"].head(8).iterrows():
        rows.append([r_["name"], _jpy(r_.pb_cur), _pct(r_.share_cur), _spct(r_.pb_growth, 0)])
    _table(s, rows, [2.75, 1.1, 1.0, 1.1], MX, 1.7, row_h=0.42, size=11, right_cols=(1, 2, 3), name="Top PB sales table")
    _text(s, 6.95, 1.3, 6, 0.3, f"PB売上成長率 上位・下位(前期PB売上{_jpy(res.thresholds.min_base_pb_sales)}以上)", 14, DARK, bold=True)
    rows = [["サブカテ", "成長率", "増減額", "区分"]]
    for _, r_ in rk["growth_top"].iterrows():
        rows.append([r_["name"], _spct(r_.pb_growth, 0), _jpy(r_.d_pb), "上位"])
    for _, r_ in rk["growth_bottom"].iterrows():
        rows.append([r_["name"], _spct(r_.pb_growth, 0), _jpy(r_.d_pb), "下位"])
    _table(s, rows, [2.5, 0.9, 1.1, 1.2], 6.95, 1.7, row_h=0.38, size=11, right_cols=(1, 2), name="Growth ranking table")
    wl = res.watchlist if res.watchlist is not None else []
    if len(wl):
        w0 = wl.iloc[0]
        _text(s, MX, 5.5, 6.2, 1.2, f"要注視: {w0['name']}はPBシェアが2期連続で低下({_pct(w0.share_prev)}→{_pct(w0.share_cur)}→{_pct(w0.share_latest)})。"
              f"{res.period_latest}期は途中期のため、直近の低下は参考値です。", 12, MUTED)
    _notes(s, "ランキングはPB売上、PB売上成長率、PBシェアに基づきます。PB売上が小さいサブカテは成長率ランキングから除外しています。")

    # ---------------------------------------------------------------- 5. direction overview
    s = prs.slides.add_slide(L_TONLY)
    _title(s, "サブカテ別の方向性(提案)")
    order = [SCALE, REPLICATE, RECAPTURE, TEST, REVIEW, DEFEND]
    cw, ch_ = (SW - 2 * MX - 0.5) / 3, 2.4
    for i, dn in enumerate(order):
        x, y = MX + (i % 3) * (cw + 0.25), 1.5 + (i // 3) * (ch_ + 0.25)
        x_df = res.directions[dn]
        _box(s, x, y, cw, ch_, name=f"Direction card {i + 1}")
        c = _box(s, x + 0.2, y + 0.2, 0.55, 0.55, fill=DIR_COLOR[dn], shape=MSO_SHAPE.OVAL, name=f"Count {i + 1}")
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(len(x_df))
        _style(r, 16, "FFFFFF", bold=True)
        _text(s, x + 0.95, y + 0.28, cw - 1.15, 0.4, dn, 18, DARK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        _text(s, x + 0.25, y + 0.85, cw - 0.5, 0.85, DIR_BLURB[dn], 10, MUTED)
        tops = "、".join(x_df["name"].head(2).tolist()) if len(x_df) else "現在のルールでは該当なし"
        _text(s, x + 0.25, y + 1.78, cw - 0.5, 0.5, "主な対象: " + tops, 10.5, TEXT, bold=True)
    nd, nm = len(res.directions[DEPRIORITIZE]), len(res.directions[MONITOR])
    _text(s, MX, 6.95, SW - 2 * MX, 0.4, f"このほか、PB未展開で小規模または縮小中の市場が{nd}件({DEPRIORITIZE})、明確なシグナルのないものが{nm}件({MONITOR})。", 11, MUTED)
    _notes(s, f"方向性はPBシェアの水準・推移と、市場の規模・成長に基づいて分類しています。{DEPRIORITIZE}: " +
           "、".join(res.directions[DEPRIORITIZE]["name"].tolist()))

    # ---------------------------------------------------------------- 6-11. one slide per direction
    hdr = ["サブカテ", f"売上 {cur}", "非PB", f"PBシェア {prev}→{cur}" + (f"→{res.period_latest}" if res.period_latest else ""), "市場前年比"]
    wd = [4.6, 1.7, 1.7, 2.9, 1.23]
    for dn in order:
        x_df = res.directions[dn]
        if x_df.empty:
            continue
        s = prs.slides.add_slide(L_TONLY)
        _title(s, DIR_TITLE[dn])
        _text(s, MX, 1.3, 10.3, 0.6, DIR_BLURB[dn], 13, MUTED)
        chip = _box(s, SW - MX - 1.7, 1.3, 1.7, 0.42, fill=DIR_COLOR[dn], name="Count chip")
        tf = chip.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = f"{len(x_df)}サブカテ"
        _style(r, 12, "FFFFFF", bold=True)
        show = x_df.head(7)
        rows = [hdr]
        for _, r_ in show.iterrows():
            lat = f"→{r_.share_latest * 100:.1f}%" if r_.share_latest == r_.share_latest else ""
            rows.append([r_["name"], _jpy(r_.sales), _jpy(r_.nb_cur),
                         f"{r_.share_prev * 100:.1f}%→{r_.share_cur * 100:.1f}%{lat}", _spct(r_.yoy)])
        _table(s, rows, wd, MX, 2.05, row_h=0.42, size=12, right_cols=(1, 2, 4), name="Direction table")
        if len(x_df) > len(show):
            _text(s, MX, 2.05 + 0.42 * len(rows) + 0.08, 8, 0.28, f"ほか{len(x_df) - len(show)}件(PDFレポートの付録参照)", 11, MUTED)
        # summary strip for the whole direction group
        tot_sales, tot_nb, tot_pb = float(x_df.sales.sum()), float(x_df.nb_cur.sum()), float(x_df.pb_cur.sum())
        up_add = float(res.upside.loc[res.upside.direction == dn, "add"].sum()) if not res.upside.empty else 0.0
        strip = [("非PB売上の合計", _jpy(tot_nb), f"{len(x_df)}サブカテの合計"),
                 ("現在のPB売上", _jpy(tot_pb), f"PBシェア {_pct(tot_pb / tot_sales if tot_sales else None)}(全体 {_jpy(tot_sales)})"),
                 ("伸びしろの試算", _jpy(up_add) if up_add > 0 else "対象外",
                  "試算スライド参照" if up_add > 0 else "この方向性は試算の対象外")]
        sw_ = (SW - 2 * MX - 0.5) / 3
        for i, (lab, val, sub_) in enumerate(strip):
            xx = MX + i * (sw_ + 0.25)
            _box(s, xx, 5.8, sw_, 1.2, name=f"Summary {i + 1}")
            _text(s, xx + 0.25, 5.92, sw_ - 0.5, 0.28, lab, 12, MUTED)
            _text(s, xx + 0.25, 6.2, sw_ - 0.5, 0.45, val, 24, DARK, bold=True)
            _text(s, xx + 0.25, 6.68, sw_ - 0.5, 0.25, sub_, 11, MUTED)
        _notes(s, DIR_BLURB[dn] + " 対象: " + "、".join(x_df["name"].tolist()))

    # ---------------------------------------------------------------- upside
    if not res.upside.empty:
        s = prs.slides.add_slide(L_TONLY)
        u = res.upside_total
        t_ = res.thresholds
        _title(s, f"伸びしろの試算: PB売上 約{_jpy(u['add'])}増")
        rows = [["サブカテ", "方向性", "PBシェア 現在→目標", "PB売上の増加額"]]
        for _, r_ in res.upside.head(8).iterrows():
            rows.append([r_["name"], r_.direction, f"{r_.share_cur * 100:.1f}%→{r_.target * 100:.1f}%", _jpy(r_["add"])])
        _table(s, rows, [3.6, 1.35, 2.2, 1.45], MX, 1.55, row_h=0.55, size=13, right_cols=(2, 3), name="Upside table")
        _box(s, 9.55, 1.55, SW - 9.55 - MX, 2.35, fill=DARK, name="Upside callout")
        _text(s, 9.8, 1.75, 3.0, 0.3, "PB売上の増加額(合計)", 12, "BFD9DA")
        _text(s, 9.8, 2.1, 2.75, 0.7, "+" + _jpy(u["add"]), 32, "FFFFFF", bold=True)
        _text(s, 9.8, 2.9, 2.75, 0.3, f"PB比率 {_pct(k['share_cur'])}→{_pct(u['share_after'])}", 14, ACCENT, bold=True)
        _text(s, 9.8, 3.25, 2.75, 0.3, f"現在のPB売上比 {_spct(u['pct_of_pb'], 0)}", 12, "BFD9DA")
        _text(s, 9.55, 4.25, SW - 9.55 - MX, 2.6,
              f"予測ではなく仮定です。{SCALE}: 現在のシェアに直近の上昇幅{t_.upside_scale_momentum_years:g}年分を加算(上限{t_.upside_scale_cap:.0%})。"
              f"{RECAPTURE}: 前期シェアに回復。{REPLICATE}: シェア{t_.upside_replicate_share:.0%}。{TEST}は対象外。"
              "売上はNBから獲得し全体売上は横ばいと仮定。粗利は考慮していません。", 12, MUTED)
        _notes(s, "計画用の仮定による試算です。")

    # ---------------------------------------------------------------- basis
    s = prs.slides.add_slide(L_TONLY)
    _title(s, "前提・注意事項")
    notes_ = key_notes(res)
    _bullets(s, MX, 1.6, SW - 2 * MX, 5.4, notes_, size=18, gap=16, name="Notes")
    _notes(s, "\n".join(notes_))

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
