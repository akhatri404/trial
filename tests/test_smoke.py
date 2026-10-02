"""Run with:  pytest -q   (set PB_SAMPLE to the real workbook path to test against real data)."""
import io
import os

import pandas as pd
import pytest

from pbreport import generate_report, WorkbookError

SAMPLE = os.environ.get("PB_SAMPLE")


@pytest.mark.skipif(not SAMPLE or not os.path.exists(SAMPLE or ""), reason="PB_SAMPLE not set")
def test_real_workbook_makes_pdf():
    pdf, res = generate_report(SAMPLE)
    assert pdf[:4] == b"%PDF" and len(pdf) > 20_000
    assert res.kpi["sku_cur"] > 0


@pytest.mark.skipif(not SAMPLE or not os.path.exists(SAMPLE or ""), reason="PB_SAMPLE not set")
def test_real_workbook_makes_native_pptx():
    from pptx import Presentation
    data, res = generate_report(SAMPLE, fmt="pptx")
    prs = Presentation(io.BytesIO(data))
    assert len(prs.slides) >= 10
    assert all(s.shapes.title is not None and s.shapes.title.text_frame.text for s in prs.slides)
    kinds = {sh.shape_type for s in prs.slides for sh in s.shapes}
    assert any(sh.has_chart for s in prs.slides for sh in s.shapes)   # native charts, not pictures
    assert any(sh.has_table for s in prs.slides for sh in s.shapes)   # native tables


def test_unknown_format_rejected():
    with pytest.raises(ValueError):
        generate_report(io.BytesIO(b""), fmt="docx")


def test_wrong_file_gives_friendly_error():
    buf = io.BytesIO()
    pd.DataFrame({"a": [1]}).to_excel(buf, index=False)
    buf.seek(0)
    with pytest.raises(WorkbookError):
        generate_report(buf)


def test_garbage_gives_friendly_error():
    with pytest.raises(WorkbookError):
        generate_report(io.BytesIO(b"not an excel file"))


def _mini_workbook(third=False, note=True):
    """Two-period workbook with only the sub-category sheet (no 方向性 / カテゴリー sheets, no partial period)."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "サブカテ"
    hdr = ["部門", "カテゴリ", "サブカテ"]
    for p in (("25/6", "26/6", "27/6") if third else ("25/6", "26/6")):
        hdr += [f"{p}期\n全体売上", f"{p}期\n全体売れ数", f"{p}期\nPB比率\n(売上)", f"{p}期\nPB比率\n(売れ数)", f"{p}期\nPB SKU数"]
    ws.append([None, None, None])
    ws.append([None] + hdr)
    rows = [
        ("0001 A", "0001 X", "0001 x1", 1e9, 1e6, .10, .12, 3, 1.2e9, 1.1e6, .15, .17, 5),
        ("0001 A", "0001 X", "0002 x2", 5e8, 9e5, .00, .00, 0, 5.2e8, 9e5, .00, .00, 0),
        ("0002 B", "0002 Y", "0001 y1", 8e8, 2e6, .60, .70, 6, 8.1e8, 2e6, .58, .69, 6),
    ]
    if third:   # a newest period with far fewer sales than the others; the 方向性 sheet may carry a 「※Nヶ月分」 note
        extra = [(2.2e8, 3e5, .19, .21, 6), (0.9e8, 2.2e5, .0, .0, 0), (1.3e8, 5e5, .52, .60, 6)]
        rows = [r + e for r, e in zip(rows, extra)]
        if note:
            wb.create_sheet("方向性")["A1"] = "※2ヶ月分"
    for r in rows:
        ws.append([None] + list(r))
    b = io.BytesIO()
    wb.save(b)
    b.seek(0)
    return b


def test_two_period_minimal_workbook_all_formats():
    from pptx import Presentation
    pdf, res = generate_report(_mini_workbook(), fmt="pdf")
    assert pdf[:4] == b"%PDF" and res.period_latest is None
    prs_bytes, _ = generate_report(_mini_workbook(), fmt="pptx")
    assert len(Presentation(io.BytesIO(prs_bytes)).slides) >= 8
    # bridge identities hold exactly
    b = res.bridge
    assert abs(b["market"] + b["share"] - b["d_pb"]) < 1
    assert abs(b["rate_pt"] + b["mix_pt"] - b["d_ratio_pt"]) < 1e-6


@pytest.mark.skipif(not SAMPLE or not os.path.exists(SAMPLE or ""), reason="PB_SAMPLE not set")
def test_real_workbook_reconciles_with_source_totals():
    _, res = generate_report(SAMPLE)
    assert len(res.recon) > 0 and res.recon.ok.all()
    # headline PB ratio must equal the source's own total, not a sub-category-only figure
    tot = res.recon[(res.recon.scope.str.contains("計")) & (res.recon.period == res.periods_full[1])].iloc[0]
    assert abs(res.kpi["share_cur"] - tot.rep_ratio) < 1e-6


def test_english_pdf_has_no_japanese_report_text():
    import re
    pdf, res = generate_report(_mini_workbook(), fmt="pdf_en")
    assert pdf[:4] == b"%PDF" and res.lang == "en"
    texts = res.findings + res.caveats + [m for _, m in res.dq] + [res.scope_note]
    # 期 and the workbook's own labels (「方向性」, 「PB SKU数」) stay Japanese by design
    cjk = re.compile(r"[\u3040-\u30ff\u4e00-\u9fff]")
    bad = [t for t in texts if cjk.search(re.sub(r"期|「方向性」|「PB SKU数」|「カテゴリー」", "", t))]
    assert not bad, bad


def test_department_view_adds_up_to_the_company_total():
    _, res = generate_report(_mini_workbook(), fmt="pdf")
    dp, b = res.depts, res.bridge
    assert set(dp["dept"]) == {"A", "B"}
    assert abs(dp.pb_cur.sum() - b["pb_cur"]) < 1
    assert abs(dp.market_eff.sum() - b["market"]) < 1 and abs(dp.share_eff.sum() - b["share"]) < 1
    assert abs(dp.rate_pt.sum() - b["rate_pt"]) < 1e-6 and abs(dp.mix_pt.sum() - b["mix_pt"]) < 1e-6
    assert all(res.dept_findings[d] for d in dp["dept"])


@pytest.mark.skipif(not SAMPLE or not os.path.exists(SAMPLE or ""), reason="PB_SAMPLE not set")
def test_real_workbook_department_view():
    from pptx import Presentation
    _, res = generate_report(SAMPLE)
    dp, b = res.depts, res.bridge
    assert len(dp) >= 2 and abs(dp.pb_cur.sum() - b["pb_cur"]) < 1
    assert abs(dp.rate_pt.sum() - b["rate_pt"]) < 1e-6 and abs(dp.mix_pt.sum() - b["mix_pt"]) < 1e-6
    data, _ = generate_report(SAMPLE, fmt="pptx")
    titles = [sl.shapes.title.text_frame.text for sl in Presentation(io.BytesIO(data)).slides]
    assert all(any(d in t for t in titles) for d in dp["dept"])      # one slide per department


def test_shorter_newest_period_is_shown_alongside_without_being_scaled_or_compared():
    from pptx import Presentation
    pdf, res = generate_report(_mini_workbook(third=True), fmt="pdf")
    assert res.period_latest == "27/6" and res.periods_full == ("25/6", "26/6")   # growth, bridge and directions stay on the comparable periods
    assert not hasattr(res, "annual_factor")                                       # nothing is scaled up to a full period
    assert abs(res.depts.pb_latest.sum() - res.kpi["pb_latest"]) < 1               # departments add up to the total
    assert pdf[:4] == b"%PDF" and generate_report(_mini_workbook(third=True), fmt="pdf_en")[0][:4] == b"%PDF"
    deck = Presentation(io.BytesIO(generate_report(_mini_workbook(third=True), fmt="pptx")[0]))
    assert any("3期" in sl.shapes.title.text_frame.text for sl in deck.slides)
    assert not any(sh.name.startswith("Stat card") for sl in deck.slides for sh in sl.shapes)   # the three cards are gone


def test_shorter_period_is_found_from_the_sales_even_without_a_note():
    _, res = generate_report(_mini_workbook(third=True, note=False), fmt="pdf")
    assert res.period_latest == "27/6"


def test_report_text_never_names_a_number_of_months():
    import re
    _, ja = generate_report(_mini_workbook(third=True), fmt="pdf")
    for lang_fmt in ("pdf", "pdf_en"):
        res = generate_report(_mini_workbook(third=True), fmt=lang_fmt)[1]
        texts = res.findings + res.caveats + [m for _, m in res.dq] + sum(res.dept_findings.values(), [])
        assert not [t for t in texts if re.search(r"ヶ月|か月|months?|年換算|annualis", t)], texts


def test_share_trend_tags_the_newest_period_against_the_last_full_year():
    from pbreport.analysis import T_DOWN, T_UP
    _, res = generate_report(_mini_workbook(third=True), fmt="pdf")
    tr = dict(zip(res.subs["name"], res.subs["trend"]))
    assert tr["X / x1"] == T_UP and tr["Y / y1"] == T_DOWN      # +6 pt and -6 pt vs 26/6; both have PB sales above the minimum base
    assert tr["X / x2"] == ""                                      # no PB sales: not tagged
