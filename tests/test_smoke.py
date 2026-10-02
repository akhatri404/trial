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


def _mini_workbook():
    """Two-period workbook with only the sub-category sheet (no 方向性 / カテゴリー sheets, no partial period)."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "サブカテ"
    hdr = ["部門", "カテゴリ", "サブカテ"]
    for p in ("25/6", "26/6"):
        hdr += [f"{p}期\n全体売上", f"{p}期\n全体売れ数", f"{p}期\nPB比率\n(売上)", f"{p}期\nPB比率\n(売れ数)", f"{p}期\nPB SKU数"]
    ws.append([None, None, None])
    ws.append([None] + hdr)
    rows = [
        ("0001 A", "0001 X", "0001 x1", 1e9, 1e6, .10, .12, 3, 1.2e9, 1.1e6, .15, .17, 5),
        ("0001 A", "0001 X", "0002 x2", 5e8, 9e5, .00, .00, 0, 5.2e8, 9e5, .00, .00, 0),
        ("0001 A", "0002 Y", "0001 y1", 8e8, 2e6, .60, .70, 6, 8.1e8, 2e6, .58, .69, 6),
    ]
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
