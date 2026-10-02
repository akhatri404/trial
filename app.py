"""Upload the Excel workbook -> choose PDF or PowerPoint -> download the generated report.

Run locally:  streamlit run app.py
"""
import hashlib
from pathlib import Path

import streamlit as st

from pbreport import FORMATS, Thresholds, WorkbookError, analyze, generate_report, load_workbook_data

# st.cache_data only hashes build()'s own source, not the pbreport package it calls, so without this a redeploy
# keeps serving reports made by the old code for a file uploaded before.
CODE_VERSION = hashlib.sha256(b"".join(p.read_bytes() for p in sorted((Path(__file__).parent / "pbreport").glob("*.py")))).hexdigest()

st.set_page_config(page_title="PB戦略レポート", page_icon="📄", layout="centered")
st.title("PB戦略レポート")
st.caption("ペットカテゴリーのExcelファイルをアップロードしてください。選択した形式でレポートを自動作成します。")

FORMAT_LABELS = {"pdf": "PDF (日本語)", "pdf_en": "PDF (English)", "pptx": "PowerPoint (日本語)"}


@st.cache_data(show_spinner=False, max_entries=8)
def build(file_bytes: bytes, name: str, fmt: str, code_version: str):
    """Cached so switching format or re-running does not repeat work for the same file."""
    import io
    blob, res = generate_report(io.BytesIO(file_bytes), Thresholds(), source_name=name, fmt=fmt)
    if res.lang != "ja":   # the notes shown on this page stay Japanese even when the report is English
        res = analyze(load_workbook_data(io.BytesIO(file_bytes)), Thresholds())
    recon_ok = bool(len(res.recon) and res.recon.ok.all())
    return blob, res.caveats, res.dq, recon_ok


up = st.file_uploader("Excelファイル (.xlsx)", type=["xlsx"], accept_multiple_files=False)

fmt_label = st.radio("レポート形式", list(FORMAT_LABELS.values()), horizontal=True,
                     help="PDFは閲覧・共有向けです。PDF (English)は英語版です(カテゴリー名などExcelの項目は日本語のまま)。PowerPointはグラフ・表・ノートをすべて編集できます。")
fmt = next(k for k, v in FORMAT_LABELS.items() if v == fmt_label)

if up is not None:
    try:
        with st.spinner(f"分析して{FORMAT_LABELS[fmt]}レポートを作成しています…"):
            blob, caveats, dq, recon_ok = build(up.getvalue(), up.name, fmt, CODE_VERSION)
    except WorkbookError as e:
        st.error(f"このファイルは処理できませんでした: {e}")
        st.stop()
    except Exception as e:
        st.error("レポートの作成中に予期しないエラーが発生しました。ファイルの形式を確認して、もう一度お試しください。")
        st.exception(e)
        st.stop()

    st.success("レポートが完成しました。")

    label, mime, ext = FORMATS[fmt]
    base = up.name.rsplit(".", 1)[0]
    st.download_button(f"⬇ {FORMAT_LABELS[fmt]}をダウンロード", data=blob, file_name=f"{base}_PB_report_en{ext}" if fmt == "pdf_en" else f"{base}_PBレポート{ext}", mime=mime,
                       type="primary")
    if recon_ok:
        st.caption("✅ 合計は元資料の部門合計と一致しています。")
    else:
        st.warning("合計が元資料の部門合計と完全には一致しませんでした。下のデータチェックを確認してください。")
    with st.expander("データチェックと注意点"):
        for sev, m in dq:
            st.write(("⚠️ " if sev == "warn" else "• ") + m)
        for c in caveats:
            st.write("• " + c)
else:
    st.info("ファイルのアップロードを待っています。")
