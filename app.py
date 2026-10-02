"""Upload the Excel workbook -> choose PDF or PowerPoint -> download the generated report.

Run locally:  streamlit run app.py
"""
from dataclasses import fields, asdict

import streamlit as st

from pbreport import FORMATS, Thresholds, WorkbookError, generate_report
from pbreport.analysis import THRESHOLD_LABELS

st.set_page_config(page_title="PB戦略レポート", page_icon="📄", layout="centered")
st.title("PB戦略レポート")
st.caption("ペットカテゴリーのExcelファイルをアップロードしてください。選択した形式でレポートを自動作成します。")

FORMAT_LABELS = {"pdf": "PDF", "pptx": "PowerPoint (.pptx)"}

with st.sidebar:
    st.header("ルールの基準値")
    st.caption("初期値は標準レポートと同じです。チームで新しいルールに合意した場合のみ変更してください。")
    defaults = Thresholds()
    custom = {}
    with st.expander("調整する", expanded=False):
        for f in fields(Thresholds):
            v = getattr(defaults, f.name)
            custom[f.name] = st.number_input(THRESHOLD_LABELS.get(f.name, f.name), value=float(v), help=f.name,
                                             format="%.4f" if abs(v) < 10 else "%.0f")
    thresholds = Thresholds(**custom)
    thr_key = tuple(sorted(asdict(thresholds).items()))


@st.cache_data(show_spinner=False, max_entries=8)
def build(file_bytes: bytes, name: str, thr_items: tuple, fmt: str):
    """Cached so switching format or re-running does not repeat work for the same file and settings."""
    import io
    blob, res = generate_report(io.BytesIO(file_bytes), Thresholds(**dict(thr_items)), source_name=name, fmt=fmt)
    recon_ok = bool(len(res.recon) and res.recon.ok.all())
    return blob, res.kpi, res.caveats, res.dq, recon_ok, res.bridge


up = st.file_uploader("Excelファイル (.xlsx)", type=["xlsx"], accept_multiple_files=False)

fmt_label = st.radio("レポート形式", list(FORMAT_LABELS.values()), horizontal=True,
                     help="PDFは閲覧・共有向けです。PowerPointはグラフ・表・ノートをすべて編集できます。")
fmt = next(k for k, v in FORMAT_LABELS.items() if v == fmt_label)

if up is not None:
    try:
        with st.spinner(f"分析して{FORMAT_LABELS[fmt]}レポートを作成しています…"):
            blob, k, caveats, dq, recon_ok, bridge = build(up.getvalue(), up.name, thr_key, fmt)
    except WorkbookError as e:
        st.error(f"このファイルは処理できませんでした: {e}")
        st.stop()
    except Exception as e:
        st.error("レポートの作成中に予期しないエラーが発生しました。ファイルの形式を確認して、もう一度お試しください。")
        st.exception(e)
        st.stop()

    st.success("レポートが完成しました。")
    c1, c2, c3 = st.columns(3)
    c1.metric("PB売上", f"{k['pb_cur'] / 1e8:,.2f}億円", f"{k['pb_yoy'] * 100:+.1f}%")
    c2.metric("PB比率", f"{k['share_cur'] * 100:.2f}%", f"{(k['share_cur'] - k['share_prev']) * 100:+.2f}pt")
    per = k.get("per_sku_chg")
    c3.metric("SKU当たりPB売上", f"{k['per_sku_cur'] / 1e4:,.0f}万円", f"{per * 100:+.0f}%" if per == per else None)

    label, mime, ext = FORMATS[fmt]
    base = up.name.rsplit(".", 1)[0]
    st.download_button(f"⬇ {FORMAT_LABELS[fmt]}をダウンロード", data=blob, file_name=f"{base}_PBレポート{ext}", mime=mime,
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
