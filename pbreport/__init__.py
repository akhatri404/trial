"""PB (private brand) report generator for the pet-category comparison workbook."""
from .loader import load_workbook_data, WorkbookError
from .analysis import analyze, Thresholds
from .pdf import build_pdf
from .pptx_report import build_pptx

FORMATS = {"pdf": ("PDF", "application/pdf", ".pdf"),
           "pdf_en": ("PDF (English)", "application/pdf", ".pdf"),
           "pptx": ("PowerPoint", "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx")}


def generate_report(xlsx, thresholds=None, source_name="", fmt="pdf"):
    """xlsx: path or file-like object; fmt: 'pdf' (日本語), 'pdf_en' (English) or 'pptx'. Returns (file_bytes, result)."""
    if fmt not in FORMATS:
        raise ValueError(f"Unknown format {fmt!r}; use one of {sorted(FORMATS)}")
    data = load_workbook_data(xlsx)
    result = analyze(data, thresholds or Thresholds(), lang="en" if fmt == "pdf_en" else "ja")
    builder = build_pptx if fmt == "pptx" else build_pdf
    return builder(result, source_name), result


__all__ = ["load_workbook_data", "WorkbookError", "analyze", "Thresholds", "build_pdf", "build_pptx",
           "generate_report", "FORMATS"]
