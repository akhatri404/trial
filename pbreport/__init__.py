"""PB (private brand) report generator for the pet-category comparison workbook."""
from .loader import load_workbook_data, WorkbookError
from .analysis import analyze, Thresholds
from .pdf import build_pdf
from .pptx_report import build_pptx

FORMATS = {"pdf": ("PDF", "application/pdf", ".pdf"),
           "pptx": ("PowerPoint", "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx")}


def generate_report(xlsx, thresholds=None, source_name="", fmt="pdf"):
    """xlsx: path or file-like object; fmt: 'pdf' or 'pptx'. Returns (file_bytes, result)."""
    if fmt not in FORMATS:
        raise ValueError(f"Unknown format {fmt!r}; use one of {sorted(FORMATS)}")
    data = load_workbook_data(xlsx)
    result = analyze(data, thresholds or Thresholds())
    builder = build_pdf if fmt == "pdf" else build_pptx
    return builder(result, source_name), result


__all__ = ["load_workbook_data", "WorkbookError", "analyze", "Thresholds", "build_pdf", "build_pptx",
           "generate_report", "FORMATS"]
