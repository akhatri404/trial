"""CLI:  python -m pbreport input.xlsx [-o report] [-f pdf|pptx|both]   |   python -m pbreport --watch inbox/ outbox/ [-f both]"""
import argparse
import sys
import time
from pathlib import Path

from . import generate_report, WorkbookError


def _one(src: Path, out: Path, fmt: str):
    data, _ = generate_report(str(src), source_name=src.name, fmt=fmt)
    out.write_bytes(data)
    print(f"OK  {src.name} -> {out}")


def _targets(base: Path, fmt: str):
    """base is the path without a meaningful extension; returns [(path, fmt)]."""
    fmts = ["pdf", "pptx"] if fmt == "both" else [fmt]
    return [(base.with_suffix("." + f), f) for f in fmts]


def main(argv=None):
    ap = argparse.ArgumentParser(prog="pbreport", description="ExcelファイルからPB戦略レポート(PDF / PowerPoint)を作成します。")
    ap.add_argument("input", nargs="?", help="Excelファイル(.xlsx)。--watch 指定時は受信フォルダ")
    ap.add_argument("output", nargs="?", help="出力先のパス。--watch 指定時は出力フォルダ")
    ap.add_argument("-o", "--out", help="出力先のパス(単一ファイル時)。拡張子は --format で決まります")
    ap.add_argument("-f", "--format", choices=["pdf", "pptx", "both"], default="pdf", help="レポート形式(既定: pdf)")
    ap.add_argument("--watch", action="store_true", help="受信フォルダを監視し、新しい .xlsx ごとにレポートを作成")
    ap.add_argument("--interval", type=float, default=5.0, help="--watch 時のフォルダ確認間隔(秒)")
    a = ap.parse_args(argv)
    if not a.input:
        ap.error("入力ファイルを指定してください")
    try:
        if a.watch:
            inbox, outbox = Path(a.input), Path(a.output or "reports")
            outbox.mkdir(parents=True, exist_ok=True)
            seen = {}
            print(f"{inbox} を監視中(Ctrl+C で停止)")
            while True:
                for f in inbox.glob("*.xlsx"):
                    if f.name.startswith("~$"):
                        continue
                    m = f.stat().st_mtime
                    if seen.get(f) != m:
                        time.sleep(1)  # let the upload finish
                        try:
                            for path, fm in _targets(outbox / f.stem, a.format):
                                _one(f, path, fm)
                        except WorkbookError as e:
                            print(f"ERR {f.name}: {e}", file=sys.stderr)
                        seen[f] = m
                time.sleep(a.interval)
        src = Path(a.input)
        for path, fm in _targets(Path(a.out or a.output or src.with_suffix("")), a.format):
            _one(src, path, fm)
    except WorkbookError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        return 0
    return 0


if __name__ == "__main__":
    sys.exit(main())
