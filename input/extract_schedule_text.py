#!/usr/bin/env python3
"""
Extract raw schedule text from Excel (or PDF) files -- NO AI
===========================================================

This is the deterministic "reader" half of a two-step workflow:

    1. THIS script  -> open the file with a real parser library
                       (openpyxl for .xlsx, pdfplumber for .pdf) and dump
                       every cell / line as plain text. No AI, no network,
                       fully reproducible.
    2. an AI later  -> take that plain text and analyse it / summarise it /
                       turn it into whatever result you need.

Keeping step 1 free of AI means the text handed to the model is exactly
what is in the file -- nothing hallucinated, nothing skipped.

Usage
-----
    # default: read ../zim_laemchabang_shanghai_schedule.xlsx next to the repo root
    python input/extract_schedule_text.py

    # one explicit file
    python input/extract_schedule_text.py zim_laemchabang_shanghai_schedule.xlsx

    # a whole folder (every *.xlsx / *.xls / *.pdf inside it)
    python input/extract_schedule_text.py exports/

    # several paths at once, choose where the .txt goes
    python input/extract_schedule_text.py a.xlsx b.pdf --out-dir output/

For each input file the script:
  * prints the extracted text to the console, and
  * writes it next to the source as "<name>.extracted.txt"
    (or into --out-dir if given).

Requirements
------------
    pip install openpyxl          # for .xlsx / .xlsm  (already in requirements.txt)
    pip install pdfplumber        # only needed if you point it at a .pdf
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path
from typing import Iterable

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
PDF_SUFFIXES = {".pdf"}
SUPPORTED_SUFFIXES = EXCEL_SUFFIXES | PDF_SUFFIXES

DEFAULT_TARGET = "zim_laemchabang_shanghai_schedule.xlsx"


# --------------------------------------------------------------------------- #
# cell / value formatting                                                     #
# --------------------------------------------------------------------------- #
def fmt_cell(value) -> str:
    """Render a single cell value as clean text (no AI, just rules)."""
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        # drop a midnight time component so dates read as plain dates
        if value.hour == value.minute == value.second == 0:
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, dt.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


# --------------------------------------------------------------------------- #
# Excel                                                                       #
# --------------------------------------------------------------------------- #
def extract_excel(path: Path) -> str:
    try:
        import openpyxl
    except ModuleNotFoundError:
        sys.exit("openpyxl is not installed. Run:  pip install openpyxl")

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out: list[str] = []
    out.append(f"FILE: {path.name}")
    out.append(f"TYPE: Excel workbook  |  SHEETS: {', '.join(wb.sheetnames)}")
    out.append("")

    for ws in wb.worksheets:
        rows = [
            [fmt_cell(c) for c in row]
            for row in ws.iter_rows(values_only=True)
        ]
        # trim fully empty trailing rows
        while rows and not any(cell for cell in rows[-1]):
            rows.pop()

        out.append("=" * 70)
        out.append(f"SHEET: {ws.title}   ({len(rows)} non-empty rows)")
        out.append("=" * 70)

        if not rows:
            out.append("(empty sheet)")
            out.append("")
            continue

        # 1) a readable aligned table
        col_count = max(len(r) for r in rows)
        rows = [r + [""] * (col_count - len(r)) for r in rows]
        widths = [
            max(len(rows[r][c]) for r in range(len(rows)))
            for c in range(col_count)
        ]
        for i, r in enumerate(rows):
            line = " | ".join(cell.ljust(widths[c]) for c, cell in enumerate(r))
            out.append(line.rstrip())
            if i == 0:  # header underline
                out.append("-+-".join("-" * w for w in widths))
        out.append("")

        # 2) a machine-friendly TSV block (handy to paste elsewhere)
        out.append("--- TSV ---")
        for r in rows:
            out.append("\t".join(r))
        out.append("")

    wb.close()
    return "\n".join(out)


# --------------------------------------------------------------------------- #
# PDF                                                                         #
# --------------------------------------------------------------------------- #
def extract_pdf(path: Path) -> str:
    try:
        import pdfplumber
    except ModuleNotFoundError:
        sys.exit("pdfplumber is not installed. Run:  pip install pdfplumber")

    out: list[str] = [f"FILE: {path.name}", "TYPE: PDF document", ""]
    with pdfplumber.open(path) as pdf:
        out[1] = f"TYPE: PDF document  |  PAGES: {len(pdf.pages)}"
        for pageno, page in enumerate(pdf.pages, start=1):
            out.append("=" * 70)
            out.append(f"PAGE {pageno}")
            out.append("=" * 70)

            text = page.extract_text() or ""
            out.append(text.strip() or "(no extractable text on this page)")
            out.append("")

            for tnum, table in enumerate(page.extract_tables(), start=1):
                out.append(f"--- PAGE {pageno} TABLE {tnum} (TSV) ---")
                for row in table:
                    out.append("\t".join((c or "").strip() for c in row))
                out.append("")
    return "\n".join(out)


# --------------------------------------------------------------------------- #
# driver                                                                      #
# --------------------------------------------------------------------------- #
def extract_file(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in EXCEL_SUFFIXES:
        return extract_excel(path)
    if suffix in PDF_SUFFIXES:
        return extract_pdf(path)
    raise ValueError(f"unsupported file type: {path.name} ({suffix or 'no suffix'})")


def collect_paths(raw_paths: Iterable[str]) -> list[Path]:
    found: list[Path] = []
    for raw in raw_paths:
        p = Path(raw)
        if p.is_dir():
            for child in sorted(p.iterdir()):
                if child.is_file() and child.suffix.lower() in SUPPORTED_SUFFIXES:
                    found.append(child)
        elif p.is_file():
            found.append(p)
        else:
            print(f"! not found, skipping: {p}", file=sys.stderr)
    return found


def main() -> None:
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent

    parser = argparse.ArgumentParser(
        description="Extract raw text from Excel/PDF schedule files (no AI). "
                    "Prints the text and saves it as <name>.extracted.txt."
    )
    parser.add_argument(
        "paths", nargs="*",
        help="Files or folders to read (.xlsx / .xlsm / .pdf). "
             f"Default: {DEFAULT_TARGET} at the repo root.",
    )
    parser.add_argument(
        "--out-dir", default=None,
        help="Directory for the .extracted.txt files (default: next to each source file).",
    )
    parser.add_argument(
        "--stdout-only", action="store_true",
        help="Only print to the console, do not write .txt files.",
    )
    args = parser.parse_args()

    raw_paths = args.paths
    if not raw_paths:
        default_path = repo_root / DEFAULT_TARGET
        if not default_path.exists():
            # fall back to any .xlsx sitting at the repo root
            candidates = sorted(repo_root.glob("*.xlsx"))
            if not candidates:
                sys.exit(f"No path given and {default_path} does not exist.")
            raw_paths = [str(candidates[0])]
            print(f"(no path given -> using {candidates[0].name})", file=sys.stderr)
        else:
            raw_paths = [str(default_path)]

    files = collect_paths(raw_paths)
    if not files:
        sys.exit("Nothing to do: no supported files found.")

    out_dir = Path(args.out_dir) if args.out_dir else None
    if out_dir:
        out_dir.mkdir(parents=True, exist_ok=True)

    for path in files:
        try:
            text = extract_file(path)
        except Exception as exc:  # noqa: BLE001 - surface the file that failed
            print(f"! failed on {path}: {exc}", file=sys.stderr)
            continue

        print(text)
        print()

        if not args.stdout_only:
            target_dir = out_dir if out_dir else path.parent
            out_path = target_dir / f"{path.stem}.extracted.txt"
            out_path.write_text(text, encoding="utf-8")
            print(f"-> wrote {out_path}", file=sys.stderr)


if __name__ == "__main__":
    main()
