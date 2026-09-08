#!/usr/bin/env python3
"""
Read every Excel file in the KMTC/ folder, merge them, dump plain text -- NO AI
=============================================================================

Two-step workflow (same idea as input/extract_schedule_text.py):

    1. THIS script  -> open each workbook with a real parser library
                       (pandas + xlrd for .xls, pandas + openpyxl for .xlsx),
                       flatten the 2-row KMTC header, stack every file into one
                       table, drop exact-duplicate sailings, and write the
                       result out as plain text / CSV.  No AI, no network,
                       fully reproducible.
    2. an AI later  -> take that plain text and analyse it (build the final
                       Laem Chabang -> Shanghai schedule, spot gaps, compare
                       against other carriers, whatever is needed).

Keeping step 1 AI-free means the text handed to the model is exactly what is
in the files -- nothing hallucinated, nothing skipped.

Note on pdfplumber
------------------
pdfplumber only reads PDFs.  The KMTC exports are real .xls workbooks (OLE2 /
BIFF), so they are read with pandas' Excel engines instead:
    * .xls           -> xlrd
    * .xlsx / .xlsm  -> openpyxl
The rest of the pipeline (deterministic reader -> AI analyser) is unchanged.

Usage
-----
    # default: read ../KMTC next to the repo root, write into ../output/
    python input/kmtc_extract_text.py

    # point it at another folder or explicit files
    python input/kmtc_extract_text.py path/to/folder
    python input/kmtc_extract_text.py KMTC/"LCH_SHA_LegSchedule 10.xls" KMTC/"LCH_SHA_LegSchedule 11.xls"

    # choose the output location / only print
    python input/kmtc_extract_text.py --out-dir output/
    python input/kmtc_extract_text.py --stdout-only

Outputs (unless --stdout-only)
------------------------------
    <out-dir>/kmtc_combined.extracted.txt   human-readable: per-file dump + merged table
    <out-dir>/kmtc_combined.csv             the merged table as CSV (deterministic, AI-friendly)

Requirements
------------
    pip install pandas xlrd openpyxl        # all already in requirements.txt except xlrd
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable

try:
    import pandas as pd
except ModuleNotFoundError:
    sys.exit("pandas is not installed. Run:  pip install pandas xlrd openpyxl")

EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm", ".xltx", ".xltm"}
DEFAULT_FOLDER = "KMTC"
OUT_STEM = "kmtc_combined"

# The KMTC "LCH_SHA_LegSchedule" export has a 2-row header.  Row 0 carries the
# group labels (Departure / T/S / Arrival ...), row 1 the sub labels.  This is
# the flattened, stable column order we normalise every file to.
CANONICAL_COLUMNS = [
    "Vessel/Voyage",
    "Departure Terminal",
    "Departure ETA",
    "Departure ETD",
    "T/S Place",
    "T/S ETD",
    "Arrival Terminal",
    "Arrival ETA",
    "T/T",
    "Document Closing Time",
    "Cargo Closing Time",
]


# --------------------------------------------------------------------------- #
# cell cleaning (rules only, no AI)                                           #
# --------------------------------------------------------------------------- #
def clean_cell(value) -> str:
    """Render a single cell as tidy one-line text."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value)
    # KMTC glues the ETA date and a stray newline together ("...18:20\n")
    text = text.replace("\r", " ").replace("\n", " ")
    text = " ".join(text.split())
    if text in {"-", "nan", "NaN", "None"}:
        return "" if text != "-" else "-"
    return text


def engine_for(path: Path) -> str | None:
    if path.suffix.lower() == ".xls":
        return "xlrd"
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return "openpyxl"
    return None


# --------------------------------------------------------------------------- #
# read one workbook                                                          #
# --------------------------------------------------------------------------- #
def read_workbook_raw(path: Path) -> dict[str, "pd.DataFrame"]:
    """Return {sheet_name: raw DataFrame (header=None, everything as text)}."""
    try:
        xls = pd.ExcelFile(path, engine=engine_for(path))
    except ImportError as exc:  # missing engine
        sys.exit(f"Cannot read {path.name}: {exc}")
    sheets: dict[str, pd.DataFrame] = {}
    for name in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=name, header=None, dtype=object)
        # pandas >=2.1: element-wise is .map (applymap removed in 3.0)
        raw = raw.map(clean_cell) if hasattr(raw, "map") else raw.applymap(clean_cell)
        # drop fully empty rows / columns
        raw = raw.loc[~(raw == "").all(axis=1)]
        raw = raw.loc[:, ~(raw == "").all(axis=0)]
        sheets[name] = raw.reset_index(drop=True)
    return sheets


def flatten_kmtc_sheet(raw: "pd.DataFrame") -> "pd.DataFrame | None":
    """
    Turn a raw KMTC leg-schedule sheet (2-row header + data) into a tidy frame
    with CANONICAL_COLUMNS.  Returns None if it does not look like that layout.
    """
    if raw.shape[0] < 3 or raw.shape[1] < 9:
        return None
    row0 = list(raw.iloc[0])
    row1 = list(raw.iloc[1])
    if clean_cell(row0[0]).lower() not in {"vessel/voyage", "vessel / voyage"}:
        return None

    # forward-fill the group labels across the merged cells of row 0
    filled: list[str] = []
    last = ""
    for cell in row0:
        cell = clean_cell(cell)
        if cell:
            last = cell
        filled.append(last)

    names: list[str] = []
    for grp, sub in zip(filled, row1):
        sub = clean_cell(sub)
        grp = clean_cell(grp)
        if grp and sub and grp.lower() != sub.lower():
            names.append(f"{grp} {sub}")
        else:
            names.append(sub or grp)

    body = raw.iloc[2:].copy()
    body.columns = names[: body.shape[1]]

    # keep / order the columns we know; ignore anything unexpected
    keep = [c for c in CANONICAL_COLUMNS if c in body.columns]
    body = body[keep]
    for col in CANONICAL_COLUMNS:
        if col not in body.columns:
            body[col] = ""
    body = body[CANONICAL_COLUMNS]

    # drop rows with no vessel and repeated header rows
    body = body[body["Vessel/Voyage"].astype(bool)]
    body = body[body["Vessel/Voyage"].str.lower() != "vessel/voyage"]
    return body.reset_index(drop=True)


# --------------------------------------------------------------------------- #
# text rendering                                                             #
# --------------------------------------------------------------------------- #
def render_table(df: "pd.DataFrame") -> list[str]:
    """Aligned table + TSV block for a DataFrame."""
    cols = list(df.columns)
    rows = [[clean_cell(v) for v in row] for row in df.itertuples(index=False, name=None)]
    grid = [cols] + rows
    widths = [max(len(r[c]) for r in grid) for c in range(len(cols))]

    out: list[str] = []
    for i, r in enumerate(grid):
        out.append(" | ".join(cell.ljust(widths[c]) for c, cell in enumerate(r)).rstrip())
        if i == 0:
            out.append("-+-".join("-" * w for w in widths))
    out.append("")
    out.append("--- TSV ---")
    out.append("\t".join(cols))
    for r in rows:
        out.append("\t".join(r))
    out.append("")
    return out


def parse_dt(series: "pd.Series") -> "pd.Series":
    """Best-effort '2026.Sep.02 14:30' -> datetime, for sorting only."""
    return pd.to_datetime(series, format="%Y.%b.%d %H:%M", errors="coerce")


# --------------------------------------------------------------------------- #
# driver                                                                     #
# --------------------------------------------------------------------------- #
def collect_files(raw_paths: Iterable[str]) -> list[Path]:
    found: list[Path] = []
    for raw in raw_paths:
        p = Path(raw)
        if p.is_dir():
            found += sorted(
                c for c in p.iterdir()
                if c.is_file() and c.suffix.lower() in EXCEL_SUFFIXES
            )
        elif p.is_file():
            found.append(p)
        else:
            print(f"! not found, skipping: {p}", file=sys.stderr)
    # de-dup, keep order
    seen: set[Path] = set()
    unique: list[Path] = []
    for f in found:
        rp = f.resolve()
        if rp not in seen:
            seen.add(rp)
            unique.append(f)
    return unique


def main() -> None:
    script_dir = Path(__file__).resolve().parent
    repo_root = script_dir.parent

    ap = argparse.ArgumentParser(
        description="Merge every Excel file in the KMTC folder and dump plain text (no AI)."
    )
    ap.add_argument(
        "paths", nargs="*",
        help=f"Folders or Excel files to read. Default: {DEFAULT_FOLDER}/ at the repo root.",
    )
    ap.add_argument("--out-dir", default=None,
                    help="Where to write kmtc_combined.* (default: <repo>/output).")
    ap.add_argument("--stdout-only", action="store_true",
                    help="Only print, do not write files.")
    args = ap.parse_args()

    raw_paths = args.paths or [str(repo_root / DEFAULT_FOLDER)]
    files = collect_files(raw_paths)
    if not files:
        sys.exit("Nothing to do: no .xls/.xlsx files found.")

    lines: list[str] = []
    lines.append("KMTC Laem Chabang -> Shanghai leg schedules -- merged extract (no AI)")
    lines.append(f"SOURCE FILES ({len(files)}): " + ", ".join(f.name for f in files))
    lines.append("")

    merged_parts: list[pd.DataFrame] = []

    for path in files:
        lines.append("#" * 78)
        lines.append(f"# FILE: {path.name}")
        lines.append("#" * 78)
        try:
            sheets = read_workbook_raw(path)
        except Exception as exc:  # noqa: BLE001
            lines.append(f"! failed to read: {exc}")
            lines.append("")
            print(f"! failed on {path}: {exc}", file=sys.stderr)
            continue

        for sheet_name, raw in sheets.items():
            lines.append("")
            lines.append(f"SHEET: {sheet_name}   ({raw.shape[0]} rows x {raw.shape[1]} cols, raw)")
            lines.append("-" * 70)

            tidy = flatten_kmtc_sheet(raw)
            if tidy is None:
                lines.append("(not a recognised KMTC leg-schedule layout -- raw cell dump)")
                lines += render_table(raw)
                continue

            lines += render_table(tidy)
            part = tidy.copy()
            part.insert(0, "Source File", path.name)
            merged_parts.append(part)

    # ------------------------------------------------------------------ #
    # merged table                                                       #
    # ------------------------------------------------------------------ #
    combined: pd.DataFrame | None = None
    if merged_parts:
        combined = pd.concat(merged_parts, ignore_index=True)

        # drop exact-duplicate sailings that repeat across monthly exports
        dedup_cols = [c for c in combined.columns if c != "Source File"]
        before = len(combined)
        combined = combined.drop_duplicates(subset=dedup_cols, keep="first")
        removed = before - len(combined)

        # chronological order by departure ETD (falls back to input order)
        combined["_sort"] = parse_dt(combined["Departure ETD"])
        combined = combined.sort_values(
            "_sort", kind="stable", na_position="last"
        ).drop(columns="_sort").reset_index(drop=True)

        lines.append("")
        lines.append("=" * 78)
        lines.append(f"= MERGED TABLE  ({len(combined)} unique sailings, "
                     f"{removed} duplicate rows removed)")
        lines.append("=" * 78)
        lines += render_table(combined)

    text = "\n".join(lines)
    print(text)

    if not args.stdout_only:
        out_dir = Path(args.out_dir) if args.out_dir else (repo_root / "output")
        out_dir.mkdir(parents=True, exist_ok=True)

        txt_path = out_dir / f"{OUT_STEM}.extracted.txt"
        txt_path.write_text(text, encoding="utf-8")
        print(f"\n-> wrote {txt_path}", file=sys.stderr)

        if combined is not None:
            csv_path = out_dir / f"{OUT_STEM}.csv"
            combined.to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"-> wrote {csv_path}", file=sys.stderr)


if __name__ == "__main__":
    main()
