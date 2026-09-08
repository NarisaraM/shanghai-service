#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tslines_schedule.py
====================

Query T.S. Lines' public "Port To Port Schedule" service
(https://www.tslines.com/th/schedule) for a given port pair and date
range, and export the result to an Excel file.

Default behaviour (no arguments): looks up the LAEM CHABANG (Thailand) ->
SHANGHAI (China) sailing schedule from the first day of the current month
through 31 December of the current year, and saves an .xlsx file.

Usage
-----
    python3 tslines_schedule.py
    python3 tslines_schedule.py --pol THLCB --pod CNSHA
    python3 tslines_schedule.py --date-from 2026-09-01 --date-to 2026-12-31
    python3 tslines_schedule.py --output tslines_lch_sha.xlsx

How it works
-------------
tslines.com's "Port To Port Schedule" page (a Nuxt.js single page app)
calls a public JSON API directly:

    POST https://epd.tslines.com/v1/LongTermSchedule
    Content-Type: application/json
    {
        "PORT_OF_LOADING": "THLCB",
        "PORT_OF_DISCHARGE": "CNSHA",
        "POL_ETD_START_AT": "YYYY/MM/DD",
        "POL_ETD_END_AT": "YYYY/MM/DD"
    }

No login/cookie is required (confirmed: the endpoint lives on a different
sub-domain, epd.tslines.com, than the website itself, so the browser does
not even send it any www.tslines.com cookies). This script calls that
endpoint directly instead of driving a browser, which is faster and far
more reliable than scripting the web page's UI.

Notes
-----
* POL/POD are T.S. Lines' own port codes (UN/LOCODE-style), e.g.
  "THLCB" = Laem Chabang, "CNSHA" = Shanghai. The defaults below cover
  the Laem Chabang -> Shanghai lane; pass --pol/--pod to use a different
  pair if you know the code (check the site's dropdown, or open
  https://www.tslines.com/th/schedule and inspect the port codes shown
  in the results table's POL/POD columns).
* T.S. Lines' own disclaimer applies: the schedule is for reference
  only, please confirm with your local agent / booking staff before
  relying on it for cut-off planning.
"""

import argparse
import datetime as dt
import sys

import requests

API_URL = "https://epd.tslines.com/v1/LongTermSchedule"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Content-Type": "application/json",
    "Origin": "https://www.tslines.com",
    "Referer": "https://www.tslines.com/",
    "Accept": "application/json, text/plain, */*",
}

DEFAULT_POL = "THLCB"  # Laem Chabang, Thailand
DEFAULT_POD = "CNSHA"  # Shanghai, China


def default_date_range() -> tuple[str, str]:
    """First day of the current month -> 31 December of the current year."""
    today = dt.date.today()
    date_from = today.replace(day=1)
    date_to = dt.date(today.year, 12, 31)
    return date_from.isoformat(), date_to.isoformat()


def to_slash_date(iso_date: str) -> str:
    """Convert YYYY-MM-DD -> YYYY/MM/DD (the format T.S. Lines' API expects)."""
    return dt.date.fromisoformat(iso_date).strftime("%Y/%m/%d")


def fetch_schedule(pol: str, pod: str, date_from: str, date_to: str, timeout: int = 30) -> list[dict]:
    """Call T.S. Lines' LongTermSchedule API and return the list of sailing rows."""
    payload = {
        "PORT_OF_LOADING": pol,
        "PORT_OF_DISCHARGE": pod,
        "POL_ETD_START_AT": to_slash_date(date_from),
        "POL_ETD_END_AT": to_slash_date(date_to),
    }
    resp = requests.post(API_URL, json=payload, headers=HEADERS, timeout=timeout)

    if resp.status_code != 200:
        raise RuntimeError(
            f"T.S. Lines API returned HTTP {resp.status_code}: {resp.text[:300]}"
        )

    data = resp.json()
    return data.get("docs") or []


def rows_to_table(rows: list[dict]) -> list[dict]:
    """Reshape the raw API rows into a clean, human-friendly table."""
    table = []
    for r in rows:
        table.append(
            {
                "Service": r.get("SERVICE", ""),
                "Vessel Name": r.get("VESSEL_NAME", ""),
                "Voyage (POL)": r.get("VOYAGE", ""),
                "POL": r.get("POL", ""),
                "POL Terminal": r.get("TERMINAL_L", ""),
                "ETD": r.get("ETD", ""),
                "CY Cut-off": r.get("CUTOFF", ""),
                "VGM Cut-off": r.get("VGM_CUT", ""),
                "POD": r.get("POD", ""),
                "POD Terminal": r.get("TERMINAL_D", ""),
                "Voyage (POD)": r.get("POD_VOYAGE", ""),
                "ETA": r.get("ETA", ""),
                "Transit (days)": r.get("SALILNG", ""),
            }
        )

    def sort_key(row):
        try:
            return dt.datetime.strptime(row["ETD"], "%Y/%m/%d %H:%M")
        except (ValueError, TypeError):
            return dt.datetime.max

    table.sort(key=sort_key)
    return table


def export_to_excel(table: list[dict], output_path: str, title: str) -> None:
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover
        raise SystemExit(
            "This script needs pandas + openpyxl to write Excel files.\n"
            "Install them with:  pip install pandas openpyxl"
        ) from exc

    df = pd.DataFrame(table)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Schedule", startrow=1)
        sheet = writer.sheets["Schedule"]
        sheet["A1"] = title
        sheet["A1"].font = sheet["A1"].font.copy(bold=True, size=12)

        for col_cells in sheet.columns:
            length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col_cells)
            col_letter = col_cells[0].column_letter
            sheet.column_dimensions[col_letter].width = min(max(length + 2, 12), 48)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Query T.S. Lines' Port To Port schedule and export it to Excel."
    )
    parser.add_argument("--pol", default=DEFAULT_POL, help=f"Port of Loading code (default: {DEFAULT_POL} = Laem Chabang)")
    parser.add_argument("--pod", default=DEFAULT_POD, help=f"Port of Discharge code (default: {DEFAULT_POD} = Shanghai)")
    parser.add_argument(
        "--date-from",
        dest="date_from",
        default=None,
        help="Start date YYYY-MM-DD (default: 1st of the current month)",
    )
    parser.add_argument(
        "--date-to",
        dest="date_to",
        default=None,
        help="End date YYYY-MM-DD (default: 31 Dec of the current year)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output .xlsx path (default: TSLines_Schedule_<POL>_<POD>_<dates>.xlsx)",
    )
    args = parser.parse_args()

    default_from, default_to = default_date_range()
    date_from = args.date_from or default_from
    date_to = args.date_to or default_to

    print(f"Querying T.S. Lines schedule: {args.pol} -> {args.pod}, {date_from} to {date_to} ...")
    try:
        rows = fetch_schedule(args.pol, args.pod, date_from, date_to)
    except requests.exceptions.RequestException as exc:
        sys.exit(f"Network error while calling T.S. Lines' schedule API: {exc}")
    except RuntimeError as exc:
        sys.exit(str(exc))

    if not rows:
        print("No sailings found for this port pair / date range.")
        print("Tip: double check the POL/POD port codes (e.g. THLCB, CNSHA) against the dropdown on https://www.tslines.com/th/schedule")
        return

    table = rows_to_table(rows)
    print(f"Found {len(table)} sailing(s).\n")

    header = f"{'Vessel':<20}{'Voy.':<10}{'ETD':<18}{'ETA':<18}{'Transit':<9}{'Svc':<5}"
    print(header)
    print("-" * len(header))
    for row in table:
        print(
            f"{row['Vessel Name']:<20}{str(row['Voyage (POL)']):<10}{row['ETD']:<18}"
            f"{row['ETA']:<18}{str(row['Transit (days)']):<9}{row['Service']:<5}"
        )

    output_path = args.output or (
        f"TSLines_Schedule_{args.pol}_{args.pod}_{date_from}_to_{date_to}.xlsx"
    )
    title = f"T.S. Lines Port To Port Schedule: {args.pol} -> {args.pod}  ({date_from} to {date_to})"
    export_to_excel(table, output_path, title)
    print(f"\nSaved: {output_path}")


if __name__ == "__main__":
    main()
