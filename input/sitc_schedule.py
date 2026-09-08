#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sitc_schedule.py
=================

Query SITC (Container Lines) Thailand's public "Shipping Schedule" service
(the same lookup exposed on https://ebusiness.sitcline.com under
Schedule > Shipping Schedule) for a given port pair and date range, and
export the result to an Excel file.

Default behaviour (no arguments): looks up the LAEM CHABANG (Thailand) ->
SHANGHAI (China) sailing schedule from the first day of the current month
through 31 December of the current year, and saves an .xlsx file.

Usage
-----
    python3 sitc_schedule.py
    python3 sitc_schedule.py --pol "LAEM CHABANG" --pod "SHANGHAI"
    python3 sitc_schedule.py --date-from 2026-09-01 --date-to 2026-12-31
    python3 sitc_schedule.py --output sitc_lch_sha.xlsx

How it works
-------------
The SITC e-Business site (https://ebusiness.sitcline.com) is a Vue single
page app. Its "Shipping Schedule" search box calls a public JSON API
directly:

    POST https://ebusiness.sitcline.com/api/equery/voyageInfo/searchPlan
         ?pol=<POL NAME>&pod=<POD NAME>&datefrom=YYYY-MM-DD&dateto=YYYY-MM-DD&type=false

No login/cookie is required. This script calls that endpoint directly
instead of driving a browser, which is faster and far more reliable than
scripting the web page's UI.

Notes
-----
* Port names must be given the way SITC's own port master lists them
  (e.g. "LAEM CHABANG", "SHANGHAI"). If you are not sure of the exact
  spelling, use --list-ports to search SITC's port autocomplete first.
* `type=false` queries by ETD (departure-based); pass --by-eta to switch
  to `type=true` (arrival-based) if you ever need that instead.
* SITC's disclaimer applies: the schedule is for reference only, please
  confirm with your local agent / booking staff before relying on it.
"""

import argparse
import calendar
import datetime as dt
import sys

import requests

BASE_URL = "https://ebusiness.sitcline.com"
SCHEDULE_URL = f"{BASE_URL}/api/equery/voyageInfo/searchPlan"
PORT_AUTOCOMPLETE_URL = f"{BASE_URL}/api/ebusiness/autocomplete/queryListPort"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Referer": f"{BASE_URL}/",
    "Accept": "application/json, text/plain, */*",
}


def default_date_range() -> tuple[str, str]:
    """First day of the current month -> 31 December of the current year."""
    today = dt.date.today()
    date_from = today.replace(day=1)
    date_to = dt.date(today.year, 12, 31)
    return date_from.isoformat(), date_to.isoformat()


def search_ports(keyword: str, timeout: int = 20) -> list[dict]:
    """Look up SITC's port master list for a keyword (helper for --list-ports)."""
    params = {
        "currentPage": 1,
        "pageSize": 20,
        "portCode": f"%{keyword}%",
        "portName": f"%{keyword}%",
        "portNameCn": f"%{keyword}%",
    }
    resp = requests.get(
        PORT_AUTOCOMPLETE_URL, params=params, headers=HEADERS, timeout=timeout
    )
    resp.raise_for_status()
    payload = resp.json()
    if not payload.get("success", True):
        raise RuntimeError(f"Port lookup failed: {payload}")
    rows = payload.get("data")
    if isinstance(rows, dict):
        rows = rows.get("rows") or rows.get("list") or []
    return rows or []


def fetch_schedule(
    pol: str,
    pod: str,
    date_from: str,
    date_to: str,
    by_eta: bool = False,
    timeout: int = 30,
) -> list[dict]:
    """Call SITC's searchPlan API and return the list of voyage rows."""
    params = {
        "pol": pol,
        "pod": pod,
        "datefrom": date_from,
        "dateto": date_to,
        "type": "true" if by_eta else "false",
    }
    resp = requests.post(SCHEDULE_URL, params=params, headers=HEADERS, timeout=timeout)
    resp.raise_for_status()
    payload = resp.json()

    if not payload.get("success", True):
        msg = payload.get("errorMessages") or payload.get("message") or payload
        raise RuntimeError(f"SITC API returned an error: {msg}")

    data = payload.get("data") or {}
    rows = data.get("list") or []
    return rows


def rows_to_table(rows: list[dict]) -> list[dict]:
    """Reshape the raw API rows into a clean, human-friendly table."""
    table = []
    for r in rows:
        table.append(
            {
                "Service Line": r.get("serviceLineCode", ""),
                "Vessel Name": r.get("vesselName", ""),
                "Voyage No.": r.get("voyageNo", ""),
                "POL": r.get("polName", ""),
                "POL Terminal": r.get("polTerminalName", ""),
                "ETD": r.get("weekEtd") or r.get("etd", ""),
                "POD": r.get("podName", ""),
                "POD Terminal": r.get("podTerminalName", ""),
                "ETA": r.get("weekEta") or r.get("eta", ""),
                "Transit (days)": r.get("underway", ""),
            }
        )
    # Sort by ETD date so the sailing calendar reads chronologically.
    def sort_key(row):
        try:
            return dt.date.fromisoformat(row["ETD"])
        except (ValueError, TypeError):
            return dt.date.max

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

        # Auto-fit column widths (rough heuristic).
        for col_cells in sheet.columns:
            length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col_cells)
            col_letter = col_cells[0].column_letter
            sheet.column_dimensions[col_letter].width = min(max(length + 2, 12), 45)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Query SITC Thailand's shipping schedule and export it to Excel."
    )
    parser.add_argument("--pol", default="LAEM CHABANG", help="Port of Loading (default: LAEM CHABANG)")
    parser.add_argument("--pod", default="SHANGHAI", help="Port of Discharge (default: SHANGHAI)")
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
        "--by-eta",
        action="store_true",
        help="Query by ETA (arrival) instead of the default ETD (departure) based search",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output .xlsx path (default: SITC_Schedule_<POL>_<POD>_<dates>.xlsx)",
    )
    parser.add_argument(
        "--list-ports",
        metavar="KEYWORD",
        help="Instead of querying a schedule, search SITC's port list for KEYWORD and print matches (useful to confirm exact port name spelling)",
    )
    args = parser.parse_args()

    if args.list_ports:
        matches = search_ports(args.list_ports)
        if not matches:
            print(f"No ports found matching '{args.list_ports}'.")
            return
        print(f"{'Port Code':<10} {'Country':<8} Port Name")
        for m in matches:
            print(f"{m.get('portCode',''):<10} {m.get('portCountryCode') or m.get('countryCode',''):<8} {m.get('portName','')}")
        return

    default_from, default_to = default_date_range()
    date_from = args.date_from or default_from
    date_to = args.date_to or default_to

    print(f"Querying SITC schedule: {args.pol} -> {args.pod}, {date_from} to {date_to} ...")
    try:
        rows = fetch_schedule(args.pol, args.pod, date_from, date_to, by_eta=args.by_eta)
    except requests.exceptions.RequestException as exc:
        sys.exit(f"Network error while calling SITC's schedule API: {exc}")
    except RuntimeError as exc:
        sys.exit(str(exc))

    if not rows:
        print("No sailings found for this port pair / date range.")
        print("Tip: run with --list-ports \"<keyword>\" to confirm the exact port name SITC expects.")
        return

    table = rows_to_table(rows)
    print(f"Found {len(table)} sailing(s).\n")

    # Console preview
    header = f"{'Vessel':<20}{'Voy.':<8}{'ETD':<12}{'ETA':<12}{'Transit':<9}{'Line':<6}"
    print(header)
    print("-" * len(header))
    for row in table:
        print(
            f"{row['Vessel Name']:<20}{str(row['Voyage No.']):<8}{row['ETD']:<12}"
            f"{row['ETA']:<12}{str(row['Transit (days)']):<9}{row['Service Line']:<6}"
        )

    output_path = args.output or (
        f"SITC_Schedule_{args.pol.replace(' ', '')}_{args.pod.replace(' ', '')}_"
        f"{date_from}_to_{date_to}.xlsx"
    )
    title = f"SITC Shipping Schedule: {args.pol} -> {args.pod}  ({date_from} to {date_to})"
    export_to_excel(table, output_path, title)
    print(f"\nSaved: {output_path}")


if __name__ == "__main__":
    main()
