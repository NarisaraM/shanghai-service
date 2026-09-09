#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
COSCO SHIPPING Lines - Sailing Schedule fetcher
================================================

Queries the "Find Schedule by City Pairs" search on the COSCO SHIPPING Lines
e-Business site (https://elines.coscoshipping.com) directly through the
same public JSON API the website itself calls, and exports the result to
an Excel file.

Default route : Laem Chabang, Thailand  ->  Shanghai, China
Default period: today  ->  the last day of October (same year)

Usage
-----
    python cosco_schedule.py
    python cosco_schedule.py --origin "Laem Chabang" --destination "Shanghai"
    python cosco_schedule.py --from 2026-09-01 --to 2026-10-31
    python cosco_schedule.py --output my_schedule.xlsx

Requirements
------------
    pip install requests openpyxl

Notes
-----
- This script talks to COSCO's own public endpoints, the same ones the
  browser uses when you search on the website - no login is required.
- It must be run from a machine that has normal internet access to
  elines.coscoshipping.com (a locked-down / proxied sandbox may block it).
"""

import argparse
import calendar
import datetime as dt
import sys

import requests

BASE = "https://elines.coscoshipping.com"
CITY_LOOKUP_URL = f"{BASE}/ebbase/public/general/findCityDistrictByPrefix"
SCHEDULE_URL = f"{BASE}/ebschedule/public/purpoShipmentWs"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
    "Referer": f"{BASE}/ebusiness/sailingSchedule/searchByCity",
}


def find_city(session: requests.Session, prefix: str) -> dict:
    """Look up a city on COSCO's city-search endpoint and return the best match.

    Raises SystemExit with a helpful message if nothing (or nothing usable)
    is found, so the caller can correct the spelling.
    """
    params = {"prefix": prefix, "timestamp": int(dt.datetime.now().timestamp() * 1000)}
    resp = session.get(CITY_LOOKUP_URL, params=params, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    payload = resp.json()
    candidates = (payload.get("data") or {}).get("content") or []
    if not candidates:
        raise SystemExit(f'No city found on COSCO for "{prefix}". Check the spelling.')

    # Prefer an exact (case-insensitive) match on the city name; otherwise take the first hit.
    exact = [c for c in candidates if c.get("cityLocName", "").lower() == prefix.lower()]
    match = exact[0] if exact else candidates[0]

    if len(candidates) > 1 and not exact:
        print(f'Multiple matches for "{prefix}", using: {match["fullFormate"]}', file=sys.stderr)

    return match


def fetch_schedule(
    session: requests.Session,
    origin: dict,
    destination: dict,
    date_from: dt.date,
    date_to: dt.date,
    cargo_nature: str = "All",
) -> list:
    """Call the sailing-schedule API and return the list of voyage records."""
    body = {
        "fromDate": date_from.isoformat(),
        "toDate": date_to.isoformat(),
        "pickup": "B",       # Both (OB Haulage: CY/DOOR)
        "delivery": "B",     # Both (IB Haulage: CY/DOOR)
        "estimateDate": "D",  # "D" = Earliest Departure Date (as opposed to "A" = Arrival)
        "originCityUuid": origin["cityUuid"],
        "destinationCityUuid": destination["cityUuid"],
        "originCity": (
            f'{origin["cityLocName"]}, {origin.get("county") or ""},'
            f'{origin["state"]},{origin["country"]},{origin["unloCode"]}'
        ),
        "destinationCity": (
            f'{destination["cityLocName"]}, {destination.get("county") or ""},'
            f'{destination["state"]},{destination["country"]},{destination["unloCode"]}'
        ),
        "cargoNature": cargo_nature,
    }

    resp = session.post(SCHEDULE_URL, json=body, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    payload = resp.json()

    if str(payload.get("code")) != "200":
        raise SystemExit(f"COSCO API returned an error: {payload}")

    records = ((payload.get("data") or {}).get("content") or {}).get("data") or []
    return records


def to_rows(records: list) -> list:
    """Flatten the raw API records into the columns we care about, sorted by ETD."""
    rows = []
    for r in records:
        rows.append(
            {
                "Cargo Nature": r.get("cargoNature"),
                "Cut Off": r.get("cutOff"),
                "Vessel": r.get("vessel"),
                "Service/Voyage": f'{r.get("service")}/{r.get("extVoyage")}',
                "POL": r.get("pol"),
                "ETD": r.get("etd"),
                "POD": r.get("pod"),
                "ETA": r.get("eta"),
                "Transit (days)": r.get("transitTime"),
                "OB Haulage": r.get("outboundHaulage"),
                "IB Haulage": r.get("inboundHaulage"),
            }
        )
    rows.sort(key=lambda x: x["ETD"] or "")
    return rows


def default_to_date(today: dt.date, target_month: int) -> dt.date:
    """Last calendar day of `target_month`, rolling into next year if that
    month has already passed this year."""
    year = today.year if target_month >= today.month else today.year + 1
    last_day = calendar.monthrange(year, target_month)[1]
    return dt.date(year, target_month, last_day)


def export_excel(rows: list, path: str, origin_label: str, destination_label: str,
                  date_from: dt.date, date_to: dt.date) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Sailing Schedule"

    ws.merge_cells("A1:K1")
    ws["A1"] = f"COSCO Sailing Schedule: {origin_label} -> {destination_label}"
    ws["A1"].font = Font(bold=True, size=13)

    ws.merge_cells("A2:K2")
    ws["A2"] = f"Departure window: {date_from.isoformat()} to {date_to.isoformat()}"
    ws["A2"].font = Font(italic=True, size=10)

    header_row = 4
    headers = list(rows[0].keys()) if rows else [
        "Cargo Nature", "Cut Off", "Vessel", "Service/Voyage", "POL", "ETD",
        "POD", "ETA", "Transit (days)", "OB Haulage", "IB Haulage",
    ]
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    for col_idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=col_name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row_idx, row in enumerate(rows, start=header_row + 1):
        for col_idx, col_name in enumerate(headers, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row.get(col_name))

    for col_idx, col_name in enumerate(headers, start=1):
        width = max(14, len(col_name) + 4)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate
    wb.save(path)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--origin", default="Laem Chabang", help='Origin city (default: "Laem Chabang")')
    parser.add_argument("--destination", default="Shanghai", help='Destination city (default: "Shanghai")')
    parser.add_argument("--from", dest="date_from", default=None,
                         help="Start date YYYY-MM-DD (default: today)")
    parser.add_argument("--to", dest="date_to", default=None,
                         help="End date YYYY-MM-DD (default: last day of October, current year)")
    parser.add_argument("--cargo-nature", default="All", help='Cargo nature filter (default: "All")')
    parser.add_argument("--output", default="cosco_schedule.xlsx", help="Output .xlsx file path")
    args = parser.parse_args()

    today = dt.date.today()
    date_from = dt.date.fromisoformat(args.date_from) if args.date_from else today
    date_to = dt.date.fromisoformat(args.date_to) if args.date_to else default_to_date(today, target_month=10)

    if date_to < date_from:
        raise SystemExit(f"--to ({date_to}) must not be before --from ({date_from}).")

    session = requests.Session()

    print(f'Looking up "{args.origin}" ...')
    origin = find_city(session, args.origin)
    print(f'  -> {origin["fullFormate"]}  [{origin["unloCode"]}]')

    print(f'Looking up "{args.destination}" ...')
    destination = find_city(session, args.destination)
    print(f'  -> {destination["fullFormate"]}  [{destination["unloCode"]}]')

    print(f"Fetching sailing schedule {date_from} -> {date_to} ...")
    records = fetch_schedule(session, origin, destination, date_from, date_to, args.cargo_nature)
    rows = to_rows(records)

    if not rows:
        print("No sailings found for this route/date range.")
        return

    print(f"Found {len(rows)} sailing(s).\n")
    col_widths = {k: max(len(k), *(len(str(r[k])) for r in rows)) for k in rows[0]}
    header_line = " | ".join(k.ljust(col_widths[k]) for k in rows[0])
    print(header_line)
    print("-" * len(header_line))
    for r in rows:
        print(" | ".join(str(r[k]).ljust(col_widths[k]) for k in r))

    export_excel(rows, args.output, origin["cityLocName"], destination["cityLocName"], date_from, date_to)
    print(f"\nSaved: {args.output}")


if __name__ == "__main__":
    main()
