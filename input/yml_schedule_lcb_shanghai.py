#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YML (Yang Ming) Point-to-Point schedule fetcher
=================================================
Laem Chabang (THLCB) -> Shanghai (CNSHA)

This script calls the same JSON API that powers the "Point-to-Point Search"
page on the Yang Ming website:

    https://www.yangming.com/en/esolution/schedule/point_to_point_search

Endpoint used:
    GET https://www.yangming.com/api/P2P/GetP2PRoutes

The website itself limits every query to a maximum 30-day window
("*End date must be within 30 days from start date."), so this script
automatically splits the requested period (current month -> end of October)
into <=30-day chunks, calls the API once per chunk, merges + de-duplicates
the results, and exports everything to a single Excel/CSV file.

Requirements:
    pip install requests openpyxl

Usage:
    python yml_schedule_lcb_shanghai.py
    python yml_schedule_lcb_shanghai.py --start 2026-09-09 --end 2026-10-31
    python yml_schedule_lcb_shanghai.py --from THLCB --to CNSHA --format csv
"""

import argparse
import calendar
import csv
import datetime as dt
import sys
import time

import requests

API_URL = "https://www.yangming.com/api/P2P/GetP2PRoutes"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://www.yangming.com/en/esolution/schedule/point_to_point_search",
}

MAX_WINDOW_DAYS = 30  # hard limit enforced by the YML website


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def month_end(year: int, month: int) -> dt.date:
    """Last calendar day of `year`-`month`."""
    last_day = calendar.monthrange(year, month)[1]
    return dt.date(year, month, last_day)


def date_windows(start: dt.date, end: dt.date, max_days: int = MAX_WINDOW_DAYS):
    """Split [start, end] into consecutive windows of at most `max_days` days."""
    windows = []
    cur = start
    while cur <= end:
        chunk_end = min(cur + dt.timedelta(days=max_days - 1), end)
        windows.append((cur, chunk_end))
        cur = chunk_end + dt.timedelta(days=1)
    return windows


def fetch_window(origin: str, dest: str, start: dt.date, end: dt.date,
                  service_term_from: str = "Y", service_term_to: str = "Y",
                  priority: str = "ALL", date_definition: str = "DEP",
                  timeout: int = 30):
    """Call the YML P2P API for a single (<=30 day) window and return the list of sailings."""
    params = {
        "locationCodeFrom": origin,
        "serviceTermFrom": service_term_from,
        "locationCodeTo": dest,
        "serviceTermTo": service_term_to,
        "priorityWay": priority,
        "dateDefinition": date_definition,
        "startDate": start.strftime("%Y%m%d"),
        "endDate": end.strftime("%Y%m%d"),
    }
    resp = requests.get(API_URL, params=params, headers=HEADERS, timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    return data if isinstance(data, list) else data.get("data", [])


def fetch_schedule(origin: str, dest: str, start: dt.date, end: dt.date, **kwargs):
    """Fetch the full schedule for [start, end], transparently paging around
    the website's 30-day-per-query limit, and de-duplicate the results."""
    all_rows = []
    seen = set()
    for win_start, win_end in date_windows(start, end):
        print(f"  Fetching {win_start:%Y-%m-%d} -> {win_end:%Y-%m-%d} ...")
        try:
            rows = fetch_window(origin, dest, win_start, win_end, **kwargs)
        except requests.RequestException as exc:
            print(f"    !! request failed: {exc}", file=sys.stderr)
            continue
        for row in rows:
            key = (row.get("masterVoyageCode"), row.get("masterETD"), row.get("masterVesselCode"))
            if key in seen:
                continue
            seen.add(key)
            all_rows.append(row)
        time.sleep(0.3)  # be polite to the API
    all_rows.sort(key=lambda r: (r.get("masterETD") or "", r.get("masterVoyageCode") or ""))
    return all_rows


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------
COLUMN_MAP = [
    ("masterVesselName", "Vessel"),
    ("masterVoyageCode", "Voyage Code"),
    ("masterComnVoyage", "Voy (Common)"),
    ("placeOfReceipt", "From"),
    ("masterETD", "ETD"),
    ("placeOfDelivery", "To"),
    ("masterETA", "ETA"),
    ("transitDays", "Transit Days"),
    ("transshipment", "Transshipment"),
    ("cutoffCY", "CY Cutoff"),
    ("cutoffSI", "SI Cutoff"),
    ("cutoffVGM", "VGM Cutoff"),
]


def rows_to_table(rows):
    headers = [h for _, h in COLUMN_MAP]
    table = []
    for r in rows:
        table.append([r.get(k, "") if r.get(k) is not None else "" for k, _ in COLUMN_MAP])
    return headers, table


def print_table(headers, table):
    if not table:
        print("No sailings found for the requested period.")
        return
    widths = [len(h) for h in headers]
    for row in table:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], len(str(cell)))
    line = "  ".join(h.ljust(widths[i]) for i, h in enumerate(headers))
    print(line)
    print("-" * len(line))
    for row in table:
        print("  ".join(str(cell).ljust(widths[i]) for i, cell in enumerate(row)))


def save_csv(path, headers, table):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(table)


def save_xlsx(path, headers, table):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except ImportError:
        print("openpyxl not installed (pip install openpyxl) - saving CSV instead.")
        save_csv(path.rsplit(".", 1)[0] + ".csv", headers, table)
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "YML Schedule"
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in table:
        ws.append(row)
    for col in ws.columns:
        length = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(length + 2, 10), 40)
    wb.save(path)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def parse_args():
    today = dt.date.today()
    default_start = today
    default_end_year = today.year if today.month <= 10 else today.year + 1
    default_end = month_end(default_end_year, 10)

    p = argparse.ArgumentParser(description="Fetch YML (Yang Ming) sailing schedule.")
    p.add_argument("--from", dest="origin", default="THLCB",
                    help="Origin port UNLOCODE (default: THLCB = Laem Chabang)")
    p.add_argument("--to", dest="dest", default="CNSHA",
                    help="Destination port UNLOCODE (default: CNSHA = Shanghai)")
    p.add_argument("--start", default=default_start.isoformat(),
                    help=f"Start date YYYY-MM-DD (default: today, {default_start.isoformat()})")
    p.add_argument("--end", default=default_end.isoformat(),
                    help=f"End date YYYY-MM-DD (default: end of October, {default_end.isoformat()})")
    p.add_argument("--priority", default="ALL", choices=["ALL", "D", "T"],
                    help="ALL / D=Direct only / T=Transshipment only (default: ALL)")
    p.add_argument("--date-definition", default="DEP", choices=["DEP", "ARR"],
                    help="Filter by Departure (DEP) or Arrival (ARR) date (default: DEP)")
    p.add_argument("--format", default="xlsx", choices=["xlsx", "csv"],
                    help="Output file format (default: xlsx)")
    p.add_argument("--out", default=None, help="Output file path (default: auto-generated name)")
    return p.parse_args()


def main():
    args = parse_args()
    start_date = dt.date.fromisoformat(args.start)
    end_date = dt.date.fromisoformat(args.end)
    if end_date < start_date:
        sys.exit("End date must not be before start date.")

    print(f"YML schedule: {args.origin} -> {args.dest}, {start_date} to {end_date}")
    rows = fetch_schedule(
        args.origin, args.dest, start_date, end_date,
        priority=args.priority, date_definition=args.date_definition,
    )

    headers, table = rows_to_table(rows)
    print()
    print_table(headers, table)

    out_path = args.out or (
        f"YML_Schedule_{args.origin}_{args.dest}_{start_date:%Y%m%d}_{end_date:%Y%m%d}.{args.format}"
    )
    if args.format == "xlsx":
        save_xlsx(out_path, headers, table)
    else:
        save_csv(out_path, headers, table)
    print(f"\nSaved {len(table)} sailings to: {out_path}")


if __name__ == "__main__":
    main()
