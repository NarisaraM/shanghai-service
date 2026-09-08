#!/usr/bin/env python3
"""
ZIM Point-to-Point Schedule Scraper
====================================

Automates the ZIM ("The Z Factor") Point-to-Point schedule search at
https://www.zim.com/schedules/point-to-point and exports the sailing
schedule between two ports to an Excel file.

Default lane: Laem Chabang, Thailand (THLEM) -> Shanghai, China (CNSNH)
Default date range: today -> 31 December of the current year

IMPORTANT - publishing horizon: ZIM (like most carriers) only publishes
vessel schedules a limited number of weeks ahead of the REAL current date
(observed to be ~12 weeks / the site's own "weeks ahead" maximum). Asking
for a date range that runs past that horizon (e.g. "today through end of
year" when today is only in September) will only return sailings up to
where ZIM's data actually stops -- the site itself returns an "Oops...
currently unavailable" response beyond that point, no matter which start
date is searched. This script detects that and stops automatically rather
than treating it as an error; simply re-run it every few weeks to keep
extending your saved schedule as ZIM publishes further sailings.

Why a browser instead of plain requests?
-----------------------------------------
The ZIM site renders schedule results client-side inside a Web Component
(<schedule-by-route-v1>) using Shadow DOM, after fetching data through
internal calls that are not simple/stable public APIs. The most robust way
to reproduce "search a lane -> read the results" is to drive a real
(headless) browser with Playwright and read the rendered Shadow DOM,
rather than reverse-engineer a private API that can change without notice.

Usage
-----
    python3 zim_schedule_scraper.py
    python3 zim_schedule_scraper.py --origin THLEM --destination CNSNH
    python3 zim_schedule_scraper.py --start 2026-09-08 --end 2026-12-31
    python3 zim_schedule_scraper.py --output my_schedule.xlsx --headed

Requirements
------------
    pip install playwright pandas openpyxl
    playwright install chromium   # (already preinstalled in some environments)
"""

import argparse
import datetime as dt
import re
import sys
import time
from typing import Optional

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

BASE_URL = "https://www.zim.com/schedules/point-to-point"

# Port codes as used by ZIM's schedule tool (UN/LOCODE-ish + terminal suffix).
# These were captured from the site's own search for "Laem Chabang" / "Shanghai".
DEFAULT_ORIGIN_CODE = "THLEM;10"      # Laem Chabang, Thailand
DEFAULT_DEST_CODE = "CNSNH;10"        # Shanghai (SH), China

MAX_WEEKS_AHEAD = 12  # the site's own dropdown tops out at 12 weeks per search

# JavaScript executed in the page to pull structured rows out of the
# <schedule-by-route-v1> web component's Shadow DOM. Class names are CSS
# Modules (e.g. "_cardContainer_1cul9_390"), so we match on the stable
# human-readable fragment with a wildcard `[class*="..."]` selector, which
# survives the random per-build hash suffix.
EXTRACT_JS = """
() => {
    const host = document.querySelector('schedule-by-route-v1');
    if (!host || !host.shadowRoot) return {rows: [], noData: false, raw: null};
    const root = host.shadowRoot;
    const bodyText = root.textContent || '';
    const noData = /No data found/i.test(bodyText);
    // ZIM only publishes sailings up to a limited horizon from the REAL
    // current date (observed: ~12 weeks). Searching further ahead than
    // that renders an "Oops... currently unavailable" panel instead of
    // results or a plain "no data" message.
    const unavailable = /Oops|currently unavailable/i.test(bodyText);

    const cards = root.querySelectorAll('[class*="cardContainer"]');
    const rows = [];

    cards.forEach(card => {
        const spans = card.querySelectorAll('[class*="wrapper"] span');
        const distance = spans[0] ? spans[0].textContent.trim() : null;
        const co2 = spans[1] ? spans[1].textContent.trim() : null;

        const depBlock = card.querySelector('[class*="departure"]');
        const arrBlock = card.querySelector('[class*="arrival"]');
        const depDate = depBlock
            ? (depBlock.querySelector('[class*="timeValue"]')?.textContent || '').trim()
            : null;
        const arrDate = arrBlock
            ? (arrBlock.querySelector('[class*="timeValue"]')?.textContent || '').trim()
            : null;

        const indicator = card.querySelector('[class*="scheduleIndicator"]');
        let transitType = null;
        let transitDays = null;
        if (indicator) {
            const divs = indicator.querySelectorAll('div');
            if (divs.length >= 1) transitType = divs[0].textContent.trim();
            if (divs.length >= 1) transitDays = divs[divs.length - 1].textContent.trim();
        }

        const vesselTitle = card.querySelector('[class*="vesselTitle"]');
        const vesselText = vesselTitle ? vesselTitle.textContent.trim() : null;

        rows.push({
            distance: distance,
            co2_wtw: co2,
            departure: depDate,
            arrival: arrDate,
            transit_type: transitType,
            transit_days: transitDays,
            vessel_voyage: vesselText,
        });
    });

    return {rows: rows, noData: noData, unavailable: unavailable, raw: bodyText.slice(0, 200)};
}
"""


def build_url(origin_code: str, dest_code: str, from_date: dt.date, weeks_ahead: int) -> str:
    return (
        f"{BASE_URL}?portcode={origin_code}&portdestinationcode={dest_code}"
        f"&direction=true&fromdate={from_date.strftime('%d-%b-%Y')}"
        f"&weeksahead={weeks_ahead}&cargotype=true&emissionstype=true&toggleemission=true"
    )


def fetch_window(page, origin_code: str, dest_code: str, from_date: dt.date,
                  weeks_ahead: int = MAX_WEEKS_AHEAD, retries: int = 2):
    url = build_url(origin_code, dest_code, from_date, weeks_ahead)
    last_err = None
    for attempt in range(retries + 1):
        try:
            page.goto(url, wait_until="networkidle", timeout=45000)
            # Wait until the web component has attached its shadow root.
            page.wait_for_function(
                "() => { const h = document.querySelector('schedule-by-route-v1'); "
                "return !!(h && h.shadowRoot); }",
                timeout=20000,
            )
            # Give the async data fetch inside the component a moment to render.
            page.wait_for_timeout(2500)
            data = page.evaluate(EXTRACT_JS)
            # If nothing rendered yet (slow network call), retry a bit before giving up.
            if not data["rows"] and not data["noData"] and not data["unavailable"]:
                page.wait_for_timeout(2500)
                data = page.evaluate(EXTRACT_JS)
            return data
        except PlaywrightTimeoutError as exc:
            last_err = exc
            print(f"  ! timeout on {from_date} (attempt {attempt + 1}/{retries + 1}), retrying...",
                  file=sys.stderr)
    print(f"  ! giving up on window starting {from_date}: {last_err}", file=sys.stderr)
    return {"rows": [], "noData": False, "unavailable": False, "raw": None}


def parse_date(s: Optional[str]) -> Optional[dt.date]:
    if not s:
        return None
    s = s.strip()
    try:
        return dt.datetime.strptime(s, "%d-%b-%Y").date()
    except ValueError:
        return None


VESSEL_RE = re.compile(r"^(?P<name>.+?)\s*\((?P<code>[^)]+)\)\s*/\s*(?P<voyage>[^/]+)\s*/\s*(?P<flag>.+)$")


def split_vessel_voyage(text: Optional[str]):
    if not text:
        return {"vessel_name": None, "vessel_code": None, "voyage": None, "flag": None}
    m = VESSEL_RE.match(text)
    if not m:
        return {"vessel_name": text, "vessel_code": None, "voyage": None, "flag": None}
    d = m.groupdict()
    return {
        "vessel_name": d["name"].strip(),
        "vessel_code": d["code"].strip(),
        "voyage": d["voyage"].strip(),
        "flag": d["flag"].strip(),
    }


def scrape_schedule(origin_code: str, dest_code: str, start_date: dt.date, end_date: dt.date,
                     headless: bool = True, verbose: bool = True):
    """Repeats the site's own search (12-week windows) until the whole
    [start_date, end_date] range is covered, de-duplicating sailings that
    show up in more than one overlapping window."""
    all_rows = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page()
        cur = start_date
        window_num = 0
        while cur <= end_date:
            window_num += 1
            if verbose:
                print(f"[{window_num}] searching {origin_code} -> {dest_code} "
                      f"from {cur.strftime('%d-%b-%Y')} (+{MAX_WEEKS_AHEAD} weeks)...")
            data = fetch_window(page, origin_code, dest_code, cur, MAX_WEEKS_AHEAD)
            if data.get("unavailable"):
                if verbose:
                    print(f"    ZIM has not published sailings this far ahead yet "
                          f"(search starting {cur.strftime('%d-%b-%Y')} came back "
                          f"'currently unavailable'). Stopping here -- re-run this "
                          f"script closer to that date to pick up the rest.")
                break
            if data.get("noData") and verbose:
                print("    (no sailings found in this window)")
            for r in data["rows"]:
                key = (r.get("vessel_voyage"), r.get("departure"))
                all_rows[key] = r
            cur = cur + dt.timedelta(weeks=MAX_WEEKS_AHEAD)
        browser.close()
    return list(all_rows.values())


def to_records(raw_rows, start_date: dt.date, end_date: dt.date):
    records = []
    for r in raw_rows:
        dep = parse_date(r.get("departure"))
        arr = parse_date(r.get("arrival"))
        if dep is None:
            continue
        if dep < start_date or dep > end_date:
            continue
        vv = split_vessel_voyage(r.get("vessel_voyage"))
        records.append({
            "Departure": dep,
            "Arrival": arr,
            "Transit Type": r.get("transit_type"),
            "Transit Time (Days)": r.get("transit_days"),
            "Vessel Name": vv["vessel_name"],
            "Vessel Code": vv["vessel_code"],
            "Voyage": vv["voyage"],
            "Total Distance": r.get("distance"),
            "Total CO2 WTW": r.get("co2_wtw"),
        })
    records.sort(key=lambda x: (x["Departure"], x["Vessel Name"] or ""))
    return records


def main():
    today = dt.date.today()
    default_end = dt.date(today.year, 12, 31)

    parser = argparse.ArgumentParser(
        description="Scrape the ZIM Point-to-Point schedule and export it to Excel."
    )
    parser.add_argument("--origin", default=DEFAULT_ORIGIN_CODE,
                         help=f"Origin port code (default: {DEFAULT_ORIGIN_CODE} = Laem Chabang, TH)")
    parser.add_argument("--destination", default=DEFAULT_DEST_CODE,
                         help=f"Destination port code (default: {DEFAULT_DEST_CODE} = Shanghai, CN)")
    parser.add_argument("--start", default=today.strftime("%Y-%m-%d"),
                         help="Start date YYYY-MM-DD (default: today)")
    parser.add_argument("--end", default=default_end.strftime("%Y-%m-%d"),
                         help="End date YYYY-MM-DD (default: 31 Dec of current year)")
    parser.add_argument("--output", default="zim_schedule.xlsx",
                         help="Output .xlsx file path")
    parser.add_argument("--headed", action="store_true",
                         help="Run the browser headed (useful for debugging)")
    args = parser.parse_args()

    start_date = dt.datetime.strptime(args.start, "%Y-%m-%d").date()
    end_date = dt.datetime.strptime(args.end, "%Y-%m-%d").date()
    if end_date < start_date:
        parser.error("--end must be on or after --start")

    print(f"ZIM schedule: {args.origin} -> {args.destination}, "
          f"{start_date.isoformat()} to {end_date.isoformat()}")

    raw_rows = scrape_schedule(args.origin, args.destination, start_date, end_date,
                                headless=not args.headed)
    records = to_records(raw_rows, start_date, end_date)

    if not records:
        print("No sailings found for this lane/date range.")
        sys.exit(0)

    import pandas as pd
    df = pd.DataFrame(records)
    df["Departure"] = pd.to_datetime(df["Departure"])
    df["Arrival"] = pd.to_datetime(df["Arrival"])

    with pd.ExcelWriter(args.output, engine="openpyxl", datetime_format="dd-mmm-yyyy") as writer:
        df.to_excel(writer, index=False, sheet_name="Schedule")
        ws = writer.sheets["Schedule"]
        widths = [12, 12, 12, 18, 22, 12, 10, 14, 16]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    print(f"Saved {len(records)} sailings to {args.output}")
    last_departure = records[-1]["Departure"]
    if last_departure < end_date:
        print(f"Note: ZIM has only published sailings up to {last_departure.isoformat()} so far "
              f"(you asked for up to {end_date.isoformat()}). Re-run this script closer to "
              f"those later dates to fill in the rest of the year as ZIM publishes it.")


if __name__ == "__main__":
    main()
