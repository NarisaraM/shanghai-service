#!/usr/bin/env python3
"""
hal_schedule_export.py

Automates checking the Heung-A Line (HAL) e-Service vessel sailing schedule
(https://ebiz.heungaline.com/Schedule) for a given loading/discharging port
pair across one or more months, and exports each month's result to an
Excel (.xlsx) file — the same "Search -> Excel export" workflow you'd do
by hand on the website, driven by Playwright instead.

Requirements:
    pip install playwright
    playwright install chromium      # downloads the browser Playwright drives

Usage examples:
    # Default: LAEM CHABANG -> SHANGHAI, current month through October 2026
    python3 hal_schedule_export.py

    # Custom lane / months
    python3 hal_schedule_export.py --pol "LAEM CHABANG" --pod "SHANGHAI" \\
        --months 2026-09 2026-10 2026-11

    # Watch it run instead of headless
    python3 hal_schedule_export.py --no-headless

Output:
    One .xlsx file per month is downloaded into --outdir (default: ./downloads),
    named however the site names them (typically something like
    "HAL Schedule (LAEM CHABANG) (SHANGHAI) 2026-09.xlsx").
"""

import argparse
import sys
from datetime import date
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

SCHEDULE_URL = "https://ebiz.heungaline.com/Schedule"
MONTH_ABBR = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}


def month_range(start: str, end: str):
    """Yield 'YYYY-MM' strings from start to end (inclusive), both 'YYYY-MM'."""
    sy, sm = (int(x) for x in start.split("-"))
    ey, em = (int(x) for x in end.split("-"))
    y, m = sy, sm
    while (y, m) <= (ey, em):
        yield f"{y:04d}-{m:02d}"
        m += 1
        if m == 13:
            m = 1
            y += 1


def close_notice_popup(page):
    """The site shows a 'Recent Notice' modal on load; close it if present."""
    for selector in [
        'button[aria-label="Close" i]',
        "button.close",
        'button:has-text("close")',
        '[role="dialog"] .close',
    ]:
        try:
            btn = page.locator(selector).first
            if btn.is_visible(timeout=1500):
                btn.click()
                page.wait_for_timeout(300)
                return
        except PlaywrightTimeoutError:
            continue
        except Exception:
            continue
    # Not fatal if there's nothing to close (notice content varies over time).


def select_port(page, field_id: str, query: str, expect_label_substr: str = None):
    """
    Click a port field (#searchPol or #searchPod), type a search query into it,
    and click the matching option that appears in the dropdown/modal list.
    """
    field = page.locator(f"#{field_id}")
    field.click()
    field.fill("")
    field.type(query, delay=30)
    page.wait_for_timeout(400)

    needle = expect_label_substr or query
    option = page.get_by_text(needle, exact=False).first
    option.wait_for(state="visible", timeout=5000)
    option.click()
    page.wait_for_timeout(300)


def select_month(page, yyyy_mm: str):
    """Open the month picker on #searchMonth and click the target month.

    The picker opens already showing the current year, which covers this
    script's normal use (searching the next few months of the current
    year). If you need a month in a different year, click the picker's
    year prev/next control yourself the first time to confirm its
    selector, or just widen the fallback below.
    """
    year, month = yyyy_mm.split("-")
    month = int(month)
    abbr = MONTH_ABBR[month]

    field = page.locator("#searchMonth")
    field.click()
    page.wait_for_timeout(300)

    picker = page.locator("#ui-monthpicker-div")
    try:
        picker.wait_for(state="visible", timeout=5000)
        if year in picker.inner_text():
            picker.get_by_text(abbr, exact=True).first.click()
            page.wait_for_timeout(300)
    except PlaywrightTimeoutError:
        pass  # fall through to the direct-fill fallback below

    # Sanity check the field actually shows the month we wanted; if the
    # picker click didn't land right (or the year didn't match), set the
    # value directly instead.
    value = field.input_value()
    if value != yyyy_mm:
        field.fill(yyyy_mm)
        page.keyboard.press("Enter")
        page.wait_for_timeout(300)


def run_search(page):
    """Click the visible schedule 'Search' button (the page also has a few
    hidden 'Search' buttons belonging to unrelated sidebar panels — Tracking,
    B/K & B/L, Invoice, Dem/Det — so we pick the first one that's actually
    visible on screen)."""
    candidates = page.locator('button:has-text("Search")')
    visible_btn = None
    for i in range(candidates.count()):
        candidate = candidates.nth(i)
        if candidate.is_visible():
            visible_btn = candidate
            break
    if visible_btn is None:
        raise RuntimeError("Could not find a visible Search button on the page.")
    visible_btn.click()
    # Wait for the results table area to (re)populate.
    page.wait_for_timeout(1500)
    try:
        page.wait_for_selector("table", timeout=10000)
    except PlaywrightTimeoutError:
        pass


def switch_to_list_view(page):
    try:
        page.get_by_text("List", exact=True).first.click()
        page.wait_for_timeout(500)
    except Exception:
        pass  # Excel export works from either view; not fatal.


def export_excel(page, outdir: Path, month_label: str):
    """Click the DataTables Excel export button and save the resulting download."""
    excel_btn = page.locator(".dt-button.buttons-excel").first
    excel_btn.wait_for(state="visible", timeout=10000)
    with page.expect_download() as download_info:
        excel_btn.click()
    download = download_info.value
    suggested = download.suggested_filename or f"HAL_Schedule_{month_label}.xlsx"
    dest = outdir / suggested
    # Avoid clobbering a same-named file from a previous run.
    if dest.exists():
        stem, suffix = dest.stem, dest.suffix
        n = 1
        while dest.exists():
            dest = outdir / f"{stem} ({n}){suffix}"
            n += 1
    download.save_as(str(dest))
    return dest


def main():
    parser = argparse.ArgumentParser(description="Export HAL e-Service vessel schedules to Excel.")
    parser.add_argument("--pol", default="LAEM CHABANG", help="Loading port search text (default: LAEM CHABANG)")
    parser.add_argument("--pod", default="SHANGHAI", help="Discharging port search text (default: SHANGHAI)")
    parser.add_argument(
        "--months", nargs="+", default=None,
        help="Explicit list of months as YYYY-MM (default: current month through 2026-10)",
    )
    parser.add_argument("--outdir", default="downloads", help="Directory to save exported .xlsx files into")
    parser.add_argument("--no-headless", action="store_true", help="Show the browser window while it runs")
    args = parser.parse_args()

    if args.months:
        months = args.months
    else:
        today = date.today()
        months = list(month_range(f"{today.year:04d}-{today.month:02d}", "2026-10"))

    outdir = Path(args.outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)

    print(f"Lane: {args.pol} -> {args.pod}")
    print(f"Months: {', '.join(months)}")
    print(f"Output directory: {outdir}")

    saved_files = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.no_headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto(SCHEDULE_URL, wait_until="domcontentloaded")
        close_notice_popup(page)

        # Outbound (ETD) is the default-selected radio; make sure it's checked.
        try:
            page.locator('input[name="bnd"][value="O"]').check(timeout=3000)
        except Exception:
            pass

        select_port(page, "searchPol", args.pol)
        select_port(page, "searchPod", args.pod)

        for i, ym in enumerate(months):
            print(f"\n[{i + 1}/{len(months)}] Searching {ym} ...")
            select_month(page, ym)
            run_search(page)
            switch_to_list_view(page)
            try:
                dest = export_excel(page, outdir, ym)
                print(f"  Saved: {dest}")
                saved_files.append(dest)
            except PlaywrightTimeoutError:
                print(f"  WARNING: no Excel export button / download detected for {ym} "
                      f"(there may be no sailings that month).")

        browser.close()

    print("\nDone.")
    if saved_files:
        print("Files saved:")
        for f in saved_files:
            print(f"  - {f}")
    else:
        print("No files were saved — check the port names and month range.")

    return 0 if saved_files else 1


if __name__ == "__main__":
    sys.exit(main())
