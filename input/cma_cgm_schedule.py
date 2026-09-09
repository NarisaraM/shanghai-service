"""
CMA CGM Schedule Scraper
========================

Automates the CMA CGM "Get a schedule" page
(https://www.cma-cgm.com/ebusiness/schedules/routing-finder) with Playwright:
searches sailings from Laem Chabang (THLCH) to Shanghai (CNSHA), covering
today's date through the end of a target month (October by default), and
exports the results to an Excel file.

SETUP (run once on your own computer):
    pip install playwright pandas openpyxl
    playwright install chromium

RUN:
    python cma_cgm_schedule.py

NOTES ON THE WEBSITE:
- The site runs Google/Akamai-style bot protection (DataDome). Running with
  headless=False (the default here) opens a real, visible Chrome window,
  which is far less likely to be challenged than a headless run. If a
  CAPTCHA/verification screen appears, solve it by hand in that window --
  the script will keep waiting (up to WAIT_FOR_RESULTS_MS) and then continue.
- The "Period (weeks)" control on the site caps out at 10 weeks per search.
  The script automatically computes how many weeks are needed to cover
  today -> the end of END_MONTH, and caps the request at 10. If your date
  range needs more than 10 weeks, rerun with a later START date, or extend
  the script to loop multiple searches (see `weeks_needed`).
"""

import sys
from datetime import date, datetime
from calendar import monthrange

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

# ----------------------------------------------------------------------------
# Configuration -- edit these if you want a different lane / month
# ----------------------------------------------------------------------------
POL_SEARCH_TEXT = "Laem Chabang"
POL_CODE = "THLCH"
POD_SEARCH_TEXT = "Shanghai"
POD_CODE = "CNSHA"

END_MONTH = 10          # October
END_YEAR = None         # None => current year

URL = "https://www.cma-cgm.com/ebusiness/schedules/routing-finder"
WAIT_FOR_RESULTS_MS = 45_000  # generous, in case a bot-check screen appears
HEADLESS = False  # keep a visible browser window -- see notes above


def weeks_needed(start: date, end: date) -> int:
    """How many 1-week steps (1-10, the site's own limits) are needed to cover start..end."""
    days = (end - start).days
    weeks = -(-days // 7)  # ceil division
    return max(1, min(10, weeks))


def set_autocomplete(page, input_id: str, search_text: str, code: str):
    """Click an Origin/Destination field, type a search term and pick the PORT
    suggestion whose code matches `code` (falls back to the first suggestion)."""
    field = page.locator(f"#{input_id}")
    field.click()
    page.keyboard.press("Control+A")
    field.type(search_text, delay=40)
    page.wait_for_timeout(900)

    suggestions = page.locator("li.place-suggestion")
    try:
        suggestions.first.wait_for(state="visible", timeout=8000)
    except PWTimeoutError:
        raise RuntimeError(
            f"No autocomplete suggestions appeared for '{search_text}'. "
            "The site's markup may have changed, or a bot-check is blocking input."
        )

    count = suggestions.count()
    for i in range(count):
        item = suggestions.nth(i)
        text = item.inner_text().upper()
        if code.upper() in text and "PORT" in text:
            item.click()
            return
    # Fallback: just take the first suggestion offered.
    suggestions.first.click()


def set_search_on_departure(page):
    """Make sure the 'Search on' selector is set to Departure (index 0)."""
    page.evaluate(
        """
        () => {
            try {
                if (window.jQuery) {
                    const dd = jQuery('#IsDeparture').data('kendoDropDownList');
                    if (dd) { dd.select(0); dd.trigger('change'); }
                }
            } catch (e) {}
        }
        """
    )


def set_period_weeks(page, weeks: int):
    """Use the page's own increasePeriod()/decreasePeriod() JS helpers to set
    the 'Period (weeks)' field, since it's driven by onclick handlers rather
    than a plain input."""
    page.evaluate(
        """
        (n) => {
            const field = document.getElementById('field');
            field.value = 1;
            for (let i = 1; i < n; i++) { increasePeriod(); }
        }
        """,
        weeks,
    )
    page.wait_for_timeout(300)


def extract_results(page):
    return page.evaluate(
        """
        () => {
            function textOf(el) { return el ? el.textContent.trim() : ''; }
            function placeName(container) {
                if (!container) return '';
                const clone = container.cloneNode(true);
                const badge = clone.querySelector('.capsule');
                if (badge) badge.remove();
                return clone.textContent.replace(/\\s+/g, ' ').trim();
            }
            function extractCard(card) {
                const dep = card.querySelector('.DepartureDatesCls .date');
                const arr = card.querySelector('.ArrivalDatesCls .date');
                const transitEl = card.querySelector('.Transitcls');
                const routeType = card.querySelector('.transit.direct, .transit.transship, .transit.transshipment');
                const polNameEl = card.querySelector('.DepartureDatesCls .capsule-container');
                const podNameEl = card.querySelector('.ArrivalDatesCls .capsule-container');
                const polTermEl = card.querySelector('.DepartureDatesCls .state-row a:last-child');
                const podTermEl = card.querySelector('.ArrivalDatesCls .state-row a:last-child');

                let vessel = '', voyage = '', cutoff = '';
                card.querySelectorAll('.more-infos.vessel dt').forEach(dt => {
                    const label = dt.textContent.trim();
                    const dd = dt.nextElementSibling;
                    const val = dd ? dd.textContent.trim() : '';
                    if (label.includes('Main vessel')) vessel = val;
                    else if (label.includes('Voyage')) voyage = val;
                    else if (label.includes('Cut-off')) cutoff = val;
                });

                return {
                    departure_date: textOf(dep),
                    pol: placeName(polNameEl),
                    pol_terminal: textOf(polTermEl),
                    port_cutoff: cutoff,
                    vessel: vessel,
                    voyage_ref: voyage,
                    arrival_date: textOf(arr),
                    pod: placeName(podNameEl),
                    pod_terminal: textOf(podTermEl),
                    transit_time: textOf(transitEl).replace(/\\s+/g, ' '),
                    routing: textOf(routeType),
                };
            }
            return Array.from(document.querySelectorAll('li.cardelem')).map(extractCard);
        }
        """
    )


def parse_site_date(s: str):
    """'Monday, 14-SEP-2026' -> date(2026, 9, 14)"""
    try:
        return datetime.strptime(s.split(", ")[-1].strip(), "%d-%b-%Y").date()
    except Exception:
        return None


def main():
    today = date.today()
    end_year = END_YEAR or today.year
    last_day = monthrange(end_year, END_MONTH)[1]
    target_end = date(end_year, END_MONTH, last_day)

    if target_end < today:
        print(f"Target end date {target_end.isoformat()} is before today; nothing to search.")
        sys.exit(1)

    period_weeks = weeks_needed(today, target_end)
    if (target_end - today).days > period_weeks * 7:
        print(
            f"Note: the site limits searches to 10 weeks. Results will run through "
            f"roughly {period_weeks} week(s) from today; some late-{target_end.strftime('%B')} "
            f"sailings might fall just outside this single search."
        )

    print(
        f"Searching {POL_SEARCH_TEXT} ({POL_CODE}) -> {POD_SEARCH_TEXT} ({POD_CODE}), "
        f"{today.isoformat()} to {target_end.isoformat()} ({period_weeks} week period)..."
    )

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=80)
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
            ),
            locale="en-US",
            viewport={"width": 1366, "height": 900},
        )
        page = context.new_page()
        page.goto(URL, wait_until="domcontentloaded", timeout=60000)

        try:
            page.wait_for_selector("#AutoCompletePOL", timeout=WAIT_FOR_RESULTS_MS)
        except PWTimeoutError:
            print(
                "The search form didn't appear in time. This usually means a "
                "bot-check/CAPTCHA is showing. Solve it in the opened browser "
                "window, then rerun the script."
            )
            browser.close()
            sys.exit(1)

        page.wait_for_timeout(1500)  # let widgets finish initializing

        set_autocomplete(page, "AutoCompletePOL", POL_SEARCH_TEXT, POL_CODE)
        page.wait_for_timeout(400)
        set_autocomplete(page, "AutoCompletePOD", POD_SEARCH_TEXT, POD_CODE)
        page.wait_for_timeout(400)

        set_search_on_departure(page)
        set_period_weeks(page, period_weeks)

        page.click("#searchSchedules")
        try:
            page.wait_for_selector("li.cardelem", timeout=WAIT_FOR_RESULTS_MS)
        except PWTimeoutError:
            print("No results appeared after searching -- check the browser window for errors.")
            browser.close()
            sys.exit(1)
        page.wait_for_timeout(1500)

        results = extract_results(page)
        browser.close()

    if not results:
        print("No sailings found for this lane/period.")
        sys.exit(0)

    for r in results:
        r["departure_parsed"] = parse_site_date(r["departure_date"])

    # Keep only sailings that depart on/before the requested end date.
    results = [r for r in results if r["departure_parsed"] and r["departure_parsed"] <= target_end]
    for r in results:
        del r["departure_parsed"]

    try:
        import pandas as pd
    except ImportError:
        print("pandas is required to export to Excel. Install it with: pip install pandas openpyxl")
        print("Raw results:")
        for r in results:
            print(r)
        sys.exit(1)

    df = pd.DataFrame(results, columns=[
        "departure_date", "pol", "pol_terminal", "port_cutoff",
        "vessel", "voyage_ref", "arrival_date", "pod", "pod_terminal",
        "transit_time", "routing",
    ])

    out_name = (
        f"CMA_CGM_Schedule_{POL_CODE}_{POD_CODE}_"
        f"{today.strftime('%Y%m%d')}_to_{target_end.strftime('%Y%m%d')}.xlsx"
    )
    df.to_excel(out_name, index=False)
    print(f"Saved {len(df)} sailing(s) to {out_name}")
    print(df.to_string(index=False))


if __name__ == "__main__":
    main()
