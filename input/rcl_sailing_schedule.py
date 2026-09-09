#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rcl_sailing_schedule.py
========================

ดึงตารางเรือ (Sailing Schedule) จากเว็บไซต์ RCL Group
สำหรับเส้นทาง Laem Chabang (THLCH) -> Shanghai (CNSHA)
ตั้งแต่ "เดือนปัจจุบัน" ถึง "เดือนตุลาคม" แล้ว export เป็นไฟล์ Excel

เว็บไซต์ต้นทาง (หน้าเว็บที่สคริปต์นี้ทำงานด้วย):
    https://eservice.rclgroup.com/SailingSchedule/

วิธีการทำงาน
------------
หน้าเว็บนี้เป็น Single Page App ที่มี Cloudflare Turnstile (ระบบตรวจสอบว่าไม่ใช่บอท)
ฝังอยู่ ดังนั้นสคริปต์นี้จึงไม่ได้ยิง HTTP request ตรง ๆ ด้วย `requests`
(ซึ่งจะโดน Cloudflare บล็อกเกือบทันที) แต่ใช้ Playwright เปิดเบราว์เซอร์ Chromium
จริง ๆ ขึ้นมา แล้วทำตามขั้นตอนเดียวกับที่คนกดเอง คือ:

    1. เปิดหน้า Sailing Schedule
    2. พิมพ์ต้นทาง "LCH" แล้วเลือกรายการ "THLCH-LAEM CHABANG" จาก autocomplete
    3. พิมพ์ปลายทาง "SHA" แล้วเลือกรายการ "CNSHA-SHANGHAI" จาก autocomplete
    4. กด "Search"
    5. ดักจับ (intercept) คำตอบ JSON ที่หน้าเว็บเรียกจาก API ภายใน:
           POST https://eservice.rclgroup.com/SailingSchedule/sailingApi/sailingSchedule
       ซึ่งเป็น API เดียวกับที่ตัวหน้าเว็บใช้วาดปฏิทินเที่ยวเรือ (ตรวจสอบแล้วว่าคำตอบ
       จาก API ครั้งเดียวครอบคลุมล่วงหน้าหลายเดือน เช่น ก.ย./ต.ค./พ.ย. ในคำตอบเดียว
       จึงไม่จำเป็นต้องกด "เดือนถัดไป" ทีละเดือน)
    6. กรองเฉพาะเที่ยวเรือที่อยู่ในช่วงวันที่ที่ต้องการ (เดือนปัจจุบัน -> 31 ตุลาคม)
    7. บันทึกผลลัพธ์เป็นไฟล์ .xlsx

ข้อควรทราบเกี่ยวกับ Cloudflare
------------------------------
ระบบตรวจสอบของ Cloudflare (Turnstile) ที่หน้านี้ปกติจะ "ผ่านอัตโนมัติ" แบบไม่ต้อง
คลิกอะไร (invisible challenge) เมื่อเปิดด้วยเบราว์เซอร์จริง แต่ถ้าเปิดด้วยโหมด
headless หรือเปิดถี่เกินไป Cloudflare อาจขึ้นภาพให้ติ๊กยืนยันตัวตน
เพราะฉะนั้น:
    - แนะนำให้รันครั้งแรกแบบ "เห็นหน้าต่างเบราว์เซอร์" (headless=False, ค่าเริ่มต้น)
      ถ้ามีภาพให้ติ๊กยืนยันขึ้นมา ให้คลิกยืนยันเองหนึ่งครั้ง
    - สคริปต์จะบันทึกสถานะเซสชัน/คุกกี้ไว้ในไฟล์ rcl_storage_state.json
      รันครั้งต่อไปจะได้ไม่ต้องผ่านการตรวจสอบซ้ำบ่อย ๆ

การติดตั้ง (ทำครั้งเดียว)
--------------------------
    pip install playwright pandas openpyxl
    playwright install chromium

วิธีใช้งาน
----------
    # ค่าเริ่มต้น: Laem Chabang -> Shanghai, เดือนปัจจุบัน ถึง 31 ตุลาคมของปีปัจจุบัน
    python3 rcl_sailing_schedule.py

    # กำหนดต้นทาง/ปลายทาง หรือช่วงวันเอง
    python3 rcl_sailing_schedule.py --pol LCH --pod SHA --end-month 10

    # รันแบบไม่เปิดหน้าต่างเบราว์เซอร์ (อาจโดน Cloudflare ถามยืนยันตัวตนบ่อยกว่า)
    python3 rcl_sailing_schedule.py --headless
"""

import argparse
import calendar
import sys
from datetime import date, datetime
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except ImportError:
    sys.exit(
        "ไม่พบไลบรารี playwright กรุณาติดตั้งก่อน:\n"
        "    pip install playwright pandas openpyxl\n"
        "    playwright install chromium"
    )

try:
    import pandas as pd
except ImportError:
    sys.exit("ไม่พบไลบรารี pandas กรุณาติดตั้งก่อน:\n    pip install pandas openpyxl")


BASE_URL = "https://eservice.rclgroup.com/SailingSchedule/"
API_URL_FRAGMENT = "sailingApi/sailingSchedule"
STORAGE_STATE_FILE = Path(__file__).with_name("rcl_storage_state.json")


def select_from_autocomplete(page, input_selector: str, list_selector: str, query: str, label: str):
    """
    พิมพ์คำค้นหาลงในช่อง input (Origin หรือ Destination) แล้วเลือกรายการแรก
    ที่ตรงกับคำค้นจาก dropdown autocomplete ที่ขึ้นมา

    input_selector: CSS selector ของช่อง input เช่น "#originInput"
    list_selector:  CSS selector ของ <ul> ที่มีรายการแนะนำ เช่น "ul.list1"
    query:          ข้อความที่จะพิมพ์ค้นหา เช่น "LCH" หรือ "SHA"
    """
    box = page.locator(input_selector)
    box.click()
    box.fill("")
    box.type(query, delay=100)  # พิมพ์ทีละตัวอักษรเพื่อ trigger event ของหน้าเว็บ

    options = page.locator(f"{list_selector} li")
    try:
        options.first.wait_for(state="visible", timeout=8000)
    except PlaywrightTimeoutError:
        raise RuntimeError(
            f"ไม่พบรายการแนะนำสำหรับ {label} ('{query}') "
            f"ลองตรวจสอบว่าพิมพ์รหัสท่าเรือ/ชื่อท่าเรือถูกต้องหรือไม่"
        )

    count = options.count()
    texts = [options.nth(i).inner_text().strip() for i in range(count)]

    # หารายการที่ข้อความมี query อยู่ (ไม่สนตัวพิมพ์เล็ก/ใหญ่)
    match_index = next(
        (i for i, t in enumerate(texts) if query.strip().upper() in t.upper()),
        0,  # ถ้าไม่เจอเป๊ะ ๆ ให้เลือกรายการแรกที่ขึ้นมา
    )
    chosen_text = texts[match_index]
    options.nth(match_index).click()
    print(f"  {label}: เลือก '{chosen_text}' (จากตัวเลือกทั้งหมด {count} รายการ: {texts})")
    return chosen_text


def fetch_schedule(pol_query: str, pod_query: str, start_date: date, headless: bool) -> dict:
    """
    เปิดเบราว์เซอร์ กรอกฟอร์ม Sailing Schedule แล้วดักจับ JSON ที่ได้จาก API
    คืนค่าเป็น dict ของ JSON response ดิบจาก RCL
    """
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            args=["--disable-blink-features=AutomationControlled"],
        )

        context_kwargs = {"locale": "en-US"}
        if STORAGE_STATE_FILE.exists():
            context_kwargs["storage_state"] = str(STORAGE_STATE_FILE)

        context = browser.new_context(**context_kwargs)
        page = context.new_page()

        print(f"เปิดหน้า {BASE_URL} ...")
        page.goto(BASE_URL, wait_until="domcontentloaded", timeout=45000)

        # ให้เวลา Cloudflare Turnstile ตรวจสอบอัตโนมัติสักครู่
        page.wait_for_timeout(2500)

        print("กรอกท่าเรือต้นทาง/ปลายทาง ...")
        select_from_autocomplete(page, "#originInput", "ul.list1", pol_query, "ต้นทาง (Origin)")
        select_from_autocomplete(page, "#destiInput", "ul.list2", pod_query, "ปลายทาง (Destination)")

        # ตั้งวันที่เริ่มค้นหา (ช่อง <input type="date"> ต้องการรูปแบบ YYYY-MM-DD)
        page.fill("#date-input", start_date.strftime("%Y-%m-%d"))

        print("กด Search และรอผลลัพธ์จาก API ...")
        try:
            with page.expect_response(
                lambda r: API_URL_FRAGMENT in r.url and r.request.method == "POST",
                timeout=25000,
            ) as response_info:
                page.click("#sailSearch")
            response = response_info.value
        except PlaywrightTimeoutError:
            raise RuntimeError(
                "ไม่ได้รับคำตอบจากระบบค้นหาภายในเวลาที่กำหนด "
                "อาจมีหน้าต่างยืนยันตัวตน (Cloudflare) ค้างอยู่ "
                "กรุณารันสคริปต์แบบไม่ใส่ --headless แล้วยืนยันตัวตนด้วยตนเองหนึ่งครั้ง"
            )

        if response.status != 200:
            raise RuntimeError(f"API ตอบกลับด้วยสถานะ {response.status}: {response.text()[:500]}")

        data = response.json()

        # บันทึกสถานะเซสชัน/คุกกี้ไว้ใช้ครั้งถัดไป จะได้ไม่ต้องผ่าน Cloudflare ซ้ำบ่อย ๆ
        context.storage_state(path=str(STORAGE_STATE_FILE))

        browser.close()
        return data


def parse_thai_or_iso_date(value: str) -> date:
    """แปลงวันที่รูปแบบ DD/MM/YYYY (ที่ RCL ใช้) เป็น datetime.date"""
    return datetime.strptime(value, "%d/%m/%Y").date()


def flatten_schedule(raw: dict, start_date: date, end_date: date) -> list:
    """
    แปลงโครงสร้าง JSON ที่ได้จาก RCL ให้เป็น list of dict แถวเดียวต่อหนึ่งเที่ยวเรือ
    พร้อมกรองเฉพาะเที่ยวที่อยู่ในช่วง [start_date, end_date]

    โครงสร้างที่พบจากการตรวจสอบ API จริง (อาจมีบางฟิลด์เพิ่ม/ขาดได้ในอนาคต
    หากทาง RCL ปรับปรุงเว็บไซต์):
        {
          "pol": "THLCH",
          "pod": "CNSHA|CNMCT",
          "calendar": [
            {
              "loadingPortArrival": "08/09/2026",
              "sailingHeader": [
                {
                  "voyNo": "...", "vesselNameNo": "...",
                  "pol": "THLCH", "pod": "CNSHA",
                  "loadingPortArrival": "...", "destinationArrival": "...",
                  "sailingDetail": [ { ... รายละเอียดเที่ยวเรือ/ต่อเรือ ... } ]
                }
              ]
            }, ...
          ]
        }
    """
    rows = []
    for day_entry in raw.get("calendar", []):
        for header in day_entry.get("sailingHeader", []):
            details = header.get("sailingDetail") or [header]
            for detail in details:
                loading_arrival = detail.get("loadingPortArrival") or header.get(
                    "loadingPortArrival"
                ) or day_entry.get("loadingPortArrival")

                try:
                    sail_date = parse_thai_or_iso_date(loading_arrival)
                except (TypeError, ValueError):
                    sail_date = None

                if sail_date is not None and not (start_date <= sail_date <= end_date):
                    continue

                rows.append(
                    {
                        "Sailing Date (Loading Port Arrival)": loading_arrival,
                        "Loading Port Departure": detail.get("loadingPortDeparture"),
                        "Vessel": detail.get("vesselNameNo") or header.get("vesselNameNo"),
                        "Voyage No": detail.get("voyNo") or header.get("voyNo"),
                        "POL": detail.get("loadPortArrival") or header.get("pol"),
                        "POD": detail.get("loadPortDeparture") or header.get("pod"),
                        "Destination Arrival": detail.get("destinationArrival")
                        or header.get("destinationArrival"),
                        "Transit Time": detail.get("transitTime"),
                        "Vessel Flag": detail.get("vesselFlag"),
                    }
                )
    return rows


def main():
    today = date.today()

    parser = argparse.ArgumentParser(
        description="ดึงตารางเรือ RCL (Sailing Schedule) และ export เป็น Excel"
    )
    parser.add_argument("--pol", default="LCH", help="ท่าเรือต้นทาง เช่น LCH (Laem Chabang) [ค่าเริ่มต้น: LCH]")
    parser.add_argument("--pod", default="SHA", help="ท่าเรือปลายทาง เช่น SHA (Shanghai) [ค่าเริ่มต้น: SHA]")
    parser.add_argument(
        "--start-date",
        default=date(today.year, today.month, 1).strftime("%Y-%m-%d"),
        help=(
            "วันที่เริ่มค้นหา รูปแบบ YYYY-MM-DD "
            "[ค่าเริ่มต้น: วันที่ 1 ของเดือนปัจจุบัน เพื่อให้ครอบคลุมทั้งเดือน]"
        ),
    )
    parser.add_argument(
        "--end-month",
        type=int,
        default=10,
        help="เดือนสุดท้ายที่ต้องการ (1-12) [ค่าเริ่มต้น: 10 = ตุลาคม]",
    )
    parser.add_argument(
        "--end-year",
        type=int,
        default=None,
        help="ปีของเดือนสุดท้าย (ค.ศ.) ถ้าไม่ระบุจะคำนวณให้อัตโนมัติ",
    )
    parser.add_argument("--output", default=None, help="ชื่อไฟล์ Excel ผลลัพธ์ (.xlsx)")
    parser.add_argument(
        "--headless",
        action="store_true",
        help="รันแบบไม่เปิดหน้าต่างเบราว์เซอร์ (ค่าเริ่มต้นคือเปิดหน้าต่างให้เห็น)",
    )
    args = parser.parse_args()

    start_date = datetime.strptime(args.start_date, "%Y-%m-%d").date()

    end_year = args.end_year
    if end_year is None:
        # ถ้าเดือนปัจจุบันเลยเดือนสิ้นสุดที่ต้องการไปแล้ว ให้เข้าใจว่าหมายถึงปีถัดไป
        end_year = start_date.year if start_date.month <= args.end_month else start_date.year + 1
    last_day = calendar.monthrange(end_year, args.end_month)[1]
    end_date = date(end_year, args.end_month, last_day)

    if end_date < start_date:
        sys.exit("ช่วงวันที่ไม่ถูกต้อง: เดือนสิ้นสุดอยู่ก่อนวันที่เริ่มต้น")

    output_path = args.output or (
        f"RCL_Schedule_{args.pol.upper()}_{args.pod.upper()}_"
        f"{start_date:%Y%m%d}_to_{end_date:%Y%m%d}.xlsx"
    )

    print(f"ค้นหาตารางเรือ RCL: {args.pol.upper()} -> {args.pod.upper()}")
    print(f"ช่วงวันที่: {start_date:%d/%m/%Y} ถึง {end_date:%d/%m/%Y}")
    print("-" * 60)

    raw = fetch_schedule(args.pol, args.pod, start_date, headless=args.headless)
    rows = flatten_schedule(raw, start_date, end_date)

    if not rows:
        print(
            "ไม่พบเที่ยวเรือในช่วงวันที่ที่กำหนด "
            "(อาจเป็นเพราะ API เปลี่ยนโครงสร้าง หรือไม่มีเที่ยวเรือจริง ๆ ในช่วงนี้)"
        )
        raw_dump_path = Path(output_path).with_suffix(".raw.json")
        raw_dump_path.write_text(
            __import__("json").dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"บันทึกข้อมูลดิบที่ได้จาก API ไว้ที่: {raw_dump_path} (ใช้ตรวจสอบโครงสร้างเพิ่มเติมได้)")
        return

    df = pd.DataFrame(rows)
    # เรียงตามวันที่ออกเรือ
    df["_sort_key"] = df["Sailing Date (Loading Port Arrival)"].apply(
        lambda v: parse_thai_or_iso_date(v) if v else date.max
    )
    df = df.sort_values("_sort_key").drop(columns="_sort_key")

    df.to_excel(output_path, index=False)
    print("-" * 60)
    print(f"พบเที่ยวเรือทั้งหมด {len(df)} เที่ยว")
    print(f"บันทึกผลลัพธ์ไปที่: {output_path}")


if __name__ == "__main__":
    main()
