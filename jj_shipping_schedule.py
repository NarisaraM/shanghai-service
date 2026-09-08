#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JJ Shipping Thailand - Sailing Schedule Scraper
================================================

ดึงตารางการเดินเรือ (Sailing Schedule) จากเว็บไซต์ JJ Shipping Thailand
(https://yourjjshipping.app/en/sailing-schedules/) โดย default จะค้นหาเส้นทาง
Laem Chabang -> Shanghai ตั้งแต่วันนี้ (เดือนปัจจุบัน) ไปจนถึงวันที่ 31 ธันวาคม
ของปีปัจจุบัน แล้ว export ผลลัพธ์ทั้งหมดออกเป็นไฟล์ Excel (.xlsx)

วิธีการทำงาน
------------
หน้าเว็บ "ตารางการเดินเรือ" อนุญาตให้ค้นหาได้ครั้งละไม่เกิน 3 สัปดาห์
("Within next" -> "3 weeks") ต่อการค้นหา 1 ครั้ง สคริปต์นี้จึงเลื่อนวันที่
เริ่มค้นหาไปทีละ 21 วัน วนซ้ำไปจนถึงวันสิ้นสุดที่กำหนด (ปลายปี) เก็บผลลัพธ์
ของแต่ละรอบมารวมกัน แล้วตัดรายการที่ซ้ำกันออก (เทียบจากชื่อเรือ/เที่ยวเรือ/
วันที่ออกเรือ) ก่อน export

ข้อควรทราบเกี่ยวกับหน้าเว็บ (จากการทดสอบจริง)
-----------------------------------------------
- ถ้าไม่พบเที่ยวเรือในช่วงที่ค้นหา เว็บจะแสดงข้อความ "Schedule not found"
- บางเที่ยวเรือที่แสดงผล อาจไม่มีข้อมูลวันที่/ท่าเรือปลายทางครบถ้วน (ทางเว็บ
  เองบันทึกข้อมูลไม่ครบ) สคริปต์จะเก็บค่าที่มีเป็น None/ว่างไว้ตามจริง ไม่เดา
- ท่าเรือ "ต้นทาง" ที่แสดงในผลลัพธ์อาจเป็นท่าที่เรือออกจริง (เช่น
  "Bangkok : PAT1") ซึ่งอาจไม่ตรงกับท่าที่เลือกค้นหาเป๊ะๆ เนื่องจากสินค้าจาก
  Laem Chabang บางเที่ยวจะถูกขนย้าย (feeder) ไปลงเรือใหญ่ที่ท่าเรืออื่นก่อน
  ถือเป็นข้อมูลจริงจากเว็บ ไม่ใช่บั๊กของสคริปต์

ติดตั้งก่อนใช้งาน
------------------
    pip install playwright pandas openpyxl
    playwright install chromium

วิธีใช้งาน
----------
    # ค่า default: Laem Chabang -> Shanghai, วันนี้ -> 31 ธ.ค. ปีปัจจุบัน
    python3 jj_shipping_schedule.py

    # กำหนดเส้นทาง / ช่วงวันที่ / ไฟล์ผลลัพธ์เอง
    python3 jj_shipping_schedule.py --origin "Laem Chabang" --destination "Shanghai" \\
        --start 2026-09-08 --end 2026-12-31 --out schedule.xlsx

    # ดูหน้าเว็บไปด้วยระหว่างรัน (ไม่ headless) เผื่อต้องการ debug
    python3 jj_shipping_schedule.py --no-headless
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Optional

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing dependency: pip install pandas openpyxl")

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
except ImportError:
    sys.exit("Missing dependency: pip install playwright && playwright install chromium")


BASE_URL = "https://yourjjshipping.app/en/sailing-schedules/"

# ค่า value ของ <option> ในหน้าเว็บจริง (ตรวจสอบจากหน้าเว็บโดยตรงแล้ว)
ORIGIN_VALUES = {
    "thailand (all ports)": "0,0",
    "laem chabang": "4,7,7",
    "bangkok": "7,4,4",
}

DESTINATION_VALUES = {
    "sihanoukville": "309-4-10",
    "shanghai": "309-4-12",
    "xiamen": "309-4-13",
    "nansha": "327-7-11",
    "lianyungang": "327-7-16",
    "qingdao": "327-7-17,335-7-15",
    "hong kong": "335-7-10",
    "hakata": "327-7-12",
    "osaka": "327-7-13,335-7-13",
    "kobe": "327-7-14,335-7-14",
    "yokohama": "335-7-11",
    "nagoya": "335-7-12",
    "busan": "327-7-15",
    "ho chi minh city": "309-4-11,327-7-10",
}

MAX_WEEKS_PER_SEARCH = 3      # ค่าสูงสุดที่ dropdown "Within next" เลือกได้
WINDOW_STEP_DAYS = 21         # = 3 สัปดาห์ ต่อการค้นหา 1 ครั้ง


@dataclass
class Sailing:
    search_origin: str
    search_destination: str
    search_window_start: str
    closing: Optional[str] = None
    departure_date: Optional[str] = None
    departure_port: Optional[str] = None
    vessel: Optional[str] = None
    voyage: Optional[str] = None
    tcf_code: Optional[str] = None
    transit_days: Optional[str] = None
    arrival_date: Optional[str] = None
    arrival_port: Optional[str] = None
    raw_text: str = field(default="", repr=False)

    def dedupe_key(self):
        return (self.vessel, self.voyage, self.departure_date, self.departure_port)


def daterange_windows(start: date, end: date, step_days: int = WINDOW_STEP_DAYS):
    """สร้างรายการวันที่เริ่มต้นของแต่ละรอบค้นหา ตั้งแต่ start จนกว่าจะเลย end"""
    cur = start
    while cur <= end:
        yield cur
        cur += timedelta(days=step_days)


def parse_sailing_card(text: str, search_origin: str, search_destination: str,
                        window_start: date) -> Sailing:
    """
    แปลงข้อความจากการ์ดผลลัพธ์ 1 ใบ (element .sailing-card) ให้เป็น Sailing
    โครงสร้างบรรทัดที่พบจริงบนเว็บ (เรียงตามลำดับ) คือ:
        Closing {yyyy/mm/dd hh:mm}
        {วันที่ออกเรือ}                 เช่น "Mon 14 September 2026"
        {ท่าเรือต้นทาง}                 เช่น "Bangkok : PAT1"
        Vessel {ชื่อเรือ}
        VOY {เที่ยวเรือ}
        TCF Code {รหัส (อาจว่าง)}
        {N} DAYS
        {วันที่ถึงปลายทาง}   (อาจไม่มี ถ้าเว็บไม่มีข้อมูล)
        {ท่าเรือปลายทาง}     (อาจไม่มี ถ้าเว็บไม่มีข้อมูล)
        Get A Quote
    """
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    sailing = Sailing(
        search_origin=search_origin,
        search_destination=search_destination,
        search_window_start=window_start.isoformat(),
        raw_text=" | ".join(lines),
    )

    if not lines:
        return sailing

    if lines[0].lower().startswith("closing"):
        sailing.closing = lines[0][len("closing"):].strip()
    if len(lines) > 1:
        sailing.departure_date = lines[1]
    if len(lines) > 2:
        sailing.departure_port = lines[2]

    days_idx = None
    for i, ln in enumerate(lines):
        low = ln.lower()
        if low.startswith("vessel"):
            sailing.vessel = ln[len("vessel"):].strip(" :")
        elif low.startswith("voy"):
            sailing.voyage = ln[len("voy"):].strip(" :")
        elif low.startswith("tcf code"):
            code = ln[len("tcf code"):].strip(" :")
            sailing.tcf_code = code or None
        else:
            m = re.match(r"^(\d+)\s*DAYS?$", ln, re.IGNORECASE)
            if m:
                sailing.transit_days = m.group(1)
                days_idx = i

    if days_idx is not None:
        tail = [ln for ln in lines[days_idx + 1:] if ln.lower() != "get a quote"]
        if len(tail) >= 1:
            sailing.arrival_date = tail[0]
        if len(tail) >= 2:
            sailing.arrival_port = tail[1]

    return sailing


def set_date_field(page, target: date) -> None:
    """
    ช่อง 'Date' ใช้ jQuery UI Datepicker (id="date") การพิมพ์ข้อความตรงๆ
    อาจเปิด popup ปฏิทินค้างไว้ จึงตั้งค่าผ่าน jQuery datepicker API แทน
    ซึ่งจะอัพเดตค่าในช่องและ trigger การเปลี่ยนแปลงให้ทันที
    """
    page.evaluate(
        """(d) => {
            const $ = window.jQuery;
            const input = document.querySelector('#date');
            if ($ && $.fn && $.fn.datepicker) {
                $('#date').datepicker('setDate', new Date(d.y, d.m, d.day));
            } else if (input) {
                input.value = d.iso;
                input.dispatchEvent(new Event('input', {bubbles: true}));
                input.dispatchEvent(new Event('change', {bubbles: true}));
            }
        }""",
        {"y": target.year, "m": target.month - 1, "day": target.day, "iso": target.isoformat()},
    )


def search_one_window(page, origin_value: str, destination_value: str,
                       window_start: date, weeks: str = str(MAX_WEEKS_PER_SEARCH)) -> list[str]:
    """ตั้งค่าฟอร์มค้นหาสำหรับ 1 รอบ แล้วคืนค่า inner_text ของการ์ดผลลัพธ์ทั้งหมด"""

    # 1) เลือกต้นทางก่อนเสมอ เพราะ dropdown ปลายทางจะถูกโหลดใหม่ตามต้นทางที่เลือก
    page.select_option("#origin", origin_value)
    # รอให้ dropdown ปลายทางโหลดตัวเลือกใหม่เสร็จ (ยิง AJAX ภายในหน้าเว็บ)
    page.wait_for_timeout(1500)

    # 2) เลือกปลายทาง
    page.select_option("#options", destination_value)

    # 3) ตั้งวันที่เริ่มค้นหา
    set_date_field(page, window_start)

    # 4) เลือกช่วงเวลาแสดงผล (สูงสุด 3 สัปดาห์ต่อการค้นหา)
    page.select_option("#week", weeks)

    # 5) กดค้นหา แล้วรอผลลัพธ์
    page.click("#submit")
    try:
        page.wait_for_selector(
            "#schedule-results .sailing-card, #schedule-results >> text=Schedule not found",
            timeout=15000,
        )
    except PWTimeoutError:
        pass
    # เผื่อ DOM render ไม่ทันหลัง selector เจอ element แรก
    page.wait_for_timeout(1500)

    results = page.locator("#schedule-results")
    if "schedule not found" in results.inner_text().lower():
        return []

    cards = page.locator("#schedule-results .sailing-card")
    return [cards.nth(i).inner_text() for i in range(cards.count())]


def run(origin: str, destination: str, start: date, end: date, out_path: str,
        headless: bool = True) -> pd.DataFrame:
    origin_key = origin.strip().lower()
    dest_key = destination.strip().lower()

    if origin_key not in ORIGIN_VALUES:
        sys.exit(f"ไม่รู้จักต้นทาง '{origin}'. ตัวเลือกที่มี: {', '.join(ORIGIN_VALUES)}")
    if dest_key not in DESTINATION_VALUES:
        sys.exit(f"ไม่รู้จักปลายทาง '{destination}'. ตัวเลือกที่มี: {', '.join(DESTINATION_VALUES)}")

    origin_value = ORIGIN_VALUES[origin_key]
    destination_value = DESTINATION_VALUES[dest_key]

    all_sailings: list[Sailing] = []
    seen_keys = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page()
        page.goto(BASE_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        for window_start in daterange_windows(start, end):
            print(f"[ค้นหา] {origin} -> {destination}  ตั้งแต่ {window_start.isoformat()} "
                  f"(+{MAX_WEEKS_PER_SEARCH} สัปดาห์) ...")
            try:
                card_texts = search_one_window(page, origin_value, destination_value, window_start)
            except PWTimeoutError:
                print("   หมดเวลารอผลลัพธ์ ข้ามรอบนี้ไป")
                continue

            new_count = 0
            for text in card_texts:
                sailing = parse_sailing_card(text, origin, destination, window_start)
                key = sailing.dedupe_key()
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                all_sailings.append(sailing)
                new_count += 1
            print(f"   พบ {len(card_texts)} เที่ยวเรือ (ใหม่ {new_count} รายการ)")

        browser.close()

    df = pd.DataFrame([s.__dict__ for s in all_sailings])
    if not df.empty:
        column_order = [
            "search_origin", "search_destination", "closing",
            "departure_date", "departure_port", "vessel", "voyage",
            "tcf_code", "transit_days", "arrival_date", "arrival_port",
            "search_window_start", "raw_text",
        ]
        df = df[[c for c in column_order if c in df.columns]]
        df.sort_values(by=["departure_date"], inplace=True, na_position="last")
        df.to_excel(out_path, index=False)
        print(f"\nบันทึกผลลัพธ์ {len(df)} เที่ยวเรือ ลงไฟล์: {out_path}")
    else:
        print("\nไม่พบเที่ยวเรือใดๆ ในช่วงที่ค้นหา ไม่ได้สร้างไฟล์ผลลัพธ์")

    return df


def parse_args():
    today = date.today()
    default_end = date(today.year, 12, 31)

    parser = argparse.ArgumentParser(
        description="ดึงตารางการเดินเรือจาก JJ Shipping Thailand แล้ว export เป็น Excel",
    )
    parser.add_argument("--origin", default="Laem Chabang",
                         help=f"ท่าเรือต้นทาง (default: Laem Chabang) ตัวเลือก: {', '.join(ORIGIN_VALUES)}")
    parser.add_argument("--destination", default="Shanghai",
                         help=f"ท่าเรือปลายทาง (default: Shanghai) ตัวเลือก: {', '.join(DESTINATION_VALUES)}")
    parser.add_argument("--start", default=today.isoformat(),
                         help="วันที่เริ่มค้นหา รูปแบบ YYYY-MM-DD (default: วันนี้)")
    parser.add_argument("--end", default=default_end.isoformat(),
                         help="วันที่สิ้นสุด รูปแบบ YYYY-MM-DD (default: 31 ธ.ค. ปีปัจจุบัน)")
    parser.add_argument("--out", default="jj_shipping_schedule.xlsx",
                         help="ชื่อไฟล์ Excel ผลลัพธ์ (default: jj_shipping_schedule.xlsx)")
    parser.add_argument("--no-headless", dest="headless", action="store_false",
                         help="เปิดหน้าต่างเบราว์เซอร์ให้เห็นระหว่างรัน (ปกติจะรันแบบ headless)")
    parser.set_defaults(headless=True)
    return parser.parse_args()


def main():
    args = parse_args()
    start = date.fromisoformat(args.start)
    end = date.fromisoformat(args.end)
    if end < start:
        sys.exit("วันที่สิ้นสุด (--end) ต้องไม่มาก่อนวันที่เริ่มต้น (--start)")

    run(
        origin=args.origin,
        destination=args.destination,
        start=start,
        end=end,
        out_path=args.out,
        headless=args.headless,
    )


if __name__ == "__main__":
    main()
