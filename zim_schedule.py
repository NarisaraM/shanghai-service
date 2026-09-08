#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ZIM Point-to-Point Sailing Schedule Fetcher
=============================================

ดึงตารางการเดินเรือ (Point To Point Schedule) จากเว็บไซต์ ZIM
(https://www.zim.com) โดย default จะค้นหาเส้นทาง Laem Chabang, Thailand ->
Shanghai (SH), China ตั้งแต่วันนี้ (เดือนปัจจุบัน) ไปจนถึงวันที่ 31 ธันวาคม
ของปีปัจจุบัน แล้ว export ผลลัพธ์ทั้งหมดออกเป็นไฟล์ Excel (.xlsx)

เบื้องหลังการทำงาน
-------------------
วิดเจ็ต "Schedule By Route" บนหน้าเว็บ ZIM (ปุ่ม Schedules ที่หน้าแรก หรือเมนู
Schedules > Point to Point) เรียกข้อมูลจาก REST API สาธารณะนี้โดยตรง (ไม่ต้อง
login, ไม่ต้องใช้เบราว์เซอร์/Playwright เลย):

    GET https://apigw.zim.com/digitalSchedules/PointToPoint/v2

พารามิเตอร์หลักที่ใช้ (ตรวจสอบจากหน้าเว็บจริงผ่าน DevTools > Network ขณะค้นหา):
    - subscription-key : API key สาธารณะที่ฝังอยู่ในหน้าเว็บ (ดูหมายเหตุด้านล่าง)
    - PortCode          : รหัสท่าเรือต้นทาง เช่น "THLEM;10" (Laem Chabang)
    - PortDestinationCode : รหัสท่าเรือปลายทาง เช่น "CNSNH;10" (Shanghai)
    - Direction          : "true" = ค้นหาตามวันออกเรือ (Departure)
    - FromDate            : วันที่เริ่มค้นหา รูปแบบ "DD-Month-YYYY" เช่น "08-September-2026"
    - WeeksAhead          : จำนวนสัปดาห์ที่จะค้นหาต่อการเรียก 1 ครั้ง (สูงสุด 12 สัปดาห์)

หน้าเว็บค้นหาได้ครั้งละไม่เกิน 12 สัปดาห์ต่อการเรียก 1 ครั้ง สคริปต์นี้จึงเลื่อน
วันที่เริ่มค้นหาไปทีละ 12 สัปดาห์ วนซ้ำจนถึงวันสิ้นสุดที่กำหนด (ปลายปี) แล้วรวม
ผลลัพธ์ทั้งหมด ตัดรายการซ้ำออก

หมายเหตุเกี่ยวกับ subscription-key
------------------------------------
ค่านี้เป็น API key แบบสาธารณะที่ฝังอยู่ในโค้ด JavaScript ของหน้าเว็บ (ทุกคนที่
เปิดหน้า Schedule ของ ZIM จะเรียก API ด้วยคีย์เดียวกันนี้) ไม่ใช่ข้อมูลส่วนตัว
ของผู้ใช้ แต่ ZIM อาจเปลี่ยนคีย์นี้เมื่อมีการอัพเดตเว็บไซต์ในอนาคต หากรันสคริปต์
แล้วได้ error 401/403 ให้เปิดหน้า https://www.zim.com -> Schedules ->
Point to Point -> ค้นหาเส้นทางใดก็ได้ -> เปิด DevTools (F12) > Network ->
กรองคำว่า "PointToPoint" -> ดู query parameter "subscription-key" ตัวใหม่
แล้วนำมาแทนที่ค่า DEFAULT_SUBSCRIPTION_KEY ด้านล่าง

รหัสท่าเรือ (Port Code) ที่ใช้บ่อย
-----------------------------------
    Laem Chabang, Thailand -> THLEM;10
    Shanghai (SH), China   -> CNSNH;10

หากต้องการท่าเรืออื่น ให้เปิดหน้า Point To Point บนเว็บ ZIM พิมพ์ชื่อท่าเรือ
ในช่อง From/To แล้วเปิด DevTools ดู query parameter PortCode/PortDestinationCode
ที่ระบบส่งไปตอนกด Search

ติดตั้งก่อนใช้งาน
------------------
    pip install requests pandas openpyxl

วิธีใช้งาน
----------
    # ค่า default: Laem Chabang -> Shanghai, วันนี้ -> 31 ธ.ค. ปีปัจจุบัน
    python3 zim_schedule.py

    # กำหนดเส้นทาง / ช่วงวันที่ / ไฟล์ผลลัพธ์เอง
    python3 zim_schedule.py --origin-code "THLEM;10" --origin-name "Laem Chabang, Thailand" \\
        --dest-code "CNSNH;10" --dest-name "Shanghai (SH), China" \\
        --start 2026-09-08 --end 2026-12-31 --out schedule.xlsx
"""

from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Optional

try:
    import requests
except ImportError:
    sys.exit("Missing dependency: pip install requests")

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing dependency: pip install pandas openpyxl")


API_URL = "https://apigw.zim.com/digitalSchedules/PointToPoint/v2"

# API key สาธารณะที่ฝังอยู่ในหน้าเว็บ ZIM ณ เวลาที่ตรวจสอบ (ก.ย. 2026)
# ถ้าใช้ไม่ได้แล้ว ให้หาค่าใหม่ตามวิธีในหมายเหตุด้านบนของไฟล์นี้
DEFAULT_SUBSCRIPTION_KEY = "9d63cf020a4c4708a7b0ebfe39578300"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.zim.com/",
    "Accept": "application/json, text/plain, */*",
}

MAX_WEEKS_PER_SEARCH = 12
WINDOW_STEP_DAYS = MAX_WEEKS_PER_SEARCH * 7


@dataclass
class Sailing:
    vessel_voyage_chain: Optional[str] = None
    num_legs: int = 1
    direct_or_ts: Optional[str] = None
    pol_name: Optional[str] = None
    pod_name: Optional[str] = None
    etd: Optional[str] = None
    eta: Optional[str] = None
    transit_days: Optional[str] = None
    line: Optional[str] = None
    doc_closing: Optional[str] = None
    vgm_closing: Optional[str] = None
    container_closing: Optional[str] = None
    hazardous_cutoff: Optional[str] = None

    def dedupe_key(self):
        return (self.vessel_voyage_chain, self.etd, self.pol_name, self.pod_name)


def fmt_iso(dt_str: Optional[str]) -> Optional[str]:
    """แปลง ISO datetime string (เช่น '2026-09-08T06:00:00+03:00') -> 'YYYY-MM-DD HH:MM'"""
    if not dt_str:
        return None
    try:
        d = datetime.fromisoformat(dt_str)
        return d.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return dt_str.replace("T", " ")[:16]


def month_step_dates(start: date, end: date, step_days: int = WINDOW_STEP_DAYS):
    cur = start
    while cur <= end:
        yield cur
        cur += timedelta(days=step_days)


def fetch_window(session: requests.Session, subscription_key: str, origin_code: str,
                  dest_code: str, from_date: date, weeks_ahead: int = MAX_WEEKS_PER_SEARCH) -> list[list[dict]]:
    params = {
        "subscription-key": subscription_key,
        "PortCode": origin_code,
        "PortDestinationCode": dest_code,
        "Direction": "true",  # true = search by date of departure
        "FromDate": from_date.strftime("%d-%B-%Y"),
        "WeeksAhead": str(weeks_ahead),
        "CountryCode": "US",
        "CargoType": "true",
        "EmissionsType": "true",
    }
    resp = session.get(API_URL, params=params, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data.get("routes", []) or []


def to_sailing(route_legs: list[dict]) -> Sailing:
    """route_legs = รายการ leg ของเที่ยวเรือ 1 เที่ยว (leg เดียว = Direct,
    มากกว่า 1 leg = มีการเปลี่ยนเรือ/ทรานชิป)"""
    first, last = route_legs[0], route_legs[-1]
    chain = " -> ".join(
        f"{leg.get('vesselName', '')} {leg.get('voyageNumber', '')}".strip()
        for leg in route_legs
    )
    num_legs = len(route_legs)

    return Sailing(
        vessel_voyage_chain=chain,
        num_legs=num_legs,
        direct_or_ts="Direct" if num_legs == 1 else f"T/S ({num_legs - 1} transshipment)",
        pol_name=first.get("portDepartureName"),
        pod_name=last.get("portArrivalName"),
        etd=fmt_iso(first.get("departureDate")),
        eta=fmt_iso(last.get("arrivalDate")),
        transit_days=sum(leg.get("daysAtSea") or 0 for leg in route_legs) or None,
        line=first.get("line"),
        doc_closing=fmt_iso(first.get("docClosingDate")),
        vgm_closing=fmt_iso(first.get("vgmClosingDate")),
        container_closing=fmt_iso(first.get("containerClosingDate")),
        hazardous_cutoff=fmt_iso(first.get("hazardousDocsCutOff")),
    )


def run(subscription_key: str, origin_code: str, origin_name: str,
        dest_code: str, dest_name: str, start: date, end: date, out_path: str) -> pd.DataFrame:

    session = requests.Session()
    all_sailings: list[Sailing] = []
    seen = set()

    for window_start in month_step_dates(start, end):
        print(f"[ค้นหา] {origin_name} -> {dest_name}  ตั้งแต่ {window_start.isoformat()} "
              f"(+{MAX_WEEKS_PER_SEARCH} สัปดาห์) ...")
        try:
            routes = fetch_window(session, subscription_key, origin_code, dest_code, window_start)
        except requests.RequestException as exc:
            print(f"   เรียก API ไม่สำเร็จ: {exc}  ข้ามช่วงนี้ไป")
            continue

        new_count = 0
        for route_legs in routes:
            if not route_legs:
                continue
            sailing = to_sailing(route_legs)
            key = sailing.dedupe_key()
            if key in seen:
                continue
            seen.add(key)
            all_sailings.append(sailing)
            new_count += 1
        print(f"   พบ {len(routes)} เที่ยวเรือ (ใหม่ {new_count} รายการ)")
        time.sleep(0.5)

    df = pd.DataFrame([s.__dict__ for s in all_sailings])
    if not df.empty:
        # กรองเฉพาะเที่ยวเรือที่ออกเรืออยู่ในช่วงวันที่ที่ผู้ใช้กำหนดจริงๆ
        # (การเรียกแต่ละครั้งอาจได้ข้อมูลเลยขอบเขตวันที่ไปบ้าง เพราะ WeeksAhead คงที่ที่ 12)
        end_str = end.strftime("%Y-%m-%d") + " 23:59"
        start_str = start.strftime("%Y-%m-%d") + " 00:00"
        df = df[(df["etd"].fillna("") >= start_str) & (df["etd"].fillna("9999") <= end_str)]

        column_order = [
            "vessel_voyage_chain", "direct_or_ts", "line",
            "pol_name", "etd", "pod_name", "eta", "transit_days",
            "doc_closing", "vgm_closing", "container_closing", "hazardous_cutoff",
        ]
        df = df[[c for c in column_order if c in df.columns]]
        df.sort_values(by=["etd"], inplace=True, na_position="last")
        df.to_excel(out_path, index=False)
        print(f"\nบันทึกผลลัพธ์ {len(df)} เที่ยวเรือ ลงไฟล์: {out_path}")
    else:
        print("\nไม่พบเที่ยวเรือใดๆ ในช่วงที่ค้นหา ไม่ได้สร้างไฟล์ผลลัพธ์")

    return df


def parse_args():
    today = date.today()
    default_end = date(today.year, 12, 31)

    parser = argparse.ArgumentParser(
        description="ดึงตารางการเดินเรือจาก ZIM (Point to Point) แล้ว export เป็น Excel",
    )
    parser.add_argument("--subscription-key", default=DEFAULT_SUBSCRIPTION_KEY,
                         help="ZIM API subscription key (ดูหมายเหตุในไฟล์นี้ถ้าคีย์เก่าใช้ไม่ได้แล้ว)")
    parser.add_argument("--origin-code", default="THLEM;10", help="รหัสท่าเรือต้นทาง (default: Laem Chabang)")
    parser.add_argument("--origin-name", default="Laem Chabang, Thailand", help="ชื่อท่าเรือต้นทาง (สำหรับแสดงผล)")
    parser.add_argument("--dest-code", default="CNSNH;10", help="รหัสท่าเรือปลายทาง (default: Shanghai)")
    parser.add_argument("--dest-name", default="Shanghai (SH), China", help="ชื่อท่าเรือปลายทาง (สำหรับแสดงผล)")
    parser.add_argument("--start", default=today.isoformat(), help="วันที่เริ่มค้นหา YYYY-MM-DD (default: วันนี้)")
    parser.add_argument("--end", default=default_end.isoformat(), help="วันที่สิ้นสุด YYYY-MM-DD (default: 31 ธ.ค. ปีปัจจุบัน)")
    parser.add_argument("--out", default="zim_schedule.xlsx", help="ชื่อไฟล์ Excel ผลลัพธ์")
    return parser.parse_args()


def main():
    args = parse_args()
    start = date.fromisoformat(args.start)
    end = date.fromisoformat(args.end)
    if end < start:
        sys.exit("วันที่สิ้นสุด (--end) ต้องไม่มาก่อนวันที่เริ่มต้น (--start)")

    run(
        subscription_key=args.subscription_key,
        origin_code=args.origin_code,
        origin_name=args.origin_name,
        dest_code=args.dest_code,
        dest_name=args.dest_name,
        start=start,
        end=end,
        out_path=args.out,
    )


if __name__ == "__main__":
    main()
