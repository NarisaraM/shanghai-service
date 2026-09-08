#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
e-KMTC Sailing Schedule Fetcher
================================

ดึงตารางการเดินเรือ (Sailing Schedule) จากเว็บไซต์ e-KMTC
(https://www.ekmtc.com) โดย default จะค้นหาเส้นทาง Laem Chabang (LCH), Thailand
-> Shanghai (SHA), China ตั้งแต่เดือนปัจจุบันไปจนถึงเดือนธันวาคมของปีปัจจุบัน
แล้ว export ผลลัพธ์ทั้งหมดออกเป็นไฟล์ Excel (.xlsx)

เบื้องหลังการทำงาน
-------------------
หน้า Schedule ของเว็บ e-KMTC (Menu: Schedule > Corridor) เรียกข้อมูลจาก REST API
สาธารณะตัวนี้โดยตรง (ไม่ต้อง login และไม่ต้องใช้เบราว์เซอร์/Playwright เลย):

    GET https://api.ekmtc.com/schedule/schedule/leg/search-schedule

โดยส่ง query parameter เป็นรหัสท่าเรือต้นทาง/ปลายทาง + เดือน/ปีที่ต้องการค้นหา
(ตรวจสอบโดยการเปิด DevTools ดู Network request ขณะค้นหาบนเว็บจริง) หน้าเว็บค้นหา
ได้ครั้งละ 1 เดือนต่อการเรียก 1 ครั้ง สคริปต์นี้จึงวนเรียกทีละเดือนตั้งแต่เดือน
เริ่มต้นถึงเดือนสิ้นสุดที่กำหนด แล้วรวมผลลัพธ์ทั้งหมด ตัดรายการซ้ำออก (เทียบจาก
รหัสเรือ + เที่ยวเรือ + วันที่ออกเรือ)

รหัสท่าเรือ (place code) ที่ใช้บ่อย
-----------------------------------
    Laem Chabang, Thailand      -> LCH / TH
    Bangkok, Thailand           -> BKK / TH
    Shanghai, China             -> SHA / CN
    Xiamen, China               -> XMN / CN
    Ningbo, China                -> NGB / CN

หากต้องการท่าเรืออื่น ให้เปิดหน้า https://www.ekmtc.com > Schedule > Corridor
พิมพ์ชื่อท่าเรือในช่อง Departure/Arrival แล้วดูรหัสในวงเล็บที่ระบบแนะนำ (เช่น
"Shanghai, China (SHA)" -> รหัสคือ SHA) แล้วนำมาใส่ใน --origin-code/--dest-code

ติดตั้งก่อนใช้งาน
------------------
    pip install requests pandas openpyxl

วิธีใช้งาน
----------
    # ค่า default: Laem Chabang -> Shanghai, เดือนปัจจุบัน -> ธันวาคม ปีปัจจุบัน
    python3 ekmtc_schedule.py

    # กำหนดเส้นทาง / ช่วงเดือน / ไฟล์ผลลัพธ์เอง
    python3 ekmtc_schedule.py \\
        --origin-code LCH --origin-name "Laem Chabang, Thailand (LCH)" --origin-country TH \\
        --dest-code SHA --dest-name "Shanghai, China (SHA)" --dest-country CN \\
        --start-year 2026 --start-month 9 --end-year 2026 --end-month 12 \\
        --out schedule.xlsx
"""

from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime
from typing import Optional

try:
    import requests
except ImportError:
    sys.exit("Missing dependency: pip install requests")

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing dependency: pip install pandas openpyxl")


API_URL = "https://api.ekmtc.com/schedule/schedule/leg/search-schedule"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.ekmtc.com/index.html",
    "Accept": "application/json, text/plain, */*",
}

# ค่าฟิลเตอร์เริ่มต้น: รวมทั้งเที่ยวตรง (Direct) และเรือต่อ (Transship: T/S)
DEFAULT_PARAMS = {
    "pointChangeYN": "",
    "bound": "O",              # O = Outbound / export (ค้นหาโดยวันออกเรือจากต้นทาง)
    "filterPolCd": "",
    "pointLength": "",
    "filterYn": "N",
    "searchYN": "Y",
    "filterPodCd": "",
    "hiddestPlcCd": "",
    "polTrmlStr": "",
    "podTrmlStr": "",
    "rteCd": "",
    "filterTs": "Y",           # รวมเที่ยวที่มีการเปลี่ยนเรือ (T/S)
    "filterDirect": "Y",       # รวมเที่ยวตรง (Direct)
    "filterTranMax": "0",
    "filterTranMin": "0",
    "hidstartPlcCd": "",
    "main": "N",
    "legIdx": "0",
    "vslType01": "01",
    "vslType03": "03",
    "unno": "",
    "commodityCd": "",
    "eiCatCd": "O",            # O = ค้นหาตามวันออกเรือ (Date of Departure)
    "calendarOrList": "L",     # L = list view (ข้อมูลตารางแบบละเอียด)
    "cpYn": "N",
    "promotionChk": "N",
    "vslCd": "",
    "voyNo": "",
}


@dataclass
class Sailing:
    vessel: Optional[str] = None
    voyage: Optional[str] = None
    route_code: Optional[str] = None
    route_name: Optional[str] = None
    direct_or_ts: Optional[str] = None
    pol_code: Optional[str] = None
    pol_name: Optional[str] = None
    pol_terminal: Optional[str] = None
    pol_berth: Optional[str] = None      # วันที่/เวลาเรือเทียบท่าต้นทาง
    etd: Optional[str] = None            # วันที่/เวลาออกเรือจากต้นทาง
    pod_code: Optional[str] = None
    pod_name: Optional[str] = None
    pod_terminal: Optional[str] = None
    eta: Optional[str] = None            # วันที่/เวลาถึงปลายทาง
    transit_time: Optional[str] = None
    doc_closing: Optional[str] = None    # Document Closing
    cargo_closing: Optional[str] = None  # Full Container Gate-In Closing

    def dedupe_key(self):
        return (self.vessel, self.voyage, self.etd, self.pol_code, self.pod_code)


def fmt_datetime(date_str: str, time_str: str = "") -> Optional[str]:
    """
    แปลงวันที่/เวลาจากรูปแบบของ API เป็น 'YYYY-MM-DD HH:MM'

    รองรับ 2 กรณี:
      - date_str='YYYYMMDD' แยก time_str='HHMM' ต่างหาก (เช่น etd/etdTm)
      - date_str='YYYYMMDDHHMM' รวมกันมาในค่าเดียว (เช่น bkgDocCls/bkgCgoCls)
    คืนค่า None ถ้าไม่มีข้อมูล
    """
    if not date_str or len(date_str) < 8:
        return None
    try:
        d = datetime.strptime(date_str[:8], "%Y%m%d")
    except ValueError:
        return date_str

    if not time_str and len(date_str) >= 12:
        time_str = date_str[8:12]

    if time_str and len(time_str) == 4:
        return f"{d.strftime('%Y-%m-%d')} {time_str[:2]}:{time_str[2:]}"
    return d.strftime("%Y-%m-%d")


def month_iter(start_year: int, start_month: int, end_year: int, end_month: int):
    y, m = start_year, start_month
    while (y, m) <= (end_year, end_month):
        yield y, m
        m += 1
        if m > 12:
            m = 1
            y += 1


def fetch_month(session: requests.Session, origin_code: str, origin_name: str,
                 origin_country: str, dest_code: str, dest_name: str,
                 dest_country: str, year: int, month: int) -> list[dict]:
    params = dict(DEFAULT_PARAMS)
    params.update({
        "startPlcCd": origin_code,
        "startPlcName": origin_name,
        "startCtrCd": origin_country,
        "destPlcCd": dest_code,
        "destPlcName": dest_name,
        "destCtrCd": dest_country,
        "searchMonth": f"{month:02d}",
        "searchYear": str(year),
    })
    resp = session.get(API_URL, params=params, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data.get("listSchedule", []) or []


def to_sailing(rec: dict) -> Sailing:
    return Sailing(
        vessel=rec.get("vslNm"),
        voyage=rec.get("voyNo"),
        route_code=rec.get("rteCd"),
        route_name=rec.get("rteCdNm"),
        direct_or_ts="T/S" if rec.get("ts") == "Y" else "Direct",
        pol_code=rec.get("pol"),
        pol_name=rec.get("polNm"),
        pol_terminal=rec.get("otrmlNm") or rec.get("polTml"),
        pol_berth=fmt_datetime(rec.get("polEtb", ""), rec.get("polEtbTm", "")),
        etd=fmt_datetime(rec.get("etd", ""), rec.get("etdTm", "")),
        pod_code=rec.get("pod"),
        pod_name=rec.get("podNm"),
        pod_terminal=rec.get("itrmlNm") or rec.get("podTml"),
        eta=fmt_datetime(rec.get("eta", ""), rec.get("etaTm", "")),
        transit_time=rec.get("transitTime"),
        doc_closing=fmt_datetime(rec.get("bkgDocCls", "")),
        cargo_closing=fmt_datetime(rec.get("bkgCgoCls", "")),
    )


def run(origin_code: str, origin_name: str, origin_country: str,
        dest_code: str, dest_name: str, dest_country: str,
        start_year: int, start_month: int, end_year: int, end_month: int,
        out_path: str) -> pd.DataFrame:

    session = requests.Session()
    all_sailings: list[Sailing] = []
    seen = set()

    for year, month in month_iter(start_year, start_month, end_year, end_month):
        print(f"[ค้นหา] {origin_name} -> {dest_name}  เดือน {year}-{month:02d} ...")
        try:
            records = fetch_month(session, origin_code, origin_name, origin_country,
                                   dest_code, dest_name, dest_country, year, month)
        except requests.RequestException as exc:
            print(f"   เรียก API ไม่สำเร็จ: {exc}  ข้ามเดือนนี้ไป")
            continue

        new_count = 0
        for rec in records:
            sailing = to_sailing(rec)
            key = sailing.dedupe_key()
            if key in seen:
                continue
            seen.add(key)
            all_sailings.append(sailing)
            new_count += 1
        print(f"   พบ {len(records)} เที่ยวเรือ (ใหม่ {new_count} รายการ)")
        time.sleep(0.5)  # เว้นจังหวะเรียก API เล็กน้อย ไม่ยิงรัวเกินไป

    df = pd.DataFrame([s.__dict__ for s in all_sailings])
    if not df.empty:
        column_order = [
            "vessel", "voyage", "direct_or_ts", "route_code", "route_name",
            "pol_code", "pol_name", "pol_terminal", "pol_berth", "etd",
            "pod_code", "pod_name", "pod_terminal", "eta", "transit_time",
            "doc_closing", "cargo_closing",
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

    parser = argparse.ArgumentParser(
        description="ดึงตารางการเดินเรือจาก e-KMTC แล้ว export เป็น Excel",
    )
    parser.add_argument("--origin-code", default="LCH", help="รหัสท่าเรือต้นทาง (default: LCH = Laem Chabang)")
    parser.add_argument("--origin-name", default="Laem Chabang, Thailand (LCH)",
                         help="ชื่อท่าเรือต้นทางแบบเต็มตามที่เว็บใช้")
    parser.add_argument("--origin-country", default="TH", help="รหัสประเทศต้นทาง (default: TH)")

    parser.add_argument("--dest-code", default="SHA", help="รหัสท่าเรือปลายทาง (default: SHA = Shanghai)")
    parser.add_argument("--dest-name", default="Shanghai, China (SHA)",
                         help="ชื่อท่าเรือปลายทางแบบเต็มตามที่เว็บใช้")
    parser.add_argument("--dest-country", default="CN", help="รหัสประเทศปลายทาง (default: CN)")

    parser.add_argument("--start-year", type=int, default=today.year, help="ปีเริ่มต้นค้นหา (default: ปีปัจจุบัน)")
    parser.add_argument("--start-month", type=int, default=today.month, help="เดือนเริ่มต้นค้นหา (default: เดือนปัจจุบัน)")
    parser.add_argument("--end-year", type=int, default=today.year, help="ปีสิ้นสุดค้นหา (default: ปีปัจจุบัน)")
    parser.add_argument("--end-month", type=int, default=12, help="เดือนสิ้นสุดค้นหา (default: 12 = ธันวาคม)")

    parser.add_argument("--out", default="ekmtc_schedule.xlsx", help="ชื่อไฟล์ Excel ผลลัพธ์")
    return parser.parse_args()


def main():
    args = parse_args()
    if (args.end_year, args.end_month) < (args.start_year, args.start_month):
        sys.exit("เดือน/ปีสิ้นสุด ต้องไม่มาก่อนเดือน/ปีเริ่มต้น")

    run(
        origin_code=args.origin_code,
        origin_name=args.origin_name,
        origin_country=args.origin_country,
        dest_code=args.dest_code,
        dest_name=args.dest_name,
        dest_country=args.dest_country,
        start_year=args.start_year,
        start_month=args.start_month,
        end_year=args.end_year,
        end_month=args.end_month,
        out_path=args.out,
    )


if __name__ == "__main__":
    main()
