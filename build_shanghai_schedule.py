#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_shanghai_schedule.py
==========================

รวม (aggregate) ตารางการเดินเรือเส้นทาง **Laem Chabang -> Shanghai** จาก
สคริปต์ scraper ต้นฉบับในโฟลเดอร์ input/ ให้กลายเป็นชุดข้อมูลเดียว
พร้อมสร้าง "ปฏิทินตารางเรือ" (calendar) และ Dashboard HTML ที่อ่านง่าย

แหล่งข้อมูลที่รองรับ (สคริปต์ต้นฉบับอยู่ในโฟลเดอร์ input/)
---------------------------------------------------------
    ดึงสด (public JSON API):
    input/sitc_schedule.py        -> สายเรือ SITC
    input/tslines_schedule.py     -> สายเรือ T.S. Lines
    input/culines_ptp_schedule.py -> สายเรือ CU Lines
    input/cosco_schedule.py       -> สายเรือ COSCO
    input/yml_schedule_lcb_shanghai.py -> สายเรือ Yang Ming (YML)

    ดึงสด (เว็บ + Playwright):
    input/jj_shipping_schedule.py -> SJJ  (NVOCC JJ Shipping)

    อ่านจากไฟล์ที่ดาวน์โหลดเอง (เว็บบล็อก IP นอกไทย):
    KMTC     -> ไฟล์ .xls รายเดือนในโฟลเดอร์  KMTC/
    ZIM      -> ไฟล์  zim_*.xlsx           ที่รากโปรเจกต์

หมายเหตุ: เว็บ KMTC (ekmtc.com) และ ZIM (zim.com) ใช้ Akamai กันบอต — IP นอกไทย
จะโดนบล็อกทั้งโดเมน จึงใช้วิธี "ดาวน์โหลดไฟล์เอง": รันสคริปต์ input/*.py ของสายนั้น
จากเครือข่ายในไทย แล้ววางไฟล์ผลลัพธ์ไว้ตามที่ระบุข้างบน ถ้าไม่มีไฟล์ Dashboard จะ
ขึ้นสถานะ "ถูกบล็อก" พร้อมวิธีแก้

การทำงาน
--------
1. เรียกฟังก์ชันดึงข้อมูลของแต่ละ scraper (แหล่งไหนล่ม/โดนบล็อก ก็ข้ามไป
   แล้วรายงานสถานะไว้บน Dashboard)
2. แปลงข้อมูลทุกแหล่งให้อยู่ในรูปแบบมาตรฐานเดียวกัน (normalize)
3. **รวมเที่ยวเรือที่ซ้ำกัน** — ถ้าเป็นเรือลำเดียวกันและออกเรือวันเดียวกัน
   (คลาดเคลื่อนได้ไม่เกิน 1 วัน) จะยุบเป็นรายการเดียว แล้วแจกแจงด้านในว่า
   จองผ่านสายเรือ/ผู้ให้บริการรายใดได้บ้าง พร้อม ETD/ETA/Cut-off ของแต่ละราย
4. Export ผลลัพธ์:
       output/combined_schedule.xlsx          (3 ชีต: Merged / All Sailings / Sources)
       output/shanghai_schedule_dashboard.html (ปฏิทิน + Dashboard)
       output/shanghai_schedule.ics           (ไฟล์ปฏิทินนำเข้า Google/Outlook ได้)
       output/_cache.json                      (แคชข้อมูล ใช้กับโหมด --offline)

ติดตั้งก่อนใช้งาน
----------------
    pip install requests pandas openpyxl playwright
    playwright install chromium

วิธีใช้งาน
---------
    python build_shanghai_schedule.py
    python build_shanghai_schedule.py --start 2026-09-01 --end 2026-12-31
    python build_shanghai_schedule.py --only sitc,tslines
    python build_shanghai_schedule.py --skip sjj
    python build_shanghai_schedule.py --offline     # ไม่ยิงเน็ต ใช้ข้อมูลจาก _cache.json
"""

from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import re
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
INPUT_DIR = HERE / "input"          # โฟลเดอร์เก็บสคริปต์ scraper ต้นฉบับ
OUT_DIR = HERE / "output"
CACHE_FILE = OUT_DIR / "_cache.json"

LANE = "Laem Chabang, Thailand  \u2192  Shanghai, China"

# สี/ลำดับ ของแต่ละแหล่งข้อมูล (ใช้บน Dashboard + ปฏิทิน)
SOURCE_META = {
    "SITC":        {"color": "#2563eb", "label": "SITC"},
    "T.S. Lines":  {"color": "#16a34a", "label": "T.S. Lines"},
    "KMTC":        {"color": "#db2777", "label": "KMTC"},
    "ZIM":         {"color": "#f59e0b", "label": "ZIM"},
    "CU Lines":    {"color": "#0891b2", "label": "CU Lines"},
    "COSCO":       {"color": "#0f766e", "label": "COSCO"},
    "Yang Ming":   {"color": "#65a30d", "label": "Yang Ming (YML)"},
    "SJJ":         {"color": "#7c3aed", "label": "SJJ"},
}
DEFAULT_COLOR = "#64748b"


# --------------------------------------------------------------------------- #
#  helpers                                                                     #
# --------------------------------------------------------------------------- #
def norm_vessel(name: str | None) -> str:
    """ปรับชื่อเรือให้เทียบกันได้: ตัวพิมพ์ใหญ่ / ตัดอักขระพิเศษ / ยุบช่องว่าง"""
    if not name:
        return ""
    s = str(name).upper()
    s = s.replace("M.V.", " ").replace("MV ", " ").replace("M/V", " ")
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


_MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June", "July",
                "August", "September", "October", "November", "December"]
MONTHS = {name.lower(): i for i, name in enumerate(_MONTH_NAMES)}
MONTHS.update({name.lower()[:3]: i for i, name in enumerate(_MONTH_NAMES) if name})


def parse_dt(value) -> tuple[str | None, dt.date | None]:
    """
    รับค่า วันที่/เวลา ได้หลายรูปแบบจาก scraper แต่ละตัว แล้วคืน
        (ข้อความ 'YYYY-MM-DD [HH:MM]', datetime.date หรือ None)
    รองรับ:
        2026-09-14 05:00 / 2026-09-14
        2026/09/14 05:00 / 2026/09/14
        08-September-2026
        Mon 14 September 2026  /  14 September 2026
        2026-09-14T05:00:00+03:00
    """
    if value is None:
        return None, None
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "-", "n/a"}:
        return None, None

    s = s.replace("T", " ")
    # ISO offset เช่น +03:00 ท้ายสุด
    s = re.sub(r"([+\-]\d{2}:?\d{2})$", "", s).strip()
    # ตัดชื่อวันขึ้นต้น เช่น 'Monday, 14-SEP-2026' (รูปแบบของ CMA CGM)
    s = re.sub(r"^[A-Za-z]{3,},\s*", "", s).strip()

    # 1) YYYY-MM-DD หรือ YYYY/MM/DD (อาจมีเวลา)
    m = re.match(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ ](\d{1,2}):(\d{2}))?", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            date = dt.date(y, mo, d)
        except ValueError:
            return s, None
        if m.group(4) is not None:
            return f"{date.isoformat()} {int(m.group(4)):02d}:{m.group(5)}", date
        return date.isoformat(), date

    # 1b) YYYY.Mon.DD [HH:MM]  (รูปแบบไฟล์ export ของ KMTC เช่น '2026.Nov.15 23:00')
    m = re.match(r"(\d{4})\.([A-Za-z]{3,})\.(\d{1,2})(?:[ ](\d{1,2}):(\d{2}))?", s)
    if m and m.group(2).lower() in MONTHS:
        y, mo, d = int(m.group(1)), MONTHS[m.group(2).lower()], int(m.group(3))
        try:
            date = dt.date(y, mo, d)
        except ValueError:
            return s, None
        if m.group(4) is not None:
            return f"{date.isoformat()} {int(m.group(4)):02d}:{m.group(5)}", date
        return date.isoformat(), date

    # 1c) DD/MM/YYYY  (รูปแบบของ RCL)
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})(?:[ ](\d{1,2}):(\d{2}))?", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            date = dt.date(y, mo, d)
        except ValueError:
            return s, None
        if m.group(4) is not None:
            return f"{date.isoformat()} {int(m.group(4)):02d}:{m.group(5)}", date
        return date.isoformat(), date

    # 2) DD-Month-YYYY  /  DD-SEP-2026
    m = re.match(r"(\d{1,2})[-\s]([A-Za-z]+)[-\s](\d{4})", s)
    if m and m.group(2).lower() in MONTHS:
        d, mo, y = int(m.group(1)), MONTHS[m.group(2).lower()], int(m.group(3))
        try:
            date = dt.date(y, mo, d)
            return date.isoformat(), date
        except ValueError:
            return s, None

    # 3) [Ddd] DD Month YYYY  (รูปแบบของ SJJ / JJ Shipping)
    m = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", s)
    if m and m.group(2).lower() in MONTHS:
        d, mo, y = int(m.group(1)), MONTHS[m.group(2).lower()], int(m.group(3))
        try:
            date = dt.date(y, mo, d)
            return date.isoformat(), date
        except ValueError:
            return s, None

    return s, None


def to_float(value) -> float | None:
    if value is None:
        return None
    m = re.search(r"-?\d+(?:\.\d+)?", str(value))
    return float(m.group()) if m else None


def clean(value) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none"}:
        return None
    return s


class SourceBlocked(RuntimeError):
    """ยกขึ้นเมื่อแหล่งข้อมูลบล็อกการเข้าถึงที่ระดับ IP/ภูมิภาค (เช่น Akamai)"""


def http_retry(call, *, tries: int = 3, delay: float = 2.0):
    """
    เรียกฟังก์ชัน HTTP ซ้ำได้สูงสุด `tries` ครั้ง (หน่วงเพิ่มขึ้นเรื่อย ๆ)
    - ถ้าเจอ 401/403 พร้อมหน้า "Access Denied" ของ Akamai จะเลิก retry ทันที
      แล้วยก SourceBlocked (retry ไปก็โดนบล็อกเหมือนเดิม เพราะบล็อกที่ IP)
    """
    import requests

    last = None
    for attempt in range(1, tries + 1):
        try:
            return call()
        except requests.HTTPError as exc:
            last = exc
            resp = getattr(exc, "response", None)
            code = getattr(resp, "status_code", None)
            body = (getattr(resp, "text", "") or "")[:2000]
            if code in (401, 403) and ("Access Denied" in body or "edgesuite" in body
                                       or "denied" in body.lower()):
                raise SourceBlocked(
                    f"ถูกบล็อกที่ระดับ IP/ภูมิภาค (HTTP {code} – Akamai Access Denied). "
                    f"ลองรันสคริปต์นี้จากเครื่อง/เครือข่ายอื่น (เช่นเน็ตออฟฟิศในไทย) "
                    f"หรือผ่าน VPN แล้วแหล่งนี้จะกลับมาใช้ได้"
                ) from exc
            time.sleep(delay * attempt)
        except requests.RequestException as exc:
            last = exc
            time.sleep(delay * attempt)
    raise last if last else RuntimeError("http_retry: unknown failure")


# --------------------------------------------------------------------------- #
#  แต่ละแหล่งข้อมูล: ดึง + แปลงเป็นรูปแบบมาตรฐาน                                 #
#  คืนค่า list[dict] ที่มี key: source, carrier, service, vessel, voyage,       #
#        etd, eta, etd_date, pol, pod, pol_terminal, pod_terminal,             #
#        transit_days, direct_or_ts, doc_cutoff, cy_cutoff, vgm_cutoff         #
# --------------------------------------------------------------------------- #
def _rec(source, **kw) -> dict:
    etd_txt, etd_date = parse_dt(kw.get("etd"))
    eta_txt, _ = parse_dt(kw.get("eta"))
    doc_txt, _ = parse_dt(kw.get("doc_cutoff"))
    cy_txt, _ = parse_dt(kw.get("cy_cutoff"))
    vgm_txt, _ = parse_dt(kw.get("vgm_cutoff"))
    vessel = clean(kw.get("vessel")) or ""
    return {
        "source": source,
        "carrier": kw.get("carrier") or source,
        "service": clean(kw.get("service")),
        "vessel": vessel,
        "vessel_key": norm_vessel(vessel),
        "voyage": clean(kw.get("voyage")),
        "etd": etd_txt,
        "eta": eta_txt,
        "etd_date": etd_date.isoformat() if etd_date else None,
        "pol": clean(kw.get("pol")),
        "pod": clean(kw.get("pod")),
        "pol_terminal": clean(kw.get("pol_terminal")),
        "pod_terminal": clean(kw.get("pod_terminal")),
        "transit_days": to_float(kw.get("transit_days")),
        "direct_or_ts": clean(kw.get("direct_or_ts")) or "Direct",
        "doc_cutoff": doc_txt,
        "cy_cutoff": cy_txt,
        "vgm_cutoff": vgm_txt,
    }


def fetch_sitc(start: dt.date, end: dt.date) -> list[dict]:
    import sitc_schedule as m
    rows = http_retry(lambda: m.fetch_schedule(
        "LAEM CHABANG", "SHANGHAI", start.isoformat(), end.isoformat()))
    out = []
    for r in rows:
        out.append(_rec(
            "SITC", carrier="SITC",
            service=r.get("serviceLineCode"),
            vessel=r.get("vesselName"), voyage=r.get("voyageNo"),
            pol=r.get("polName"), pod=r.get("podName"),
            pol_terminal=r.get("polTerminalName"), pod_terminal=r.get("podTerminalName"),
            etd=r.get("weekEtd") or r.get("etd"),
            eta=r.get("weekEta") or r.get("eta"),
            transit_days=r.get("underway"),
        ))
    return out


def fetch_tslines(start: dt.date, end: dt.date) -> list[dict]:
    import tslines_schedule as m
    rows = http_retry(lambda: m.fetch_schedule(
        m.DEFAULT_POL, m.DEFAULT_POD, start.isoformat(), end.isoformat()))
    out = []
    for r in rows:
        out.append(_rec(
            "T.S. Lines", carrier="T.S. Lines",
            service=r.get("SERVICE"),
            vessel=r.get("VESSEL_NAME"), voyage=r.get("VOYAGE"),
            pol=r.get("POL"), pod=r.get("POD"),
            pol_terminal=r.get("TERMINAL_L"), pod_terminal=r.get("TERMINAL_D"),
            etd=r.get("ETD"), eta=r.get("ETA"),
            transit_days=r.get("SALILNG"),
            cy_cutoff=r.get("CUTOFF"), vgm_cutoff=r.get("VGM_CUT"),
        ))
    return out


def fetch_jj(start: dt.date, end: dt.date) -> list[dict]:
    import jj_shipping_schedule as m
    tmp = OUT_DIR / "_jj_raw.xlsx"
    df = m.run("Laem Chabang", "Shanghai", start, end, str(tmp), headless=True)
    out = []
    if df is None or df.empty:
        return out
    for _, r in df.iterrows():
        out.append(_rec(
            "SJJ", carrier="SJJ",
            service=r.get("tcf_code"),
            vessel=r.get("vessel"), voyage=r.get("voyage"),
            pol=r.get("departure_port") or "Laem Chabang",
            pod=r.get("arrival_port") or "Shanghai",
            etd=r.get("departure_date"), eta=r.get("arrival_date"),
            transit_days=r.get("transit_days"),
            doc_cutoff=r.get("closing"),
        ))
    return out


def _name_voy(text):
    """'JARU BHUM 2634N' -> ('JARU BHUM', '2634N'); ถ้าไม่มีเลขเที่ยวต่อท้ายคืน voyage=None"""
    if not text:
        return None, None
    parts = str(text).split()
    if len(parts) > 1 and re.match(r"^\d+[A-Za-z]?$", parts[-1]):
        return " ".join(parts[:-1]), parts[-1]
    return str(text), None


# --------------------------------------------------------------------------- #
#  KMTC + ZIM: อ่านจากไฟล์ที่ผู้ใช้ดาวน์โหลดเอง (เว็บ 2 รายนี้บล็อก IP นอกไทย)    #
#  - KMTC : ไฟล์ .xls รายเดือนในโฟลเดอร์ KMTC/  (export จากหน้า Leg Schedule)     #
#  - ZIM  : ไฟล์ zim_*.xlsx ที่ราก repo (ผลจาก input/zim_schedule_scraper.py)     #
# --------------------------------------------------------------------------- #
KMTC_DIR = HERE / "KMTC"
ZIM_GLOBS = ("zim_*.xlsx", "zim_*.xls", "input/zim_*.xlsx", "exports/zim_*.xlsx")


def fetch_kmtc(start: dt.date, end: dt.date) -> list[dict]:
    import kmtc_extract_text as kx

    if not KMTC_DIR.is_dir():
        raise SourceBlocked(
            f"ไม่พบโฟลเดอร์ {KMTC_DIR.name}/ — เว็บ KMTC บล็อก IP นอกไทย ให้เปิด "
            "ekmtc.com > Schedule > Leg Schedule (LCH→SHA) กด Excel รายเดือน "
            "แล้ววางไฟล์ .xls ไว้ในโฟลเดอร์ KMTC/ จากนั้นรันใหม่")

    files = kx.collect_files([str(KMTC_DIR)])
    if not files:
        raise SourceBlocked(f"โฟลเดอร์ {KMTC_DIR.name}/ ว่าง — ต้องมีไฟล์ .xls export จาก KMTC")

    frames = []
    for path in files:
        for raw in kx.read_workbook_raw(path).values():
            tidy = kx.flatten_kmtc_sheet(raw)
            if tidy is not None and not tidy.empty:
                frames.append(tidy)
    if not frames:
        raise RuntimeError(f"อ่านไฟล์ KMTC ได้ แต่ไม่พบตารางในรูปแบบที่รู้จัก ({len(files)} ไฟล์)")

    import pandas as pd
    combined = pd.concat(frames, ignore_index=True).drop_duplicates()

    out, seen = [], set()
    for _, r in combined.iterrows():
        vv = str(r.get("Vessel/Voyage") or "")
        vessel, voyage = (vv.split("/", 1) + [""])[:2] if "/" in vv else (vv, "")
        etd_txt, etd_d = parse_dt(r.get("Departure ETD"))
        if etd_d is None or not (start <= etd_d <= end):
            continue
        key = (norm_vessel(vessel), voyage.strip(), etd_d.isoformat())
        if key in seen:
            continue
        seen.add(key)
        ts_place = clean(r.get("T/S Place"))
        out.append(_rec(
            "KMTC", carrier="KMTC",
            vessel=vessel.strip(), voyage=voyage.strip(),
            pol="LAEM CHABANG", pod="SHANGHAI",
            pol_terminal=clean(r.get("Departure Terminal")),
            pod_terminal=clean(r.get("Arrival Terminal")),
            etd=r.get("Departure ETD"), eta=r.get("Arrival ETA"),
            transit_days=r.get("T/T"),
            direct_or_ts="Direct" if not ts_place or ts_place == "-" else "T/S",
            doc_cutoff=r.get("Document Closing Time"),
            cy_cutoff=r.get("Cargo Closing Time"),
        ))
    return out


def _first_file(globs) -> "Path | None":
    for pat in globs:
        hits = sorted(HERE.glob(pat))
        if hits:
            return hits[-1]
    return None


def _xlsx_rows(path: "Path") -> list[dict]:
    """อ่านชีตแรกของ .xlsx เป็น list[dict] โดยใช้แถวแรกเป็นหัวคอลัมน์"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    header = [str(c or "").strip() for c in rows[0]]
    out = []
    for row in rows[1:]:
        if not any(c is not None and str(c).strip() for c in row):
            continue
        out.append({header[i]: row[i] for i in range(min(len(header), len(row)))})
    return out


def fetch_zim(start: dt.date, end: dt.date) -> list[dict]:
    path = _first_file(ZIM_GLOBS)
    if path is None:
        raise SourceBlocked(
            "ไม่พบไฟล์ zim_*.xlsx — เว็บ ZIM บล็อก IP นอกไทย ให้รัน "
            "`python input/zim_schedule_scraper.py` จากเครือข่ายในไทย "
            "แล้ววางไฟล์ผลลัพธ์ (เช่น zim_laemchabang_shanghai_schedule.xlsx) ไว้ที่ราก repo")

    out = []
    for r in _xlsx_rows(path):
        _, etd_d = parse_dt(r.get("Departure"))
        if etd_d is None or not (start <= etd_d <= end):
            continue
        out.append(_rec(
            "ZIM", carrier="ZIM",
            service=clean(r.get("Vessel Code")),
            vessel=r.get("Vessel Name"), voyage=r.get("Voyage"),
            pol="LAEM CHABANG", pod="SHANGHAI",
            etd=r.get("Departure"), eta=r.get("Arrival"),
            transit_days=r.get("Transit Time (Days)"),
            direct_or_ts=clean(r.get("Transit Type")) or "Direct",
        ))
    return out


def fetch_cosco(start: dt.date, end: dt.date) -> list[dict]:
    import cosco_schedule as m
    import requests
    session = requests.Session()
    try:
        origin = http_retry(lambda: m.find_city(session, "Laem Chabang"))
        dest = http_retry(lambda: m.find_city(session, "Shanghai"))
        records = http_retry(lambda: m.fetch_schedule(session, origin, dest, start, end))
    except SystemExit as exc:  # find_city/fetch_schedule ยก SystemExit เมื่อไม่พบเมือง/มี error
        raise RuntimeError(f"COSCO: {exc}") from exc

    out, seen = [], set()
    for r in m.to_rows(records):
        svc, _, voy = str(r.get("Service/Voyage") or "").partition("/")
        _, d = parse_dt(r.get("ETD"))
        if d is None or not (start <= d <= end):
            continue
        key = (norm_vessel(r.get("Vessel")), voy.strip(), d.isoformat())
        if key in seen:       # COSCO ส่งหลายแถวต่อเที่ยว (แยกตามชนิดสินค้า/haulage)
            continue
        seen.add(key)
        out.append(_rec(
            "COSCO", carrier="COSCO",
            service=svc.strip() or None,
            vessel=r.get("Vessel"), voyage=voy.strip() or None,
            pol=r.get("POL") or "LAEM CHABANG", pod=r.get("POD") or "SHANGHAI",
            etd=r.get("ETD"), eta=r.get("ETA"),
            transit_days=r.get("Transit (days)"),
            cy_cutoff=r.get("Cut Off"),
        ))
    return out


def fetch_yml(start: dt.date, end: dt.date) -> list[dict]:
    import yml_schedule_lcb_shanghai as m
    search_start = max(start, dt.date.today())   # YML API ปฏิเสธ startDate ที่เป็นอดีต (HTTP 400)
    if search_start > end:
        return []
    rows = m.fetch_schedule("THLCB", "CNSHA", search_start, end)
    out = []
    for r in rows:
        _, d = parse_dt(r.get("masterETD"))
        if d is None or not (start <= d <= end):
            continue
        ts = str(r.get("transshipment") or "").strip()
        out.append(_rec(
            "Yang Ming", carrier="Yang Ming",
            service=clean(r.get("masterVoyageCode")),
            vessel=r.get("masterVesselName"),
            voyage=r.get("masterComnVoyage") or r.get("masterVoyageCode"),
            pol=r.get("placeOfReceipt") or "LAEM CHABANG",
            pod=r.get("placeOfDelivery") or "SHANGHAI",
            etd=r.get("masterETD"), eta=r.get("masterETA"),
            transit_days=r.get("transitDays"),
            direct_or_ts="T/S" if ts and ts.upper() not in {"N", "NO", "DIRECT", "0"} else "Direct",
            cy_cutoff=r.get("cutoffCY"), vgm_cutoff=r.get("cutoffVGM"),
        ))
    return out


def fetch_culines(start: dt.date, end: dt.date) -> list[dict]:
    import culines_ptp_schedule as m
    import requests
    session = requests.Session()
    origin = http_retry(lambda: m.get_port_code(session, "Laem Chabang"))
    dest = http_retry(lambda: m.get_port_code(session, "Shanghai"))
    raw = http_retry(lambda: m.fetch_schedule(
        session, origin["locCd"], dest["locCd"], start.isoformat(), end.isoformat(), ""))
    out = []
    for rec in raw:
        n = m.normalize_record(rec)
        name, voy = _name_voy(n.get("Vessel"))
        out.append(_rec(
            "CU Lines", carrier="CU Lines",
            service=n.get("Lane"),
            vessel=name, voyage=voy or n.get("Consortium/Voyage"),
            pol=n.get("Loading Port") or "LAEM CHABANG",
            pod=n.get("Discharging Port") or "SHANGHAI",
            etd=n.get("Departure Date"), eta=n.get("Arrival Date"),
            transit_days=n.get("T/Time Total (Day)"),
            direct_or_ts=n.get("Direct/T-S") or "Direct",
            cy_cutoff=n.get("Cargo Closing Time"),
        ))
    return out


SOURCES = {
    "sitc":    ("SITC",        fetch_sitc),
    "tslines": ("T.S. Lines",  fetch_tslines),
    "kmtc":    ("KMTC",        fetch_kmtc),
    "zim":     ("ZIM",         fetch_zim),
    "culines": ("CU Lines",    fetch_culines),
    "cosco":   ("COSCO",       fetch_cosco),
    "yml":     ("Yang Ming",   fetch_yml),
    "sjj":     ("SJJ",         fetch_jj),
}


# --------------------------------------------------------------------------- #
#  รวมเที่ยวเรือซ้ำ                                                            #
# --------------------------------------------------------------------------- #
def dedupe_records(records: list[dict]) -> list[dict]:
    """ตัดแถวที่ซ้ำกันเป๊ะ ๆ จากแหล่งเดียวกัน (บาง API ส่งข้อมูลซ้ำมาเอง)"""
    out, seen = [], set()
    for r in records:
        k = (r["source"], r["vessel_key"], r.get("voyage"), r.get("etd"), r.get("eta"))
        if k in seen:
            continue
        seen.add(k)
        out.append(r)
    return out


def merge_sailings(records: list[dict]) -> list[dict]:
    """
    ยุบ record ที่เป็น 'เรือลำเดียวกัน + ออกเรือวันเดียวกัน (คลาดเคลื่อน <= 1 วัน)'
    ให้เป็นเที่ยวเรือเดียว โดยเก็บรายละเอียดของแต่ละแหล่ง/สายเรือไว้ใน 'offers'
    """
    clusters: list[dict] = []

    def dpick(rec):
        _, d = parse_dt(rec.get("etd_date") or rec.get("etd"))
        return d

    for rec in sorted(records, key=lambda r: (r["vessel_key"], r.get("etd_date") or "9999")):
        d = dpick(rec)
        placed = False
        if rec["vessel_key"] and d:
            for c in clusters:
                if c["vessel_key"] != rec["vessel_key"]:
                    continue
                if c["_date"] and abs((c["_date"] - d).days) <= 1:
                    c["offers"].append(rec)
                    if d < c["_date"]:
                        c["_date"] = d
                    placed = True
                    break
        if not placed:
            clusters.append({
                "vessel_key": rec["vessel_key"],
                "vessel": rec["vessel"],
                "_date": d,
                "offers": [rec],
            })

    merged = []
    for c in clusters:
        offers = c["offers"]
        carriers = sorted({o["carrier"] for o in offers})
        transit_vals = [o["transit_days"] for o in offers if o["transit_days"]]
        etas = [o["eta"] for o in offers if o["eta"]]
        etds = [o["etd"] for o in offers if o["etd"]]
        is_direct = any((o.get("direct_or_ts") or "").lower().startswith("direct") for o in offers)
        merged.append({
            "vessel": c["vessel"] or (offers[0]["vessel"]),
            "etd_date": c["_date"].isoformat() if c["_date"] else None,
            "etd": min(etds) if etds else None,
            "eta": min(etas) if etas else None,
            "transit_days": (min(transit_vals), max(transit_vals)) if transit_vals else None,
            "direct_or_ts": "Direct" if is_direct else "Transshipment",
            "carriers": carriers,
            "n_sources": len(carriers),
            "n_offers": len(offers),
            "offers": sorted(offers, key=lambda o: o["carrier"]),
        })
    merged.sort(key=lambda x: (x["etd_date"] or "9999", x["vessel"]))
    return merged


# --------------------------------------------------------------------------- #
#  Export: Excel                                                              #
# --------------------------------------------------------------------------- #
def transit_text(t) -> str:
    if not t:
        return ""
    lo, hi = t
    if lo == hi:
        return f"{lo:g}"
    return f"{lo:g}\u2013{hi:g}"


def write_excel(merged: list[dict], records: list[dict], statuses: list[dict], path: Path):
    import pandas as pd

    merged_rows = []
    for mrec in merged:
        merged_rows.append({
            "ETD Date": mrec["etd_date"],
            "Vessel": mrec["vessel"],
            "ETD": mrec["etd"],
            "ETA": mrec["eta"],
            "Transit (days)": transit_text(mrec["transit_days"]),
            "Type": mrec["direct_or_ts"],
            "# Lines": mrec["n_sources"],
            "Bookable via": ", ".join(mrec["carriers"]),
            "Voyages": " | ".join(
                f"{o['carrier']}:{o['voyage'] or '-'}" for o in mrec["offers"]),
        })
    df_merged = pd.DataFrame(merged_rows)

    df_all = pd.DataFrame([{
        "Source": r["source"], "Carrier": r["carrier"], "Service": r["service"],
        "Vessel": r["vessel"], "Voyage": r["voyage"],
        "POL": r["pol"], "POL Terminal": r["pol_terminal"],
        "ETD": r["etd"], "ETA": r["eta"],
        "POD": r["pod"], "POD Terminal": r["pod_terminal"],
        "Transit (days)": r["transit_days"], "Type": r["direct_or_ts"],
        "Doc Cut-off": r["doc_cutoff"], "CY Cut-off": r["cy_cutoff"],
        "VGM Cut-off": r["vgm_cutoff"],
    } for r in sorted(records, key=lambda r: (r.get("etd_date") or "9999", r["carrier"]))])

    df_src = pd.DataFrame(statuses)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_merged.to_excel(writer, index=False, sheet_name="Merged", startrow=1)
        df_all.to_excel(writer, index=False, sheet_name="All Sailings", startrow=1)
        df_src.to_excel(writer, index=False, sheet_name="Sources")
        for name, df in (("Merged", df_merged), ("All Sailings", df_all)):
            sh = writer.sheets[name]
            sh["A1"] = (f"Sailing Schedule  {LANE}   "
                        f"(generated {dt.datetime.now():%Y-%m-%d %H:%M})")
            sh["A1"].font = sh["A1"].font.copy(bold=True, size=12)
            for col in sh.columns:
                width = max((len(str(c.value)) if c.value is not None else 0) for c in col)
                sh.column_dimensions[col[0].column_letter].width = min(max(width + 2, 12), 48)


# --------------------------------------------------------------------------- #
#  Export: ICS calendar                                                      #
# --------------------------------------------------------------------------- #
def write_ics(merged: list[dict], path: Path):
    def esc(s):
        return (str(s).replace("\\", "\\\\").replace(";", "\\;")
                .replace(",", "\\,").replace("\n", "\\n"))

    lines = ["BEGIN:VCALENDAR", "VERSION:2.0",
             "PRODID:-//shanghai service//build_shanghai_schedule//TH",
             "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
             "X-WR-CALNAME:LCH \u2192 Shanghai Sailing Schedule"]
    stamp = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    for i, mrec in enumerate(merged):
        if not mrec["etd_date"]:
            continue
        d = mrec["etd_date"].replace("-", "")
        _, eta_date = parse_dt(mrec["eta"])
        summary = f"\U0001F6A2 {mrec['vessel']} \u2192 Shanghai ({', '.join(mrec['carriers'])})"
        desc_parts = [
            f"Type: {mrec['direct_or_ts']}",
            f"Transit: {transit_text(mrec['transit_days'])} days",
            f"ETD: {mrec['etd']}   ETA: {mrec['eta']}",
            "Bookable via:",
        ]
        for o in mrec["offers"]:
            desc_parts.append(
                f"  - {o['carrier']} | Voy {o['voyage'] or '-'} | "
                f"svc {o['service'] or '-'} | ETD {o['etd'] or '-'} | ETA {o['eta'] or '-'}"
                + (f" | CY {o['cy_cutoff']}" if o['cy_cutoff'] else ""))
        lines += [
            "BEGIN:VEVENT",
            f"UID:lch-sha-{d}-{i}@shanghai-service",
            f"DTSTAMP:{stamp}",
            f"DTSTART;VALUE=DATE:{d}",
            f"DTEND;VALUE=DATE:{(eta_date + dt.timedelta(days=1)).strftime('%Y%m%d') if eta_date else d}",
            f"SUMMARY:{esc(summary)}",
            f"DESCRIPTION:{esc(chr(10).join(desc_parts))}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    path.write_text("\r\n".join(lines), encoding="utf-8")


# --------------------------------------------------------------------------- #
#  Export: HTML dashboard + calendar                                         #
# --------------------------------------------------------------------------- #
CSS = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#0f172a;--muted:#64748b;--line:#e2e8f0;--accent:#0ea5e9}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
     font-family:"Segoe UI",Roboto,"Helvetica Neue",Arial,"Noto Sans Thai",sans-serif;font-size:14px}
.wrap{max-width:1180px;margin:0 auto;padding:24px}
h1{font-size:22px;margin:0 0 4px}
h2{font-size:16px;margin:28px 0 12px;padding-bottom:6px;border-bottom:2px solid var(--line)}
.sub{color:var(--muted);margin:0 0 20px}
.grid{display:grid;gap:14px}
.kpis{grid-template-columns:repeat(auto-fit,minmax(150px,1fr))}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px}
.kpi .v{font-size:26px;font-weight:700}
.kpi .l{color:var(--muted);font-size:12px;margin-top:2px}
.srcgrid{grid-template-columns:repeat(auto-fit,minmax(200px,1fr))}
.src{display:flex;align-items:flex-start;gap:10px}
.dot{width:12px;height:12px;border-radius:50%;margin-top:3px;flex:none}
.src .name{font-weight:600}
.src .meta{color:var(--muted);font-size:12px}
.ok{color:#16a34a;font-weight:600}.bad{color:#dc2626;font-weight:600}.warn{color:#d97706;font-weight:600}
.blocked{color:#b45309;font-weight:600}
.bar{display:flex;align-items:center;gap:8px;margin:6px 0}
.bar .lab{width:150px;font-size:12px;color:var(--muted);text-align:right;flex:none}
.bar .track{flex:1;background:#eef2f7;border-radius:6px;overflow:hidden}
.bar .fill{height:16px;border-radius:6px}
.bar .num{width:34px;font-size:12px;font-weight:600;flex:none}
table{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);
      border-radius:12px;overflow:hidden}
th,td{padding:9px 10px;text-align:left;border-bottom:1px solid var(--line);font-size:13px;vertical-align:top}
th{background:#f1f5f9;font-size:12px;text-transform:uppercase;letter-spacing:.03em;color:var(--muted)}
tr.main{cursor:pointer}
tr.main:hover{background:#f8fafc}
tr.details{display:none;background:#fbfcfe}
tr.details.show{display:table-row}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;color:#fff;margin:1px 2px}
.pill{display:inline-block;padding:1px 7px;border-radius:6px;font-size:11px;background:#eef2f7;color:#334155;margin-right:4px}
.tag-direct{background:#dcfce7;color:#166534}.tag-ts{background:#fef9c3;color:#854d0e}
.caltabs{display:flex;flex-wrap:wrap;gap:6px;margin:4px 0 14px}
.caltab{border:1px solid var(--line);background:var(--card);border-radius:8px;padding:6px 12px;
        cursor:pointer;font:inherit;font-size:13px;font-weight:600;color:var(--muted)}
.caltab:hover{border-color:var(--accent);color:var(--ink)}
.caltab.active{background:var(--accent);color:#fff;border-color:var(--accent)}
.caltab .cnt{display:inline-block;margin-left:7px;background:rgba(0,0,0,.12);border-radius:999px;
             padding:0 7px;font-size:11px}
.caltab.active .cnt{background:rgba(255,255,255,.28)}
.calpanel[hidden]{display:none}
.cal{display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin-bottom:18px}
.cal .h{font-size:11px;color:var(--muted);text-align:center;padding:4px 0;font-weight:600}
.cal .cell{background:var(--card);border:1px solid var(--line);border-radius:8px;min-height:78px;padding:4px}
.cal .cell.dim{background:#f1f5f9}
.cal .d{font-size:11px;color:var(--muted);font-weight:600}
.cal .ev{font-size:10px;line-height:1.25;margin-top:2px;padding:1px 3px;border-radius:4px;
         color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.legend{display:flex;flex-wrap:wrap;gap:12px;margin:6px 0 14px}
.legend span{font-size:12px;color:var(--muted);display:flex;align-items:center;gap:5px}
.foot{color:var(--muted);font-size:12px;margin-top:30px;text-align:center}
"""

JS = """
document.querySelectorAll('tr.main').forEach(function(r){
  r.addEventListener('click',function(){
    var n=r.nextElementSibling;
    if(n&&n.classList.contains('details'))n.classList.toggle('show');
  });
});
document.querySelectorAll('.caltab').forEach(function(t){
  t.addEventListener('click',function(){
    document.querySelectorAll('.caltab').forEach(function(x){x.classList.remove('active');});
    document.querySelectorAll('.calpanel').forEach(function(x){x.hidden=true;});
    t.classList.add('active');
    var el=document.getElementById(t.dataset.t);
    if(el)el.hidden=false;
  });
});
"""


def _bars(counts: dict, colors: dict, keep_order: bool = False) -> str:
    if not counts:
        return "<p class='sub'>ไม่มีข้อมูล</p>"
    mx = max(counts.values()) or 1
    items = list(counts.items()) if keep_order else sorted(counts.items(), key=lambda kv: -kv[1])
    out = []
    for k, v in items:
        pct = v / mx * 100
        col = colors.get(k, DEFAULT_COLOR)
        out.append(
            f"<div class='bar'><div class='lab'>{html.escape(str(k))}</div>"
            f"<div class='track'><div class='fill' style='width:{pct:.0f}%;background:{col}'></div></div>"
            f"<div class='num'>{v}</div></div>")
    return "".join(out)


def _calendar(merged: list[dict], start: dt.date, end: dt.date) -> str:
    """ปฏิทินเดียว + ปุ่มเลือกเดือน (แสดงทีละเดือน)"""
    by_day: dict[str, list[dict]] = {}
    for m in merged:
        if m["etd_date"]:
            by_day.setdefault(m["etd_date"], []).append(m)

    def carrier_color(m):
        return SOURCE_META.get(m["carriers"][0], {}).get("color", DEFAULT_COLOR)

    dows = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    months: list[dt.date] = []
    cur = dt.date(start.year, start.month, 1)
    last = dt.date(end.year, end.month, 1)
    while cur <= last:
        months.append(cur)
        cur = dt.date(cur.year + (cur.month == 12), (cur.month % 12) + 1, 1)

    def month_count(mo: dt.date) -> int:
        pre = f"{mo:%Y-%m}"
        return sum(len(v) for k, v in by_day.items() if k.startswith(pre))

    # เดือนเริ่มต้น = เดือนแรกที่มีเที่ยวเรือ (ถ้าไม่มีเลย ใช้เดือนแรก)
    default_idx = next((i for i, mo in enumerate(months) if month_count(mo)), 0)

    tabs, panels = [], []
    for i, mo in enumerate(months):
        mid = f"cal-{mo:%Y-%m}"
        is_def = (i == default_idx)
        tabs.append(
            f"<button class='caltab{' active' if is_def else ''}' data-t='{mid}'>"
            f"{mo:%b %Y}<span class='cnt'>{month_count(mo)}</span></button>")

        first_dow = mo.weekday()
        nxt = dt.date(mo.year + (mo.month == 12), (mo.month % 12) + 1, 1)
        ndays = (nxt - mo).days
        cells = ["<div class='cell dim'></div>"] * first_dow
        for day in range(1, ndays + 1):
            iso = dt.date(mo.year, mo.month, day).isoformat()
            evs = by_day.get(iso, [])
            ev_html = "".join(
                f"<div class='ev' style='background:{carrier_color(e)}' "
                f"title='{html.escape(e['vessel'])} — {html.escape(', '.join(e['carriers']))} "
                f"| ETD {html.escape(e['etd'] or iso)}'>"
                f"{html.escape(e['vessel'][:18])}</div>"
                for e in evs)
            cells.append(f"<div class='cell'><div class='d'>{day}</div>{ev_html}</div>")
        head = "".join(f"<div class='h'>{d}</div>" for d in dows)
        panels.append(
            f"<div class='calpanel' id='{mid}'{'' if is_def else ' hidden'}>"
            f"<div class='cal'>{head}{''.join(cells)}</div></div>")

    return f"<div class='caltabs'>{''.join(tabs)}</div>{''.join(panels)}"


def write_html(merged, records, statuses, start, end, path: Path):
    total = len(merged)
    vessels = len({m["vessel"] for m in merged})
    carriers_all = sorted({c for m in merged for c in m["carriers"]})
    multi = sum(1 for m in merged if m["n_sources"] > 1)
    transit_all = [t for m in merged if m["transit_days"] for t in m["transit_days"]]
    avg_transit = f"{sum(transit_all) / len(transit_all):.1f}" if transit_all else "—"

    per_carrier: dict[str, int] = {}
    for m in merged:
        for c in m["carriers"]:
            per_carrier[c] = per_carrier.get(c, 0) + 1

    week_counts: dict[tuple, int] = {}
    for m in merged:
        if m["etd_date"]:
            d = dt.date.fromisoformat(m["etd_date"])
            monday = d - dt.timedelta(days=d.weekday())
            week_counts[(monday, d.isocalendar().week)] = week_counts.get((monday, d.isocalendar().week), 0) + 1
    per_week = {f"W{wk:02d} ({mon:%d %b})": c
               for (mon, wk), c in sorted(week_counts.items())}

    # source status cards
    src_cards = []
    for st in statuses:
        meta = SOURCE_META.get(st["Source"], {})
        col = meta.get("color", DEFAULT_COLOR)
        if st["Status"] == "ok":
            state = f"<span class='ok'>&#10003; ทำงานปกติ</span> &middot; {st['Sailings']} เที่ยว"
        elif st["Status"] == "empty":
            state = "<span class='warn'>&#9888; เชื่อมต่อได้ แต่ไม่พบเที่ยวเรือ</span>"
        elif st["Status"] == "blocked":
            state = "<span class='blocked'>&#9940; ถูกบล็อกการเข้าถึง (IP/ภูมิภาค)</span>"
        elif st["Status"] == "skipped":
            state = "<span class='warn'>&#8210; ข้ามไว้ (--skip)</span>"
        else:
            state = "<span class='bad'>&#10007; ดึงข้อมูลไม่สำเร็จ</span>"
        note = f"<div class='meta'>{html.escape(st.get('Detail') or '')}</div>" if st.get("Detail") else ""
        src_cards.append(
            f"<div class='card src'><div class='dot' style='background:{col}'></div>"
            f"<div><div class='name'>{html.escape(str(SOURCE_META.get(st['Source'],{}).get('label', st['Source'])))}</div>"
            f"<div class='meta'>{state}</div>{note}</div></div>")

    legend = "".join(
        f"<span><i class='dot' style='display:inline-block;background:{m['color']}'></i>{html.escape(m['label'])}</span>"
        for m in SOURCE_META.values())

    # detail table
    rows = []
    for i, m in enumerate(merged):
        badges = "".join(
            f"<span class='badge' style='background:{SOURCE_META.get(c,{}).get('color',DEFAULT_COLOR)}'>{html.escape(c)}</span>"
            for c in m["carriers"])
        tag = "tag-direct" if m["direct_or_ts"] == "Direct" else "tag-ts"
        sub = []
        for o in m["offers"]:
            sub.append(
                "<tr><td>" + html.escape(o["carrier"]) + "</td>"
                "<td>" + html.escape(o["voyage"] or "—") + "</td>"
                "<td>" + html.escape(o["service"] or "—") + "</td>"
                "<td>" + html.escape(o["etd"] or "—") + "</td>"
                "<td>" + html.escape(o["eta"] or "—") + "</td>"
                "<td>" + html.escape(f"{o['transit_days']:g}" if o["transit_days"] else "—") + "</td>"
                "<td>" + html.escape(o["pol_terminal"] or "—") + "</td>"
                "<td>" + html.escape(" / ".join(x for x in [o["doc_cutoff"], o["cy_cutoff"], o["vgm_cutoff"]] if x) or "—") + "</td></tr>")
        detail = (
            "<tr class='details'><td colspan='7'><table>"
            "<tr><th>สายเรือ/ผู้ให้บริการ</th><th>Voyage</th><th>Service</th><th>ETD</th>"
            "<th>ETA</th><th>Transit</th><th>ท่าต้นทาง</th><th>Cut-off (Doc/CY/VGM)</th></tr>"
            + "".join(sub) + "</table></td></tr>")
        rows.append(
            f"<tr class='main'><td>{html.escape(m['etd_date'] or '—')}</td>"
            f"<td><b>{html.escape(m['vessel'])}</b></td>"
            f"<td>{html.escape(m['etd'] or '—')}</td>"
            f"<td>{html.escape(m['eta'] or '—')}</td>"
            f"<td>{html.escape(transit_text(m['transit_days']) or '—')}</td>"
            f"<td><span class='pill {tag}'>{html.escape(m['direct_or_ts'])}</span></td>"
            f"<td>{badges}</td></tr>{detail}")

    doc = f"""<!doctype html><html lang="th"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ตารางเรือ Laem Chabang → Shanghai</title><style>{CSS}</style></head>
<body><div class="wrap">
<h1>&#128674; ตารางการเดินเรือ &mdash; {html.escape(LANE)}</h1>
<p class="sub">ช่วงข้อมูล {start:%d %b %Y} &ndash; {end:%d %b %Y}
&nbsp;|&nbsp; สร้างเมื่อ {dt.datetime.now():%Y-%m-%d %H:%M}
&nbsp;|&nbsp; รวมข้อมูลจาก {sum(1 for s in statuses if s['Status']=='ok')}/{len(statuses)} แหล่ง</p>

<h2>สถานะแหล่งข้อมูล</h2>
<div class="grid srcgrid">{''.join(src_cards)}</div>

<h2>ภาพรวม</h2>
<div class="grid kpis">
  <div class="card kpi"><div class="v">{total}</div><div class="l">เที่ยวเรือ (รวมซ้ำแล้ว)</div></div>
  <div class="card kpi"><div class="v">{len(records)}</div><div class="l">รายการดิบจากทุกแหล่ง</div></div>
  <div class="card kpi"><div class="v">{vessels}</div><div class="l">จำนวนเรือ (ลำ)</div></div>
  <div class="card kpi"><div class="v">{len(carriers_all)}</div><div class="l">สายเรือ/ผู้ให้บริการ</div></div>
  <div class="card kpi"><div class="v">{multi}</div><div class="l">เที่ยวที่จองได้ &gt;1 สาย</div></div>
  <div class="card kpi"><div class="v">{avg_transit}</div><div class="l">Transit เฉลี่ย (วัน)</div></div>
</div>

<h2>จำนวนเที่ยวเรือ แยกตามสายเรือ</h2>
<div class="card">{_bars(per_carrier, {k: v['color'] for k, v in SOURCE_META.items()})}</div>

<h2>จำนวนเที่ยวเรือ แยกตามสัปดาห์ (ตาม ETD)</h2>
<div class="card">{_bars(per_week, {}, keep_order=True)}</div>

<h2>ปฏิทินตารางเรือ (ตามวัน ETD)</h2>
<div class="legend">{legend}</div>
{_calendar(merged, start, end)}

<h2>รายละเอียดเที่ยวเรือ (คลิกแถวเพื่อดูรายสายเรือ)</h2>
<table>
<tr><th>วัน ETD</th><th>เรือ</th><th>ETD</th><th>ETA</th><th>Transit (วัน)</th><th>ประเภท</th><th>จองผ่าน</th></tr>
{''.join(rows)}
</table>

<p class="foot">ข้อมูลเพื่อการอ้างอิงเบื้องต้นเท่านั้น โปรดยืนยัน schedule / cut-off กับสายเรือหรือตัวแทนอีกครั้งก่อนใช้งานจริง<br>
สร้างโดย build_shanghai_schedule.py</p>
</div><script>{JS}</script></body></html>"""
    path.write_text(doc, encoding="utf-8")


# --------------------------------------------------------------------------- #
#  main                                                                      #
# --------------------------------------------------------------------------- #
def main():
    today = dt.date.today()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--start", default=today.replace(day=1).isoformat(),
                    help="วันเริ่ม YYYY-MM-DD (ค่าเริ่มต้น: วันที่ 1 ของเดือนปัจจุบัน)")
    ap.add_argument("--end", default=dt.date(today.year, 12, 31).isoformat(),
                    help="วันสิ้นสุด YYYY-MM-DD (ค่าเริ่มต้น: 31 ธ.ค. ปีปัจจุบัน)")
    ap.add_argument("--only", default="", help="ใช้เฉพาะบางแหล่ง เช่น --only sitc,tslines")
    ap.add_argument("--skip", default="", help="ข้ามบางแหล่ง เช่น --skip sjj,zim")
    ap.add_argument("--offline", action="store_true",
                    help="ไม่ยิงเน็ต ใช้ข้อมูลจาก output/_cache.json ที่เคยรันไว้")
    args = ap.parse_args()

    start = dt.date.fromisoformat(args.start)
    end = dt.date.fromisoformat(args.end)
    if end < start:
        sys.exit("--end ต้องไม่มาก่อน --start")

    OUT_DIR.mkdir(exist_ok=True)
    if not INPUT_DIR.is_dir():
        sys.exit(f"ไม่พบโฟลเดอร์ {INPUT_DIR} (ต้องมีสคริปต์ scraper ต้นฉบับอยู่ข้างใน)")
    sys.path.insert(0, str(INPUT_DIR))
    sys.path.insert(0, str(HERE))

    only = {s.strip().lower() for s in args.only.split(",") if s.strip()}
    skip = {s.strip().lower() for s in args.skip.split(",") if s.strip()}

    records: list[dict] = []
    statuses: list[dict] = []

    if args.offline:
        if not CACHE_FILE.exists():
            sys.exit(f"ไม่พบไฟล์แคช {CACHE_FILE} — ต้องรันแบบออนไลน์อย่างน้อย 1 ครั้งก่อน")
        cached = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        records = cached["records"]
        statuses = cached["statuses"]
        start = dt.date.fromisoformat(cached.get("start", args.start))
        end = dt.date.fromisoformat(cached.get("end", args.end))
        print(f"[offline] โหลดจากแคช: {len(records)} รายการ")
    else:
        for key, (label, fn) in SOURCES.items():
            if only and key not in only:
                continue
            if key in skip:
                statuses.append({"Source": label, "Status": "skipped", "Sailings": 0, "Detail": "ถูกข้ามด้วย --skip"})
                continue
            print(f"[{label}] กำลังดึงข้อมูล {start} -> {end} ...")
            try:
                recs = fn(start, end)
                for r in recs:
                    records.append(r)
                st = "ok" if recs else "empty"
                statuses.append({"Source": label, "Status": st, "Sailings": len(recs), "Detail": ""})
                print(f"   -> {len(recs)} รายการ")
            except SourceBlocked as exc:
                short = str(exc).splitlines()[0][:200]
                statuses.append({"Source": label, "Status": "blocked", "Sailings": 0, "Detail": short})
                print(f"   !! ถูกบล็อก: {short}")
            except Exception as exc:  # noqa: BLE001 - แหล่งเดียวล่ม ไม่ควรล้มทั้งงาน
                short = str(exc).splitlines()[0][:180]
                statuses.append({"Source": label, "Status": "failed", "Sailings": 0, "Detail": short})
                print(f"   !! ล้มเหลว: {short}")

    if not records:
        print("\nไม่ได้ข้อมูลจากแหล่งใดเลย — ยังสร้าง Dashboard ให้ (จะว่าง) เพื่อดูสถานะแหล่งข้อมูล")

    records = dedupe_records(records)
    merged = merge_sailings(records)

    xlsx = OUT_DIR / "combined_schedule.xlsx"
    htmlf = OUT_DIR / "shanghai_schedule_dashboard.html"
    ics = OUT_DIR / "shanghai_schedule.ics"

    write_excel(merged, records, statuses, xlsx)
    write_ics(merged, ics)
    write_html(merged, records, statuses, start, end, htmlf)
    CACHE_FILE.write_text(json.dumps(
        {"start": start.isoformat(), "end": end.isoformat(),
         "records": records, "statuses": statuses,
         "generated": dt.datetime.now().isoformat()},
        ensure_ascii=False, indent=1), encoding="utf-8")

    ok = sum(1 for s in statuses if s["Status"] == "ok")
    print(f"\n=== สรุป ===")
    print(f"แหล่งข้อมูลที่ใช้ได้ : {ok}/{len(statuses)}")
    for s in statuses:
        print(f"  - {s['Source']:<12} {s['Status']:<8} {s['Sailings']} เที่ยว  {s.get('Detail') or ''}")
    print(f"รายการดิบรวม        : {len(records)}")
    print(f"เที่ยวเรือหลังรวมซ้ำ : {len(merged)}")
    print(f"\nไฟล์ผลลัพธ์:")
    print(f"  {xlsx}")
    print(f"  {htmlf}")
    print(f"  {ics}")


if __name__ == "__main__":
    main()
