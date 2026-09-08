# Shanghai Service – Sailing Schedule Tools

เครื่องมือดึงและรวมตารางการเดินเรือเส้นทาง **Laem Chabang → Shanghai** จากหลายสายเรือ

## โครงสร้างโปรเจกต์

```
build_shanghai_schedule.py   ตัวรวมข้อมูล + สร้าง Dashboard
input/                       สคริปต์ scraper ต้นฉบับรายสายเรือ
output/                      ผลลัพธ์ที่สร้างขึ้น (ไม่ commit เข้า git)
```

## สคริปต์ดึงข้อมูลรายสายเรือ (`input/`)

| ไฟล์ | แหล่งข้อมูล | วิธีดึง |
|---|---|---|
| `input/sitc_schedule.py` | SITC | public JSON API |
| `input/tslines_schedule.py` | T.S. Lines | public JSON API |
| `input/culines_ptp_schedule.py` | CU Lines | public JSON API |
| `input/jj_shipping_schedule.py` | JJ Shipping (NVOCC) | เว็บ + Playwright |
| `input/kmtc_extract_text.py` | KMTC | อ่านไฟล์ `.xls` ในโฟลเดอร์ `KMTC/` |
| `input/extract_schedule_text.py` | ZIM | อ่านไฟล์ `zim_*.xlsx` ที่ราก repo |

> **KMTC / ZIM:** เว็บสองรายนี้ใช้ Akamai กันบอต IP นอกไทยจะโดนบล็อกทั้งโดเมน
> จึงใช้วิธี *ดาวน์โหลดไฟล์เอง*:
> - **KMTC** — เปิด ekmtc.com → Schedule → Leg Schedule (LCH→SHA) กดปุ่ม Excel
>   ทีละเดือน วางไฟล์ `.xls` ทั้งหมดไว้ในโฟลเดอร์ `KMTC/`
>   (หรือรัน `python input/kmtc_schedule_export.py` จากเครือข่ายในไทย)
> - **ZIM** — รัน `python input/zim_schedule_scraper.py` จากเครือข่ายในไทย แล้ววาง
>   `zim_laemchabang_shanghai_schedule.xlsx` ไว้ที่ราก repo
>
> `build_shanghai_schedule.py` จะอ่านไฟล์เหล่านี้เข้ามารวมเอง ถ้าไม่มีไฟล์จะขึ้น
> สถานะ "ถูกบล็อก" พร้อมวิธีแก้บน Dashboard

แต่ละไฟล์รันเดี่ยว ๆ ได้ และ export เป็น Excel เช่น

```bash
python input/sitc_schedule.py --date-from 2026-09-01 --date-to 2026-12-31
```

## ตัวรวมข้อมูล + Dashboard

`build_shanghai_schedule.py` เรียกสคริปต์ทุกตัว รวมข้อมูลเป็นชุดเดียว รวมเที่ยวเรือ
ที่ซ้ำกัน (เรือลำเดียวกัน + ออกเรือวันเดียวกัน) แล้วแจกแจงว่าจองผ่านสายเรือใดได้บ้าง
พร้อมสร้างปฏิทิน (เลือกดูทีละเดือน) และ Dashboard HTML

```bash
python build_shanghai_schedule.py
python build_shanghai_schedule.py --start 2026-09-01 --end 2026-12-31
python build_shanghai_schedule.py --only sitc,tslines
python build_shanghai_schedule.py --offline      # ใช้ข้อมูลจาก output/_cache.json
```

ผลลัพธ์ (โฟลเดอร์ `output/`, ไม่ commit เข้า git):

- `shanghai_schedule_dashboard.html` – ปฏิทิน + Dashboard
- `combined_schedule.xlsx` – 3 ชีต: Merged / All Sailings / Sources
- `shanghai_schedule.ics` – ไฟล์ปฏิทินนำเข้า Google/Outlook

## ติดตั้ง

```bash
pip install -r requirements.txt
playwright install chromium
```
