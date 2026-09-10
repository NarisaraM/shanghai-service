# Shanghai Service – Sailing Schedule Tools

เครื่องมือดึงและรวมตารางการเดินเรือเส้นทาง **Laem Chabang → Shanghai** จากหลายสายเรือ

**เว็บ (GitHub Pages):** <https://narisaram.github.io/shanghai-service/>

## โครงสร้างโปรเจกต์

```
build_shanghai_schedule.py   ตัวรวมข้อมูล + สร้าง Dashboard
input/                       สคริปต์ scraper ต้นฉบับรายสายเรือ
output/                      ผลลัพธ์ที่สร้างขึ้น (ไม่ commit เข้า git)
docs/                        หน้าเว็บที่ GitHub Pages เผยแพร่
  index.html                 หน้า "Schedule Service" รวมลิงก์บริการ
  laem-chabang-shanghai/     Dashboard ของเส้นทางนี้ (สร้างใหม่ทุกครั้งที่รันสคริปต์)
```

## สคริปต์ดึงข้อมูลรายสายเรือ (`input/`)

| ไฟล์ | แหล่งข้อมูล | วิธีดึง |
|---|---|---|
| `input/sitc_schedule.py` | SITC | public JSON API (ดึงสด) |
| `input/tslines_schedule.py` | T.S. Lines | public JSON API (ดึงสด) |
| `input/culines_ptp_schedule.py` | CU Lines | public JSON API (ดึงสด) |
| `input/cosco_schedule.py` | COSCO | public JSON API (ดึงสด) |
| `input/yml_schedule_lcb_shanghai.py` | Yang Ming (YML) | public JSON API (ดึงสด) |
| `input/jj_shipping_schedule.py` | SJJ (NVOCC JJ Shipping) | เว็บ + Playwright (ดึงสด) |
| `input/kmtc_extract_text.py` + `KMTC/*.xls` | KMTC | อ่านไฟล์ที่ดาวน์โหลดเอง |
| `input/extract_schedule_text.py` + `zim_*.xlsx` | ZIM | อ่านไฟล์ที่ดาวน์โหลดเอง |
| `input/hal_schedule_export.py` + `HAL Schedule*.xlsx` | HAL (Heung-A) | อ่านไฟล์ที่ดาวน์โหลดเอง |

> **KMTC / ZIM / HAL:** เว็บกลุ่มนี้กันบอตที่ระดับ IP — IP นอกไทยจะโดนบล็อก
> จึงต้อง *ดาวน์โหลดไฟล์เอง* จากเครือข่ายในไทย แล้ววางไว้ที่ราก repo:
> - **KMTC** — เปิด ekmtc.com → Leg Schedule (LCH→SHA) กด Excel ทีละเดือน วางใน `KMTC/`
>   (หรือรัน `python input/kmtc_schedule_export.py`)
> - **ZIM** — `python input/zim_schedule_scraper.py` → วาง `zim_*.xlsx`
> - **HAL** — `python input/hal_schedule_export.py --no-headless` → วางไฟล์ `HAL Schedule*.xlsx` รายเดือน
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

## GitHub Pages

เว็บเผยแพร่จาก branch `main` โฟลเดอร์ `/docs`

- **อัปเดต Dashboard** — รัน `python build_shanghai_schedule.py` (เขียนทับ
  `docs/laem-chabang-shanghai/index.html` ให้เอง) แล้ว commit + push `docs/`
- **เพิ่มบริการใหม่** — สร้างโฟลเดอร์ `docs/<ชื่อบริการ>/` พร้อม `index.html`
  ข้างใน แล้วเพิ่ม `<li>` ในรายการของ `docs/index.html` (มีคอมเมนต์บอกวิธีไว้)
  โดยตั้งชื่อลิงก์เป็นชื่อ Service นั้น ๆ

## ให้คนอื่นส่งข้อมูลเที่ยวเรือเข้ามา

ทั้งหน้า `docs/index.html` และ Dashboard มีปุ่ม **"ส่งข้อมูลเที่ยวเรือเข้ามา"**
ที่ลิงก์ไปยัง GitHub Issue form (`.github/ISSUE_TEMPLATE/new-sailing.yml`)

- ผู้ส่งต้องมีบัญชี GitHub (ล็อกอินฟรี) แล้วกรอกฟอร์ม: สายเรือ / ชื่อเรือ / voyage /
  ETD Laem Chabang / ETA Shanghai / ที่มาข้อมูล
- ทุกครั้งที่ส่ง จะเปิดเป็น Issue ใหม่ใน repo ติด label `sailing-submission`
  ดูรวมได้ที่ <https://github.com/NarisaraM/shanghai-service/issues?q=label%3Asailing-submission>
- นำข้อมูลที่ยืนยันแล้วไปใส่ในสคริปต์ scraper / ไฟล์ที่เกี่ยวข้อง แล้วรัน
  `build_shanghai_schedule.py` ใหม่เพื่ออัปเดต Dashboard
- แก้ฟิลด์ในฟอร์มได้ที่ `.github/ISSUE_TEMPLATE/new-sailing.yml` (ต้องอยู่บน `main`
  ถึงจะมีผล) — ถ้าเปลี่ยนชื่อ label ให้แก้ทั้งไฟล์นี้และปุ่มลิงก์ใน `docs/`
