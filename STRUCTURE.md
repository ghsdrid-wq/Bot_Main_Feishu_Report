# STRUCTURE — Bot_Main_Feishu_Report (Auto Report Feishu — Enterprise Console v10.0)

> ⚠️ **กฎการดูแลไฟล์นี้ (สำคัญ)**
> ทุกครั้งที่แก้ไขโค้ดใน repo นี้ — เพิ่ม/ลบ/ย้ายไฟล์, เปลี่ยน logic ฟังก์ชัน/คลาส, เปลี่ยนโครงสร้าง config section, เพิ่ม/แก้ขั้นใน pipeline, หรือเปลี่ยน endpoint DWS/JMS/Feishu — **ต้องอัปเดต STRUCTURE.md นี้ให้ตรงกับโค้ดเสมอ**

## ภาพรวม
แอป GUI (**CustomTkinter**) ชื่อ *"Auto Report Feishu Enterprise Console v10.0"* — **รวมทั้ง pipeline งานรายงานไว้ในแอปเดียว**:
1. **Bot Export (Raw DWS/JMS)** — ดึงข้อมูลจากระบบ DWS local + JMS J&T ออกเป็นไฟล์ Excel (DWS9-11, JMS Auto建包扫描, JMS PDA卸车扫描, Realtime DB)
2. **Bot Chat (Excel → รูป → Feishu)** — เปิด Excel ที่จัดไว้ (Workbook Manager) แคปรูป PNG แล้วส่งเข้า **Feishu Chat ตาม Chat ID** (+แนบไฟล์ Excel ได้)

ถ้าเลือกทั้งคู่: รัน **Export → Chat** ตามลำดับ มีหน้า Dashboard แสดง pipeline 6 ขั้น (DWS → JMS Auto → JMS PDA → Realtime → Excel Image → Feishu)

> เดิม repo นี้เป็นแค่ Workbook Manager (Excel→รูป→ส่ง Feishu ผ่าน webhook) — **v10.0 เปลี่ยนเป็น Enterprise Console** ที่รวม raw export เข้ามา และเปลี่ยน Feishu เป็น **Chat ID mode** (เลิกใช้ webhook/secret)

## วิธีรัน / Entry point
- รัน: `python bot_main.py` → คลาส `App(ctk.CTk)` (หน้าต่างคงที่ 1280×900)
- 5 หน้า (sidebar): **หน้าหลัก** (home) / **ไฟล์ Excel** (workbooks) / **ส่งออกข้อมูล** (dws_jms) / **จัดการไฟล์** (output_manager) / **ตั้งค่า** (settings)

## โครงสร้างไฟล์
| ไฟล์ | หน้าที่ |
|------|---------|
| `bot_main.py` | **ตัวหลัก (~2,600 บรรทัด)** — UI 5 หน้า, scheduler/worker/watchdog, raw export (`run_dws`, `run_jms_auto`, `run_jms_pda`, `run_realtime_db`, `_export_jms`), `run_process` (pipeline Export→Chat), ส่งไฟล์ Excel เข้า Feishu (`send_selected_excel_files_to_feishu`), การ์ด Workbook/Sheet |
| `Createphoto.py` | **Excel → PNG** ด้วย win32com (Workbook Manager: `WorkbookItem`/`ExportItem`, `run_create()`, `migrate_old_export_config`) + ซ่อนหน้าต่าง Excel (`hide_excel_from_taskbar`, `_move_excel_offscreen`, `_pump_excel_messages`) |
| `Botmessage.py` | **ส่ง Feishu แบบ Chat ID** — `get_token`, `upload_image`, `send_image_to_chat` (`receive_id_type=chat_id`), `run_send()` อัปโหลดขนาน (ตัด webhook/HMAC ออกแล้ว) |
| `BOT_Main_Report.spec` | สเปก PyInstaller |
| `config.ini` | **ไม่ถูก track ใน git** (มี token) — สร้าง/แก้ผ่าน UI |

## Pipeline (`App.run_process`)
1. ถ้า **bot_export** เปิด → วน raw steps: `run_dws` → `run_jms_auto` (建包扫描) → `run_jms_pda` (卸车扫描) → `run_realtime_db` (สร้าง export job บน `jmsgw.jtexpress.co.th` → poll → download)
2. ถ้า **bot_chat** เปิด → `run_create()` (Excel→PNG) → `run_send()` (ส่งรูปเข้า chat) → `send_selected_excel_files_to_feishu()` (แนบไฟล์ Excel ตาม checkbox)
- `scheduler_loop()` รันตาม `run_minute` + ช่วง `start_hour`/`end_hour` (อิง business date)

## โครงสร้าง config (`config.ini`)
- `[PATH]` `output_dir` (โฟลเดอร์เก็บรูป)
- `[FEISHU]` `APP_ID`, `APP_SECRET`, `CHAT_ID` *(เลิกใช้ `WEBHOOK`/`SECRET` — ถูกลบอัตโนมัติใน `ensure_config`)*
- `[TIME]` `run_minute`, `start_hour`, `end_hour`
- `[DWS_JMS]` `dws_url`, `dws_token`, `jms_token`, `name_dws`/`name_auto`/`name_dwspda`/`name_realtime_db`, `raw_path`, `download_size`, `start_date`/`end_date`/`start_hour`/`end_hour`, `enabled`, `send_*_file` (แนบไฟล์ Excel เข้า Feishu)
- `[WORKBOOKS]`/`[WORKBOOK:<key>]` + `[EXPORTS]`/`[EXPORT:<n>]` — โมเดล Workbook↔Export (เหมือน `Bot-Feishu`: 1 Excel → หลาย Sheet/Range/PNG)

## Dependencies / บริการภายนอก
- `customtkinter`, `tkcalendar`, `pandas`, `pywin32` (win32com/win32gui — **ต้องมี Microsoft Excel**), `requests`
- DWS local API, JMS J&T (`jmsgw.jtexpress.co.th` — `operatingplatform` 建包/卸车 + realtime export), Feishu OpenAPI (Chat ID mode)

## ข้อควรระวัง
- ต้องรันบน Windows + Excel
- token DWS/JMS หมดอายุได้ → แก้ในหน้า "ส่งออกข้อมูล"; ต้องตั้ง `CHAT_ID` ในหน้า "ตั้งค่า"
- ไฟล์ `bot_main.py`/`Createphoto.py`/`Botmessage.py` ของ repo นี้ **ต่างจาก `Bot-Feishu` อย่างมาก** (คนละ generation) — อย่า merge ข้าม repo โดยไม่ diff
