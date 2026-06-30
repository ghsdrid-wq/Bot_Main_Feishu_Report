# STRUCTURE — Bot_Main_Feishu_Report

> ⚠️ **กฎการดูแลไฟล์นี้ (สำคัญ)**
> ทุกครั้งที่แก้ไขโค้ดใน repo นี้ — เพิ่ม/ลบ/ย้ายไฟล์, เปลี่ยน logic ฟังก์ชัน/คลาส, เปลี่ยนโครงสร้าง config, เปลี่ยน flow refresh/export/ส่ง — **ต้องอัปเดต STRUCTURE.md นี้ให้ตรงกับโค้ดเสมอ**

## ภาพรวม
เป็น **เวอร์ชันต่อยอดจาก `Bot-Feishu` (Workbook Manager)** — GUI CustomTkinter ที่ Excel → PNG → ส่ง Feishu ตามรอบเวลา โครงสร้างและ flow เหมือนกัน ต่างกันที่รายละเอียดการ build/แคปรูปและการส่ง

## วิธีรัน / Entry point
- รัน: `python bot_main.py` → คลาส `App(ctk.CTk)`
- มี `BOT_Main_Report.spec` สำหรับ build เป็น .exe ด้วย PyInstaller

## โครงสร้างไฟล์
| ไฟล์ | หน้าที่ |
|------|---------|
| `bot_main.py` | UI + scheduler + จัดการ config (โครงเดียวกับ `Bot-Feishu/bot_main.py`; เพิ่ม import `json`, `mimetypes`) |
| `Createphoto.py` | Excel → PNG ด้วย win32com (เวอร์ชันนี้เพิ่ม `win32gui/win32con/win32process` สำหรับจัดการหน้าต่าง Excel ระหว่างแคปรูป) |
| `Botmessage.py` | ส่ง Feishu (เวอร์ชันนี้ตัดการใช้ HMAC sign บางส่วน — ต่างจาก `Bot-Feishu`) |
| `BOT_Main_Report.spec` | สเปก PyInstaller |

## ความต่างจาก `Bot-Feishu` (อ้างอิงเทียบโค้ด)
- `Createphoto.py`: เพิ่มการเรียก Win32 GUI (จัดการ/ซ่อนหน้าต่าง Excel ตอน CopyPicture)
- `Botmessage.py`: เปลี่ยนวิธี auth/ส่ง (ตัด `hmac/base64/hashlib` ออกจาก import)
> รายละเอียด config sections / โมเดล Workbook↔Export ให้ยึดตาม `Bot-Feishu/STRUCTURE.md` เป็นหลัก แล้วบันทึกจุดที่ต่างไว้ที่นี่เมื่อแก้

## Dependencies / บริการภายนอก
- `customtkinter`, `pywin32` (ต้องมี Microsoft Excel), `requests`
- Feishu OpenAPI

## ข้อควรระวัง
- ต้องรันบน Windows + Excel
- ถ้าแก้ logic ร่วมกับ `Bot-Feishu` ให้ระวังว่าทั้งสอง repo มีไฟล์ชื่อเดียวกันแต่ **เนื้อในต่างกัน** — อย่าก๊อปทับโดยไม่ diff
