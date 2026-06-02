# Auto Report Feishu Enterprise Console v10.0

## Overview

Auto Report Feishu Enterprise Console v10.0 คือระบบ Automation สำหรับงาน Warehouse Operation และ Dashboard Monitoring ที่รวมระบบเดิมหลายตัวเข้าด้วยกันภายในโปรแกรมเดียว

โปรเจกต์นี้เกิดจากการรวมความสามารถของ:

- Bot Export DWS 9-11
- Bot Export JMS
- Bot Export PDA
- Bot Export Realtime DB
- Dashboard Image Generator
- Feishu Chat Sender
- Scheduler System
- Workbook Manager

ให้กลายเป็น Enterprise Console ที่สามารถควบคุมทุก Workflow ได้จากหน้าจอเดียว

---

# Core Objectives

ระบบถูกออกแบบมาเพื่อ

- ลดงาน Manual Export Report
- ลดเวลา Refresh Dashboard
- ลดขั้นตอนการส่งรายงานเข้ากลุ่ม Feishu
- รวมหลายระบบไว้ใน Application เดียว
- เพิ่มความเสถียรของงาน Scheduled Report
- รองรับการขยายจำนวน Dashboard ในอนาคต

---

# Main Features

## Report Export

- Export DWS Report
- Export JMS Report
- Export PDA Report
- Export Realtime DB Report
- Auto Download Source Data
- Configurable Date Range
- Configurable Business Hour

## Dashboard Processing

- Excel Auto Refresh
- Business Date Update
- Dynamic Column Delete
- Dynamic Range Shift
- Dashboard Rendering
- PNG Export

## Feishu Integration

- Tenant Access Token
- Chat ID Messaging
- Image Upload API
- Multi Image Broadcast
- Parallel Upload Worker
- Retry System

## Scheduler

- Run Now
- Auto Scheduler
- Cross-Day Schedule
- Stop Auto
- Stop Running Task

## User Interface

- Enterprise Console UI
- Workbook Manager
- Export Manager
- Real-Time Log Viewer
- Progress Tracking
- Runtime Lock Protection
- Watchdog Recovery

---

# System Architecture

```text
                   Scheduler
                       │
                       ▼
                 Task Queue
                       │
        ┌──────────────┼──────────────┐
        ▼                             ▼
   DWS/JMS Export              Workbook Manager
        │                             │
        ▼                             ▼
 Download Excel                Refresh Workbook
        │                             │
        ▼                             ▼
 Generate Source Data         Export Dashboard PNG
        └──────────────┬──────────────┘
                       ▼
                 Feishu Sender
                       ▼
                Group Chat
```

---

# Project Structure

```text
project/
│
├── bot_main.py
├── Createphoto.py
├── Botmessage.py
├── config.ini
│
├── output/
│   ├── AUTOREALTIME.png
│   ├── DWSREALTIME.png
│   ├── AUTO_PDA.png
│   ├── DWS_PDA.png
│   └── REALTIME_DB.png
│
└── logs/
```

---

# Module Breakdown

## bot_main.py

Main Application Controller

Responsibilities:

- Enterprise GUI
- Scheduler Engine
- Queue Manager
- Workbook Manager
- Runtime Lock System
- Watchdog Thread
- Progress Tracking
- Configuration Management

---

## Createphoto.py

Dashboard Processing Engine

Responsibilities:

- Excel COM Automation
- Workbook Refresh
- Business Date Processing
- Dynamic Delete Logic
- Dynamic Range Shift
- Dashboard Rendering
- PNG Export

---

## Botmessage.py

Feishu Communication Layer

Responsibilities:

- Get Tenant Access Token
- Upload Images
- Send Chat Message
- Retry Requests
- Parallel Upload Worker

---

# Workbook Manager

Version 10 ใช้ Dynamic Workbook Architecture

รองรับ:

- เพิ่ม Workbook หลายไฟล์
- เพิ่ม Export ได้ไม่จำกัด
- เปิด/ปิด Workbook แยกได้
- เปิด/ปิด Export แยกได้
- เปิด/ปิด Feishu Send แยกได้

ตัวอย่าง

```text
Workbook A
├── Export 1
├── Export 2
└── Export 3

Workbook B
├── Export 4
└── Export 5
```

---

# Export Item Configuration

แต่ละ Export สามารถกำหนดได้

| Setting | Description |
|----------|-------------|
| Use | เปิด/ปิด Export |
| Send | ส่งเข้า Feishu |
| Delete | ใช้ Dynamic Delete |
| Sheet | ชื่อ Worksheet |
| Range | Export Range |
| File | Output PNG |

---

# Business Date Engine

ระบบใช้ Business Date แทน Calendar Date

ตัวอย่าง

```text
Start Hour = 15

Current Time = 08:00

Business Date = Yesterday
```

ช่วยให้ Dashboard กลางคืนและรายงานข้ามวันแสดงผลได้ถูกต้อง

---

# Dynamic Dashboard Logic

Workflow

```text
Business Hour
      ↓
Delete Old Columns
      ↓
Shift Export Range
      ↓
Refresh Dashboard
      ↓
Render Image
      ↓
Export PNG
```

---

# Excel Processing Engine

ระบบใช้

```python
win32.DispatchEx("Excel.Application")
```

สำหรับ

- Open Workbook
- Refresh Query
- Calculate Workbook
- Copy Picture
- Export Chart
- Generate PNG

Functions ที่ใช้งานหลัก

```python
RefreshAll()
CalculateFull()
CopyPicture()
Chart.Export()
```

---

# Blank Image Protection

มีระบบป้องกันภาพขาวจาก Excel

ประกอบด้วย

- Clipboard Pump
- Export Validation
- Retry Export
- Render Verification

ช่วยลดปัญหา PNG ว่างจาก Excel COM

---

# Feishu Workflow

```text
Generate PNG
      ↓
Upload Image
      ↓
Receive image_key
      ↓
Send Chat Message
      ↓
Complete
```

รองรับส่งหลายภาพภายในรอบเดียว

---

# Parallel Upload System

ใช้

```python
ThreadPoolExecutor(max_workers=3)
```

เพื่อ

- Upload หลายไฟล์พร้อมกัน
- ลดเวลารอ
- เพิ่ม Throughput

---

# Retry System

รองรับ Retry สำหรับ

- Get Token
- Upload Image
- Send Message
- Download Request
- API Request

พร้อม Exponential Backoff

---

# Scheduler System

รองรับ

- Auto Run
- Run Now
- Cross-Day Schedule
- Business Hour Window
- Stop Scheduler
- Stop Current Task

ตัวอย่าง

```text
Start Hour = 15
End Hour = 12
Run Minute = 5
```

ระบบจะทำงานทุกชั่วโมงภายในช่วงเวลาที่กำหนด

---

# Queue System

ใช้

```python
queue.Queue()
```

สำหรับ

- Sequential Processing
- Prevent Duplicate Run
- Background Task Execution

---

# Runtime Lock System

ระหว่างระบบกำลังทำงาน

จะทำการ Lock

- Workbook Settings
- Export Settings
- Scheduler Settings
- Feishu Settings

ป้องกัน

- Config Corruption
- Runtime Conflict
- Invalid State

---

# Watchdog Recovery

Watchdog Thread ทำหน้าที่

- ตรวจจับ Process ค้าง
- ตรวจจับ Excel Hang
- ตรวจจับ Deadlock

เมื่อพบปัญหา

```text
Stop Current Task
Reset Runtime State
Unlock UI
```

---

# Enterprise Console UI

## Home

แสดง

- Current Time
- Business Date
- Current Mode
- Scheduler Status
- Progress Bar
- Live Log
- Pipeline Status

## Workbook Manager

จัดการ

- Workbook
- Export Items
- Output Files
- Enable / Disable

## Export Manager

จัดการ

- DWS
- JMS
- PDA
- Realtime DB

## Settings

จัดการ

- App ID
- App Secret
- Chat ID
- Output Folder
- Scheduler Config

---

# Pipeline Monitoring

แสดงสถานะของ

```text
DWS
JMS AUTO
JMS PDA
Realtime DB
Excel Image
Feishu
```

สถานะที่รองรับ

```text
READY
RUNNING
DONE
ERROR
```

---

# Logging System

ตัวอย่าง

```text
[OPEN] Workbook
[REFRESH] Dashboard
[DELETE] AUTO REALTIME
[EXPORT] AUTOREALTIME.png
[TOKEN] OK
[UPLOAD] Image
[SEND] Feishu
[DONE]
```

---

# Configuration

ใช้ไฟล์

```text
config.ini
```

เก็บข้อมูล

- Workbook Settings
- Export Settings
- Scheduler Settings
- Feishu Settings
- DWS/JMS Settings

---

# Installation

## Install Dependencies

```bash
pip install customtkinter
pip install requests
pip install pandas
pip install pywin32
pip install tkcalendar
pip install pillow
```

## Run

```bash
python bot_main.py
```

---

# Build EXE

OneDir

```bash
pyinstaller --onedir --windowed --name Auto_Report_Feishu bot_main.py
```

OneFile

```bash
pyinstaller --onefile --windowed --name Auto_Report_Feishu bot_main.py
```

---

# Supported Workflows

## DWS Export

```text
Download DWS
      ↓
Save Excel
      ↓
Generate Dashboard
```

## JMS Export

```text
Create Export Job
      ↓
Wait Complete
      ↓
Download XLSX
```

## Dashboard Export

```text
Refresh Workbook
      ↓
Render Dashboard
      ↓
Export PNG
```

## Feishu Delivery

```text
Upload Image
      ↓
Send Chat Message
```

---

# Version 10.0 Highlights

- Enterprise Console UI
- Dynamic Workbook Manager
- Multi Workbook Support
- Multi Export Support
- Runtime Lock Protection
- Watchdog Recovery
- Queue Processing
- Business Date Engine
- Cross-Day Scheduler
- Parallel Upload
- Improved Excel Rendering
- Improved Feishu Integration
- Unified DWS + JMS + Dashboard + Feishu Workflow

---

# License

MIT License

---

# Author

Developed for Warehouse Operations, Dashboard Monitoring, DWS/JMS Automation and Feishu Report Broadcasting.
