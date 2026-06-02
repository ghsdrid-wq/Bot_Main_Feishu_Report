# Auto Report Feishu Enterprise Console v10.0

## Table of Contents

1. Project Overview
2. Problem Statement
3. Solution Overview
4. System Architecture
5. Core Features
6. Application Modules
7. Enterprise Console UI
8. Workbook Manager
9. Export Engine
10. Excel Rendering Engine
11. Business Date Engine
12. Scheduler Engine
13. Queue Processing
14. Runtime Lock System
15. Watchdog Recovery
16. Feishu Integration
17. DWS Export Workflow
18. JMS Export Workflow
19. Dashboard Workflow
20. Configuration Reference
21. Threading Architecture
22. Logging System
23. Error Handling
24. Installation
25. Build EXE
26. Deployment Guide
27. Security Notes
28. Performance Notes
29. Troubleshooting
30. Roadmap
31. Changelog
32. License

---

# Project Overview

Auto Report Feishu Enterprise Console เป็นระบบ Automation สำหรับ Warehouse Operation และ Dashboard Monitoring

โปรเจกต์นี้รวมระบบเดิมหลายตัวเข้าด้วยกัน:

- DWS Export Bot
- JMS Export Bot
- PDA Export Bot
- Realtime DB Export Bot
- Excel Dashboard Generator
- Feishu Chat Sender
- Scheduler System
- Workbook Manager

เป้าหมายคือให้ Operator สามารถควบคุมทุก Workflow ได้จากหน้าจอเดียว

---

# Problem Statement

ก่อนพัฒนาโปรเจกต์นี้

- Export รายงานหลายระบบแยกกัน
- Refresh Dashboard ด้วยมือ
- Capture รูปด้วยมือ
- ส่ง Feishu ด้วยมือ
- ตรวจสอบสถานะหลายหน้าจอ

ส่งผลให้

- ใช้เวลามาก
- เกิด Human Error
- ส่งรายงานล่าช้า
- ตรวจสอบย้อนหลังลำบาก

---

# Solution Overview

ระบบใหม่รวมทุกขั้นตอนเป็น Pipeline เดียว

```text
Download Data
      ↓
Refresh Dashboard
      ↓
Generate PNG
      ↓
Upload Feishu
      ↓
Broadcast Report
```

---

# System Architecture

```text
                     Scheduler
                         │
                         ▼
                    Task Queue
                         │
        ┌────────────────┼────────────────┐
        ▼                                 ▼
   DWS/JMS Export                  Workbook Manager
        │                                 │
        ▼                                 ▼
 Download XLSX                    Refresh Dashboard
        │                                 │
        └──────────────┬──────────────────┘
                       ▼
                Excel Renderer
                       ▼
                  PNG Output
                       ▼
                 Feishu Sender
                       ▼
                 Group Chat
```

---

# Core Features

## Data Export

- DWS Export
- JMS Export
- PDA Export
- Realtime DB Export
- Scheduled Export
- Manual Export

## Dashboard

- Excel Refresh
- Business Date Update
- Dynamic Delete Columns
- Dynamic Range Shift
- PNG Export

## Feishu

- Chat ID Messaging
- Image Upload API
- Multi Image Broadcast
- Retry Logic

## Stability

- Runtime Lock
- Queue Processing
- Watchdog Recovery
- Stop Process
- Error Recovery

---

# Application Modules

## bot_main.py

Main Controller

Responsibilities:

- GUI
- Scheduler
- Queue Manager
- Workbook Manager
- Runtime Lock
- Watchdog

## Createphoto.py

Dashboard Engine

Responsibilities:

- Excel COM
- Workbook Refresh
- Date Processing
- Dashboard Export

## Botmessage.py

Feishu Layer

Responsibilities:

- Authentication
- Upload Images
- Send Messages
- Retry Requests

---

# Enterprise Console UI

## Home

Displays

- Current Time
- Business Date
- Scheduler State
- Progress Bar
- Pipeline Status
- Live Log

## Workbook Manager

Displays

- Workbook Cards
- Export Items
- Enable / Disable Controls

## Export Manager

Displays

- DWS Settings
- JMS Settings
- Download Configuration

## Settings

Displays

- App ID
- App Secret
- Chat ID
- Output Folder

---

# Workbook Manager

Dynamic Workbook Architecture

```text
Workbook
   ├── Export 1
   ├── Export 2
   ├── Export 3
   └── Export N
```

Capabilities

- Add Workbook
- Remove Workbook
- Add Export
- Enable Workbook
- Disable Workbook
- Enable Send
- Disable Send

---

# Export Engine

Each Export Item contains

| Field | Description |
|---------|------------|
| Use | Enable Export |
| Send | Send To Feishu |
| Sheet | Worksheet |
| Range | Export Range |
| File | Output File |
| Delete | Dynamic Delete |

---

# Excel Rendering Engine

Uses

```python
win32.DispatchEx("Excel.Application")
```

Operations

```python
RefreshAll()
CalculateFull()
CopyPicture()
Chart.Export()
```

Purpose

- Refresh Dashboard
- Calculate Workbook
- Generate PNG

---

# Business Date Engine

Example

```text
Start Hour = 15
Current Time = 08:00
Business Date = Previous Day
```

Benefits

- Correct Night Shift Reporting
- Cross-Day Dashboard Support

---

# Scheduler Engine

Supports

- Run Now
- Auto Run
- Hourly Schedule
- Cross-Day Schedule
- Stop Scheduler

Example

```text
Run Minute = 5
Start Hour = 15
End Hour = 12
```

---

# Queue Processing

Uses

```python
queue.Queue()
```

Purpose

- Prevent Duplicate Run
- Sequential Execution
- Background Processing

---

# Runtime Lock System

Locks UI During Runtime

Protected Areas

- Workbook Config
- Export Config
- Scheduler Config
- Feishu Config

Benefits

- Prevent Invalid State
- Prevent Runtime Corruption

---

# Watchdog Recovery

Monitors

- Excel Hang
- Deadlock
- Frozen Process

Recovery

```text
Stop Current Task
Reset Runtime State
Unlock UI
```

---

# Feishu Integration

Workflow

```text
Get Token
      ↓
Upload Image
      ↓
Receive Image Key
      ↓
Send Chat Message
```

Supports

- Chat ID
- Image Message
- Multi Image Send

---

# DWS Export Workflow

```text
Request Data
      ↓
Generate File
      ↓
Download XLSX
      ↓
Save Output
```

---

# JMS Export Workflow

```text
Create Job
      ↓
Wait Complete
      ↓
Download XLSX
      ↓
Save Output
```

---

# Dashboard Workflow

```text
Open Workbook
      ↓
Refresh Workbook
      ↓
Calculate Workbook
      ↓
Render Dashboard
      ↓
Export PNG
```

---

# Configuration Reference

## FEISHU

```ini
APP_ID=
APP_SECRET=
CHAT_ID=
```

## TIME

```ini
run_minute=5
start_hour=15
end_hour=12
```

## PATH

```ini
output_dir=
```

---

# Threading Architecture

```text
Main UI Thread
      │
      ├── Scheduler Thread
      ├── Queue Worker
      ├── Watchdog Thread
      └── Upload Workers
```

---

# Logging System

Example

```text
[OPEN] Workbook
[REFRESH] Dashboard
[EXPORT] Image
[TOKEN] OK
[UPLOAD] Image
[SEND] Feishu
[DONE]
```

---

# Error Handling

Supported

- Excel Timeout
- Workbook Missing
- Upload Failure
- Invalid Token
- Invalid Chat ID
- Scheduler Error
- Runtime Cancellation

---

# Installation

```bash
pip install customtkinter
pip install requests
pip install pandas
pip install pywin32
pip install tkcalendar
pip install pillow
```

Run

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

# Deployment Guide

1. Install Python
2. Install Dependencies
3. Configure config.ini
4. Test Manual Run
5. Test Feishu Connection
6. Configure Scheduler
7. Build EXE
8. Deploy To Production

---

# Security Notes

- Do not commit APP_SECRET
- Do not commit CHAT_ID if private
- Restrict access to config.ini
- Use dedicated Feishu application

---

# Performance Notes

Recommended

- Windows 10/11
- Excel Desktop Installed
- SSD Storage
- 8GB RAM+

---

# Troubleshooting

## Blank PNG

Check

- Excel Installed
- Workbook Path
- Sheet Name
- Export Range

## Feishu Send Failed

Check

- APP_ID
- APP_SECRET
- CHAT_ID
- Internet Connection

## Scheduler Not Running

Check

- Time Window
- Run Minute
- Auto Mode Enabled

---

# Roadmap

- PDF Export
- Dashboard History
- Web Dashboard
- Telegram Integration
- Slack Integration
- Multi Group Broadcast
- Database Storage
- User Permissions

---

# Changelog

## v10.0

- Enterprise Console UI
- Workbook Manager
- Dynamic Export System
- Runtime Lock
- Watchdog Recovery
- Parallel Upload
- Queue Processing

## Previous Versions

- v1-v3 DWS Export
- v4-v5 JMS Integration
- v6 Dashboard Export
- v7 Feishu Sender
- v8 Scheduler Upgrade
- v9 Workbook Migration
- v10 Enterprise Console

---

# License

MIT License

---

# Author

Developed for Warehouse Operations, Dashboard Monitoring, DWS/JMS Automation and Feishu Report Broadcasting.
