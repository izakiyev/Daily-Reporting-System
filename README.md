# Daily Reporting System 

## Overview
The **Daily Reporting System** is a monolithic, enterprise-grade desktop application suite designed for engineering operations management. Built with a robust technology stack including Python, PyQt6, and SQLite, the suite provides seamless daily task tracking, executive oversight, and automated reporting workflows (PDF and Excel).

The system is designed with offline-first capabilities, utilizing a smart-switching database architecture that synchronizes with a network-shared drive when available, and falls back to local storage when offline.

## System Components

### 1. Daily Report Automate (Client Application)
- **Primary File:** `app4.py`
- **Purpose:** The main interface for engineers and operational staff to log daily activities, morning plans, and evening closures.
- **Key Features:**
  - Enter and track daily tasks with dependencies and planned times.
  - Update task status (Planned, In Progress, Completed, Delayed, Cancelled).
  - High-contrast PDF generation for official records.
  - Asynchronous COM interop for Windows integration.

### 2. Operations Intelligence (Executive Dashboard)
- **Primary File:** `manager.py`
- **Purpose:** A read-only analytical dashboard for management to oversee operations across all departments.
- **Key Features:**
  - Real-time data synchronization from the shared SQLite database.
  - Filtering by department, employee, date range, and text search.
  - KPI metrics (Completion rates, Delayed items).
  - Advanced timeline view for tracking the lifecycle of specific tasks.
  - Executive summary PDF generation using ReportLab.

### 3. Utilities
- **`fix_databese.py`**: Database schema management and recovery utility.
- **`config.json`**: Application configuration allowing dynamic path resolution for the hub database.

## Architecture & Technology Stack
- **UI Framework:** PyQt6 (Desktop UI)
- **Database:** SQLite3 (Local & Network Shared)
- **Data Processing:** Pandas
- **Reporting:** ReportLab (PDFs), Matplotlib (Charts), XlsxWriter (Excel)
- **Windows Integration:** pywin32 (COM interop), ctypes

## Requirements & Setup

### Dependencies
Ensure you have Python 3.8+ installed. Install the required dependencies using pip:
```bash
pip install PyQt6 pandas xlsxwriter pywin32 matplotlib reportlab
```

### Configuration
The applications rely on `config.json` to locate the centralized database. By default, it looks for the database on a configured shared network drive. If the server is unreachable, it seamlessly falls back to the user's local AppData directory.

## Licensing
**Proprietary** - Internal Enterprise Application.
