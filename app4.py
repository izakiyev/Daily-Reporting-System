# -*- coding: utf-8 -*-
"""
================================================================================
APP NAME:    Daily Report Automate (Titanium Edition)
VERSION:     4.0.0-ENT
AUTHOR:      Senior Systems Architect
DATE:        2025-12-21
LICENSE:     Proprietary (KTIB Azerbaijan LLC)
FRAMEWORK:   PyQt6 + Win32COM + Pandas + Matplotlib

DESCRIPTION:
    A monolithic, enterprise-grade desktop application for engineering 
    operations management. Features asynchronous COM interop, robust 
    JSON persistence, and automated reporting workflows.
================================================================================
"""
import ctypes  # Required for the Windows Taskbar icon fix
import sys
import os
import json
import shutil
import logging
import traceback
import re
import urllib.parse
from enum import Enum
from dataclasses import dataclass, asdict, field
from datetime import datetime, date, timedelta
from typing import List, Dict, Optional, Any, Tuple, Union


# ==============================================================================
# SECTION 1: THIRD-PARTY DEPENDENCIES
# ==============================================================================
# We wrap imports in try-except blocks to provide user-friendly error messages
# rather than cryptic stack traces upon startup failure.
try:
    import pandas as pd
    import xlsxwriter
    import pythoncom
    import win32com.client as win32
    
    # Matplotlib Configuration for PyQt6 Integration
    import matplotlib
    matplotlib.use('QtAgg')
    from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure
    import matplotlib.pyplot as plt
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    # PyQt6 Core Components
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QTabWidget, QLabel, QLineEdit, QPushButton, QComboBox,
        QTableWidget, QTableWidgetItem, QHeaderView, QTimeEdit,
        QMessageBox, QFrame, QSplitter, QTextEdit, QSlider,
        QFileDialog, QGridLayout, QScrollArea, QStatusBar, QMenuBar,
        QMenu, QProgressBar, QAbstractItemView, QDialog, QFormLayout,
        QSystemTrayIcon, QStyle, QSizePolicy, QCheckBox , QDateEdit,QGraphicsDropShadowEffect, QSpinBox,
        
    )
    from PyQt6.QtCore import (
        Qt, QUrl, QTime, QSize, QTimer, QSettings, pyqtSignal, 
        QThread, QRunnable, QThreadPool, QObject, QStandardPaths, QPoint, QDate, QDateTime
    )
    from PyQt6.QtGui import (
        QIcon, QFont, QColor, QAction, QPalette, QBrush, 
        QPixmap, QPainter, QLinearGradient, QGradient 
    )

except ImportError as e:
    # This block ensures the user knows exactly what to install.
    print("!" * 80)
    print("CRITICAL STARTUP ERROR: MISSING DEPENDENCIES")
    print(f"Details: {e}")
    print("-" * 80)
    print("Please execute the following command in your terminal:")
    print("pip install PyQt6 pandas xlsxwriter pywin32 matplotlib")
    print("!" * 80)
    sys.exit(1)
from PyQt6 import QtSvg  # REQUIRED for SVG icons in EXE
# ==============================================================================
# SECTION 2: GLOBAL CONFIGURATION & CONSTANTS
# ==============================================================================
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
# 
# ==============================================================================
# SECTION 2: GLOBAL CONFIGURATION & CONSTANTS
# ==============================================================================

class Config:
    APP_NAME = "KTIB_Report_Automate"
    APP_VERSION = "4.1.0"
    ORG_NAME = "KTIB Azerbaijan LLC"
    
    # --- PATH DETERMINATION (Handles EXE vs Python Script) ---
    if getattr(sys, 'frozen', False):
        EXE_LOCATION = os.path.dirname(sys.executable)
    else:
        EXE_LOCATION = os.path.dirname(os.path.abspath(__file__))

    # --- CONFIGURATION LOADER ---
    # 1. Define the default hardcoded path
    DEFAULT_HUB_DIR = r"H:\Public AZ\DailyReportAutomate"
    
    # 2. define where the config file should be
    CONFIG_FILE_PATH = os.path.join(EXE_LOCATION, "config.json")
    
    # 3. Initialize with default, then try to override
    SERVER_HUB_DIR = DEFAULT_HUB_DIR
    
    try:
        if os.path.exists(CONFIG_FILE_PATH):
            with open(CONFIG_FILE_PATH, 'r') as f:
                config_data = json.load(f)
                # If the key exists in JSON, use it. Otherwise keep default.
                SERVER_HUB_DIR = config_data.get("server_hub_dir", DEFAULT_HUB_DIR)
    except Exception:
        # If JSON is broken or unreadable, fail silently and use the H: drive default
        pass

    HUB_DB_NAME = "ktib_ops_database.db"

    # UI Options
    DEPARTMENTS = ["OPS", "MAINT", "ELEC", "INSTR", "HSE", "ADMIN", "LOGISTICS", "MGMT"]
    STATUS_OPTS = ["Planned", "In Progress", "Completed", "Delayed", "Cancelled"]
    PRIORITY_OPTS = ["High", "Medium", "Low"]

    # Local Fallback (C: Drive AppData)
    STORAGE_DIR = os.path.join(
        QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppLocalDataLocation),
        APP_NAME
    )
    
    # Directory Definitions
    BACKUP_DIR = os.path.join(STORAGE_DIR, "backups")
    REPORTS_DIR = os.path.join(EXE_LOCATION, "daily_reports")
    LOG_PATH = os.path.join(STORAGE_DIR, "application.log")

    @staticmethod
    def get_db_path():
        """
        Smart-Switching DB Path:
        1. Checks Configured Server Path (H: or JSON path).
        2. Falls back to Local AppData if Server is unreachable.
        """
        server_db = os.path.join(Config.SERVER_HUB_DIR, Config.HUB_DB_NAME)
        local_db = os.path.join(Config.STORAGE_DIR, "ops_local.db")

        # Check if the folder exists (Server is reachable)
        if os.path.exists(Config.SERVER_HUB_DIR):
            return server_db
        return local_db

    @staticmethod
    def initialize_directories():
        """Creates folders on both Local and Server Drive."""
        # Create Local folders
        for folder in [Config.STORAGE_DIR, Config.BACKUP_DIR, Config.REPORTS_DIR]:
            os.makedirs(folder, exist_ok=True)
        
        # Create Server folder if reachable
        # We check the parent directory or the path itself to see if the network drive is mapped
        try:
            if os.path.exists(os.path.dirname(Config.SERVER_HUB_DIR)):
                os.makedirs(Config.SERVER_HUB_DIR, exist_ok=True)
        except Exception:
            pass # Server likely offline or permission denied

# IMPORTANT: Also make sure you call initialize_directories() 
# at the very bottom of your script in the if __name__ == "__main__": block!

class Colors:
    """Centralized Color Palette."""
    PRIMARY = "#064e3b"
    PRIMARY_HOVER = "#065f46"
    SECONDARY = "#0f172a"
    ACCENT = "#10b981"
    BACKGROUND = "#f8fafc"
    SURFACE = "#ffffff"
    BORDER = "#e2e8f0"
    TEXT_MAIN = "#334155"
    TEXT_MUTED = "#64748b"
    SUCCESS = "#059669"
    WARNING = "#d97706"
    DANGER = "#e11d48"
    DANGER_BG = "#fee2e2"
    SUCCESS_BG = "#dcfce7"

# Initialize Logging & Folders
# This block ensures the folders exist before the app starts
for folder in [Config.BACKUP_DIR, Config.REPORTS_DIR]:
    if not os.path.exists(folder):
        os.makedirs(folder, exist_ok=True)

logging.basicConfig(
    filename=Config.LOG_PATH,
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(module)s:%(lineno)d | %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# ==============================================================================
# SECTION 3: DATA MODELS & ENUMS
# ==============================================================================

# ==============================================================================
# SECTION 4: INFRASTRUCTURE -> PDF ENGINE (FULL DATA DUMP)
# ==============================================================================
# ==============================================================================
# SECTION 4: INFRASTRUCTURE -> PDF ENGINE (FULL DATA DUMP - HIGH CONTRAST)
# ==============================================================================
class PDFEngine:
    """
    Generates an 'Official Record' PDF containing EVERY data point.
    Style: High Contrast (Black Headers on Light Gray).
    """
    
    @staticmethod
    def generate_analytics_report(filepath: str, tasks: List[TaskItem], profile: UserProfile, start_date: str, end_date: str) -> bool:
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch

            # 1. DOCUMENT CONFIGURATION (Landscape A4)
            doc = SimpleDocTemplate(
                filepath, 
                pagesize=landscape(A4),
                rightMargin=20, leftMargin=20, 
                topMargin=30, bottomMargin=30
            )
            
            elements = []
            styles = getSampleStyleSheet()
            
            # 2. STATISTICS CALCULATION
            total_tasks = len(tasks)
            completed_tasks = len([t for t in tasks if t.status == 'Completed'])
            delayed_tasks = len([t for t in tasks if t.status == 'Delayed'])
            efficiency = int((completed_tasks / total_tasks * 100)) if total_tasks > 0 else 0
            
            # 3. REPORT HEADER
            # Title
            title_style = ParagraphStyle(
                'Title', parent=styles['Heading1'], 
                fontSize=18, textColor=colors.black, # Changed to Black
                spaceAfter=5, alignment=1 
            )
            elements.append(Paragraph(f"FULL OPERATIONAL ANALYTICS REPORT", title_style))
            elements.append(Paragraph(f"Period: <b>{start_date}</b> to <b>{end_date}</b>", 
                ParagraphStyle('Sub', parent=styles['Normal'], alignment=1, fontSize=11)))
            elements.append(Spacer(1, 15))

            # 4. SUMMARY BOX
            summary_html = f"""
            <font size="10">
            <b>OPERATOR:</b> {profile.name} &nbsp;|&nbsp; 
            <b>DEPT:</b> {profile.department} &nbsp;|&nbsp; 
            <b>GENERATED:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}<br/><br/>
            <b>TASKS:</b> {total_tasks} &nbsp;|&nbsp; 
            <b>SUCCESS RATE:</b> {efficiency}% &nbsp;|&nbsp; 
            <b>DELAYED:</b> <font color='red'>{delayed_tasks}</font>
            </font>
            """
            sum_data = [[Paragraph(summary_html, styles['Normal'])]]
            sum_table = Table(sum_data, colWidths=[10.8*inch])
            sum_table.setStyle(TableStyle([
                ('BOX', (0,0), (-1,-1), 1, colors.black),
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
                ('PADDING', (0,0), (-1,-1), 8),
            ]))
            elements.append(sum_table)
            elements.append(Spacer(1, 15))

            # 5. DATA TABLE CONSTRUCTION
            headers = ["Meta Data", "Task Configuration", "Status & Progress", "Execution Log (Notes, Results, Next Steps)"]
            data = [headers]
            
            sorted_tasks = sorted(tasks, key=lambda x: x.date)
            
            # Styles for cell content
            s_norm = ParagraphStyle('CellNorm', parent=styles['Normal'], fontSize=9, leading=11, spaceAfter=3, textColor=colors.black)
            
            for t in sorted_tasks:
                # --- COLUMN 1: META ---
                meta_html = f"""
                <b>Date:</b> {t.date}<br/>
                <b>ID:</b> {t.id}<br/>
                <b>Time:</b> {t.planned_time}
                """
                
                # --- COLUMN 2: CONFIGURATION ---
                config_html = f"""
                <b>Activity:</b> {t.description}<br/>
                <b>Priority:</b> {t.priority}<br/>
                <b>Dependencies:</b> {t.dependencies if t.dependencies else 'None'}
                """
                
                # --- COLUMN 3: STATUS ---
                color = "black"
                if t.status == "Completed": color = "green"
                elif t.status == "Delayed": color = "red"
                elif t.status == "In Progress": color = "orange"
                
                status_html = f"""
                <b>Status:</b> <font color='{color}'>{t.status}</font><br/>
                <b>Progress:</b> {t.percent_complete}%
                """
                
                # --- COLUMN 4: FULL LOG ---
                log_parts = []
                if t.morning_notes:
                    log_parts.append(f"<b>[Plan Note]:</b> {t.morning_notes}")
                if t.results:
                    log_parts.append(f"<b>[Result]:</b> {t.results}")
                if t.reason_delay:
                    log_parts.append(f"<b>[DELAY CAUSE]:</b> <font color='red'>{t.reason_delay}</font>")
                if t.next_steps:
                    log_parts.append(f"<b>[Next Steps]:</b> {t.next_steps}")
                
                if not log_parts:
                    log_html = "<i>No logs entered.</i>"
                else:
                    log_html = "<br/><br/>".join(log_parts)
                
                row = [
                    Paragraph(meta_html, s_norm),
                    Paragraph(config_html, s_norm),
                    Paragraph(status_html, s_norm),
                    Paragraph(log_html, s_norm)
                ]
                data.append(row)

            # 6. TABLE STYLING (UPDATED FOR BLACK HEADERS)
            col_widths = [1.2*inch, 2.5*inch, 1.3*inch, 6.0*inch]
            
            t = Table(data, colWidths=col_widths, repeatRows=1)
            
            tbl_style = TableStyle([
                # --- HEADER STYLING CHANGED HERE ---
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#e2e8f0")), # Light Gray Background
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),               # Black Text
                # -----------------------------------
                
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('TOPPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('ROWBACKGROUNDS', (1, 0), (-1, -1), [colors.white, colors.whitesmoke]),
            ])
            t.setStyle(tbl_style)
            elements.append(t)

            # 7. FOOTER
            elements.append(Spacer(1, 20))
            footer_text = f"OFFICIAL DATA DUMP | {Config.APP_NAME} | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            elements.append(Paragraph(footer_text, ParagraphStyle('Footer', fontSize=7, textColor=colors.grey, alignment=1)))

            doc.build(elements)
            return True

        except Exception as e:
            logging.error(f"PDF Analytics Error: {traceback.format_exc()}")
            return False
        
@dataclass
class UserProfile:
    """Represents the active user configuration."""
    name: str = ""
    department: str = "OPS"
    manager_name: str = ""    # <--- ADD THIS LINE
    manager_email: str = ""
    theme: str = "Light"
    include_summary: bool = True

    def is_valid(self) -> bool:
        return bool(self.name and self.department and self.manager_email)

@dataclass
class TaskItem:
    """Represents a single unit of work."""
    id: str
    date: str  # ISO Format YYYY-MM-DD
    employee: str
    department: str
    description: str
    priority: str = "Medium"
    planned_time: str = "17:00"
    dependencies: str = ""
    morning_notes: str = ""
    sort_order: int = 0  # <--- Add this at the end

    # Evening Update Fields
    status: str = "Planned"
    percent_complete: int = 0
    results: str = ""
    reason_delay: str = ""
    next_steps: str = ""
    
    def to_dict(self) -> Dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict):
        return cls(**data)

# ==============================================================================
# SECTION 4: INFRASTRUCTURE (BACKEND LOGIC)
# ==============================================================================

class Utils:
    """Static utility helpers."""
    
    @staticmethod
    def get_timestamp() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def get_today_iso() -> str:
        return date.today().isoformat()

    @staticmethod
    def validate_email(email: str) -> bool:
        """Regex validation for email addresses."""
        pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
        return re.match(pattern, email) is not None

import threading

import sqlite3
import threading
import os
from contextlib import contextmanager

class DatabaseEngine:
    """
    TITANIUM SQLITE ENGINE (Network & Manager Ready)
    
    Optimized for:
    - Multi-user concurrency via WAL mode.
    - Server-side shared folder deployment.
    - Composite Key Integrity (ID + Date + Employee).
    - Future-proof for Manager App statistical queries.
    """
    
    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super(DatabaseEngine, cls).__new__(cls)
                    cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self.init_db()
        self._initialized = True

    @contextmanager
    def connection(self):
        """Thread-safe connection to the H: Drive database."""
        db_file = Config.get_db_path()
        
        # timeout=60 is critical for network drives so the app doesn't 
        # crash if two employees save at the same exact second.
        conn = sqlite3.connect(db_file, timeout=60, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        try:
            # TRUNCATE mode is the most stable for H: drives (network shares)
            conn.execute("PRAGMA journal_mode=TRUNCATE")
            conn.execute("PRAGMA synchronous=NORMAL")
            yield conn
            conn.commit()
        except Exception as e:
            conn.rollback()
            logging.error(f"Database Error: {e}")
            raise e
        finally:
            conn.close()



    def init_db(self):
        """Initializes tables and performs schema migrations."""
        with self.connection() as conn:
            # 1. Profile Table (Added manager_name to the schema)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS profile (
                    os_user TEXT PRIMARY KEY,
                    full_name TEXT,
                    department TEXT,
                    manager_name TEXT,
                    manager_email TEXT,
                    include_summary INTEGER DEFAULT 1
                )
            """)
            
            # SCHEMA MIGRATION: Ensure 'manager_name' exists in older DB files
            try:
                conn.execute("ALTER TABLE profile ADD COLUMN manager_name TEXT")
                logging.info("Migration: Added manager_name to profile table.")
            except sqlite3.OperationalError:
                pass # Column already exists

            # SCHEMA MIGRATION: Ensure 'include_summary' exists in older DB files
            try:
                conn.execute("ALTER TABLE profile ADD COLUMN include_summary INTEGER DEFAULT 1")
                logging.info("Migration: Added include_summary to profile table.")
            except sqlite3.OperationalError:
                pass # Column already exists

            # 2. Tasks Table
            conn.execute("""
                CREATE TABLE IF NOT EXISTS tasks (
                    task_id TEXT, report_date TEXT, os_user TEXT,
                    employee TEXT, department TEXT, description TEXT,
                    priority TEXT, planned_time TEXT, dependencies TEXT,
                    morning_notes TEXT, status TEXT, percent_complete INTEGER,
                    results TEXT, reason_delay TEXT, next_steps TEXT,
                    sort_order INTEGER DEFAULT 0, 
                    PRIMARY KEY (task_id, report_date, os_user)
                )
            """)
            try:
                conn.execute("ALTER TABLE tasks ADD COLUMN sort_order INTEGER DEFAULT 0")
                logging.info("Migration: Added sort_order to tasks table.")
            except sqlite3.OperationalError:
                pass 

    def get_tasks(self, date_filter: str = None) -> List[TaskItem]:
        user_key = os.getlogin()
        with self.connection() as conn:
            query = "SELECT * FROM tasks WHERE os_user = ?"
            params = [user_key]
            
            if date_filter:
                query += " AND report_date LIKE ?"
                params.append(f"{date_filter}%")
            
            # --- MODIFIED SORTING ---
            # Sort by sort_order first, then by task_id
            query += " ORDER BY sort_order ASC, task_id ASC" 
            
            cursor = conn.execute(query, params)
            return [self._map_row_to_task(row) for row in cursor]

    def upsert_task(self, t: TaskItem):
        user_key = os.getlogin()
        with self.connection() as conn:
            # Check if sort_order needs initialization (if it's 0 or None)
            current_sort = getattr(t, 'sort_order', 0)
            if current_sort == 0:
                # Find max sort_order for this day and add 1
                row = conn.execute("SELECT MAX(sort_order) as m FROM tasks WHERE report_date = ? AND os_user = ?", 
                                 (t.date, user_key)).fetchone()
                current_sort = (row['m'] or 0) + 1
            
            # Ensure TaskItem dataclass has sort_order, or handle it manually here
            # We are updating the SQL to include the 16th parameter (sort_order)
            conn.execute("""
                REPLACE INTO tasks VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                t.id, t.date, user_key, t.employee, t.department, t.description, 
                t.priority, t.planned_time, t.dependencies, t.morning_notes, 
                t.status, t.percent_complete, t.results, t.reason_delay, t.next_steps,
                current_sort # <--- Insert sort_order
            ))

    def swap_task_positions(self, task_id_1: str, task_id_2: str, date_iso: str):
        """New method to swap sort_order of two tasks"""
        user_key = os.getlogin()
        with self.connection() as conn:
            # Get current orders
            t1 = conn.execute("SELECT sort_order FROM tasks WHERE task_id=? AND report_date=? AND os_user=?", 
                              (task_id_1, date_iso, user_key)).fetchone()
            t2 = conn.execute("SELECT sort_order FROM tasks WHERE task_id=? AND report_date=? AND os_user=?", 
                              (task_id_2, date_iso, user_key)).fetchone()
            
            if t1 and t2:
                # Swap values
                conn.execute("UPDATE tasks SET sort_order=? WHERE task_id=? AND report_date=? AND os_user=?", 
                             (t2['sort_order'], task_id_1, date_iso, user_key))
                conn.execute("UPDATE tasks SET sort_order=? WHERE task_id=? AND report_date=? AND os_user=?", 
                             (t1['sort_order'], task_id_2, date_iso, user_key))

    def _map_row_to_task(self, r) -> TaskItem:
        # Create task and dynamically attach sort_order since it might not be in the dataclass definition yet
        t = TaskItem(
            id=r['task_id'], date=r['report_date'], employee=r['employee'],
            department=r['department'], description=r['description'],
            priority=r['priority'], planned_time=r['planned_time'],
            dependencies=r['dependencies'], morning_notes=r['morning_notes'],
            status=r['status'], percent_complete=r['percent_complete'],
            results=r['results'], reason_delay=r['reason_delay'],
            next_steps=r['next_steps']
        )
        t.sort_order = r['sort_order'] if 'sort_order' in r.keys() else 0
        return t
        
    def get_profile(self) -> UserProfile:
        """Fetches the profile and maps the new manager_name field."""
        user_key = os.getlogin()
        with self.connection() as conn:
            row = conn.execute("SELECT * FROM profile WHERE os_user = ?", (user_key,)).fetchone()
            if row:
                return UserProfile(
                    name=row['full_name'] or "",
                    department=row['department'] or "OPS",
                    manager_name=row['manager_name'] or "",  # <-- LOAD THIS FIELD
                    manager_email=row['manager_email'] or "",
                    include_summary=bool(row['include_summary']) if row['include_summary'] is not None else True
                )
            return UserProfile(name=user_key)

    def set_profile(self, p: UserProfile):
        """Saves the profile including the manager_name field."""
        user_key = os.getlogin()
        with self.connection() as conn:
            # Updated to 6 columns to match the new schema
            conn.execute("""
                REPLACE INTO profile (os_user, full_name, department, manager_name, manager_email, include_summary)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                user_key, 
                p.name, 
                p.department, 
                p.manager_name,  # <-- SAVE THIS FIELD
                p.manager_email, 
                int(p.include_summary)
            ))

    # ==========================================================================
    # TASK MANAGEMENT (Manager-App Ready)
    # ==========================================================================

    # Change in DatabaseEngine.get_tasks
    def get_tasks(self, date_filter: str = None) -> List[TaskItem]:
        user_key = os.getlogin()
        with self.connection() as conn:
            query = "SELECT * FROM tasks WHERE os_user = ?"
            params = [user_key]
            
            if date_filter:
                # --- CHANGE 1: ALLOW FLEXIBLE TYPING ---
                # Changed '=' to 'LIKE' to allow partial matching 
                # (e.g., user types "2025-01" to get whole month)
                query += " AND report_date LIKE ?"
                # We add '%' to ensure it matches starts-with
                params.append(f"{date_filter}%") 
            
            # --- CHANGE 2: KEEP THE ORDER FIX ---
            query += " ORDER BY task_id ASC" 
            
            cursor = conn.execute(query, params)
            return [self._map_row_to_task(row) for row in cursor]
    
    # Change in DatabaseEngine.upsert_task
    def upsert_task(self, t: TaskItem):
        user_key = os.getlogin()
        with self.connection() as conn:
            conn.execute("""
                REPLACE INTO tasks VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                t.id, t.date, user_key, t.employee, t.department, t.description, 
                t.priority, t.planned_time, t.dependencies, t.morning_notes, 
                t.status, t.percent_complete, t.results, t.reason_delay, t.next_steps
            ))

    def delete_task(self, task_id: str, task_date: str):
        user_key = os.getlogin() # Change from employee_name to user_key
        with self.connection() as conn:
            conn.execute("""
                DELETE FROM tasks 
                WHERE task_id = ? AND report_date = ? AND os_user = ?
            """, (task_id, task_date, user_key))

    def get_previous_available_tasks(self, current_date_iso: str) -> List[TaskItem]:
        user_key = os.getlogin() # Change from employee_name to user_key
        with self.connection() as conn:
            row = conn.execute("""
                SELECT report_date FROM tasks 
                WHERE report_date < ? AND os_user = ?
                ORDER BY report_date DESC LIMIT 1
            """, (current_date_iso, user_key)).fetchone()
            
            if row:
                return self.get_tasks(row['report_date'])
            return []

    def _map_row_to_task(self, r) -> TaskItem:
        return TaskItem(
            id=r['task_id'], date=r['report_date'], employee=r['employee'],
            department=r['department'], description=r['description'],
            priority=r['priority'], planned_time=r['planned_time'],
            dependencies=r['dependencies'], morning_notes=r['morning_notes'],
            status=r['status'], percent_complete=r['percent_complete'],
            results=r['results'], reason_delay=r['reason_delay'],
            next_steps=r['next_steps']
        )
    
class OutlookEngine:
    """
    COM Interface Wrapper for Outlook. 
    Must be instantiated inside a Worker thread for UI responsiveness.
    """
    @staticmethod
    def create_draft(to: str, subject: str, body: str, attachments: List[str] = None) -> bool:
        try:
            # Initialize COM library for this thread
            pythoncom.CoInitialize()
            
            try:
                outlook = win32.GetActiveObject("Outlook.Application")
            except:
                outlook = win32.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0) # olMailItem
            mail.To = to
            mail.Subject = subject
            mail.HTMLBody = body
            
            if attachments:
                for path in attachments:
                    if os.path.exists(path):
                        mail.Attachments.Add(os.path.abspath(path))
                    else:
                        logging.warning(f"Attachment missing: {path}")
            
            mail.Display() # Open window
            return True
        except Exception as e:
            logging.error(f"Outlook COM Error: {e}")
            raise e
        finally:
            pythoncom.CoUninitialize()

class ExcelEngine:
    """
    Generates Excel reports with specific formatting for Morning and Evening phases.
    Entry point: generate_report
    """

    @staticmethod
    def generate_report(filepath: str, tasks: List[TaskItem], mode: str = "morning") -> bool:
        """
        Unified entry point called by the UI. 
        It fetches the profile automatically to match the required signatures.
        """
        db = DatabaseEngine()
        profile = db.get_profile()
        
        if mode == "morning":
            return ExcelEngine.generate_morning(filepath, tasks, profile)
        else:
            return ExcelEngine.generate_evening(filepath, tasks, profile)

    @staticmethod
    def generate_morning(filepath: str, tasks: List[TaskItem], profile: UserProfile) -> bool:
        try:
            data = []
            for t in tasks:
                data.append({
                    "Date": t.date,
                    "Employee": t.employee,
                    "Department": t.department,
                    "Task ID": t.id,
                    "Task Description": t.description,
                    "Priority (High/Medium/Low)": t.priority,
                    "Planned Completion Time": t.planned_time,
                    "Dependencies": t.dependencies,
                    "Morning Notes": t.morning_notes
                })
            df = pd.DataFrame(data)
            
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Morning Plan')
                wb = writer.book
                ws = writer.sheets['Morning Plan']
                
                # Formats
                header_fmt = wb.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#F2F2F2', 'text_wrap': True})
                center_fmt = wb.add_format({'border': 1, 'align': 'center', 'valign': 'top'})
                body_left = wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True})
                
                # Column Widths
                ws.set_column('A:D', 15)
                ws.set_column('E:E', 50)
                ws.set_column('F:I', 20)
                
                # Write Headers
                for i, col in enumerate(df.columns):
                    ws.write(0, i, col, header_fmt)
                    
                # Write Data
                for r, row in df.iterrows():
                    idx = r + 1
                    ws.write(idx, 0, row['Date'], center_fmt)
                    ws.write(idx, 1, row['Employee'], center_fmt)
                    ws.write(idx, 2, row['Department'], center_fmt)
                    ws.write(idx, 3, row['Task ID'], center_fmt)
                    ws.write(idx, 4, row['Task Description'], body_left)
                    ws.write(idx, 5, row['Priority (High/Medium/Low)'], center_fmt)
                    ws.write(idx, 6, row['Planned Completion Time'], center_fmt)
                    ws.write(idx, 7, row['Dependencies'], body_left)
                    ws.write(idx, 8, row['Morning Notes'], body_left)
            return True
        except Exception as e:
            logging.error(f"Morning Excel Error: {traceback.format_exc()}")
            return False

    @staticmethod
    def generate_evening(filepath: str, tasks: List[TaskItem], profile: UserProfile) -> bool:
        try:
            evening_data = []
            for t in tasks:
                # Clean ID to just the numeric part
                simple_id = t.id.split('-')[-1] if '-' in t.id else t.id
                try:
                    clean_id = int(simple_id)
                except:
                    clean_id = simple_id

                display_pct = t.percent_complete
                if t.status == "Completed":
                    display_pct = 100

                evening_data.append({
                    "Date": t.date,
                    "Employee": t.employee,
                    "Task ID": clean_id,
                    "Task Description": t.description,
                    "Status (Completed/In Progress/Delayed)": t.status,
                    "% Completed": f"{display_pct}%", # Use calculated value
                    "Results / Notes": t.results,
                    "Reason for Delay": t.reason_delay,
                    "Next Steps / Move To (Date)": t.next_steps
                })
            
            cols = ["Date", "Employee", "Task ID", "Task Description", "Status (Completed/In Progress/Delayed)", 
                    "% Completed", "Results / Notes", "Reason for Delay", "Next Steps / Move To (Date)"]
            df = pd.DataFrame(evening_data, columns=cols)

            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                # --- SHEET 1: EVENING UPDATE ---
                df.to_excel(writer, index=False, sheet_name='Evening Update')
                wb = writer.book
                ws = writer.sheets['Evening Update']
                
                header_fmt = wb.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#F2F2F2', 'text_wrap': True})
                center_fmt = wb.add_format({'border': 1, 'align': 'center', 'valign': 'top'})
                left_fmt = wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True})
                
                # Column widths for Screenshot 2
                ws.set_column('A:B', 15)
                ws.set_column('C:C', 10)
                ws.set_column('D:D', 45)
                ws.set_column('E:I', 20)
                
                for i, col in enumerate(df.columns):
                    ws.write(0, i, col, header_fmt)
                    
                for r, row in df.iterrows():
                    idx = r + 1
                    ws.write(idx, 0, row['Date'], center_fmt)
                    ws.write(idx, 1, row['Employee'], center_fmt)
                    ws.write(idx, 2, row['Task ID'], center_fmt)
                    ws.write(idx, 3, row['Task Description'], left_fmt)
                    ws.write(idx, 4, row['Status (Completed/In Progress/Delayed)'], center_fmt)
                    ws.write(idx, 5, row['% Completed'], center_fmt)
                    ws.write(idx, 6, row['Results / Notes'], left_fmt)
                    ws.write(idx, 7, row['Reason for Delay'], center_fmt)
                    ws.write(idx, 8, row['Next Steps / Move To (Date)'], center_fmt)

                # --- SHEET 2: SUMMARY DASHBOARD ---
                ws_dash = wb.add_worksheet('Summary Dashboard')
                total = len(tasks)
                completed = len([t for t in tasks if t.status == 'Completed'])
                delayed = len([t for t in tasks if t.status == 'Delayed'])
                rate = (completed / total) if total > 0 else 0
                
                d_head = wb.add_format({'bold': True, 'border': 1, 'font_color': '#064e3b', 'bg_color': '#FFFFFF'})
                d_label = wb.add_format({'bold': False, 'border': 1})
                d_val = wb.add_format({'bold': True, 'border': 1, 'align': 'right'})
                
                ws_dash.write('A1', 'Date to Analyse', d_head); ws_dash.write('B1', Utils.get_today_iso(), d_val)
                ws_dash.write('A3', 'Metric', d_head); ws_dash.write('B3', 'Value', d_head)
                ws_dash.write('A4', 'Total Tasks', d_label); ws_dash.write('B4', total, d_val)
                ws_dash.write('A5', 'Completed Tasks', d_label); ws_dash.write('B5', completed, d_val)
                ws_dash.write('A6', 'Delayed Tasks', d_label); ws_dash.write('B6', delayed, d_val)
                ws_dash.write('A7', 'Completion %', d_label); ws_dash.write('B7', f"{rate:.1%}", d_val)
                
                ws_dash.write('A9', 'Department Workload', d_head)
                ws_dash.write('A10', 'Employee Name', d_head); ws_dash.write('B10', 'Task Count', d_head)
                ws_dash.write('A11', profile.name, d_label); ws_dash.write('B11', total, d_val)
                
                ws_dash.set_column('A:A', 35)
                ws_dash.set_column('B:B', 15)
                
            return True
        except Exception as e:
            logging.error(f"Evening Excel Error: {traceback.format_exc()}")
            return False
# ==============================================================================
# SECTION 5: UI COMPONENT LIBRARY (DESIGN SYSTEM)
# ==============================================================================

class UI:
    """Static factory for standardized UI styling."""
    
    STYLE_QSS = f"""
        QMainWindow {{ background-color: {Colors.BACKGROUND}; }}
        QWidget {{ color: {Colors.TEXT_MAIN}; }}
        
        /* Cards */
        QFrame.card {{
            background-color: {Colors.SURFACE};
            border: 1px solid {Colors.BORDER};
            border-radius: 10px;
        }}
        
        /* Inputs */
        QLineEdit, QComboBox, QTimeEdit {{
            border: 1px solid {Colors.BORDER};
            border-radius: 6px;
            padding: 10px;
            background-color: {Colors.SURFACE};
            font-size: 13px;
        }}
        QLineEdit:focus, QComboBox:focus {{
            border: 2px solid {Colors.PRIMARY};
        }}
        
        /* Buttons */
        QPushButton {{
            background-color: {Colors.SURFACE};
            border: 1px solid {Colors.BORDER};
            border-radius: 6px;
            padding: 10px 20px;
            font-weight: bold;
            color: {Colors.TEXT_MAIN};
        }}
        QPushButton:hover {{
            background-color: {Colors.BACKGROUND};
            border-color: {Colors.TEXT_MUTED};
        }}
        QPushButton.primary {{
            background-color: {Colors.PRIMARY};
            color: {Colors.SURFACE};
            border: none;
        }}
        QPushButton.primary:hover {{
            background-color: {Colors.PRIMARY_HOVER};
        }}
        QPushButton.danger {{
            background-color: {Colors.DANGER_BG};
            color: {Colors.DANGER};
            border: 1px solid {Colors.DANGER};
        }}
    """

class Card(QFrame):
    """A standardized container with shadow and padding."""
    def __init__(self):
        super().__init__()
        self.setProperty("class", "card")
        
        # Shadow Effect
        from PyQt6.QtWidgets import QGraphicsDropShadowEffect
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 10))
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)

class HeaderLabel(QLabel):
    """Typography standard for Headers."""
    def __init__(self, text: str, size: int = 18):
        super().__init__(text)
        self.setStyleSheet(f"font-size: {size}px; font-weight: 800; color: {Colors.SECONDARY}; letter-spacing: 0.5px;")

class CaptionLabel(QLabel):
    """Typography standard for secondary text."""
    def __init__(self, text: str):
        super().__init__(text)
        self.setStyleSheet(f"font-size: 11px; font-weight: 600; color: {Colors.TEXT_MUTED}; text-transform: uppercase;")

class PrimaryButton(QPushButton):
    """Primary Action Button."""
    def __init__(self, text: str, icon_name: str = None):
        super().__init__(text)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setProperty("class", "primary")
        if icon_name:
            self.setIcon(QIcon.fromTheme(icon_name))

# ==============================================================================
# SECTION 6: WORKER THREADS
# ==============================================================================

class AsyncWorker(QRunnable):
    """
    Generic Worker for handling background tasks to prevent UI freezing.
    """
    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
            self.signals.result.emit(result)
        except Exception:
            self.signals.error.emit(traceback.format_exc())
        finally:
            self.signals.finished.emit()

class WorkerSignals(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    result = pyqtSignal(object)

# ==============================================================================
# SECTION 7: VIEW CONTROLLERS (TABS)
# ==============================================================================

class ProfileTab(QWidget):
    """
    ENT-Grade Profile Management Tab.
    ---------------------------------
    Refinements:
    - Added 'Manager Name' field for dynamic email salutations.
    - Improved 'Save Silently' logic for the MainWindow auto-save feature.
    - Enhanced Input Validation (Email regex).
    - Modern industrial layout with live signature preview.
    """
    def __init__(self, db: DatabaseEngine):
        super().__init__()
        self.db = db
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)

        # 1. SCROLLABLE CONTAINER (Handles smaller screens)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setStyleSheet(f"background-color: {Colors.BACKGROUND};")
        
        self.content = QWidget()
        self.content.setStyleSheet(f"background-color: {Colors.BACKGROUND};")
        self.v_layout = QVBoxLayout(self.content)
        self.v_layout.setContentsMargins(0, 40, 0, 40)
        self.v_layout.setSpacing(25)

        # 2. HORIZONTAL CENTERING
        self.centering_box = QHBoxLayout()
        self.form_column = QWidget()
        self.form_column.setFixedWidth(650) # Standard enterprise form width
        self.column_layout = QVBoxLayout(self.form_column)
        self.column_layout.setSpacing(25)

        # --- SECTION A: HEADER ---
        header_stack = QVBoxLayout()
        header_stack.setSpacing(5)
        title = QLabel("Identity & Preferences")
        title.setStyleSheet(f"font-size: 30px; font-weight: 900; color: {Colors.SECONDARY};")
        subtitle = QLabel("Your credentials determine how reports are signed and dispatched.")
        subtitle.setStyleSheet(f"color: {Colors.TEXT_MUTED}; font-size: 14px;")
        header_stack.addWidget(title)
        header_stack.addWidget(subtitle)
        self.column_layout.addLayout(header_stack)

        # --- SECTION B: THE FORM CARD ---
        self.form_card = Card()
        f_layout = QVBoxLayout(self.form_card)
        f_layout.setContentsMargins(35, 35, 35, 35)
        f_layout.setSpacing(20)

        # Field Factory
        def add_input(label, placeholder, icon=""):
            lbl = QLabel(f"{icon} {label}".upper())
            lbl.setStyleSheet("font-size: 10px; font-weight: 800; color: #94a3b8; letter-spacing: 1px;")
            edit = QLineEdit()
            edit.setPlaceholderText(placeholder)
            edit.setMinimumHeight(50)
            edit.setStyleSheet(f"""
                QLineEdit {{ 
                    border: 1px solid {Colors.BORDER}; border-radius: 8px; 
                    padding: 10px 15px; font-size: 14px; background: #f8fafc;
                }}
                QLineEdit:focus {{ border: 2px solid {Colors.PRIMARY}; background: white; }}
            """)
            f_layout.addWidget(lbl)
            f_layout.addWidget(edit)
            return edit

        self.inp_name = add_input("Full Name", "e.g. Ibrahim Zekiyev", "👤")
        self.inp_dept = add_input("Department", "e.g. Operations Control Center", "🏢")
        self.inp_mgr_name = add_input("Manager Name", "e.g. Natig bey", "🤵")
        self.inp_email = add_input("Manager Email", "manager@company.com", "📧")

        self.chk_summary = QCheckBox("Include detailed HTML task summary in Outlook body")
        self.chk_summary.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_summary.setStyleSheet(f"font-weight: 800; color: {Colors.PRIMARY}; margin-top: 10px;")
        f_layout.addWidget(self.chk_summary)

        self.column_layout.addWidget(self.form_card)

        # --- SECTION C: LIVE PREVIEW ---
        self.preview_box = QFrame()
        self.preview_box.setStyleSheet(f"background: #f1f5f9; border-radius: 12px; border: 1px dashed #cbd5e1;")
        p_lyt = QVBoxLayout(self.preview_box)
        p_lyt.setContentsMargins(25, 25, 25, 25)
        
        p_header = QLabel("AUTOMATED SIGNATURE PREVIEW")
        p_header.setStyleSheet("font-size: 10px; font-weight: 900; color: #64748b; letter-spacing: 1px;")
        p_lyt.addWidget(p_header)
        
        self.lbl_preview = QLabel("Best Regards,\nUser Name\nDepartment")
        self.lbl_preview.setStyleSheet(f"font-size: 15px; color: {Colors.TEXT_MAIN}; margin-top: 10px; line-height: 150%;")
        p_lyt.addWidget(self.lbl_preview)
        self.column_layout.addWidget(self.preview_box)

        # --- SECTION D: ACTIONS ---
        self.btn_save = QPushButton("💾 UPDATE SYSTEM IDENTITY")
        self.btn_save.setMinimumHeight(60)
        self.btn_save.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_save.setStyleSheet(f"""
            QPushButton {{ 
                background: {Colors.PRIMARY}; color: white; font-weight: 900; 
                border-radius: 10px; font-size: 15px; border: none;
            }}
            QPushButton:hover {{ background: {Colors.PRIMARY_HOVER}; }}
        """)
        self.btn_save.clicked.connect(self.save_data)
        self.column_layout.addWidget(self.btn_save)

        self.column_layout.addStretch()
        
        # Assembly
        self.centering_box.addStretch()
        self.centering_box.addWidget(self.form_column)
        self.centering_box.addStretch()
        
        self.v_layout.addLayout(self.centering_box)
        self.scroll.setWidget(self.content)
        self.main_layout.addWidget(self.scroll)

        # Signals
        self.inp_name.textChanged.connect(self.update_preview)
        self.inp_dept.textChanged.connect(self.update_preview)

    def update_preview(self):
        name = self.inp_name.text().strip() or "User Name"
        dept = self.inp_dept.text().strip() or "Department"
        self.lbl_preview.setText(f"Best Regards,<br><b>{name}</b><br>{dept}")

    def load_data(self):
        p = self.db.get_profile()
        self.inp_name.setText(p.name)
        self.inp_dept.setText(p.department)
        # Handle cases where manager_name might not exist in older DB records
        self.inp_mgr_name.setText(getattr(p, 'manager_name', '')) 
        self.inp_email.setText(p.manager_email)
        self.chk_summary.setChecked(getattr(p, 'include_summary', True))
        self.update_preview()

    def save_data_silently(self):
        """Used by MainWindow to auto-save without showing popups."""
        name = self.inp_name.text().strip()
        if not name: return 

        profile = UserProfile(
            name=name,
            department=self.inp_dept.text().strip(),
            manager_name=self.inp_mgr_name.text().strip(), # <--- PASS THIS HERE
            manager_email=self.inp_email.text().strip(),
            include_summary=self.chk_summary.isChecked()
        )
        
        self.db.set_profile(profile)

    def save_data(self):
        """Standard save with UI feedback."""
        name = self.inp_name.text().strip()
        email = self.inp_email.text().strip()

        if not name or not email:
            QMessageBox.warning(self, "Data Incomplete", "Please provide at least your Name and Manager Email.")
            return

        if "@" not in email:
            QMessageBox.warning(self, "Invalid Email", "The Manager Email format appears incorrect.")
            return

        self.save_data_silently()
        
        # Visual Success Feedback
        self.btn_save.setText("✅ IDENTITY UPDATED")
        self.btn_save.setStyleSheet(self.btn_save.styleSheet().replace(Colors.PRIMARY, Colors.SUCCESS))
        QTimer.singleShot(2500, self._reset_save_button)

    def _reset_save_button(self):
        self.btn_save.setText("💾 UPDATE SYSTEM IDENTITY")
        self.btn_save.setStyleSheet(f"""
            QPushButton {{ 
                background: {Colors.PRIMARY}; color: white; font-weight: 900; 
                border-radius: 10px; font-size: 15px; border: none;
            }}
            QPushButton:hover {{ background: {Colors.PRIMARY_HOVER}; }}
        """)

class DateSelectorDialog(QDialog):
    """
    MANUAL IMPORT DIALOG
    Allows the user to type a date (YYYY-MM-DD) or a month (YYYY-MM).
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manual Data Import")
        self.setFixedWidth(350)
        
        layout = QVBoxLayout(self)
        
        header = HeaderLabel("Import Source", 14)
        layout.addWidget(header)
        
        # Helper text
        lbl_hint = QLabel("Type a Date (YYYY-MM-DD) or Month (YYYY-MM):")
        lbl_hint.setStyleSheet(f"color: {Colors.TEXT_MUTED}; font-size: 11px;")
        layout.addWidget(lbl_hint)
        
        # --- CHANGED TO MANUAL TEXT INPUT ---
        self.txt_date = QLineEdit()
        # Pre-fill with yesterday's date as a suggestion
        yesterday = QDate.currentDate().addDays(-1).toString(Qt.DateFormat.ISODate)
        self.txt_date.setText(yesterday)
        self.txt_date.setPlaceholderText("Ex: 2025-01-08")
        self.txt_date.setMinimumHeight(40)
        self.txt_date.setStyleSheet(f"""
            QLineEdit {{
                border: 2px solid {Colors.PRIMARY};
                border-radius: 6px;
                padding: 5px;
                font-size: 14px;
                font-weight: bold;
            }}
        """)
        layout.addWidget(self.txt_date)
        
        btn_layout = QHBoxLayout()
        self.btn_import = PrimaryButton("Import Data")
        self.btn_import.clicked.connect(self.accept)
        
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addWidget(self.btn_cancel)
        btn_layout.addWidget(self.btn_import)
        layout.addLayout(btn_layout)

    def get_selected_date_iso(self) -> str:
        """Returns whatever the user typed."""
        return self.txt_date.text().strip()

class MorningTab(QWidget):
    def __init__(self, db: DatabaseEngine, thread_pool: QThreadPool, main_window): 
        super().__init__()
        self.db = db
        self.thread_pool = thread_pool
        self.main_window = main_window  
        self.current_priority = "Medium"
        self.active_edit_id = None  
        
        # --- NEW: Track dates checked in this session ---
        self.processed_dates = set() 
        
        self.setup_ui()

    def get_context_date(self) -> str:
        """Fetch the date from the main window's global picker."""
        # Fix: Access the picker via the main_window reference passed in __init__
        return self.main_window.reporting_date_picker.date().toPyDate().isoformat()
    

    def get_date(self) -> str:
        return self.calendar.date().toPyDate().isoformat()
    
    def setup_ui(self):
        """Builds the Morning Planning Interface."""
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(15, 15, 15, 15)
        
        # --- 1. VISUAL HEALTH INDICATOR ---
        self.health_box = QFrame()
        self.health_box.setStyleSheet(f"background: {Colors.SURFACE}; border-radius: 10px; border: 1px solid {Colors.BORDER};")
        hb_layout = QHBoxLayout(self.health_box)
        
        gauge_label = QLabel("<b>DAILY READINESS GAUGE:</b>")
        gauge_label.setStyleSheet(f"color: {Colors.SECONDARY}; font-size: 11px;")
        hb_layout.addWidget(gauge_label)
        
        self.health_bar = QProgressBar()
        self.health_bar.setRange(0, 100)
        self.health_bar.setValue(0)
        self.health_bar.setFormat("%p% Shift Readiness")
        self.health_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 1px solid {Colors.BORDER}; border-radius: 5px;
                text-align: center; background-color: {Colors.BACKGROUND};
                height: 22px; color: {Colors.TEXT_MAIN}; font-weight: bold;
            }}
            QProgressBar::chunk {{ background-color: {Colors.ACCENT}; border-radius: 4px; }}
        """)
        hb_layout.addWidget(self.health_bar)
        main_layout.addWidget(self.health_box)

        # --- 2. TASK INPUT CARD ---
        input_card = Card()
        l_in = QVBoxLayout(input_card)
        
        h_layout = QHBoxLayout()
        h_layout.addWidget(HeaderLabel("Morning Deployment", 16))
        self.lbl_profile_status = QLabel()
        self.lbl_profile_status.setStyleSheet(f"color: {Colors.PRIMARY}; font-weight: bold; font-size: 12px;")
        h_layout.addStretch()
        h_layout.addWidget(self.lbl_profile_status)
        l_in.addLayout(h_layout)
        
        grid = QGridLayout()
        grid.setSpacing(10)
        
        self.txt_desc = QLineEdit()
        self.txt_desc.setPlaceholderText("Describe the activity (e.g., Weekly Safety Inspection...)")
        
        self.time_picker = QLineEdit()
        self.time_picker.setPlaceholderText("e.g. 17:00, ASAP, or Morning")
        self.time_picker.setText("17:00") # Default value
        
        self.txt_dep = QLineEdit()
        self.txt_dep.setPlaceholderText("Dependencies (e.g., Permit #102, PPE...)")
        
        self.txt_note = QLineEdit()
        self.txt_note.setPlaceholderText("Additional context for this task...")
        
        # Priority Buttons
        p_layout = QHBoxLayout()
        self.p_btns = []
        for p in Config.PRIORITY_OPTS:
            b = QPushButton(p)
            b.setCheckable(True)
            b.setFixedWidth(85)
            b.clicked.connect(lambda c, x=p: self.set_priority(x))
            self.p_btns.append(b)
            p_layout.addWidget(b)
        self.set_priority("Medium") 
        
        grid.addWidget(CaptionLabel("Activity Description"), 0, 0)
        grid.addWidget(self.txt_desc, 1, 0, 1, 2)
        grid.addWidget(CaptionLabel("Planned time"), 0, 2)
        grid.addWidget(self.time_picker, 1, 2)
        grid.addWidget(CaptionLabel("Urgency"), 2, 0)
        grid.addLayout(p_layout, 3, 0)
        grid.addWidget(CaptionLabel("Operational Dependencies"), 2, 1)
        grid.addWidget(self.txt_dep, 3, 1)
        grid.addWidget(CaptionLabel("Technical Notes"), 2, 2)
        grid.addWidget(self.txt_note, 3, 2)
        l_in.addLayout(grid)
        
        self.btn_add = PrimaryButton("+ Add to Today's Plan")
        self.btn_add.setMinimumHeight(45)
        self.btn_add.clicked.connect(self.add_task)
        l_in.addWidget(self.btn_add)
        main_layout.addWidget(input_card)
        
        # --- 3. TASK TABLE ---
        table_container = QHBoxLayout() # Container for Table + Order Buttons

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["ID", "Description", "Priority", "Planned Time", "Status"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table_container.addWidget(self.table)

        # NEW: Reordering Buttons Side Panel
        btn_box = QVBoxLayout()
        btn_box.addStretch()
        
        self.btn_up = QPushButton("▲")
        self.btn_up.setFixedWidth(30)
        self.btn_up.setToolTip("Move Task Up")
        self.btn_up.clicked.connect(self.move_row_up)
        
        self.btn_down = QPushButton("▼")
        self.btn_down.setFixedWidth(30)
        self.btn_down.setToolTip("Move Task Down")
        self.btn_down.clicked.connect(self.move_row_down)
        
        btn_box.addWidget(self.btn_up)
        btn_box.addWidget(self.btn_down)
        btn_box.addStretch()
        
        table_container.addLayout(btn_box)
        main_layout.addLayout(table_container)
        # Interactions
        self.table.itemClicked.connect(self.load_task_to_edit)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        
        main_layout.addWidget(self.table)
        
        # --- 4. FOOTER CONTROLS ---
        f_layout = QHBoxLayout()
        
        # This replaces the old "Paste Previous" button
        self.btn_paste = QPushButton("📋 Import From Specific Date")
        self.btn_paste.setMinimumHeight(40)
        self.btn_paste.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_paste.setStyleSheet(f"border: 1px dashed {Colors.PRIMARY}; color: {Colors.PRIMARY}; font-weight: bold;")
        self.btn_paste.clicked.connect(self.paste_previous_plan)
        
        self.btn_export_excel = QPushButton("💾 Generate Morning Excel")
        self.btn_export_excel.setMinimumHeight(40)
        self.btn_export_excel.clicked.connect(self.export_to_local_folder)
        
        self.btn_email = QPushButton("📧 Dispatch Morning Email")
        self.btn_email.setMinimumHeight(40)
        self.btn_email.setProperty("class", "primary")
        self.btn_email.clicked.connect(self.draft_email)
        
        f_layout.addWidget(self.btn_paste)
        f_layout.addWidget(self.btn_export_excel)
        f_layout.addWidget(self.btn_email)
        f_layout.addStretch()
        
        main_layout.addLayout(f_layout)
        self.refresh_data()

    def move_row_up(self):
        row = self.table.currentRow()
        if row <= 0: return # Can't move top row up
        
        curr_id = self.table.item(row, 0).text()
        prev_id = self.table.item(row - 1, 0).text()
        date_iso = self.get_context_date()
        
        # Swap in DB
        self.db.swap_task_positions(curr_id, prev_id, date_iso)
        
        # Refresh and keep selection
        self.refresh_data()
        self.table.selectRow(row - 1)

    def move_row_down(self):
        row = self.table.currentRow()
        if row < 0 or row >= self.table.rowCount() - 1: return # Can't move bottom row down
        
        curr_id = self.table.item(row, 0).text()
        next_id = self.table.item(row + 1, 0).text()
        date_iso = self.get_context_date()
        
        # Swap in DB
        self.db.swap_task_positions(curr_id, next_id, date_iso)
        
        # Refresh and keep selection
        self.refresh_data()
        self.table.selectRow(row + 1)

    def refresh_data(self):
        """Syncs the UI with the database."""
        # 1. Save current scroll position and selection
        current_scroll = self.table.verticalScrollBar().value()
        selected_row = -1
        if self.table.selectionModel().hasSelection():
            selected_row = self.table.currentRow()

        p = self.db.get_profile()
        report_date = self.get_context_date()
        
        self.lbl_profile_status.setText(f"Operator: {p.name} | Date: {report_date}")
        
        # --- EXECUTE AUTO-COPY FIRST ---
        self._execute_auto_carry_over(report_date)
        # -------------------------------

        tasks = self.db.get_tasks(report_date) 
        total = len(tasks)
        
        accomplished = len([t for t in tasks if t.status in ["Completed", "In Progress"]])
        val = int((accomplished / total) * 100) if total > 0 else 0
        self.health_bar.setValue(val)
        
        self.table.setRowCount(0)
        for r, t in enumerate(tasks):
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(t.id))
            self.table.setItem(r, 1, QTableWidgetItem(t.description))
            self.table.setItem(r, 2, QTableWidgetItem(t.priority))
            self.table.setItem(r, 3, QTableWidgetItem(t.planned_time))
            self.table.setItem(r, 4, QTableWidgetItem(t.status))

        # 2. Restore Selection and Scroll
        if selected_row >= 0 and selected_row < self.table.rowCount():
            self.table.selectRow(selected_row)
        
        self.table.verticalScrollBar().setValue(current_scroll)

    def add_task(self):
        """Unified Create/Update handler."""
        desc = self.txt_desc.text().strip()
        if not desc:
            return

        p = self.db.get_profile()
        report_date = self.get_context_date()

        if self.active_edit_id:
            # UPDATE MODE
            tasks = self.db.get_tasks(report_date)
            task = next((t for t in tasks if t.id == self.active_edit_id), None)
            if task:
                task.description = desc
                task.priority = self.current_priority
                task.planned_time = self.time_picker.text()
                task.dependencies = self.txt_dep.text()
                task.morning_notes = self.txt_note.text()
                self.db.upsert_task(task)
        else:
            # CREATE MODE
            today_tasks = self.db.get_tasks(report_date)
            # Safe ID parsing to prevent crashes from malformed data
            ids = []
            for t in today_tasks:
                parts = t.id.split('-')
                if len(parts) > 1 and parts[1].isdigit():
                    ids.append(int(parts[1]))
            next_id = max(ids) + 1 if ids else 1
            
            task = TaskItem(
                id=f"T-{next_id:03d}",
                date=report_date,
                employee=p.name,
                department=p.department,
                description=desc,
                priority=self.current_priority,
                planned_time=self.time_picker.text(),
                dependencies=self.txt_dep.text(),
                morning_notes=self.txt_note.text()
            )
            self.db.upsert_task(task)

        self.reset_input_fields()
        self.refresh_data()

    def load_task_to_edit(self, item):
        """Populates fields for editing when a row is clicked."""
        row = item.row()
        task_id = self.table.item(row, 0).text()
        tasks = self.db.get_tasks(self.get_context_date())
        task = next((t for t in tasks if t.id == task_id), None)
        
        if task:
            self.active_edit_id = task.id
            self.txt_desc.setText(task.description)
            self.txt_dep.setText(task.dependencies)
            self.txt_note.setText(task.morning_notes)
            self.set_priority(task.priority)
            self.time_picker.setText(task.planned_time)
            
            self.btn_add.setText(f"💾 Update Task {task.id}")
            self.btn_add.setStyleSheet(f"background-color: {Colors.WARNING}; color: white; font-weight: bold;")

    def handle_async_error(self, error_message):
            """Triggered if the background Excel or Outlook logic fails."""
            logging.error(f"Async Error: {error_message}")
            QMessageBox.critical(
                self, 
                "Operation Failed", 
                f"An error occurred while generating the report:\n\n{error_message[:200]}..."
            )
            
    def reset_input_fields(self):
        self.active_edit_id = None
        self.txt_desc.clear()
        self.txt_dep.clear()
        self.txt_note.clear()
        self.btn_add.setText("+ Add to Today's Plan")
        self.btn_add.setStyleSheet("") 
        self.time_picker.setText("17:00")


    def show_context_menu(self, pos):
        index = self.table.indexAt(pos)
        if not index.isValid(): return
        
        menu = QMenu()
        del_action = menu.addAction("❌ Delete Task")
        action = menu.exec(self.table.viewport().mapToGlobal(pos))
        if action == del_action:
            self.confirm_and_delete_row(index.row())

    def confirm_and_delete_row(self, row):
        task_id = self.table.item(row, 0).text()
        report_date = self.get_context_date()
        msg = f"Permanently delete task {task_id} from {report_date}?"
        if QMessageBox.question(self, "Confirm Delete", msg) == QMessageBox.StandardButton.Yes:
            self.db.delete_task(task_id, report_date)
            self.refresh_data()

    def set_priority(self, val):
        self.current_priority = val
        for b in self.p_btns:
            active = (b.text() == val)
            b.setChecked(active)
            if active:
                b.setStyleSheet(f"background-color: {Colors.SUCCESS_BG}; color: {Colors.PRIMARY}; border: 2px solid {Colors.PRIMARY}; font-weight: bold;")
            else:
                b.setStyleSheet("")

    def paste_previous_plan(self):
        """Opens a dialog to manually type a date/month and imports tasks."""
        # 1. Show the manual text dialog
        dlg = DateSelectorDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        source_input = dlg.get_selected_date_iso() # This is now raw text
        current_sel_date = self.get_context_date() 

        # Validation: Ensure they actually typed something
        if not source_input:
            QMessageBox.warning(self, "Empty Input", "Please type a date or month.")
            return

        # 2. Fetch tasks (Uses LIKE, so "2025-01" gets the whole month)
        source_tasks = self.db.get_tasks(source_input)

        if not source_tasks:
            QMessageBox.information(self, "No Data", f"No tasks found matching: '{source_input}'.")
            return

        # 3. Confirm with user
        msg = f"Found {len(source_tasks)} tasks matching '{source_input}'. Import them into your {current_sel_date} plan?"
        if QMessageBox.question(self, "Confirm Import", msg) == QMessageBox.StandardButton.No:
            return

        # 4. Import logic (Standard logic continues below...)
        today_tasks = self.db.get_tasks(current_sel_date)
        existing_desc = {t.description.strip().lower() for t in today_tasks}
        p = self.db.get_profile()
        
        # Calculate starting ID number
        ids = [int(t.id.split('-')[1]) for t in today_tasks if '-' in t.id]
        next_id_num = max(ids) + 1 if ids else 1

        added_count = 0
        for old_t in source_tasks:
            # Skip duplicates
            if old_t.description.strip().lower() in existing_desc:
                continue

            new_task = TaskItem(
                id=f"T-{next_id_num:03d}",
                date=current_sel_date, 
                employee=p.name, 
                department=p.department,
                description=old_t.description, 
                priority=old_t.priority,
                planned_time=old_t.planned_time, 
                dependencies=old_t.dependencies,
                
                # FIX 1: Changed {source_date} to {source_input}
                morning_notes=f"Imported from {source_input}", 
                
                status="Planned"
            )
            self.db.upsert_task(new_task)
            next_id_num += 1
            added_count += 1

        self.refresh_data()
        
        # FIX 2: Changed {source_date} to {source_input}
        QMessageBox.information(self, "Import Success", f"Imported {added_count} tasks from {source_input}.")

    def export_to_local_folder(self):
        report_date = self.get_context_date()
        tasks = self.db.get_tasks(report_date)
        if not tasks: return

        filename = f"Morning_Report_{report_date}.xlsx"
        full_path = os.path.join(Config.REPORTS_DIR, filename)

        if ExcelEngine.generate_report(full_path, tasks, "morning"):
            QMessageBox.information(self, "Success", f"Report saved:\n{full_path}")
            os.startfile(Config.REPORTS_DIR)

    def _execute_auto_carry_over(self, report_date):
        """
        INTERNAL: Automatically copies ONLY 'In Progress' tasks from the last working day.
        SAFETY GUARD: ONLY RUNS IF THE REPORT DATE IS ACTUALLY TODAY.
        """
        # --- 1. CRITICAL SAFETY: ONLY RUN FOR TODAY ---
        # Get the actual computer date
        today_iso = date.today().isoformat()
        
        # If the user is looking at history or a future date, STOP immediately.
        if report_date != today_iso:
            return 

        # --- 2. SESSION GUARD ---
        # Prevent running twice for the same date in one session
        if report_date in self.processed_dates:
            return 
        self.processed_dates.add(report_date)

        # --- 3. GET PREVIOUS DATA ---
        prev_tasks = self.db.get_previous_available_tasks(report_date)
        if not prev_tasks: 
            return

        # --- 4. FILTER FOR 'IN PROGRESS' ---
        ongoing_tasks = [t for t in prev_tasks if t.status == "In Progress"]
        if not ongoing_tasks: 
            return

        # --- 5. DUPLICATE PREVENTION ---
        current_tasks = self.db.get_tasks(report_date)
        existing_descriptions = {t.description.strip().lower() for t in current_tasks}

        p = self.db.get_profile()
        
        # Calculate next ID
        ids = [int(t.id.split('-')[1]) for t in current_tasks if '-' in t.id and t.id.split('-')[1].isdigit()]
        next_id_num = (max(ids) + 1) if ids else 1
        
        imported_count = 0
        
        for old_t in ongoing_tasks:
            # SKIP if we already have this task today
            if old_t.description.strip().lower() in existing_descriptions:
                continue

            new_task = TaskItem(
                id=f"T-{next_id_num:03d}",      # New ID
                date=report_date,               # Today's Date
                employee=p.name,                # Current User
                department=p.department,        # Current Dept
                
                # --- EXACT COPY OF ALL DATA ---
                description=old_t.description,
                priority=old_t.priority,
                planned_time=old_t.planned_time,
                dependencies=old_t.dependencies,
                morning_notes=old_t.morning_notes,
                status=old_t.status,                     
                percent_complete=old_t.percent_complete, 
                results=old_t.results,                   
                reason_delay=old_t.reason_delay,         
                next_steps=old_t.next_steps              
            )
            self.db.upsert_task(new_task)
            next_id_num += 1
            imported_count += 1
            
        # 6. REFRESH UI
        if imported_count > 0:
            if hasattr(self.main_window, 'status'):
                self.main_window.status.showMessage(f"🔄 Shift Handover: {imported_count} tasks carried over to Today.", 5000)
            self.refresh_data()

    def draft_email(self):
        report_date = self.get_context_date()
        tasks = self.db.get_tasks(report_date)
        if not tasks: return
        
        p = self.db.get_profile()
        fpath = os.path.join(Config.REPORTS_DIR, f"Morning_Report_{report_date}.xlsx")
        
        self.btn_email.setEnabled(False)
        self.btn_email.setText("Drafting Outlook...")
        
        worker = AsyncWorker(self._email_worker, fpath, tasks, p, report_date)
        worker.signals.error.connect(self.handle_async_error) 

        worker.signals.finished.connect(lambda: (
            self.btn_email.setEnabled(True), 
            self.btn_email.setText("📧 Dispatch Morning Email")
        ))
        self.thread_pool.start(worker)

    def _email_worker(self, fpath, tasks, profile, report_date):
        """Background worker for Morning Deployment dispatch."""
        
        # 1. FILE-LOCK PROTECTION
        # Ensures the app doesn't crash if the user has the Morning Excel open
        if os.path.exists(fpath):
            try:
                with open(fpath, 'a'): pass
            except OSError:
                raise Exception(f"DISPATCH BLOCKED: The file '{os.path.basename(fpath)}' is open in Excel. "
                                f"Please close it so the system can update the report.")

        # 2. GENERATE MORNING EXCEL
        ExcelEngine.generate_report(fpath, tasks, "morning")
        
        # 3. CONSTRUCT HTML SUMMARY
        items_html = ""
        if getattr(profile, 'include_summary', True):
            # Using <li> for a clean numbered list of planned activities
            rows = "".join([f"<li style='margin-bottom:6px;'>{t.description}</li>" for t in tasks])
            items_html = f"<p><b>Summary of planned activities:</b></p><ol>{rows}</ol>"
        
        # 4. DYNAMIC SALUTATION LOGIC
        # Pulls from the 'Manager Name' field in the Profile Tab
        mgr_name = getattr(profile, 'manager_name', '').strip()
        if not mgr_name:
            mgr_name = "Manager" # Professional fallback

        # 5. ASSEMBLE PRO-GRADE HTML BODY
        html_body = f"""
        <html>
        <body style="font-family: 'Segoe UI', Calibri, sans-serif; font-size: 11pt; color: #1e293b;">
            <p>Dear {mgr_name},</p>
            <p>Please find attached my planned deployment and activity schedule for <b>{report_date}</b>.</p>
            
            {items_html}
            
            <br>
            <p>Best Regards,<br>
            <span>{profile.name}</span><br>
        </body>
        </html>
        """

        # 6. DISPATCH DRAFT TO OUTLOOK
        OutlookEngine.create_draft(
            profile.manager_email, 
            f"Morning Planned Tasks - {report_date}", 
            html_body, 
            [fpath]
        )
    
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar

class EveningTab(QWidget):
    """
    TITANIUM EDITION: Evening Update Terminal (OCC Optimized)
    -------------------------------------------------------
    Refinements:
    - Dynamic Manager Salutation (No longer hardcoded).
    - File-Lock Detection (Prevents crashes if Excel is open).
    - Precise Efficiency KPI Calculation.
    - Safe Async Dispatch (Protects UI thread).
    """
    def __init__(self, db: DatabaseEngine, thread_pool: QThreadPool, main_window):
        super().__init__()
        self.db = db
        self.thread_pool = thread_pool
        self.main_window = main_window 
        self.active_task_id = None
        
        # --- AUTO-SAVE ENGINE (Debounced) ---
        self.save_timer = QTimer()
        self.save_timer.setSingleShot(True)
        self.save_timer.timeout.connect(self.perform_auto_save)
        
        self.setup_ui()

    def get_context_date(self) -> str:
        """Fetch the reporting date from the Global Main Window context."""
        return self.main_window.reporting_date_picker.date().toPyDate().isoformat()

    def setup_ui(self):
        """Final Compact Build: Tightened spacing and optimized horizontal alignment."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(5) # Reduced from 15

        # --- 1. COMPACT KPI HEADER ---
        self.health_card = Card()
        self.health_card.setFixedHeight(75) # Shorter header
        hc_layout = QHBoxLayout(self.health_card)
        hc_layout.setContentsMargins(15, 0, 15, 0)
        
        title_stack = QVBoxLayout()
        title_stack.setSpacing(0)
        title_stack.addWidget(QLabel("<b>TECHNICAL CLOSURE TERMINAL</b>"))
        self.lbl_shift_status = QLabel("OCC System Monitor | Data Integrity Active")
        self.lbl_shift_status.setStyleSheet("font-size: 10px; color: #64748b;")
        title_stack.addWidget(self.lbl_shift_status)
        hc_layout.addLayout(title_stack)

        hc_layout.addStretch()

        stat_stack = QVBoxLayout()
        stat_stack.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.lbl_pct = QLabel("0% EFFICIENCY")
        self.lbl_pct.setStyleSheet(f"font-weight: 800; color: {Colors.PRIMARY}; font-size: 13px;")
        self.health_bar = QProgressBar()
        self.health_bar.setFixedWidth(180); self.health_bar.setFixedHeight(6); self.health_bar.setTextVisible(False)
        stat_stack.addWidget(self.lbl_pct); stat_stack.addWidget(self.health_bar)
        hc_layout.addLayout(stat_stack)
        main_layout.addWidget(self.health_card)

        # --- 2. MASTER-DETAIL BODY ---
        content_split = QHBoxLayout()
        content_split.setSpacing(8) # Reduced distance between Left Table and Right Editor
        
        # LEFT RAIL (Activity Selection)
        rail_container = QWidget()
        rail_container.setFixedWidth(280) # Narrower rail
        rail_lyt = QVBoxLayout(rail_container)
        rail_lyt.setContentsMargins(0, 0, 0, 0)
        rail_lyt.addWidget(CaptionLabel("📂 Shift Activity Log"))
        
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["ID", "Description"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.itemClicked.connect(self.on_row_selected)
        rail_lyt.addWidget(self.table)
        content_split.addWidget(rail_container)

        # RIGHT EDITOR (The Tight Form)
        self.editor_card = Card()
        ed_layout = QVBoxLayout(self.editor_card)
        ed_layout.setContentsMargins(20, 15, 20, 15) # Tight margins
        ed_layout.setSpacing(8) # Tightened vertical distance between inputs

        # Identity Area
        id_row = QHBoxLayout()
        self.lbl_active_id = QLabel("NO SELECTION")
        self.lbl_active_id.setStyleSheet("font-weight: 800; color: #94a3b8; font-size: 9px;")
        self.lbl_save_status = QLabel("")
        self.lbl_save_status.setStyleSheet(f"color: {Colors.SUCCESS}; font-size: 10px; font-weight: bold;")
        id_row.addWidget(self.lbl_active_id); id_row.addStretch(); id_row.addWidget(self.lbl_save_status)
        ed_layout.addLayout(id_row)

        self.lbl_active_desc = QLabel("Please select a task from the log...")
        self.lbl_active_desc.setStyleSheet("font-size: 16px; font-weight: 700; color: #1e293b;")
        self.lbl_active_desc.setWordWrap(True)
        ed_layout.addWidget(self.lbl_active_desc)

        # Input Row 1: Status & Progress
        row1 = QHBoxLayout()
        row1.setSpacing(15)
        
        stat_v = QVBoxLayout()
        stat_v.setSpacing(4)
        stat_v.addWidget(CaptionLabel("Current Status"))
        self.cmb_status = QComboBox()
        self.cmb_status.addItems(Config.STATUS_OPTS); self.cmb_status.setFixedHeight(35)
        stat_v.addWidget(self.cmb_status)
        
        prog_v = QVBoxLayout()
        prog_v.setSpacing(4)
        prog_v.addWidget(CaptionLabel("Execution Progress"))
        prog_h = QHBoxLayout()
        self.sld_progress = QSlider(Qt.Orientation.Horizontal)
        self.spn_progress = QSpinBox(); self.spn_progress.setRange(0, 100); self.spn_progress.setSuffix("%")
        self.spn_progress.setFixedHeight(35); self.spn_progress.setFixedWidth(65)
        prog_h.addWidget(self.sld_progress); prog_h.addWidget(self.spn_progress)
        prog_v.addLayout(prog_h)
        
        row1.addLayout(stat_v, 1); row1.addLayout(prog_v, 1)
        ed_layout.addLayout(row1)

        # Input Group 2: Results Area
        res_v = QVBoxLayout()
        res_v.setSpacing(4)
        res_v.addWidget(CaptionLabel("📝 Technical Findings"))
        self.txt_results = QTextEdit()
        self.txt_results.setPlaceholderText("Enter technical findings here...")
        self.txt_results.setMinimumHeight(100) # Balanced height
        self.txt_results.setStyleSheet("border-radius: 6px; padding: 8px; background: white;")
        res_v.addWidget(self.txt_results)
        ed_layout.addLayout(res_v, stretch=1) 

        # Input Group 3: Contingencies (Side-by-Side)
        row2 = QHBoxLayout()
        row2.setSpacing(15)
        
        delay_v = QVBoxLayout()
        delay_v.setSpacing(4)
        delay_v.addWidget(CaptionLabel("⚠️ Root Cause of Delay"))
        self.txt_delay = QLineEdit(); self.txt_delay.setFixedHeight(35); self.txt_delay.setPlaceholderText("Bottlenecks...")
        delay_v.addWidget(self.txt_delay)
        
        next_v = QVBoxLayout()
        next_v.setSpacing(4)
        next_v.addWidget(CaptionLabel("⏭️ Next Steps / Recovery"))
        self.txt_next = QLineEdit(); self.txt_next.setFixedHeight(35); self.txt_next.setPlaceholderText("Immediate actions...")
        next_v.addWidget(self.txt_next)
        
        row2.addLayout(delay_v, 1); row2.addLayout(next_v, 1)
        ed_layout.addLayout(row2)

        content_split.addWidget(self.editor_card, stretch=1)
        main_layout.addLayout(content_split)

        # --- 3. COMPACT FOOTER ---
        footer = QHBoxLayout()
        self.btn_email = PrimaryButton("📧 DISPATCH SHIFT CLOSURE")
        self.btn_email.setFixedWidth(240); self.btn_email.setFixedHeight(42)
        self.btn_email.setStyleSheet(self.btn_email.styleSheet() + "font-size: 12px; font-weight: bold;")
        self.btn_email.clicked.connect(self.draft_email)
        footer.addStretch(); footer.addWidget(self.btn_email)
        main_layout.addLayout(footer)

        # Logic
        self.sld_progress.valueChanged.connect(self.spn_progress.setValue)
        self.spn_progress.valueChanged.connect(self.sld_progress.setValue)
        self.cmb_status.currentTextChanged.connect(self.on_status_change) 
        self._bind_auto_save()
        self.refresh_list()
        
    def on_status_change(self, text):
        """Handle visual changes and ENFORCE 100% logic."""
        is_delayed = (text == "Delayed")
        self.txt_delay.setEnabled(is_delayed)
        self.txt_delay.setStyleSheet("background: #fff1f2; border: 1px solid #e11d48;" if is_delayed else "")
        
        # --- NEW LOGIC: Force 100% if Completed ---
        if text == "Completed":
            self.spn_progress.setValue(100)
            self.sld_progress.setValue(100)
            # Disable progress editing when completed to prevent errors? Optional.
            self.spn_progress.setEnabled(False)
            self.sld_progress.setEnabled(False)
        else:
            # Re-enable if moved away from Completed
            self.spn_progress.setEnabled(True)
            self.sld_progress.setEnabled(True)
            
        # Trigger auto-save immediately to ensure DB gets the 100%
        self.trigger_save_timer()

    def _bind_auto_save(self):
        """Connects all input widgets to the debounced timer."""
        for widget in [self.txt_results, self.txt_delay, self.txt_next]:
            widget.textChanged.connect(self.trigger_save_timer)
        self.cmb_status.currentIndexChanged.connect(self.trigger_save_timer)
        self.spn_progress.valueChanged.connect(self.trigger_save_timer)

    def trigger_save_timer(self):
        if not self.active_task_id: return
        self.lbl_save_status.setText("● Local Syncing...")
        self.save_timer.start(1200) # 1.2 second debounce

    def perform_auto_save(self):
        """Commits current technical findings to the DB silently."""
        if not self.active_task_id: return
        
        report_date = self.get_context_date()
        tasks = self.db.get_tasks(report_date)
        task = next((t for t in tasks if t.id == self.active_task_id), None)
        
        if task:
            task.status = self.cmb_status.currentText()
            task.percent_complete = self.spn_progress.value()
            task.results = self.txt_results.toPlainText()
            task.reason_delay = self.txt_delay.text()
            task.next_steps = self.txt_next.text()
            
            self.db.upsert_task(task) 
            self.lbl_save_status.setText("✓ Cloud Synced")
            self.update_kpis(tasks)
            self.refresh_rail_visuals()

    def on_row_selected(self, item):
        """Loads task into editor while forcing a save of the previous task."""
        if self.save_timer.isActive():
            self.save_timer.stop()
            self.perform_auto_save()

        task_id = self.table.item(item.row(), 0).data(Qt.ItemDataRole.UserRole)
        report_date = self.get_context_date()
        
        tasks = self.db.get_tasks(report_date)
        task = next((t for t in tasks if t.id == task_id), None)
        
        if task:
            self._block_signals(True)
            self.active_task_id = task.id
            self.lbl_active_id.setText(f"DATABASE REFERENCE: {task.id}")
            self.lbl_active_desc.setText(task.description)
            self.cmb_status.setCurrentText(task.status)
            self.spn_progress.setValue(task.percent_complete)
            self.txt_results.setText(task.results)
            self.txt_delay.setText(task.reason_delay)
            self.txt_next.setText(task.next_steps)
            self.on_status_change(task.status)
            
            self.lbl_save_status.setText("✓ Data Fetched")
            self._block_signals(False)

    def refresh_list(self):
        """Rebuilds the left selection rail."""
        self.table.setRowCount(0)
        report_date = self.get_context_date()
        tasks = self.db.get_tasks(report_date)
        
        self.update_kpis(tasks)

        for r, t in enumerate(tasks):
            self.table.insertRow(r)
            id_item = QTableWidgetItem(t.id)
            id_item.setData(Qt.ItemDataRole.UserRole, t.id)
            id_item.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            self.table.setItem(r, 0, id_item)
            self.table.setItem(r, 1, QTableWidgetItem(t.description))
        self.refresh_rail_visuals()

    def update_kpis(self, tasks):
        total = len(tasks)
        if total == 0:
            self.health_bar.setValue(0)
            self.lbl_pct.setText("0% EFFICIENCY")
            self.lbl_shift_status.setText("No activities registered for this date.")
            return

        completed = len([t for t in tasks if t.status == "Completed"])
        efficiency = int((completed / total) * 100)
        self.health_bar.setValue(efficiency)
        self.lbl_pct.setText(f"{efficiency}% SHIFT EFFICIENCY")
        self.lbl_shift_status.setText(f"{completed} of {total} activities reported as finalized.")

    def on_status_change(self, text):
        is_delayed = (text == "Delayed")
        self.txt_delay.setEnabled(is_delayed)
        self.txt_delay.setStyleSheet("background: #fff1f2; border: 1px solid #e11d48;" if is_delayed else "")
        if text == "Completed":
            self.spn_progress.setValue(100)

    def refresh_rail_visuals(self):
        """Colors task IDs in the list based on their status."""
        for r in range(self.table.rowCount()):
            tid = self.table.item(r, 0).data(Qt.ItemDataRole.UserRole)
            # Find status from DB
            tasks = self.db.get_tasks(self.get_context_date())
            t = next((task for task in tasks if task.id == tid), None)
            if t:
                color = Colors.TEXT_MAIN
                if t.status == "Completed": color = Colors.SUCCESS
                elif t.status == "Delayed": color = Colors.DANGER
                elif t.status == "In Progress": color = Colors.WARNING
                self.table.item(r, 0).setForeground(QBrush(QColor(color)))

    def _block_signals(self, block: bool):
        for w in [self.txt_results, self.txt_delay, self.txt_next, self.cmb_status, self.spn_progress]:
            w.blockSignals(block)

    def draft_email(self):
        """Initiates async Excel generation and Outlook dispatch."""
        report_date = self.get_context_date()
        tasks = self.db.get_tasks(report_date)
        if not tasks:
            QMessageBox.warning(self, "No Data", "Cannot dispatch an empty shift report.")
            return

        p = self.db.get_profile()
        fpath = os.path.join(Config.REPORTS_DIR, f"Evening_Report_{report_date}.xlsx")
        
        self.btn_email.setEnabled(False)
        self.btn_email.setText("💾 GENERATING EXCEL...")
        
        worker = AsyncWorker(self._email_worker, fpath, tasks, p, report_date)
        worker.signals.error.connect(self._handle_error)
        worker.signals.finished.connect(self._reset_button)
        self.thread_pool.start(worker)

    def _email_worker(self, fpath, tasks, profile, report_date):
        # 1. FILE-LOCK PROTECTION (Ensures we don't crash if Excel is open)
        if os.path.exists(fpath):
            try:
                # Attempt to open file in append mode; fails if locked by Excel
                with open(fpath, 'a'): pass
            except OSError:
                raise Exception(f"REPORT BLOCKED: The file '{os.path.basename(fpath)}' is open in Excel. "
                                f"Please close it and try again.")

        # 2. GENERATE THE EXCEL REPORT
        ExcelEngine.generate_report(fpath, tasks, "evening")
        
        # 3. BUILD THE HTML EMAIL BODY
        summary = ""
        if getattr(profile, 'include_summary', True):
            # Create a clean bulleted list of tasks and their status
            rows = "".join([f"<li style='margin-bottom:5px;'>{t.description} — <b>{t.status}</b></li>" for t in tasks])
            summary = f"<p><b>Shift Task Summary:</b></p><ul>{rows}</ul>"

        # DYNAMIC SALUTATION FIX:
        # We now pull the manager name directly from the profile. 
        # Fallback to "Manager" if the user left the input field blank.
        mgr_salutation = getattr(profile, 'manager_name', '').strip()
        if not mgr_salutation:
            mgr_salutation = "Manager"

        body = f"""
        <html>
        <body style="font-family: 'Segoe UI', Calibri, sans-serif; font-size: 11pt; color: #1e293b;">
            <p>Dear {mgr_salutation},</p>
            <p>Please find the shift technical closure report for <b>{report_date}</b> attached for your review.</p>
            {summary}
            <br>
            <p>Best Regards,<br>
            <span><b>{profile.name}</b></span><br>
        </body>
        </html>
        """
        
        # 4. DISPATCH TO OUTLOOK
        OutlookEngine.create_draft(
            profile.manager_email, 
            f"Evening Report - {report_date}", 
            body, 
            [fpath]
        )

    def _handle_error(self, err):
        QMessageBox.critical(self, "Dispatch Error", str(err))

    def _reset_button(self):
        self.btn_email.setEnabled(True)
        self.btn_email.setText(" 📧 DISPATCH SHIFT CLOSURE ")

        
class DashboardTab(QWidget):
    """
    ENT-Grade Analytics Dashboard.
    Features: 
    - Safe metric calculation (Crash-proof)
    - Dynamic Filtering (Date & Keyword)
    - Matplotlib Integration with High-DPI support
    - PNG Export functionality
    """
    def __init__(self, db: DatabaseEngine):
        super().__init__()
        self.db = db
        self.setup_ui()
        # Delay the first refresh slightly to ensure UI is painted
        QTimer.singleShot(500, self.refresh)

    def setup_ui(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # 1. SCROLLABLE CONTAINER
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setStyleSheet(f"background-color: {Colors.BACKGROUND};")
        
        self.content_widget = QWidget()
        self.content_widget.setObjectName("DashboardContent")
        self.content_widget.setStyleSheet(f"#DashboardContent {{ background-color: {Colors.BACKGROUND}; }}")
        self.v_layout = QVBoxLayout(self.content_widget)
        self.v_layout.setContentsMargins(30, 30, 30, 30)
        self.v_layout.setSpacing(25)

        # --- HEADER SECTION ---
        header_row = QHBoxLayout()
        title_stack = QVBoxLayout()
        title_stack.addWidget(HeaderLabel("Operational Intelligence", 28))
        
        self.lbl_sync = QLabel("Analyzing real-time shift data...")
        self.lbl_sync.setStyleSheet(f"color: {Colors.TEXT_MUTED}; font-size: 13px;")
        title_stack.addWidget(self.lbl_sync)
        header_row.addLayout(title_stack)
        
        header_row.addStretch()
        
        self.btn_pdf = QPushButton("📄 Export Range Report (PDF)")
        self.btn_pdf.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pdf.setMinimumHeight(45)
        self.btn_pdf.setStyleSheet(f"""
            QPushButton {{ 
                background: {Colors.PRIMARY}; border: none; 
                border-radius: 8px; font-weight: 800; color: white;
                padding: 0 20px;
            }}
            QPushButton:hover {{ background: {Colors.PRIMARY_HOVER}; }}
        """)
        self.btn_pdf.clicked.connect(self.export_analytics_pdf) # Connects to the function below
        header_row.addWidget(self.btn_pdf)
        
        # [EXISTING] PNG BUTTON (Optional: Keep or Remove)
        self.btn_png = QPushButton("📸 Snapshot")
        self.btn_png.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_png.setMinimumHeight(45)
        self.btn_png.setStyleSheet(f"background: white; border: 1px solid {Colors.BORDER}; border-radius: 8px;")
        self.btn_png.clicked.connect(self.export_visuals)
        header_row.addWidget(self.btn_png)

        self.v_layout.addLayout(header_row)

        # --- FILTER CONTROL BAR ---
        filter_card = QFrame()
        filter_card.setStyleSheet(f"background: white; border-radius: 12px; border: 1px solid {Colors.BORDER};")
        f_layout = QHBoxLayout(filter_card)
        f_layout.setContentsMargins(20, 15, 20, 15)
        
        f_layout.addWidget(CaptionLabel("📅 Range:"))
        self.date_start = QDateEdit(calendarPopup=True)
        self.date_start.setDate(QDate.currentDate().addDays(-14))
        self.date_end = QDateEdit(calendarPopup=True)
        self.date_end.setDate(QDate.currentDate())
        
        for d in [self.date_start, self.date_end]:
            d.setMinimumHeight(38)
            d.setFixedWidth(120)
            d.setStyleSheet(f"border: 1px solid {Colors.BORDER}; border-radius: 6px; padding: 5px;")
            d.dateChanged.connect(self.refresh)
        
        f_layout.addWidget(self.date_start)
        f_layout.addWidget(QLabel("-"))
        f_layout.addWidget(self.date_end)
        f_layout.addSpacing(30)
        
        f_layout.addWidget(CaptionLabel("🔍 Search:"))
        self.txt_filter = QLineEdit()
        self.txt_filter.setPlaceholderText("Filter logs by activity or technical keywords...")
        self.txt_filter.setMinimumHeight(38)
        self.txt_filter.setStyleSheet(f"border: 1px solid {Colors.BORDER}; border-radius: 19px; padding: 0 15px; background: #f8fafc;")
        self.txt_filter.textChanged.connect(self.refresh)
        f_layout.addWidget(self.txt_filter)
        
        self.v_layout.addWidget(filter_card)

        # --- KPI METRIC CARDS ---
        kpi_row = QHBoxLayout()
        kpi_row.setSpacing(20)
        
        self.card_vol = self.make_metric_card("Volume", "Total Shift Activities", Colors.SECONDARY)
        self.card_rate = self.make_metric_card("Success", "Avg. Completion Rate", Colors.SUCCESS)
        self.card_risk = self.make_metric_card("Bottleneck", "Critical Delay Unit", Colors.DANGER)
        
        kpi_row.addWidget(self.card_vol)
        kpi_row.addWidget(self.card_rate)
        kpi_row.addWidget(self.card_risk)
        self.v_layout.addLayout(kpi_row)

        # --- VISUALIZATION GRID ---
        viz_grid = QGridLayout()
        viz_grid.setSpacing(25)
        
        # 1. Timeline Chart (Full Width)
        self.fig_trend = Figure(figsize=(10, 4), dpi=100)
        self.can_trend = FigureCanvas(self.fig_trend)
        viz_grid.addWidget(self.wrap_visual("Activity Execution Timeline", self.can_trend), 0, 0, 1, 2)
        
        # 2. Status Breakdown (Half Width)
        self.fig_status = Figure(figsize=(5, 5), dpi=100)
        self.can_status = FigureCanvas(self.fig_status)
        viz_grid.addWidget(self.wrap_visual("Global Status Distribution", self.can_status), 1, 0)
        
        # 3. Workload Distribution (Half Width)
        self.fig_dept = Figure(figsize=(5, 5), dpi=100)
        self.can_dept = FigureCanvas(self.fig_dept)
        viz_grid.addWidget(self.wrap_visual("Departmental Loading (Active Tasks)", self.can_dept), 1, 1)
        
        self.v_layout.addLayout(viz_grid)
        self.v_layout.addStretch()

        self.scroll.setWidget(self.content_widget)
        self.main_layout.addWidget(self.scroll)

    def make_metric_card(self, category, label, color):
        """Creates a standardized KPI card with a value pointer."""
        card = QFrame()
        card.setGraphicsEffect(self._get_shadow())
        card.setStyleSheet(f"background: white; border-radius: 12px; border: 1px solid {Colors.BORDER};")
        
        l = QVBoxLayout(card)
        l.setContentsMargins(20, 18, 20, 18)
        
        cat_lbl = QLabel(category.upper())
        cat_lbl.setStyleSheet(f"color: {color}; font-weight: 900; font-size: 10px; letter-spacing: 1.5px;")
        
        val_lbl = QLabel("0")
        val_lbl.setStyleSheet(f"font-size: 32px; font-weight: 800; color: {Colors.SECONDARY}; margin: 4px 0;")
        
        desc_lbl = QLabel(label)
        desc_lbl.setStyleSheet(f"color: {Colors.TEXT_MUTED}; font-size: 11px; font-weight: 600;")
        
        l.addWidget(cat_lbl)
        l.addWidget(val_lbl)
        l.addWidget(desc_lbl)
        
        card.val_ptr = val_lbl # Store pointer for refresh updates
        return card

    def wrap_visual(self, title, canvas):
        """Wraps a Matplotlib canvas into a styled QFrame."""
        wrapper = QFrame()
        wrapper.setStyleSheet(f"background: white; border-radius: 12px; border: 1px solid {Colors.BORDER};")
        l = QVBoxLayout(wrapper)
        l.setContentsMargins(20, 20, 20, 20)
        
        head = QLabel(title.upper())
        head.setStyleSheet(f"font-weight: 800; font-size: 11px; color: {Colors.TEXT_MUTED}; letter-spacing: 1px;")
        l.addWidget(head)
        l.addSpacing(10)
        l.addWidget(canvas)
        canvas.setMinimumHeight(320)
        return wrapper

    def _get_shadow(self):
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 20))
        shadow.setOffset(0, 6)
        return shadow

    def refresh(self):
        """Main Data Engine: Fetches, Filters, and Redraws Visuals."""
        try:
            # 1. FETCH DATA
            tasks = self.db.get_tasks()
            if not tasks:
                self.lbl_sync.setText("No data available in system database.")
                return
            
            df = pd.DataFrame([t.to_dict() for t in tasks])
            df['date'] = pd.to_datetime(df['date'])
            
            # 2. APPLY FILTERS
            start = pd.Timestamp(self.date_start.date().toPyDate())
            end = pd.Timestamp(self.date_end.date().toPyDate())
            df = df[(df['date'] >= start) & (df['date'] <= end)]
            
            query = self.txt_filter.text().lower().strip()
            if query:
                mask = df['description'].str.lower().str.contains(query, na=False) | \
                       df['results'].str.lower().str.contains(query, na=False)
                df = df[mask]

            # 3. UPDATE METRICS (WITH CRASH PROTECTION)
            total_count = len(df)
            self.card_vol.val_ptr.setText(str(total_count))
            
            # Safe Success Rate
            if total_count > 0:
                comp = len(df[df['status'] == 'Completed'])
                rate = int((comp / total_count) * 100)
                self.card_rate.val_ptr.setText(f"{rate}%")
            else:
                self.card_rate.val_ptr.setText("0%")

            # Safe Bottleneck (Crash-proof mode calculation)
            delayed_df = df[df['status'] == 'Delayed']
            if not delayed_df.empty:
                mode_result = delayed_df['department'].mode()
                if not mode_result.empty:
                    self.card_risk.val_ptr.setText(str(mode_result.iloc[0]))
                else:
                    self.card_risk.val_ptr.setText("N/A")
            else:
                self.card_risk.val_ptr.setText("None")

            # 4. REDRAW CHARTS
            self._update_charts(df)
            self.lbl_sync.setText(f"Feed Synced: {datetime.now().strftime('%H:%M:%S')} | Filters applied.")

        except Exception as e:
            logging.error(f"Dashboard Refresh Failure: {traceback.format_exc()}")

    def _update_charts(self, df):
        """Internal Chart Drawing Engine."""
        # Global Chart Styling
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial']

        # --- 1. TREND LINE ---
        self.fig_trend.clear()
        ax1 = self.fig_trend.add_subplot(111)
        if not df.empty:
            daily = df.groupby(df['date'].dt.date).size()
            ax1.plot(daily.index, daily.values, color=Colors.PRIMARY, lw=3, marker='o', mfc='white', mew=2)
            ax1.fill_between(daily.index, daily.values, alpha=0.1, color=Colors.PRIMARY)
        self._style_axis(ax1, "Shift Date", "Task Volume")
        self.can_trend.draw()

        # --- 2. STATUS PIE ---
        self.fig_status.clear()
        ax2 = self.fig_status.add_subplot(111)
        if not df.empty:
            counts = df['status'].value_counts()
            colors = [Colors.SUCCESS, Colors.ACCENT, Colors.DANGER, Colors.SECONDARY, Colors.TEXT_MUTED]
            ax2.pie(counts, labels=counts.index, autopct='%1.0f%%', startangle=140, 
                    colors=colors[:len(counts)], wedgeprops={'width': 0.4, 'edgecolor': 'w'})
        self.can_status.draw()

        # --- 3. DEPARTMENT BAR ---
        self.fig_dept.clear()
        ax3 = self.fig_dept.add_subplot(111)
        if not df.empty:
            depts = df['department'].value_counts()
            depts.plot(kind='bar', ax=ax3, color=Colors.ACCENT, width=0.6)
            plt.setp(ax3.get_xticklabels(), rotation=30, ha='right', fontsize=8)
        self._style_axis(ax3, "", "Activity Count")
        self.can_dept.draw()

    def _style_axis(self, ax, xlabel, ylabel):
        ax.set_xlabel(xlabel, fontsize=8, color=Colors.TEXT_MUTED)
        ax.set_ylabel(ylabel, fontsize=8, color=Colors.TEXT_MUTED)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color(Colors.BORDER)
        ax.spines['bottom'].set_color(Colors.BORDER)
        ax.tick_params(colors=Colors.TEXT_MUTED, labelsize=8)
        ax.grid(axis='y', linestyle='--', alpha=0.2)

    def export_visuals(self):
        """Captures the entire scroll area as a high-res PNG."""
        path, _ = QFileDialog.getSaveFileName(self, "Export Analytics", f"Report_{Utils.get_today_iso()}.png", "PNG (*.png)")
        if path:
            # We grab the content widget specifically to avoid scrollbars in the screenshot
            pixmap = self.content_widget.grab()
            pixmap.save(path, "PNG")
            QMessageBox.information(self, "Export Successful", f"Dashboard snapshot saved to:\n{path}")

    def export_analytics_pdf(self):
        """Filters data by the selected range and generates a PDF."""
        # 1. Get Dates from UI Pickers
        start_date_q = self.date_start.date().toPyDate()
        end_date_q = self.date_end.date().toPyDate()
        
        start_str = start_date_q.isoformat()
        end_str = end_date_q.isoformat()
        
        if start_date_q > end_date_q:
            QMessageBox.warning(self, "Invalid Range", "Start Date cannot be after End Date.")
            return

        # 2. Fetch ALL tasks from DB (User specific)
        all_tasks = self.db.get_tasks()
        
        # 3. Filter tasks by the selected Date Range
        filtered_tasks = []
        for t in all_tasks:
            try:
                # Convert string date to object for comparison
                t_date = datetime.strptime(t.date, "%Y-%m-%d").date()
                if start_date_q <= t_date <= end_date_q:
                    filtered_tasks.append(t)
            except ValueError:
                continue

        if not filtered_tasks:
            QMessageBox.information(self, "No Data", f"No activities found between {start_str} and {end_str}.")
            return

        # 4. Generate PDF
        filename = f"Analytics_Report_{start_str}_to_{end_str}.pdf"
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Analytics Report", filename, "PDF Documents (*.pdf)")
        
        if save_path:
            profile = self.db.get_profile()
            
            success = PDFEngine.generate_analytics_report(
                save_path, 
                filtered_tasks, 
                profile, 
                start_str, 
                end_str
            )
            
            if success:
                QMessageBox.information(self, "Success", f"Report generated successfully:\n{save_path}")
                try:
                    os.startfile(save_path)
                except:
                    pass
            else:
                QMessageBox.critical(self, "Error", "Could not generate PDF. Check logs.")
# ==============================================================================
# SECTION 8: MAIN WINDOW & BOOTSTRAP
# ==============================================================================

class MainWindow(QMainWindow):
    """
    TITANIUM EDITION: Main Mission Control
    --------------------------------------
    Fixes:
    - High-DPI Scaling (No more blurry UI on 4K laptops)
    - Silent Auto-save (Removed annoying popups on tab switch)
    - Robust Shift Timers (Handles multi-day app uptime)
    - Centralized Error Handling
    """
    def __init__(self):
        # 1. PRE-INITIALIZATION: High-DPI Scaling Fix
        # (This ensures the app looks sharp on high-resolution screens)
        
        super().__init__()
        
        # 2. CORE ENGINES
        self.db = DatabaseEngine()
        self.thread_pool = QThreadPool.globalInstance() # Use global instance for better resource management
        
        # 3. WINDOW IDENTITY
        self.setWindowTitle(f"{Config.APP_NAME} - Enterprise Edition v{Config.APP_VERSION}")
        self.resize(1300, 850)
        self.setMinimumSize(1100, 750)
        
        # Set Branded Icon
        icon_path = resource_path("KTIB daily report.svg")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # 4. TRACKING STATE
        self.last_notified_date = QDate.currentDate()
        self.reminder_morning_triggered = False
        self.reminder_evening_triggered = False
        self.previous_tab_index = 0 # To track where the user came from

        # 5. UI CONSTRUCTION
        self.setup_ui()
        self.setup_tray()
        self.setStyleSheet(UI.STYLE_QSS)
        
        # 6. SYSTEM TIMERS
        self.master_timer = QTimer(self)
        self.master_timer.timeout.connect(self.system_heartbeat)
        self.master_timer.start(1000) # 1-second pulse
        
        self.showMaximized()

    def setup_ui(self):
        """Builds the primary industrial interface."""
        
        # --- 1. STATUS BAR (MOVE TO TOP) ---
        # We must initialize this first so child tabs can access it during their own init
        self.status = QStatusBar()
        self.status.setStyleSheet("background: white; border-top: 1px solid #e2e8f0;")
        self.setStatusBar(self.status)
        
        self.lbl_user_identity = QLabel("Initializing User Session...")
        self.lbl_user_identity.setStyleSheet("font-weight: 700; color: #475569; padding-right: 20px;")
        self.status.addPermanentWidget(self.lbl_user_identity)

        # --- 2. CENTRAL WIDGET ---
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # --- 3. HEADER BAR ---
        header = QFrame()
        header.setFixedHeight(90)
        header.setStyleSheet(f"background-color: white; border-bottom: 2px solid {Colors.BORDER};")
        h_lyt = QHBoxLayout(header)
        h_lyt.setContentsMargins(25, 0, 25, 0)
        
        # Brand Stack
        brand_stack = QVBoxLayout()
        org_lbl = QLabel(Config.ORG_NAME.upper())
        org_lbl.setStyleSheet(f"font-size: 18px; font-weight: 900; color: {Colors.PRIMARY}; letter-spacing: 1px;")
        dept_lbl = QLabel("OPERATIONS CONTROL CENTER | SYSTEM MONITOR")
        dept_lbl.setStyleSheet(f"font-size: 10px; font-weight: 700; color: {Colors.TEXT_MUTED}; letter-spacing: 2px;")
        brand_stack.addWidget(org_lbl)
        brand_stack.addWidget(dept_lbl)
        h_lyt.addLayout(brand_stack)
        
        h_lyt.addStretch()

        # Global Date Context (Crucial for post-midnight reporting)
        date_box = QWidget()
        db_lyt = QVBoxLayout(date_box)
        db_lyt.setContentsMargins(0, 0, 0, 0)
        db_lyt.setSpacing(2)
        
        date_cap = QLabel("ACTIVE REPORTING DATE")
        date_cap.setStyleSheet(f"font-size: 9px; font-weight: 800; color: {Colors.TEXT_MUTED};")
        
        self.reporting_date_picker = QDateEdit(calendarPopup=True)
        self.reporting_date_picker.setDate(QDate.currentDate())
        self.reporting_date_picker.setMinimumHeight(40)
        self.reporting_date_picker.setFixedWidth(160)
        self.reporting_date_picker.setStyleSheet(f"""
            QDateEdit {{
                border: 2px solid {Colors.PRIMARY};
                border-radius: 8px;
                font-weight: 800;
                padding-left: 10px;
                font-size: 14px;
            }}
        """)
        self.reporting_date_picker.dateChanged.connect(self.on_global_date_changed)
        db_lyt.addWidget(date_cap)
        db_lyt.addWidget(self.reporting_date_picker)
        h_lyt.addWidget(date_box)
        
        h_lyt.addSpacing(30)
        
        # Digital Clock
        self.lbl_clock = QLabel("00:00:00")
        self.lbl_clock.setStyleSheet(f"font-family: 'Consolas', monospace; font-size: 24px; color: {Colors.SECONDARY}; font-weight: bold;")
        h_lyt.addWidget(self.lbl_clock)
        
        layout.addWidget(header)
        
        # --- 4. NAVIGATION TABS ---
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        
        # Initialize Tabs
        self.tab_profile = ProfileTab(self.db)
        # Now safe to initialize MorningTab because self.status exists
        self.tab_morning = MorningTab(self.db, self.thread_pool, self) 
        self.tab_evening = EveningTab(self.db, self.thread_pool, self)
        self.tab_dash = DashboardTab(self.db)
        
        self.tabs.addTab(self.tab_profile, "👤 IDENTITY")
        self.tabs.addTab(self.tab_morning, "☀️ MORNING PLAN")
        self.tabs.addTab(self.tab_evening, "🌙 EVENING REPORT")
        self.tabs.addTab(self.tab_dash, "📊 ANALYTICS")
        
        self.tabs.currentChanged.connect(self.on_tab_navigated)
        layout.addWidget(self.tabs)
        
        self.refresh_user_display()

    def setup_tray(self):
        """Initializes the background system tray icon."""
        self.tray = QSystemTrayIcon(self)
        icon_path = resource_path("KTIB daily report.svg")
        if os.listdir() and os.path.exists(icon_path):
            self.tray.setIcon(QIcon(icon_path))
        else:
            self.tray.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))
        
        tray_menu = QMenu()
        show_action = tray_menu.addAction("Show System")
        show_action.triggered.connect(self.showNormal)
        exit_action = tray_menu.addAction("Exit Application")
        exit_action.triggered.connect(QApplication.instance().quit)
        
        self.tray.setContextMenu(tray_menu)
        self.tray.show()

    # ==========================================================================
    # BUSINESS LOGIC & EVENTS
    # ==========================================================================

    def on_global_date_changed(self):
        """Forces all tabs to reload data based on the new selected date."""
        new_date = self.reporting_date_picker.date().toString(Qt.DateFormat.ISODate)
        self.status.showMessage(f"Switching Context to {new_date}...", 2000)
        self.on_tab_navigated(self.tabs.currentIndex())

    def system_heartbeat(self):
        """Standard 1s update for clock and shift logic."""
        now = QDateTime.currentDateTime()
        self.lbl_clock.setText(now.toString("HH:mm:ss"))
        
        # Reset reminders at midnight
        if now.date() > self.last_notified_date:
            self.last_notified_date = now.date()
            self.reminder_morning_triggered = False
            self.reminder_evening_triggered = False

        curr_time = now.time()
        # 10:00 AM Notification
        if curr_time.hour() == 10 and curr_time.minute() == 0 and not self.reminder_morning_triggered:
            self.send_notification("Morning Plan Due", "Please finalize and dispatch your morning shift deployment.")
            self.reminder_morning_triggered = True
            
        # 16:30 PM Notification
        if curr_time.hour() == 16 and curr_time.minute() == 30 and not self.reminder_evening_triggered:
            self.send_notification("Shift Ending Soon", "OCC Shift ends at 17:00. Start finalizing evening reports.")
            self.reminder_evening_triggered = True

    def on_tab_navigated(self, new_index):
        """
        Handles Logic when leaving and entering tabs.
        Crucial: Auto-saves data without annoying the user with popups.
        """
        # 1. SAVE ACTIONS (When leaving a tab)
        if self.previous_tab_index == 0: # Left Profile Tab
            self.tab_profile.save_data_silently()
            self.refresh_user_display()
        elif self.previous_tab_index == 2: # Left Evening Tab
            if self.tab_evening.save_timer.isActive():
                self.tab_evening.perform_auto_save()

        # 2. LOAD ACTIONS (When entering a tab)
        target_widget = self.tabs.widget(new_index)
        if target_widget == self.tab_morning:
            self.tab_morning.refresh_data()
        elif target_widget == self.tab_evening:
            self.tab_evening.refresh_list()
        elif target_widget == self.tab_dash:
            self.tab_dash.refresh()
        elif target_widget == self.tab_profile:
            self.tab_profile.load_data()

        self.previous_tab_index = new_index

    def refresh_user_display(self):
        """Syncs the status bar with the current user profile."""
        p = self.db.get_profile()
        user_text = f"OPERATOR: {p.name if p.name else 'UNCONFIGURED'}"
        self.lbl_user_identity.setText(user_text)

    def send_notification(self, title, msg):
        self.tray.showMessage(title, msg, QSystemTrayIcon.MessageIcon.Information, 8000)

    def closeEvent(self, event):
        """Safety check before shutdown."""
        # Ensure final auto-saves
        self.tab_profile.save_data_silently()
        if self.tab_evening.save_timer.isActive():
            self.tab_evening.perform_auto_save()

        p = self.db.get_profile()
        if not p.is_valid():
            btn = QMessageBox.warning(self, "Incomplete Profile", 
                                     "Your identity profile is not set. Data may not be attributed correctly. Exit anyway?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if btn == QMessageBox.StandardButton.No:
                event.ignore()
                return
        
        logging.info("Application closed by user.")
        event.accept()

if __name__ == "__main__":
    # 1. Initialize Directories using the logic in Config
    Config.initialize_directories()

    # 2. DPI Scaling Fix
    from PyQt6.QtCore import Qt
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

    # 3. Taskbar Icon fix
    try:
        myappid = f'ktib.dailyreport.automate.{Config.APP_VERSION}'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())