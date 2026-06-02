# -*- coding: utf-8 -*-
"""
================================================================================
APP NAME:    KTIB Operations Intelligence (Executive Edition)
VERSION:     5.4.0-PLATINUM
AUTHOR:      Senior Systems Architect
LICENSE:     Proprietary (KTIB Azerbaijan LLC)
FRAMEWORK:   PyQt6 + SQLite3 + Pandas + ReportLab

"Precision in every line, intelligence in every insight."
================================================================================
"""

import os
import sys
import sqlite3
import logging
import pandas as pd
from datetime import datetime, date, timedelta
from typing import List, Dict, Optional
from contextlib import contextmanager

# --- REPORTLAB PDF ENGINE (Matches Main App Style) ---
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
except ImportError:
    print("CRITICAL: 'reportlab' library missing. Run: pip install reportlab")
    sys.exit(1)

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QFrame, QDateEdit, QStatusBar, QAbstractItemView, 
    QMessageBox, QSplitter, QListWidget, QLineEdit, 
    QScrollArea, QGraphicsDropShadowEffect, QListWidgetItem, QStyle, 
    QFileDialog, QSizePolicy
)
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QThread, QSize, QDate
from PyQt6.QtGui import QColor, QBrush, QIcon, QFont, QPainter, QPen

# ==============================================================================
# SECTION 1: EXECUTIVE DESIGN SYSTEM
# ==============================================================================

class Config:
    APP_NAME = "KTIB_Executive_Hub"
    APP_VERSION = "5.4.0"
    ORG_NAME = "KTIB Azerbaijan LLC"
    # Update this path to match your server location
    HUB_PATH = r"H:\Public AZ\DailyReportAutomate\ktib_ops_database.db"
    
    # Fallback to local if server not found (for testing)
    if not os.path.exists(HUB_PATH):
        local_path = os.path.join(os.path.dirname(__file__), "ktib_ops_database.db")
        if os.path.exists(local_path):
            HUB_PATH = local_path

class Theme:
    """Executive Industrial Palette."""
    WHITE = "#ffffff"
    BACKGROUND = "#f8fafc" 
    SURFACE = "#ffffff"
    BORDER = "#e2e8f0" 
    PRIMARY = "#0f172a" 
    ACCENT = "#2563eb" 
    SUCCESS = "#059669" 
    DANGER = "#dc2626" 
    WARNING = "#d97706" 
    TEXT_MAIN = "#334155" 
    TEXT_DIM = "#94a3b8" 

# ==============================================================================
# SECTION 2: HIGH-CONTRAST PDF ENGINE (MATCHING APP3.PY)
# ==============================================================================

class PDFEngine:
    """Generates a professional Audit PDF using ReportLab."""
    
    @staticmethod
    def generate_report(filepath: str, df: pd.DataFrame, start_date: str, end_date: str):
        try:
            doc = SimpleDocTemplate(
                filepath, 
                pagesize=landscape(A4),
                rightMargin=20, leftMargin=20, 
                topMargin=30, bottomMargin=30
            )
            
            elements = []
            styles = getSampleStyleSheet()
            
            # 1. HEADER
            title_style = ParagraphStyle(
                'Title', parent=styles['Heading1'], 
                fontSize=18, textColor=colors.black, 
                spaceAfter=5, alignment=1 
            )
            elements.append(Paragraph(f"KTIB OPERATIONS AUDIT REPORT", title_style))
            elements.append(Paragraph(f"Period: <b>{start_date}</b> to <b>{end_date}</b>", 
                ParagraphStyle('Sub', parent=styles['Normal'], alignment=1, fontSize=11)))
            elements.append(Spacer(1, 15))

            # 2. SUMMARY BOX
            total = len(df)
            completed = len(df[df['status'] == 'Completed'])
            delayed = len(df[df['status'] == 'Delayed'])
            rate = int((completed/total)*100) if total > 0 else 0

            summary_html = f"""
            <font size="10">
            <b>GENERATED:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}<br/><br/>
            <b>TOTAL RECORDS:</b> {total} &nbsp;|&nbsp; 
            <b>COMPLETION RATE:</b> {rate}% &nbsp;|&nbsp; 
            <b>DELAYED ITEMS:</b> <font color='red'>{delayed}</font>
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

            # 3. MAIN TABLE DATA
            headers = ["Meta Data", "Task Configuration", "Status & Progress", "Execution Log (Notes & Results)"]
            data = [headers]
            
            # Sort: Name then Date
            df = df.sort_values(by=['display_name', 'report_date'], ascending=[True, False])
            
            s_norm = ParagraphStyle('CellNorm', parent=styles['Normal'], fontSize=9, leading=11, spaceAfter=3, textColor=colors.black)
            
            for _, row in df.iterrows():
                # Col 1: Meta
                meta = f"""
                <b>Date:</b> {row['report_date']}<br/>
                <b>User:</b> {row['display_name']}<br/>
                <b>ID:</b> {row['task_id']}
                """
                
                # Col 2: Config
                config = f"""
                <b>Activity:</b> {row['description']}<br/>
                <b>Plan Time:</b> {row.get('planned_time', 'N/A')}<br/>
                <b>Deps:</b> {row.get('dependencies', '')}
                """
                
                # Col 3: Status
                stat_txt = row['status']
                color = "black"
                if stat_txt == "Completed": color = "green"
                elif stat_txt == "Delayed": color = "red"
                elif stat_txt == "In Progress": color = "orange"
                
                status = f"""
                <b>Status:</b> <font color='{color}'>{stat_txt}</font><br/>
                <b>Progress:</b> {row.get('percent_complete', 0)}%
                """
                
                # Col 4: Full Log
                log_parts = []
                if row.get('morning_notes'): log_parts.append(f"<b>[Note]:</b> {row['morning_notes']}")
                if row.get('results'): log_parts.append(f"<b>[Result]:</b> {row['results']}")
                if row.get('reason_delay'): log_parts.append(f"<b>[DELAY]:</b> <font color='red'>{row['reason_delay']}</font>")
                if row.get('next_steps'): log_parts.append(f"<b>[Next]:</b> {row['next_steps']}")
                
                log_html = "<br/>".join(log_parts) if log_parts else "<i>No details logged.</i>"
                
                data.append([
                    Paragraph(meta, s_norm),
                    Paragraph(config, s_norm),
                    Paragraph(status, s_norm),
                    Paragraph(log_html, s_norm)
                ])

            # 4. TABLE STYLING (High Contrast)
            col_widths = [1.5*inch, 3.5*inch, 1.2*inch, 4.8*inch]
            
            t = Table(data, colWidths=col_widths, repeatRows=1)
            t.setStyle(TableStyle([
                # BLACK HEADER
                ('BACKGROUND', (0, 0), (-1, 0), colors.black),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                ('TOPPADDING', (0, 0), (-1, 0), 10),
                # BODY
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('ROWBACKGROUNDS', (1, 0), (-1, -1), [colors.white, colors.whitesmoke]),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 10),
                ('TOPPADDING', (0, 1), (-1, -1), 10),
            ]))
            elements.append(t)
            
            doc.build(elements)
            return True
        except Exception as e:
            return str(e)

# ==============================================================================
# SECTION 3: DATA INTELLIGENCE BRIDGE
# ==============================================================================

class HubEngine:
    @contextmanager
    def connection(self):
        """Thread-safe Read-Only bridge."""
        if not os.path.exists(Config.HUB_PATH):
            raise FileNotFoundError(f"Database not found at: {Config.HUB_PATH}")
        
        db_uri = f"file:{Config.HUB_PATH}?mode=ro"
        conn = sqlite3.connect(db_uri, uri=True, timeout=60)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    def fetch_intelligence(self, start: str, end: str) -> pd.DataFrame:
        """Fetches raw data."""
        with self.connection() as conn:
            query = """
                SELECT t.*, p.full_name as profile_name, p.department as profile_dept
                FROM tasks t
                LEFT JOIN profile p ON t.os_user = p.os_user
                WHERE t.report_date BETWEEN ? AND ?
                ORDER BY t.report_date DESC
            """
            df = pd.read_sql_query(query, conn, params=(start, end))
            # Fallback for display name if profile missing
            df['display_name'] = df['profile_name'].fillna(df['employee'])
            return df

    def get_employee_registry(self) -> List[Dict]:
        with self.connection() as conn:
            query = "SELECT DISTINCT os_user, full_name, department FROM profile WHERE full_name IS NOT NULL ORDER BY full_name ASC"
            return [dict(r) for r in conn.execute(query).fetchall()]

class DataWorker(QThread):
    sync_complete = pyqtSignal(object)
    sync_failed = pyqtSignal(str)

    def __init__(self, engine: HubEngine, start: str, end: str):
        super().__init__()
        self.engine = engine
        self.start_date = start
        self.end_date = end

    def run(self):
        try:
            df = self.engine.fetch_intelligence(self.start_date, self.end_date)
            self.sync_complete.emit(df)
        except Exception as e:
            self.sync_failed.emit(str(e))

# ==============================================================================
# SECTION 4: CUSTOM INTERFACE COMPONENTS
# ==============================================================================

class ExecutiveCard(QFrame):
    def __init__(self, bg=Theme.SURFACE):
        super().__init__()
        self.setStyleSheet(f"""
            ExecutiveCard {{ 
                background: {bg}; 
                border: 1px solid {Theme.BORDER}; 
                border-radius: 12px; 
            }}
        """)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20); shadow.setColor(QColor(0,0,0,10)); shadow.setOffset(0,4)
        self.setGraphicsEffect(shadow)

class MetricTile(ExecutiveCard):
    def __init__(self, title: str, icon: str = ""):
        super().__init__()
        self.setFixedHeight(100)
        l = QVBoxLayout(self); l.setContentsMargins(20, 15, 20, 15)
        
        self.lbl_title = QLabel(f"{icon}  {title.upper()}")
        self.lbl_title.setStyleSheet(f"color: {Theme.TEXT_DIM}; font-weight: 800; font-size: 10px; letter-spacing: 1px;")
        
        self.lbl_val = QLabel("0")
        self.lbl_val.setStyleSheet(f"color: {Theme.PRIMARY}; font-size: 28px; font-weight: 900;")
        
        l.addWidget(self.lbl_title); l.addWidget(self.lbl_val)

    def set_value(self, val, color=Theme.PRIMARY):
        self.lbl_val.setText(str(val))
        self.lbl_val.setStyleSheet(f"color: {color}; font-size: 28px; font-weight: 900;")

class TimelineItem(QWidget):
    """Visualizes a single day's status for a task in the Inspector."""
    def __init__(self, date_str, status, notes):
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 10)
        
        # Color Code
        color = Theme.TEXT_DIM
        if status == "Completed": color = Theme.SUCCESS
        elif status == "Delayed": color = Theme.DANGER
        elif status == "In Progress": color = Theme.WARNING
        
        # Dot
        dot = QLabel("●")
        dot.setStyleSheet(f"color: {color}; font-size: 14px;")
        layout.addWidget(dot)
        
        # Info
        info_layout = QVBoxLayout()
        top_line = QLabel(f"{date_str}  |  {status.upper()}")
        top_line.setStyleSheet(f"font-weight: 800; font-size: 10px; color: {Theme.TEXT_MAIN};")
        
        note_lbl = QLabel(notes if notes else "No specific notes.")
        note_lbl.setWordWrap(True)
        note_lbl.setStyleSheet(f"color: {Theme.TEXT_DIM}; font-size: 11px;")
        
        info_layout.addWidget(top_line)
        info_layout.addWidget(note_lbl)
        layout.addLayout(info_layout)

class InspectorPanel(ExecutiveCard):
    """Right-hand side details panel."""
    def __init__(self):
        super().__init__()
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(20, 20, 20, 20)
        
        self.lbl_header = QLabel("TASK INSPECTOR")
        self.lbl_header.setStyleSheet(f"color: {Theme.ACCENT}; font-size: 11px; font-weight: 900; letter-spacing: 1.5px;")
        self.layout.addWidget(self.lbl_header)
        self.layout.addSpacing(10)
        
        # Details Area
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setStyleSheet("background: transparent;")
        
        self.content = QWidget()
        self.content_lyt = QVBoxLayout(self.content)
        self.content_lyt.setSpacing(15)
        self.content_lyt.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        self.scroll.setWidget(self.content)
        self.layout.addWidget(self.scroll)
        
        self.reset_view()

    def reset_view(self):
        self.clear_content()
        lbl = QLabel("Select a task from the grid to view its full lifecycle and technical findings.")
        lbl.setWordWrap(True)
        lbl.setStyleSheet(f"color: {Theme.TEXT_DIM}; font-style: italic; margin-top: 20px;")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.content_lyt.addWidget(lbl)

    def clear_content(self):
        while self.content_lyt.count():
            item = self.content_lyt.takeAt(0)
            if item.widget(): item.widget().deleteLater()

    def add_section_header(self, title):
        lbl = QLabel(title.upper())
        lbl.setStyleSheet(f"color: {Theme.PRIMARY}; font-size: 10px; font-weight: 900; border-bottom: 1px solid {Theme.BORDER}; padding-bottom: 4px; margin-top: 10px;")
        self.content_lyt.addWidget(lbl)

    def add_detail(self, label, value, highlight=False):
        lbl = QLabel(label.upper())
        lbl.setStyleSheet(f"color: {Theme.TEXT_DIM}; font-size: 9px; font-weight: 800;")
        val = QLabel(str(value) if value else "-")
        val.setWordWrap(True)
        val.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        color = Theme.TEXT_MAIN if not highlight else Theme.PRIMARY
        val.setStyleSheet(f"color: {color}; font-size: 13px; font-weight: 500;")
        self.content_lyt.addWidget(lbl)
        self.content_lyt.addWidget(val)

    def add_timeline(self, history_df):
        self.add_section_header("LIFECYCLE TIMELINE")
        
        history_df = history_df.sort_values(by='report_date', ascending=False)
        
        for _, row in history_df.iterrows():
            note = row.get('results') or row.get('morning_notes') or ""
            item = TimelineItem(row['report_date'], row['status'], note)
            self.content_lyt.addWidget(item)

# ==============================================================================
# SECTION 5: MAIN CONTROL CENTER
# ==============================================================================

class ExecutiveDashboard(QWidget):
    def __init__(self, db: HubEngine, status_bar: QStatusBar):
        super().__init__()
        self.db = db
        self.status_bar = status_bar
        self.full_df = pd.DataFrame()
        self.current_display_df = pd.DataFrame() 
        
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self.apply_filters)
        
        self.setup_ui()
        QTimer.singleShot(500, self.initiate_sync)

    def setup_ui(self):
        """Builds the Managerial Interface."""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # --- 1. SIDEBAR ---
        self.sidebar = QFrame()
        self.sidebar.setFixedWidth(260)
        self.sidebar.setStyleSheet(f"background: {Theme.WHITE}; border-right: 1px solid {Theme.BORDER};")
        s_lyt = QVBoxLayout(self.sidebar)
        s_lyt.setContentsMargins(20, 30, 20, 20)
        
        logo = QLabel("KTIB Manager")
        logo.setStyleSheet(f"color: {Theme.PRIMARY}; font-weight: 900; font-size: 24px; letter-spacing: -0.5px;")
        s_lyt.addWidget(logo)
        
        s_lyt.addWidget(QLabel("TEAM SELECTION", 
                                styleSheet=f"color: {Theme.TEXT_DIM}; font-size: 10px; font-weight: 800; margin-top: 25px;"))
        
        self.user_list = QListWidget()
        self.user_list.setStyleSheet(f"""
            QListWidget {{ background: transparent; border: none; outline: none; margin-top: 5px; }}
            QListWidget::item {{ padding: 12px; border-radius: 8px; color: {Theme.TEXT_MAIN}; margin-bottom: 2px; font-weight: 500; }}
            QListWidget::item:selected {{ background: {Theme.ACCENT}; color: white; font-weight: bold; }}
            QListWidget::item:hover:!selected {{ background: #f1f5f9; }}
        """)
        self.user_list.itemClicked.connect(self.apply_filters)
        s_lyt.addWidget(self.user_list)
        layout.addWidget(self.sidebar)

        # --- 2. MAIN AREA ---
        main_area = QWidget()
        m_lyt = QVBoxLayout(main_area)
        m_lyt.setContentsMargins(25, 25, 25, 25)
        m_lyt.setSpacing(20)
        
        # TOOLBAR
        toolbar = QHBoxLayout()
        
        self.d1 = QDateEdit(date.today() - timedelta(days=7), calendarPopup=True)
        self.d2 = QDateEdit(date.today(), calendarPopup=True)
        for d in [self.d1, self.d2]:
            d.setFixedSize(120, 40)
            d.dateChanged.connect(self.initiate_sync)
            d.setStyleSheet(f"background: white; border: 1px solid {Theme.BORDER}; border-radius: 6px; font-weight: 600;")
            toolbar.addWidget(d)
        
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search task ID, description, or keyword...")
        self.search.setFixedHeight(40)
        self.search.textChanged.connect(lambda: self.search_timer.start(300))
        self.search.setStyleSheet(f"background: white; border: 1px solid {Theme.BORDER}; border-radius: 20px; padding: 0 15px;")
        toolbar.addWidget(self.search)
        
        # UNIQULAIZE TOGGLE (DEFAULT OFF AS REQUESTED)
        self.btn_unique = QPushButton(" UNIQULAIZE ")
        self.btn_unique.setCheckable(True)
        self.btn_unique.setChecked(False) # <--- DEFAULT IS NOW OFF (SHOW ALL HISTORY)
        self.btn_unique.clicked.connect(self.apply_filters)
        self.btn_unique.setFixedSize(130, 40)
        self.btn_unique.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_unique.setStyleSheet(f"""
            QPushButton {{ background: white; border: 1px solid {Theme.BORDER}; border-radius: 6px; font-weight: 700; color: {Theme.TEXT_DIM}; }}
            QPushButton:checked {{ background: {Theme.PRIMARY}; color: white; border: none; }}
        """)
        toolbar.addWidget(self.btn_unique)
        
        # PDF EXPORT
        self.btn_pdf = QPushButton(" PDF REPORT ")
        self.btn_pdf.clicked.connect(self.export_to_pdf)
        self.btn_pdf.setFixedSize(120, 40)
        self.btn_pdf.setStyleSheet(f"background: {Theme.ACCENT}; color: white; font-weight: 700; border-radius: 6px;")
        toolbar.addWidget(self.btn_pdf)
        
        self.btn_sync = QPushButton("↻")
        self.btn_sync.setFixedSize(40, 40)
        self.btn_sync.clicked.connect(self.initiate_sync)
        toolbar.addWidget(self.btn_sync)
        
        m_lyt.addLayout(toolbar)

        # KPIS
        metrics = QHBoxLayout()
        self.k1 = MetricTile("Displayed Records", "📋")
        self.k2 = MetricTile("Completion Rate", "🎯")
        self.k3 = MetricTile("Delayed Items", "⚠️")
        for k in [self.k1, self.k2, self.k3]: metrics.addWidget(k)
        m_lyt.addLayout(metrics)

        # SPLITTER
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        
        # TABLE
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["REPORT DATE", "EMPLOYEE", "TASK ID", "ACTIVITY DESCRIPTION", "TIME", "STATUS"])
        
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(5, 120)

        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        self.table.setStyleSheet(f"""
            QTableWidget {{ background: white; border: 1px solid {Theme.BORDER}; border-radius: 8px; font-size: 13px; }}
            QHeaderView::section {{ background: {Theme.BACKGROUND}; color: {Theme.TEXT_MAIN}; padding: 10px; font-weight: 700; border: none; font-size: 11px; }}
            QTableWidget::item {{ padding: 5px; }}
            QTableWidget::item:selected {{ background: {Theme.ACCENT}; color: white; }}
        """)
        
        splitter.addWidget(self.table)
        
        # INSPECTOR
        self.inspector = InspectorPanel()
        splitter.addWidget(self.inspector)
        
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        
        m_lyt.addWidget(splitter)
        layout.addWidget(main_area)

    def initiate_sync(self):
        self.status_bar.showMessage("Connecting to Operations Hub...")
        self.btn_sync.setEnabled(False)
        
        self.worker = DataWorker(self.db, self.d1.date().toPyDate().isoformat(), self.d2.date().toPyDate().isoformat())
        self.worker.sync_complete.connect(self.on_data_ready)
        self.worker.sync_failed.connect(self.on_sync_failure)
        self.worker.start()

    def on_sync_failure(self, err):
        self.btn_sync.setEnabled(True)
        QMessageBox.critical(self, "Sync Error", f"Could not retrieve data.\n{err}")

    def on_data_ready(self, df):
        self.full_df = df
        self.btn_sync.setEnabled(True)
        self.status_bar.showMessage(f"Data Synced: {len(df)} records found.")
        
        if self.user_list.count() <= 1:
            self.user_list.clear()
            all_it = QListWidgetItem("🏢  ALL DEPARTMENTS")
            all_it.setData(Qt.ItemDataRole.UserRole, None)
            self.user_list.addItem(all_it)
            self.user_list.setCurrentItem(all_it)
            
            # --- NEW CODE (Shows everyone) ---
            users = self.db.get_employee_registry()
            # We remove the active_users filter so we see everyone
            for u in users:
                # Just check if we have a valid name to display
                if u['full_name']: 
                    it = QListWidgetItem(f"👤  {u['full_name']}")
                    it.setData(Qt.ItemDataRole.UserRole, u['os_user'])
                    self.user_list.addItem(it)
        
        self.apply_filters()

    def apply_filters(self):
        # 1. Safety Check
        if self.full_df.empty:
            self.update_grid(pd.DataFrame())
            return

        df = self.full_df.copy()

        # 2. USER FILTER
        curr = self.user_list.currentItem()
        if curr:
            user_id = curr.data(Qt.ItemDataRole.UserRole)
            if user_id is not None:
                clean_selection = str(user_id).strip().lower()
                mask = df['os_user'].astype(str).str.strip().str.lower() == clean_selection
                df = df[mask]

        # 3. TEXT SEARCH FILTER
        q = self.search.text().lower().strip()
        if q:
            mask = (
                df['description'].str.lower().str.contains(q, na=False) | 
                df['task_id'].str.lower().str.contains(q, na=False) |
                df['display_name'].str.lower().str.contains(q, na=False) |
                df['results'].str.lower().str.contains(q, na=False)
            )
            df = df[mask]

        # 4. INTELLIGENT UNIQUE LOGIC (DESCRIPTION PRIORITY)
        if self.btn_unique.isChecked():
            # A. Create "Description Fingerprint"
            # We take the first 25 characters of the description. 
            # If the description is the same (or very close), we treat it as the same task.
            # WE IGNORE THE TASK ID HERE.
            df['fingerprint'] = df['description'].astype(str).str.strip().str.lower().str.slice(0, 25)

            # B. CALCULATE DURATION
            # Count how many unique dates this specific ACTIVITY description appears
            df['calculated_duration'] = df.groupby(['os_user', 'fingerprint'])['report_date'].transform('nunique')

            # C. Sort by Date (Newest first)
            df = df.sort_values(by=['report_date'], ascending=False)
            
            # D. Drop duplicates based on the DESCRIPTION only
            # This will merge T-006 and T-010 if they have the same description
            df = df.drop_duplicates(subset=['os_user', 'fingerprint'], keep='first')
            
            self.btn_unique.setText(" ✓ UNIQUE VIEW ")
        else:
            self.btn_unique.setText(" UNIQULAIZE ")

        # 5. FINAL VISUAL SORTING
        df = df.sort_values(by=['report_date', 'task_id'], ascending=[False, True])

        self.current_display_df = df
        self.update_grid(df)

    def update_grid(self, df):
        # Update KPI Tiles
        self.k1.set_value(len(df))
        if not df.empty:
            completed = len(df[df['status'] == 'Completed'])
            rate = int((completed / len(df)) * 100)
            self.k2.set_value(f"{rate}%", Theme.SUCCESS if rate > 80 else Theme.WARNING)
            delayed = len(df[df['status'] == 'Delayed'])
            self.k3.set_value(delayed, Theme.DANGER if delayed > 0 else Theme.TEXT_MAIN)
        else:
            self.k2.set_value("0%"); self.k3.set_value(0)

        self.table.setRowCount(0)
        
        # --- DYNAMIC HEADER CHANGE ---
        # If Unique View is ON, show "DURATION" instead of "TIME"
        is_unique_mode = self.btn_unique.isChecked()
        if is_unique_mode:
            self.table.setHorizontalHeaderItem(4, QTableWidgetItem("DURATION"))
        else:
            self.table.setHorizontalHeaderItem(4, QTableWidgetItem("TIME"))

        # --- POPULATE ROWS ---
        for _, row in df.iterrows():
            r = self.table.rowCount()
            self.table.insertRow(r)
            row_data = row.to_dict()
            
            self.table.setItem(r, 0, QTableWidgetItem(row['report_date']))
            self.table.setItem(r, 1, QTableWidgetItem(row['display_name']))
            
            id_item = QTableWidgetItem(str(row['task_id']))
            id_item.setFont(QFont("Segoe UI", -1, QFont.Weight.Bold))
            self.table.setItem(r, 2, id_item)
            
            self.table.setItem(r, 3, QTableWidgetItem(row['description']))
            
            # --- DURATION VS TIME LOGIC ---
            if is_unique_mode and 'calculated_duration' in row:
                # Show how many days this task has been carried over
                days = row['calculated_duration']
                dur_text = f"{days} Day{'s' if days > 1 else ''}"
                dur_item = QTableWidgetItem(dur_text)
                dur_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                dur_item.setFont(QFont("Segoe UI", -1, QFont.Weight.Bold))
                dur_item.setForeground(QBrush(QColor(Theme.ACCENT)))
                self.table.setItem(r, 4, dur_item)
            else:
                # Show standard planned time
                time_val = row.get('planned_time', 'N/A')
                self.table.setItem(r, 4, QTableWidgetItem(str(time_val)))
            
            # Status Column
            status = row['status']
            s_item = QTableWidgetItem(status)
            s_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            
            if status == "Completed":
                s_item.setForeground(QBrush(QColor(Theme.SUCCESS)))
                s_item.setFont(QFont("Segoe UI", -1, QFont.Weight.Bold))
            elif status == "Delayed":
                s_item.setForeground(QBrush(QColor(Theme.DANGER)))
                s_item.setFont(QFont("Segoe UI", -1, QFont.Weight.Bold))
            elif status == "In Progress":
                s_item.setForeground(QBrush(QColor(Theme.WARNING)))
            
            self.table.setItem(r, 5, s_item)
            self.table.item(r, 0).setData(Qt.ItemDataRole.UserRole, row_data)

    def on_selection_changed(self):
        rows = self.table.selectedItems()
        if not rows:
            self.inspector.reset_view()
            return
        
        # 1. Get Data from the selected row
        data = self.table.item(rows[0].row(), 0).data(Qt.ItemDataRole.UserRole)
        if not data: return
        
        self.inspector.clear_content()
        
        # --- SECTION A: DETAIL VIEW ---
        self.inspector.add_section_header("MORNING PLAN")
        self.inspector.add_detail("Reference ID", data['task_id'], highlight=True)
        self.inspector.add_detail("Activity Description", data['description'])
        self.inspector.add_detail("Operator", data['display_name'])
        self.inspector.add_detail("Department", data.get('department', 'N/A'))
        self.inspector.add_detail("Planned Time", data.get('planned_time', 'N/A'))
        self.inspector.add_detail("Dependencies", data.get('dependencies', 'None'))
        self.inspector.add_detail("Morning Notes", data.get('morning_notes', 'None'))

        self.inspector.add_section_header("EVENING CLOSURE")
        self.inspector.add_detail("Completion Status", data['status'], highlight=True)
        self.inspector.add_detail("Progress", f"{data.get('percent_complete', 0)}%")
        self.inspector.add_detail("Technical Results", data.get('results', 'Pending'))
        
        if data.get('next_steps'):
             self.inspector.add_detail("Next Steps", data['next_steps'])
        
        if data['status'] == 'Delayed':
            self.inspector.add_detail("Reason for Delay", data.get('reason_delay', 'Not specified'), highlight=True)

        # --- SECTION B: LIFECYCLE TIMELINE (STRICT MATCHING) ---
        
        # 1. Prepare Match Targets
        target_id = str(data['task_id']).strip().lower()
        target_user = str(data['os_user']).strip().lower()
        
        # Create a "Fingerprint" of the description (First 15 characters)
        current_desc = str(data['description']).strip().lower()
        target_fingerprint = current_desc[:15] if len(current_desc) > 15 else current_desc

        # 2. Filter the Full Database
        mask = (
            (self.full_df['os_user'].astype(str).str.strip().str.lower() == target_user) &
            (self.full_df['task_id'].astype(str).str.strip().str.lower() == target_id) &
            (self.full_df['description'].astype(str).str.strip().str.lower().str.startswith(target_fingerprint))
        )
        
        # 3. Get History & Sort
        history = self.full_df[mask].copy()
        history = history.sort_values(by='report_date', ascending=False) # Newest at top

        # 4. Combine Notes Logic (Plan + Result + Delay)
        def merge_notes(row):
            parts = []
            if row.get('morning_notes'): 
                parts.append(f"📋 [PLAN]: {row['morning_notes']}")
            if row.get('results'): 
                parts.append(f"✅ [RES]: {row['results']}")
            if row.get('reason_delay'):
                 parts.append(f"⚠️ [DELAY]: {row['reason_delay']}")
            
            return "\n".join(parts) if parts else "No specific notes logged."

        history['results'] = history.apply(merge_notes, axis=1)

        # --- THIS WAS MISSING ---
        self.inspector.add_timeline(history)


    def export_to_pdf(self):
        if self.current_display_df.empty:
            QMessageBox.warning(self, "No Data", "Grid is empty.")
            return
            
        path, _ = QFileDialog.getSaveFileName(self, "Export Report", f"KTIB_Audit_{date.today()}.pdf", "PDF (*.pdf)")
        if not path: return

        # Call the new ReportLab Engine
        result = PDFEngine.generate_report(
            path, 
            self.current_display_df, 
            self.d1.date().toString(Qt.DateFormat.ISODate),
            self.d2.date().toString(Qt.DateFormat.ISODate)
        )
        
        if result is True:
            QMessageBox.information(self, "Export Successful", "Professional PDF generated successfully.")
        else:
            QMessageBox.critical(self, "Export Error", str(result))

class TitaniumManagerHub(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"KTIB Management Intelligence v{Config.APP_VERSION}")
        self.resize(1600, 950)

        # --- START CHANGE: ADD ICON ---
        # Get the directory of the current script
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "manager.png")
        
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            print(f"Warning: Icon not found at {icon_path}")
        # --- END CHANGE ---
        
        if hasattr(Qt, 'AA_EnableHighDpiScaling'):
            # ... rest of your code ...
            QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
            QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

        self.db = HubEngine()
        self.setCentralWidget(ExecutiveDashboard(self.db, self.statusBar()))
        
        self.setStyleSheet(f"""
            QMainWindow {{ background: {Theme.BACKGROUND}; }}
            QWidget {{ font-family: 'Segoe UI', sans-serif; color: {Theme.TEXT_MAIN}; }}
            QStatusBar {{ background: {Theme.WHITE}; color: {Theme.TEXT_DIM}; font-size: 11px; border-top: 1px solid {Theme.BORDER}; }}
            QScrollBar:vertical {{ border: none; background: #f1f5f9; width: 10px; margin: 0px; }}
            QScrollBar::handle:vertical {{ background: #cbd5e1; min-height: 20px; border-radius: 5px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        self.statusBar().showMessage("SYSTEM SECURE | READ-ONLY ACCESS ENABLED")

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    # --- START CHANGE: WINDOWS TASKBAR FIX ---
    import ctypes
    # Arbitrary string ID to separate this app from other Python scripts
    myappid = 'ktib.operations.manager.executive.5.4.0' 
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except ImportError:
        pass # Not on Windows
    # --- END CHANGE ---

    app = QApplication(sys.argv)
    
    # Set the icon globally for the application (dialogs, etc.)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(script_dir, "manager.png")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    app.setStyle("Fusion")
    
    font = app.font()
    font.setPointSize(10)
    app.setFont(font)
    
    hub = TitaniumManagerHub()
    hub.show()
    sys.exit(app.exec())