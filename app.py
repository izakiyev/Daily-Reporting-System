import sys
import json
import os
import pandas as pd
from datetime import date, timedelta
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QTabWidget, QLabel, QLineEdit, QPushButton, QComboBox, 
    QTableWidget, QTableWidgetItem, QHeaderView, QTimeEdit, 
    QMessageBox, QFrame, QSplitter, QTextEdit, QSlider, 
    QFileDialog, QGridLayout, QScrollArea
)
from PyQt6.QtCore import Qt, QUrl, QTime
from PyQt6.QtGui import QIcon, QFont, QColor
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# ==========================================
# 0. OUTLOOK LOGIC (NEW IMPLEMENTATION)
# ==========================================
# ==========================================
# 0. INFRASTRUCTURE & UTILS
# ==========================================

class OutlookClient:
    """
    Handles connection to Microsoft Outlook Desktop.
    Now supports Attachments and HTML Body.
    """
    def __init__(self):
        self.available = False
        try:
            import win32com.client as win32
            self.outlook = win32.Dispatch('outlook.application')
            self.available = True
        except ImportError:
            print("Error: 'pywin32' library missing. Run: pip install pywin32")
        except Exception as e:
            print(f"Outlook Error: {e}")

    def create_draft(self, to_email, subject, html_body, attachments=None):
        """
        Creates an Outlook email draft.
        :param attachments: List of file paths (strings) to attach.
        """
        if not self.available:
            return False, "Outlook not available."
        
        try:
            mail = self.outlook.CreateItem(0) # 0 = MailItem
            mail.To = to_email
            mail.Subject = subject
            mail.HTMLBody = html_body
            
            # Handle Attachments
            if attachments:
                for path in attachments:
                    if os.path.exists(path):
                        mail.Attachments.Add(os.path.abspath(path))
                    else:
                        print(f"Warning: Attachment not found at {path}")
            
            mail.Display() # Opens the window so you can review before sending
            return True, "Draft opened in Outlook."
        except Exception as e:
            return False, str(e)


class ReportTemplate:
    """
    Generates the text/HTML body for emails.
    Improved to use Numbered Lists matching your request.
    """
    
    BASE_STYLE = """
    <style>
        body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000; }
        ol { margin-top: 10px; margin-bottom: 10px; }
        li { margin-bottom: 5px; }
        .signature { margin-top: 20px; color: #555; font-size: 10pt; }
    </style>
    """

    @staticmethod
    def render_morning(profile, tasks):
        """
        Generates the 'Please find below my planned tasks...' list.
        """
        list_items = ""
        for t in tasks:
            # Format: "Task Description (Priority - Time)"
            # You can remove priority/time if you want strictly just description
            desc = t.get('description', '')
            details = f"<span style='color:#666; font-size:0.9em;'>({t.get('priority')} - {t.get('planned_time')})</span>"
            
            list_items += f"<li>{desc} {details}</li>"
        
        return f"""
        <html>
        <head>{ReportTemplate.BASE_STYLE}</head>
        <body>
            <p>Dear Manager,</p>
            <p>Please find attached the detailed Excel plan.</p>
            <p><b>Please find below my planned tasks for today:</b></p>
            <ol>
                {list_items}
            </ol>
            <div class="signature">
                Best regards,<br>
                <b>{profile.get('name')}</b><br>
                {profile.get('department')} Unit
            </div>
        </body>
        </html>
        """

    @staticmethod
    def render_evening(profile, tasks, stats):
        """
        Generates the Evening summary list.
        """
        list_items = ""
        for t in tasks:
            status = t.get('status', 'Planned')
            desc = t.get('description', '')
            
            # Color code status text
            color = "black"
            if status == "Completed": color = "green"
            elif status == "Delayed": color = "red"
            
            # Add Result/Delay note if exists
            note = ""
            if t.get('results'): note = f" - <i>Result: {t.get('results')}</i>"
            if t.get('reason_delay'): note = f" - <b>Delay: {t.get('reason_delay')}</b>"
            
            list_items += f"<li><b>{desc}</b> <span style='color:{color}'>[{status.upper()}]</span>{note}</li>"

        return f"""
        <html>
        <head>{ReportTemplate.BASE_STYLE}</head>
        <body>
            <p>Dear Manager,</p>
            <p>Please find attached the Evening Closure Report (Excel).</p>
            <p><b>Summary:</b> {stats['completed']}/{stats['total']} tasks resolved.</p>
            <p><b>Task Status Details:</b></p>
            <ol>
                {list_items}
            </ol>
            <div class="signature">
                Best regards,<br>
                <b>{profile.get('name')}</b><br>
                {profile.get('department')} Unit
            </div>
        </body>
        </html>
        """

class ReportTemplate:
    """Generates Professional HTML Tables for Ops Reports."""
    
    BASE_STYLE = """
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; color: #333; }
        .header { background-color: #064e3b; color: white; padding: 15px; text-align: center; }
        .meta { background-color: #f8fafc; padding: 10px; border-bottom: 2px solid #064e3b; font-size: 12px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { background-color: #f1f5f9; color: #64748b; padding: 10px; text-align: left; border-bottom: 2px solid #cbd5e1; font-size: 11px; text-transform: uppercase; }
        td { padding: 10px; border-bottom: 1px solid #e2e8f0; font-size: 13px; }
        .status-comp { color: #059669; font-weight: bold; }
        .status-delay { color: #e11d48; font-weight: bold; }
        .footer { margin-top: 30px; font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 10px; }
    </style>
    """

    @staticmethod
    def render_morning(profile, tasks):
        rows = ""
        for t in tasks:
            rows += f"""
            <tr>
                <td><b>{t.get('id')}</b></td>
                <td>{t.get('description')}</td>
                <td>{t.get('priority')}</td>
                <td>{t.get('planned_time')}</td>
                <td>{t.get('dependencies', '-')}</td>
                <td>{t.get('morning_notes', '-')}</td>
            </tr>
            """
        
        return f"""
        <html>
        <head>{ReportTemplate.BASE_STYLE}</head>
        <body>
            <div class="header">
                <h2>MORNING DEPLOYMENT MANIFEST</h2>
            </div>
            <div class="meta">
                <b>DATE:</b> {date.today()}<br>
                <b>OPERATOR:</b> {profile.get('name')}<br>
                <b>UNIT:</b> {profile.get('department')}
            </div>
            <table>
                <tr>
                    <th>ID</th> <th>Description</th> <th>Priority</th> <th>Time</th> <th>Dependencies</th> <th>Notes</th>
                </tr>
                {rows}
            </table>
            <div class="footer">Generated by OpsCenter Desktop</div>
        </body>
        </html>
        """

    @staticmethod
    def render_evening(profile, tasks, stats):
        rows = ""
        for t in tasks:
            status_class = ""
            if t.get('status') == 'Completed': status_class = "status-comp"
            if t.get('status') == 'Delayed': status_class = "status-delay"
            
            rows += f"""
            <tr>
                <td><b>{t.get('id')}</b></td>
                <td>{t.get('description')}</td>
                <td class="{status_class}">{t.get('status').upper()}</td>
                <td>{t.get('percent')}%</td>
                <td>{t.get('results', '-')}</td>
                <td>{t.get('reason_delay', '-')}</td>
            </tr>
            """

        return f"""
        <html>
        <head>{ReportTemplate.BASE_STYLE}</head>
        <body>
            <div class="header">
                <h2>EVENING CLOSURE REPORT</h2>
            </div>
            <div class="meta">
                <b>DATE:</b> {date.today()}<br>
                <b>OPERATOR:</b> {profile.get('name')}<br>
                <b>SUMMARY:</b> {stats['completed']}/{stats['total']} Tasks Resolved
            </div>
            <table>
                <tr>
                    <th>ID</th> <th>Description</th> <th>Status</th> <th>%</th> <th>Results</th> <th>Delay Reason</th>
                </tr>
                {rows}
            </table>
            <div class="footer">Generated by OpsCenter Desktop</div>
        </body>
        </html>
        """

# ==========================================
# 1. CONFIGURATION & STYLING
# ==========================================
DATA_FILE = "ops_vault.json"
DEPARTMENTS = ["OPS", "MAINT", "ELEC", "INSTR", "HSE"]

THEME_CSS = """
    QMainWindow { background-color: #f8fafc; }
    QWidget { font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #0f172a; }
    
    QTabWidget::pane { border: 1px solid #e2e8f0; background: white; border-radius: 10px; }
    QTabBar::tab {
        background: white; color: #64748b; padding: 12px 25px;
        font-weight: bold; border-top-left-radius: 8px; border-top-right-radius: 8px;
    }
    QTabBar::tab:selected { background: #f1f5f9; color: #064e3b; border-bottom: 3px solid #064e3b; }

    QLineEdit, QComboBox, QTimeEdit, QTextEdit {
        border: 1px solid #cbd5e1; border-radius: 8px; padding: 8px; background: #ffffff;
    }
    QLineEdit:focus, QComboBox:focus { border: 2px solid #064e3b; }

    QPushButton {
        background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px;
        padding: 8px 16px; font-weight: bold; color: #334155;
    }
    QPushButton:hover { background-color: #f8fafc; border-color: #cbd5e1; }
    
    QTableWidget {
        background-color: white; border: 1px solid #e2e8f0; gridline-color: #f1f5f9;
        selection-background-color: #dcfce7; selection-color: #064e3b;
    }
    QHeaderView::section {
        background-color: #f8fafc; padding: 6px; border: none;
        font-weight: bold; color: #64748b; text-transform: uppercase;
    }
"""

# ==========================================
# 2. DATA MANAGER
# ==========================================
class DataManager:
    def __init__(self):
        self.data = self.load_data()

    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'r') as f:
                    return json.load(f)
            except:
                return self.default_structure()
        return self.default_structure()

    def default_structure(self):
        return {
            "profile": {"name": "", "department": "OPS", "manager_email": ""},
            "tasks": []
        }

    def save_data(self):
        with open(DATA_FILE, 'w') as f:
            json.dump(self.data, f, indent=4)

    def get_today_tasks(self):
        today = date.today().isoformat()
        return [t for t in self.data["tasks"] if t["date"] == today]

    def add_task(self, task):
        self.data["tasks"].append(task)
        self.save_data()

    def update_task(self, task_id, updates):
        for task in self.data["tasks"]:
            if task["id"] == task_id:
                task.update(updates)
                break
        self.save_data()

# ==========================================
# 3. TABS (UI COMPONENTS)
# ==========================================

class ProfileTab(QWidget):
    def __init__(self, data_manager):
        super().__init__()
        self.dm = data_manager
        layout = QVBoxLayout()
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        header = QLabel("Operator Profile")
        header.setStyleSheet("font-size: 24px; font-weight: 900; color: #0f172a;")
        layout.addWidget(header)

        form_layout = QGridLayout()
        form_layout.addWidget(QLabel("FULL NAME"), 0, 0)
        self.name_input = QLineEdit()
        form_layout.addWidget(self.name_input, 0, 1)

        form_layout.addWidget(QLabel("DEPARTMENT"), 1, 0)
        self.dept_input = QComboBox()
        self.dept_input.addItems(DEPARTMENTS)
        form_layout.addWidget(self.dept_input, 1, 1)

        form_layout.addWidget(QLabel("MANAGER EMAIL"), 2, 0)
        self.email_input = QLineEdit()
        form_layout.addWidget(self.email_input, 2, 1)
        layout.addLayout(form_layout)

        save_btn = QPushButton("Update Access")
        save_btn.setStyleSheet("background-color: #064e3b; color: white; border: none; font-weight: bold;")
        save_btn.clicked.connect(self.save_profile)
        layout.addWidget(save_btn)
        layout.addStretch()
        
        purge_btn = QPushButton("⚠️ Purge System Data")
        purge_btn.setStyleSheet("background-color: #fee2e2; color: #991b1b; border: 1px solid #fecaca;")
        purge_btn.clicked.connect(self.purge_data)
        layout.addWidget(purge_btn)

        self.setLayout(layout)
        self.load_profile()

    def load_profile(self):
        p = self.dm.data["profile"]
        self.name_input.setText(p.get("name", ""))
        self.dept_input.setCurrentText(p.get("department", "OPS"))
        self.email_input.setText(p.get("manager_email", ""))

    def save_profile(self):
        self.dm.data["profile"] = {
            "name": self.name_input.text(),
            "department": self.dept_input.currentText(),
            "manager_email": self.email_input.text()
        }
        self.dm.save_data()
        QMessageBox.information(self, "Success", "Profile Updated Successfully")

    def purge_data(self):
        if QMessageBox.question(self, "Danger", "Delete all data?") == QMessageBox.StandardButton.Yes:
            if os.path.exists(DATA_FILE): os.remove(DATA_FILE)
            self.dm.data = self.dm.default_structure()
            self.load_profile()

class MorningTab(QWidget):
    def __init__(self, data_manager):
        super().__init__()
        self.dm = data_manager
        self.outlook = OutlookClient() 
        self.priority_selection = "Medium"
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # --- INPUT CARD ---
        input_card = QFrame()
        input_card.setStyleSheet("background: white; border: 1px solid #e2e8f0; border-radius: 15px;")
        card_layout = QVBoxLayout(input_card)
        
        # Header
        hl = QHBoxLayout()
        hl.addWidget(QLabel("<b>MORNING PLAN REGISTRATION</b>"))
        self.lbl_profile = QLabel("")
        self.lbl_profile.setStyleSheet("color: #64748b; font-size: 11px; font-weight: bold;")
        hl.addWidget(self.lbl_profile)
        card_layout.addLayout(hl)
        
        # Form Grid
        grid = QGridLayout()
        grid.setSpacing(15)

        # Description
        self.inp_desc = QLineEdit()
        self.inp_desc.setPlaceholderText("What needs to be done today?")
        grid.addWidget(QLabel("TASK DESCRIPTION"), 0, 0)
        grid.addWidget(self.inp_desc, 1, 0, 1, 2)
        
        # Time
        self.inp_time = QTimeEdit()
        self.inp_time.setDisplayFormat("HH:mm")
        self.inp_time.setTime(QTime(17, 0))
        grid.addWidget(QLabel("PLANNED TIME"), 0, 2)
        grid.addWidget(self.inp_time, 1, 2)
        
        # Priority Buttons
        grid.addWidget(QLabel("PRIORITY"), 2, 0)
        prio_layout = QHBoxLayout()
        self.btn_group = []
        for p in ["High", "Medium", "Low"]:
            btn = QPushButton(p)
            btn.setCheckable(True)
            btn.clicked.connect(lambda c, x=p: self.set_prio(x))
            self.btn_group.append(btn)
            prio_layout.addWidget(btn)
        self.set_prio("Medium")
        grid.addLayout(prio_layout, 3, 0)
        
        # Dependencies
        self.inp_dep = QLineEdit()
        self.inp_dep.setPlaceholderText("Permits, parts, people...")
        grid.addWidget(QLabel("DEPENDENCIES"), 2, 1)
        grid.addWidget(self.inp_dep, 3, 1)
        
        # Notes
        self.inp_notes = QLineEdit()
        self.inp_notes.setPlaceholderText("SOPs, safety checks...")
        grid.addWidget(QLabel("NOTES"), 2, 2)
        grid.addWidget(self.inp_notes, 3, 2)
        
        card_layout.addLayout(grid)
        
        # Add Button
        add_btn = QPushButton("+ Add to Manifest")
        add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        add_btn.setStyleSheet("""
            QPushButton { background: #064e3b; color: white; padding: 12px; font-weight: bold; font-size: 14px; }
            QPushButton:hover { background: #065f46; }
        """)
        add_btn.clicked.connect(self.add_task)
        card_layout.addWidget(add_btn)
        
        layout.addWidget(input_card)
        
        # --- TASK TABLE ---
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["ID", "Description", "Priority", "Time", "Dept"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setStyleSheet("""
            QTableWidget { border: none; background: white; border-radius: 10px; padding: 10px; }
            QTableWidget::item { border-bottom: 1px solid #f1f5f9; padding: 5px; }
        """)
        layout.addWidget(self.table)
        
        # --- BOTTOM ACTIONS ---
        btn_layout = QHBoxLayout()
        
        # Email Button
        btn_mail = QPushButton("Draft Email (Attach Excel)")
        btn_mail.setIcon(QIcon.fromTheme("mail-message-new"))
        btn_mail.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_mail.setStyleSheet("""
            QPushButton { background: #ecfdf5; color: #064e3b; border: 1px solid #a7f3d0; padding: 10px; }
            QPushButton:hover { background: #d1fae5; }
        """)
        btn_mail.clicked.connect(self.draft_email_with_attachment)
        
        # Manual Export Button
        btn_export = QPushButton("Export CSV Only")
        btn_export.clicked.connect(self.export_csv)
        
        btn_layout.addWidget(btn_mail)
        btn_layout.addWidget(btn_export)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
        self.refresh_ui()

    def set_prio(self, name):
        self.priority_selection = name
        for btn in self.btn_group:
            btn.setChecked(btn.text() == name)
            if btn.isChecked():
                btn.setStyleSheet("background: #dcfce7; color: #064e3b; border: 1px solid #064e3b;")
            else:
                btn.setStyleSheet("background: white; color: #64748b; border: 1px solid #e2e8f0;")

    def refresh_ui(self):
        p = self.dm.data.get("profile", {})
        self.lbl_profile.setText(f"Posting as: {p.get('name', 'Unknown')} | {p.get('department', 'N/A')}")
        
        self.table.setRowCount(0)
        tasks = self.dm.get_today_tasks()
        
        for row, t in enumerate(tasks):
            self.table.insertRow(row)
            
            # ID Formatting
            id_item = QTableWidgetItem(t["id"])
            id_item.setForeground(QColor("#064e3b"))
            id_item.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
            self.table.setItem(row, 0, id_item)
            
            self.table.setItem(row, 1, QTableWidgetItem(t["description"]))
            self.table.setItem(row, 2, QTableWidgetItem(t["priority"]))
            self.table.setItem(row, 3, QTableWidgetItem(t["planned_time"]))
            self.table.setItem(row, 4, QTableWidgetItem(t["department"]))

    def add_task(self):
        desc = self.inp_desc.text()
        if not desc: 
            QMessageBox.warning(self, "Input Error", "Please provide a task description.")
            return
        
        # Robust ID Generation
        existing = self.dm.get_today_tasks()
        if existing:
            # Extract numbers from "T-001", "T-005"
            ids = [int(t['id'].split('-')[1]) for t in existing if '-' in t['id']]
            next_id = max(ids) + 1
        else:
            next_id = 1
            
        p = self.dm.data.get("profile", {})
        new_task = {
            "id": f"T-{next_id:03d}",
            "date": date.today().isoformat(),
            "employee": p.get("name"),
            "department": p.get("department"),
            "description": desc,
            "priority": self.priority_selection,
            "planned_time": self.inp_time.text(),
            "dependencies": self.inp_dep.text(),
            "morning_notes": self.inp_notes.text(),
            # Evening Fields Initialized
            "status": "Planned", "percent": 0, "results": "", "reason_delay": "", "next_steps": ""
        }
        self.dm.add_task(new_task)
        
        # Clear inputs
        self.inp_desc.clear()
        self.inp_dep.clear()
        self.inp_notes.clear()
        self.refresh_ui()

    def export_csv(self):
        tasks = self.dm.get_today_tasks()
        if tasks:
            path, _ = QFileDialog.getSaveFileName(self, "Save", f"Morning_{date.today()}.csv", "CSV (*.csv)")
            if path: pd.DataFrame(tasks).to_csv(path, index=False)

    def draft_email_with_attachment(self):
        """
        1. Saves Excel file locally.
        2. Generates HTML List body.
        3. Opens Outlook with file attached.
        """
        tasks = self.dm.get_today_tasks()
        if not tasks: 
            QMessageBox.warning(self, "Empty", "No tasks to report.")
            return

        p = self.dm.data.get("profile", {})
        
        # 1. Generate Excel File automatically
        file_name = f"Morning_Plan_{date.today()}.xlsx"
        file_path = os.path.abspath(file_name) # Save in current folder
        
        try:
            df = pd.DataFrame(tasks)
            # Filter clean columns for the manager
            cols_to_show = ["id", "description", "priority", "planned_time", "dependencies", "morning_notes"]
            # Only select cols that actually exist in df
            final_cols = [c for c in cols_to_show if c in df.columns]
            
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df[final_cols].to_excel(writer, index=False, sheet_name='Morning Plan')
                
                # Auto-adjust column width (Cosmetic)
                worksheet = writer.sheets['Morning Plan']
                worksheet.set_column(0, 5, 20) 
                
        except Exception as e:
            QMessageBox.critical(self, "Excel Error", f"Could not create attachment:\n{e}")
            return

        # 2. Render HTML Body
        subject = f"Morning Deployment - {date.today()} - {p.get('name')}"
        html_body = ReportTemplate.render_morning(p, tasks)
        
        # 3. Create Draft
        success, msg = self.outlook.create_draft(
            to_email=p.get('manager_email', ''),
            subject=subject,
            html_body=html_body,
            attachments=[file_path]
        )
        
        if not success:
            QMessageBox.critical(self, "Outlook Error", msg)

class EveningTab(QWidget):
    def __init__(self, data_manager):
        super().__init__()
        self.dm = data_manager
        self.outlook = OutlookClient()
        self.current_id = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # --- HEADER CARD ---
        h_frame = QFrame()
        h_frame.setStyleSheet("background: white; border: 1px solid #e2e8f0; border-radius: 15px;")
        hl = QHBoxLayout(h_frame)
        hl.setContentsMargins(20, 15, 20, 15)
        
        title_box = QVBoxLayout()
        lbl_title = QLabel("EVENING UPDATE & CLOSURE")
        lbl_title.setStyleSheet("font-size: 16px; font-weight: 900; color: #064e3b;")
        lbl_sub = QLabel("Log deliverables and performance metrics.")
        lbl_sub.setStyleSheet("color: #64748b; font-size: 12px;")
        title_box.addWidget(lbl_title)
        title_box.addWidget(lbl_sub)
        
        # Action Buttons
        btn_mail = QPushButton("Draft Report (Attach Excel)")
        btn_mail.setIcon(QIcon.fromTheme("mail-message-new"))
        btn_mail.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_mail.setStyleSheet("""
            QPushButton { background: #ecfdf5; color: #064e3b; border: 1px solid #a7f3d0; padding: 10px 15px; font-weight: bold; border-radius: 8px; }
            QPushButton:hover { background: #d1fae5; }
        """)
        btn_mail.clicked.connect(self.draft_email_with_attachment)
        
        btn_exp = QPushButton("Export Excel Only")
        btn_exp.clicked.connect(self.manual_export)
        
        hl.addLayout(title_box)
        hl.addStretch()
        hl.addWidget(btn_mail)
        hl.addWidget(btn_exp)
        
        layout.addWidget(h_frame)
        
        # --- MAIN SPLIT VIEW ---
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        splitter.setStyleSheet("QSplitter::handle { background: #e2e8f0; }")
        
        # LEFT: Task List
        self.list = QTableWidget()
        self.list.setColumnCount(3)
        self.list.setHorizontalHeaderLabels(["ID", "Description", "Status"])
        self.list.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.list.verticalHeader().setVisible(False)
        self.list.setShowGrid(False)
        self.list.setStyleSheet("""
            QTableWidget { background: white; border: 1px solid #e2e8f0; border-radius: 10px; }
            QTableWidget::item { padding: 12px; border-bottom: 1px solid #f1f5f9; }
            QTableWidget::item:selected { background: #f0fdf4; color: #064e3b; }
        """)
        self.list.itemClicked.connect(self.load_task)
        splitter.addWidget(self.list)
        
        # RIGHT: Editor Panel
        editor_container = QFrame()
        editor_container.setStyleSheet("background: white; border-radius: 10px; border: 1px solid #e2e8f0;")
        el = QVBoxLayout(editor_container)
        el.setContentsMargins(25, 25, 25, 25)
        el.setSpacing(20)
        
        self.lbl_task_id = QLabel("Select a task to edit...")
        self.lbl_task_id.setStyleSheet("font-size: 18px; font-weight: 900; color: #0f172a;")
        el.addWidget(self.lbl_task_id)
        
        # Form Grid
        grid = QGridLayout()
        grid.setSpacing(15)
        
        # Status Dropdown
        self.inp_status = QComboBox()
        self.inp_status.addItems(["Planned", "In Progress", "Completed", "Delayed"])
        self.inp_status.currentTextChanged.connect(self.on_status_change)
        grid.addWidget(QLabel("STATUS"), 0, 0)
        grid.addWidget(self.inp_status, 1, 0)
        
        # Progress Slider
        self.inp_pct = QSlider(Qt.Orientation.Horizontal)
        self.inp_pct.setRange(0, 100)
        self.inp_pct.setSingleStep(5)
        self.lbl_pct = QLabel("0%")
        self.lbl_pct.setStyleSheet("font-weight: bold; color: #064e3b; min-width: 40px; text-align: right;")
        self.inp_pct.valueChanged.connect(lambda v: self.lbl_pct.setText(f"{v}%"))
        
        grid.addWidget(QLabel("% COMPLETED"), 0, 1)
        sl_layout = QHBoxLayout()
        sl_layout.addWidget(self.inp_pct)
        sl_layout.addWidget(self.lbl_pct)
        grid.addLayout(sl_layout, 1, 1)
        
        # Results
        self.inp_res = QTextEdit()
        self.inp_res.setPlaceholderText("Detailed outcome of the assignment...")
        self.inp_res.setMaximumHeight(80)
        grid.addWidget(QLabel("RESULTS / NOTES"), 2, 0, 1, 2)
        grid.addWidget(self.inp_res, 3, 0, 1, 2)
        
        # Delay Reason
        self.inp_delay = QLineEdit()
        self.inp_delay.setPlaceholderText("Root cause...")
        grid.addWidget(QLabel("DELAY REASON"), 4, 0)
        grid.addWidget(self.inp_delay, 5, 0)
        
        # Next Steps
        self.inp_next = QLineEdit()
        self.inp_next.setPlaceholderText("Handover instructions...")
        grid.addWidget(QLabel("NEXT STEPS"), 4, 1)
        grid.addWidget(self.inp_next, 5, 1)
        
        el.addLayout(grid)
        
        # Save Button
        btn_save = QPushButton("Save Updates")
        btn_save.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_save.setStyleSheet("""
            QPushButton { background: #0f172a; color: white; padding: 12px; font-weight: bold; border-radius: 8px; }
            QPushButton:hover { background: #1e293b; }
        """)
        btn_save.clicked.connect(self.save)
        el.addWidget(btn_save)
        el.addStretch()
        
        splitter.addWidget(editor_container)
        splitter.setStretchFactor(0, 4) # List 40%
        splitter.setStretchFactor(1, 6) # Editor 60%
        
        layout.addWidget(splitter)
        self.setLayout(layout)

    def refresh_ui(self):
        """Reloads list from DB"""
        self.list.setRowCount(0)
        tasks = self.dm.get_today_tasks()
        
        for row, t in enumerate(tasks):
            self.list.insertRow(row)
            self.list.setItem(row, 0, QTableWidgetItem(t["id"]))
            self.list.setItem(row, 1, QTableWidgetItem(t["description"]))
            
            # Status Color Styling in List
            stat_item = QTableWidgetItem(t["status"])
            stat_font = QFont("Segoe UI", 9, QFont.Weight.Bold)
            stat_item.setFont(stat_font)
            
            if t["status"] == "Completed":
                stat_item.setForeground(QColor("#059669")) # Green
            elif t["status"] == "Delayed":
                stat_item.setForeground(QColor("#e11d48")) # Red
            else:
                stat_item.setForeground(QColor("#64748b")) # Gray
                
            self.list.setItem(row, 2, stat_item)

    def load_task(self, item):
        """Populate Editor"""
        row = item.row()
        tid = self.list.item(row, 0).text()
        tasks = self.dm.get_today_tasks()
        t = next((x for x in tasks if x["id"] == tid), None)
        
        if t:
            self.current_id = tid
            self.lbl_task_id.setText(f"{t['id']} - {t['description']}")
            self.inp_status.setCurrentText(t.get("status", "Planned"))
            self.inp_pct.setValue(int(t.get("percent", 0)))
            self.inp_res.setPlainText(t.get("results", ""))
            self.inp_delay.setText(t.get("reason_delay", ""))
            self.inp_next.setText(t.get("next_steps", ""))
            
            # Trigger logic check
            self.on_status_change(t.get("status", "Planned"))

    def on_status_change(self, txt):
        """Logic: Disable Delay input if not Delayed"""
        if txt == "Delayed":
            self.inp_delay.setEnabled(True)
            self.inp_delay.setStyleSheet("background: #fff1f2; border: 1px solid #fda4af;")
            self.inp_delay.setPlaceholderText("Required: Root cause of delay")
        else:
            self.inp_delay.setEnabled(False)
            self.inp_delay.setStyleSheet("background: #f1f5f9; border: 1px solid #e2e8f0; color: #94a3b8;")
            self.inp_delay.setPlaceholderText("N/A")
            self.inp_delay.clear()
            
        if txt == "Completed":
            self.inp_pct.setValue(100)

    def save(self):
        if self.current_id:
            self.dm.update_task(self.current_id, {
                "status": self.inp_status.currentText(),
                "percent": self.inp_pct.value(),
                "results": self.inp_res.toPlainText(),
                "reason_delay": self.inp_delay.text(),
                "next_steps": self.inp_next.text()
            })
            self.refresh_ui()
            # Brief visual feedback could go here

    def generate_excel_file(self, filename):
        """Helper to create the Colored Excel file"""
        tasks = self.dm.get_today_tasks()
        if not tasks: return False
        
        df = pd.DataFrame(tasks)
        # Select columns for the report
        cols = ["id", "description", "status", "percent", "results", "reason_delay", "next_steps"]
        # Ensure cols exist
        final_cols = [c for c in cols if c in df.columns]
        
        try:
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                df[final_cols].to_excel(writer, index=False, sheet_name='Evening Report')
                workbook = writer.book
                worksheet = writer.sheets['Evening Report']
                
                # Styles
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0f172a', 'font_color': 'white'})
                green_fmt = workbook.add_format({'bg_color': '#dcfce7', 'border': 1})
                red_fmt = workbook.add_format({'bg_color': '#fee2e2', 'border': 1})
                base_fmt = workbook.add_format({'border': 1})
                
                # Apply Header
                for col_num, value in enumerate(final_cols):
                    worksheet.write(0, col_num, value.upper(), header_fmt)
                    worksheet.set_column(col_num, col_num, 20)
                
                # Apply Row Colors based on Status
                if "status" in df.columns:
                    for row_num, status in enumerate(df["status"]):
                        row_idx = row_num + 1
                        fmt = base_fmt
                        if status == "Completed": fmt = green_fmt
                        elif status == "Delayed": fmt = red_fmt
                        
                        # Write the whole row with the specific format
                        for col_num, col_name in enumerate(final_cols):
                            val = df.iloc[row_num][col_name]
                            # Handle NaNs
                            if pd.isna(val): val = ""
                            worksheet.write(row_idx, col_num, val, fmt)
            return True
        except Exception as e:
            print(f"Excel Generation Error: {e}")
            return False

    def manual_export(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save", f"Evening_{date.today()}.xlsx", "Excel (*.xlsx)")
        if path:
            if self.generate_excel_file(path):
                QMessageBox.information(self, "Success", "File saved successfully.")

    def draft_email_with_attachment(self):
        """
        1. Generates Colored Excel.
        2. Renders Evening HTML.
        3. Opens Outlook.
        """
        tasks = self.dm.get_today_tasks()
        if not tasks: 
            QMessageBox.warning(self, "Empty", "No tasks to report.")
            return

        p = self.dm.data.get("profile", {})
        
        # 1. Create Attachment
        file_name = f"Evening_Report_{date.today()}.xlsx"
        file_path = os.path.abspath(file_name)
        if not self.generate_excel_file(file_path):
            QMessageBox.critical(self, "Error", "Failed to generate Excel attachment.")
            return

        # 2. Stats for Body
        comp = len([t for t in tasks if t.get('status') == 'Completed'])
        stats = {'completed': comp, 'total': len(tasks)}
        
        # 3. HTML Body
        html_body = ReportTemplate.render_evening(p, tasks, stats)
        subject = f"Evening Closure - {date.today()} - {p.get('name')}"
        
        # 4. Outlook
        success, msg = self.outlook.create_draft(
            to_email=p.get('manager_email', ''),
            subject=subject,
            html_body=html_body,
            attachments=[file_path]
        )
        
        if not success:
            QMessageBox.critical(self, "Outlook Error", msg)

class DashboardTab(QWidget):
    def __init__(self, data_manager):
        super().__init__()
        self.dm = data_manager
        layout = QVBoxLayout()
        hl = QHBoxLayout()
        self.combo = QComboBox()
        self.combo.addItems(["Last 7 Days", "Last 30 Days", "All Time"])
        self.combo.currentTextChanged.connect(self.refresh)
        hl.addWidget(QLabel("<b>OPS INTELLIGENCE</b>"))
        hl.addStretch()
        hl.addWidget(self.combo)
        layout.addLayout(hl)
        self.kpi_layout = QHBoxLayout()
        self.kpi_lbls = {}
        for k in ["Total", "Rate", "Delayed"]:
            f = QFrame()
            f.setStyleSheet("background: white; border-radius: 10px; border: 1px solid #e2e8f0;")
            v = QVBoxLayout(f)
            l = QLabel("0")
            l.setStyleSheet("font-size: 24px; font-weight: bold; color: #064e3b;")
            v.addWidget(l, alignment=Qt.AlignmentFlag.AlignCenter)
            v.addWidget(QLabel(k), alignment=Qt.AlignmentFlag.AlignCenter)
            self.kpi_lbls[k] = l
            self.kpi_layout.addWidget(f)
        layout.addLayout(self.kpi_layout)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        cw = QWidget()
        gl = QGridLayout(cw)
        self.fig1 = Figure(figsize=(5,3))
        self.can1 = FigureCanvas(self.fig1)
        gl.addWidget(self.can1, 0, 0, 1, 2)
        self.fig2 = Figure(figsize=(4,3))
        self.can2 = FigureCanvas(self.fig2)
        gl.addWidget(self.can2, 1, 0)
        self.fig3 = Figure(figsize=(4,3))
        self.can3 = FigureCanvas(self.fig3)
        gl.addWidget(self.can3, 1, 1)
        scroll.setWidget(cw)
        layout.addWidget(scroll)
        self.setLayout(layout)

    def refresh(self):
        tasks = self.dm.data.get("tasks", [])
        if not tasks: return
        df = pd.DataFrame(tasks)
        df['date'] = pd.to_datetime(df['date'])
        today = pd.to_datetime(date.today())
        sel = self.combo.currentText()
        if "7" in sel: df = df[df['date'] >= today - timedelta(days=6)]
        elif "30" in sel: df = df[df['date'] >= today - timedelta(days=29)]
        if df.empty: return
        total = len(df)
        comp = len(df[df['status']=='Completed'])
        dely = len(df[df['status']=='Delayed'])
        rate = int(comp/total*100) if total else 0
        self.kpi_lbls["Total"].setText(str(total))
        self.kpi_lbls["Rate"].setText(f"{rate}%")
        self.kpi_lbls["Delayed"].setText(str(dely))
        self.fig1.clear()
        ax1 = self.fig1.add_subplot(111)
        daily = df.groupby('date').size()
        ax1.plot(daily.index, daily.values, marker='o')
        ax1.set_title("Timeline")
        self.can1.draw()
        self.fig2.clear()
        ax2 = self.fig2.add_subplot(111)
        vc = df['status'].value_counts()
        ax2.pie(vc, labels=vc.index, autopct='%1.1f%%')
        ax2.set_title("Status Mix")
        self.can2.draw()
        self.fig3.clear()
        ax3 = self.fig3.add_subplot(111)
        pc = df['priority'].value_counts()
        ax3.bar(pc.index, pc.values, color='#064e3b')
        ax3.set_title("Priorities")
        self.can3.draw()

# ==========================================
# 4. APP MAIN
# ==========================================
class OpsApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OpsCenter | Engineering Management")
        self.resize(1100, 750)
        self.dm = DataManager()
        self.tabs = QTabWidget()
        self.tabs.addTab(ProfileTab(self.dm), "👤 Identity")
        self.morning = MorningTab(self.dm)
        self.tabs.addTab(self.morning, "☀️ Morning Plan")
        self.evening = EveningTab(self.dm)
        self.tabs.addTab(self.evening, "🌙 Evening Report")
        self.dash = DashboardTab(self.dm)
        self.tabs.addTab(self.dash, "📊 Intelligence")
        self.tabs.currentChanged.connect(self.on_change)
        self.setCentralWidget(self.tabs)
        self.setStyleSheet(THEME_CSS)

    def on_change(self, i):
        if i == 1: self.morning.refresh_ui()
        if i == 2: self.evening.refresh_ui()
        if i == 3: self.dash.refresh()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    w = OpsApp()
    w.show()
    sys.exit(app.exec())