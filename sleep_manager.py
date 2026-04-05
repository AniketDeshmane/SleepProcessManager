"""
Windows Power & Sleep Manager Dashboard
========================================
A real-time monitoring dashboard for diagnosing Windows sleep/wake issues.
Requires Administrator privileges to run powercfg commands.

Author: SleepProcessManager
Tech: Python 3 + PyQt5
"""

import sys
import os
import re
import ctypes
import subprocess
import signal
from datetime import datetime
from functools import partial

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QFrame, QScrollArea, QGroupBox,
    QSplitter, QMessageBox, QSizePolicy, QGraphicsDropShadowEffect,
    QSystemTrayIcon, QMenu, QAction, QToolTip
)
from PyQt5.QtCore import (
    Qt, QTimer, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve,
    QSize, QPoint
)
from PyQt5.QtGui import (
    QFont, QColor, QPalette, QIcon, QPainter, QLinearGradient,
    QBrush, QPen, QFontDatabase, QPixmap
)


# ─── Color Palette ───────────────────────────────────────────────────────────

class Colors:
    BG_PRIMARY = "#0d1117"
    BG_SECONDARY = "#161b22"
    BG_CARD = "#1c2333"
    BG_CARD_HOVER = "#222d3f"
    BORDER = "#30363d"
    BORDER_ACCENT = "#3b82f6"
    TEXT_PRIMARY = "#e6edf3"
    TEXT_SECONDARY = "#8b949e"
    TEXT_MUTED = "#6e7681"
    ACCENT_BLUE = "#3b82f6"
    ACCENT_BLUE_HOVER = "#60a5fa"
    ACCENT_GREEN = "#22c55e"
    ACCENT_GREEN_DIM = "#166534"
    ACCENT_RED = "#ef4444"
    ACCENT_RED_DIM = "#7f1d1d"
    ACCENT_YELLOW = "#f59e0b"
    ACCENT_ORANGE = "#f97316"
    ACCENT_PURPLE = "#a855f7"
    GRADIENT_START = "#1e3a5f"
    GRADIENT_END = "#0d1117"


# ─── Admin Check ─────────────────────────────────────────────────────────────

def is_admin():
    """Check if the script is running with Administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def request_admin():
    """Re-launch the script with elevated privileges."""
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )
    sys.exit(0)


# ─── Worker Thread for powercfg ──────────────────────────────────────────────

class PowerCfgWorker(QThread):
    """Background thread to run powercfg commands without freezing the UI."""
    results_ready = pyqtSignal(dict)

    def run(self):
        results = {}
        commands = {
            "requests": "powercfg /requests",
            "lastwake": "powercfg /lastwake",
            "waketimers": "powercfg /waketimers",
        }
        for key, cmd in commands.items():
            try:
                proc = subprocess.run(
                    cmd, capture_output=True, text=True, shell=True,
                    timeout=10, creationflags=subprocess.CREATE_NO_WINDOW
                )
                results[key] = proc.stdout.strip() if proc.stdout else proc.stderr.strip()
            except subprocess.TimeoutExpired:
                results[key] = "[Timeout] Command took too long."
            except Exception as e:
                results[key] = f"[Error] {e}"
        self.results_ready.emit(results)


# ─── Parsing Helpers ─────────────────────────────────────────────────────────

def parse_requests(text):
    """
    Parse powercfg /requests output into a dict of category -> list of blockers.
    Each blocker is a dict with 'raw' text and optionally 'process' name.
    """
    categories = {}
    current_cat = None

    for line in text.splitlines():
        line_stripped = line.strip()
        if not line_stripped:
            continue

        # Category headers end with ':'
        cat_match = re.match(r'^([A-Z]+):$', line_stripped)
        if cat_match:
            current_cat = cat_match.group(1)
            categories[current_cat] = []
            continue

        if current_cat is not None:
            if line_stripped.lower() == "none.":
                continue
            # Try to extract a process / exe name
            proc_match = re.search(r'(\b\w+\.exe\b)', line_stripped, re.IGNORECASE)
            blocker = {
                "raw": line_stripped,
                "process": proc_match.group(1) if proc_match else None,
                "category": current_cat,
            }
            categories[current_cat].append(blocker)

    return categories


def has_active_blockers(categories):
    """Return True if any EXECUTION or SYSTEM blockers exist."""
    for cat in ("EXECUTION", "SYSTEM"):
        if categories.get(cat):
            return True
    return False


def get_all_blockers(categories):
    """Flatten all blockers across categories."""
    all_blockers = []
    for cat, blockers in categories.items():
        all_blockers.extend(blockers)
    return all_blockers


# ─── Traffic Light Widget ────────────────────────────────────────────────────

class TrafficLight(QWidget):
    """A glowing traffic-light indicator: green = clear, red = blockers active."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._color = QColor(Colors.ACCENT_GREEN)
        self._is_red = False
        self.setFixedSize(28, 28)
        self._pulse_opacity = 1.0
        self._pulse_timer = QTimer(self)
        self._pulse_timer.timeout.connect(self._pulse_tick)
        self._pulse_dir = -0.05

    def set_status(self, is_blocked):
        self._is_red = is_blocked
        self._color = QColor(Colors.ACCENT_RED if is_blocked else Colors.ACCENT_GREEN)
        if is_blocked and not self._pulse_timer.isActive():
            self._pulse_timer.start(60)
        elif not is_blocked:
            self._pulse_timer.stop()
            self._pulse_opacity = 1.0
        self.update()

    def _pulse_tick(self):
        self._pulse_opacity += self._pulse_dir
        if self._pulse_opacity <= 0.35:
            self._pulse_dir = 0.05
        elif self._pulse_opacity >= 1.0:
            self._pulse_dir = -0.05
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Outer glow
        glow_color = QColor(self._color)
        glow_color.setAlphaF(0.25 * self._pulse_opacity)
        painter.setBrush(QBrush(glow_color))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(0, 0, 28, 28)

        # Inner circle
        inner_color = QColor(self._color)
        inner_color.setAlphaF(self._pulse_opacity)
        painter.setBrush(QBrush(inner_color))
        painter.setPen(QPen(QColor(self._color.name()), 1))
        painter.drawEllipse(5, 5, 18, 18)

        # Shine dot
        shine = QColor("#ffffff")
        shine.setAlphaF(0.35 * self._pulse_opacity)
        painter.setBrush(QBrush(shine))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(9, 8, 6, 5)

        painter.end()


# ─── Styled Card Widget ─────────────────────────────────────────────────────

class Card(QFrame):
    """A dark-themed card container with subtle border and shadow."""

    def __init__(self, title="", icon_char="", parent=None):
        super().__init__(parent)
        self.setObjectName("Card")
        self.setStyleSheet(f"""
            QFrame#Card {{
                background-color: {Colors.BG_CARD};
                border: 1px solid {Colors.BORDER};
                border-radius: 12px;
                padding: 0px;
            }}
        """)

        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(16, 14, 16, 14)
        self._layout.setSpacing(10)

        if title:
            header = QHBoxLayout()
            header.setSpacing(8)

            if icon_char:
                icon_lbl = QLabel(icon_char)
                icon_lbl.setFont(QFont("Segoe UI Emoji", 14))
                icon_lbl.setStyleSheet(f"color: {Colors.ACCENT_BLUE}; background: transparent; border: none;")
                header.addWidget(icon_lbl)

            title_lbl = QLabel(title)
            title_lbl.setFont(QFont("Segoe UI", 13, QFont.DemiBold))
            title_lbl.setStyleSheet(f"color: {Colors.TEXT_PRIMARY}; background: transparent; border: none;")
            header.addWidget(title_lbl)
            header.addStretch()

            self._layout.addLayout(header)

            # Separator line
            sep = QFrame()
            sep.setFixedHeight(1)
            sep.setStyleSheet(f"background-color: {Colors.BORDER}; border: none;")
            self._layout.addWidget(sep)

    def add_widget(self, widget):
        self._layout.addWidget(widget)

    def add_layout(self, layout):
        self._layout.addLayout(layout)


# ─── Blocker Row Widget ─────────────────────────────────────────────────────

class BlockerRow(QFrame):
    """A single row for a detected sleep blocker with action buttons."""

    kill_requested = pyqtSignal(str)
    override_requested = pyqtSignal(str)

    def __init__(self, blocker_info, parent=None):
        super().__init__(parent)
        self.blocker = blocker_info
        self.setStyleSheet(f"""
            QFrame {{
                background-color: {Colors.BG_SECONDARY};
                border: 1px solid {Colors.ACCENT_RED_DIM};
                border-radius: 8px;
                padding: 4px;
            }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(10)

        # Category badge
        cat = blocker_info.get("category", "?")
        badge_colors = {
            "EXECUTION": (Colors.ACCENT_RED, Colors.ACCENT_RED_DIM),
            "SYSTEM": (Colors.ACCENT_ORANGE, "#4a2000"),
            "DISPLAY": (Colors.ACCENT_YELLOW, "#4a3800"),
            "AWAYMODE": (Colors.ACCENT_PURPLE, "#3b1d6e"),
        }
        fg, bg = badge_colors.get(cat, (Colors.TEXT_SECONDARY, Colors.BG_CARD))

        badge = QLabel(cat)
        badge.setFont(QFont("Segoe UI", 9, QFont.Bold))
        badge.setAlignment(Qt.AlignCenter)
        badge.setFixedWidth(90)
        badge.setStyleSheet(f"""
            color: {fg};
            background-color: {bg};
            border-radius: 4px;
            padding: 3px 8px;
            border: none;
        """)
        layout.addWidget(badge)

        # Blocker text
        text_lbl = QLabel(blocker_info["raw"])
        text_lbl.setFont(QFont("Cascadia Code", 10))
        text_lbl.setStyleSheet(f"color: {Colors.TEXT_PRIMARY}; border: none; background: transparent;")
        text_lbl.setWordWrap(True)
        text_lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout.addWidget(text_lbl, stretch=1)

        process_name = blocker_info.get("process")
        if process_name:
            # Kill button
            kill_btn = QPushButton(f"⛔ Kill {process_name}")
            kill_btn.setFont(QFont("Segoe UI", 9, QFont.DemiBold))
            kill_btn.setCursor(Qt.PointingHandCursor)
            kill_btn.setToolTip(f"Terminate {process_name} via taskkill")
            kill_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {Colors.ACCENT_RED_DIM};
                    color: {Colors.ACCENT_RED};
                    border: 1px solid {Colors.ACCENT_RED};
                    border-radius: 6px;
                    padding: 5px 14px;
                }}
                QPushButton:hover {{
                    background-color: {Colors.ACCENT_RED};
                    color: #ffffff;
                }}
            """)
            kill_btn.clicked.connect(lambda: self.kill_requested.emit(process_name))
            layout.addWidget(kill_btn)

        # Override button (works for any blocker with a process name)
        if process_name:
            override_btn = QPushButton("🛡️ Override")
            override_btn.setFont(QFont("Segoe UI", 9, QFont.DemiBold))
            override_btn.setCursor(Qt.PointingHandCursor)
            override_btn.setToolTip(
                f"Run: powercfg /requestsoverride PROCESS {process_name} EXECUTION"
            )
            override_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    color: {Colors.ACCENT_BLUE};
                    border: 1px solid {Colors.ACCENT_BLUE};
                    border-radius: 6px;
                    padding: 5px 14px;
                }}
                QPushButton:hover {{
                    background-color: {Colors.ACCENT_BLUE};
                    color: #ffffff;
                }}
            """)
            override_btn.clicked.connect(lambda: self.override_requested.emit(process_name))
            layout.addWidget(override_btn)


# ─── Main Window ─────────────────────────────────────────────────────────────

class SleepManagerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("⚡ Windows Power & Sleep Manager")
        self.setMinimumSize(1080, 780)
        self.resize(1200, 860)

        self._apply_global_styles()
        self._build_ui()

        # Worker thread
        self.worker = None

        # Polling timer – every 5 seconds
        self.poll_timer = QTimer(self)
        self.poll_timer.timeout.connect(self._run_scan)
        self.poll_timer.start(5000)

        # Initial scan
        self._run_scan()

    # ── Styles ────────────────────────────────────────────────────────────

    def _apply_global_styles(self):
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {Colors.BG_PRIMARY};
            }}
            QScrollArea {{
                background-color: transparent;
                border: none;
            }}
            QScrollBar:vertical {{
                background: {Colors.BG_SECONDARY};
                width: 8px;
                border-radius: 4px;
            }}
            QScrollBar::handle:vertical {{
                background: {Colors.BORDER};
                border-radius: 4px;
                min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {Colors.TEXT_MUTED};
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
            QScrollBar:horizontal {{
                height: 0px;
            }}
            QToolTip {{
                background-color: {Colors.BG_CARD};
                color: {Colors.TEXT_PRIMARY};
                border: 1px solid {Colors.BORDER};
                padding: 6px 10px;
                border-radius: 6px;
                font-size: 12px;
            }}
        """)

    # ── Build UI ──────────────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(20, 16, 20, 16)
        root_layout.setSpacing(16)

        # ── Header ────────────────────────────────────────────────────────
        header_frame = QFrame()
        header_frame.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 {Colors.GRADIENT_START}, stop:1 {Colors.BG_PRIMARY}
                );
                border-radius: 14px;
                border: 1px solid {Colors.BORDER};
            }}
        """)
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(24, 16, 24, 16)

        # Title area
        title_area = QVBoxLayout()
        title_area.setSpacing(4)

        title_row = QHBoxLayout()
        title_row.setSpacing(12)

        self.traffic_light = TrafficLight()
        title_row.addWidget(self.traffic_light)

        app_title = QLabel("Power & Sleep Manager")
        app_title.setFont(QFont("Segoe UI", 20, QFont.Bold))
        app_title.setStyleSheet(f"color: {Colors.TEXT_PRIMARY}; background: transparent; border: none;")
        title_row.addWidget(app_title)
        title_row.addStretch()

        title_area.addLayout(title_row)

        self.status_label = QLabel("● Scanning…")
        self.status_label.setFont(QFont("Segoe UI", 11))
        self.status_label.setStyleSheet(f"color: {Colors.TEXT_SECONDARY}; background: transparent; border: none;")
        title_area.addWidget(self.status_label)

        header_layout.addLayout(title_area, stretch=1)

        # Action buttons
        btn_area = QHBoxLayout()
        btn_area.setSpacing(10)

        self.scan_btn = self._make_action_btn("🔄 Scan Now", Colors.ACCENT_BLUE)
        self.scan_btn.clicked.connect(self._run_scan)
        btn_area.addWidget(self.scan_btn)

        devmgr_btn = self._make_action_btn("🖥️ Device Manager", Colors.ACCENT_PURPLE)
        devmgr_btn.clicked.connect(self._open_device_manager)
        btn_area.addWidget(devmgr_btn)

        poweropts_btn = self._make_action_btn("⚙️ Power Options", Colors.ACCENT_YELLOW)
        poweropts_btn.clicked.connect(self._open_power_options)
        btn_area.addWidget(poweropts_btn)

        header_layout.addLayout(btn_area)
        root_layout.addWidget(header_frame)

        # ── Main Content Splitter ─────────────────────────────────────────
        splitter = QSplitter(Qt.Vertical)
        splitter.setStyleSheet(f"""
            QSplitter::handle {{
                background: {Colors.BORDER};
                height: 2px;
                margin: 6px 40px;
                border-radius: 1px;
            }}
        """)

        # Top half: 3-column monitoring panels
        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(14)

        # ── Requests Panel ─────────────
        self.requests_card = Card("Sleep Blockers", "🚫")
        self.requests_scroll = QScrollArea()
        self.requests_scroll.setWidgetResizable(True)
        self.requests_container = QWidget()
        self.requests_container.setStyleSheet("background: transparent;")
        self.requests_layout = QVBoxLayout(self.requests_container)
        self.requests_layout.setContentsMargins(0, 0, 0, 0)
        self.requests_layout.setSpacing(8)
        self.requests_layout.addStretch()
        self.requests_scroll.setWidget(self.requests_container)
        self.requests_card.add_widget(self.requests_scroll)
        top_layout.addWidget(self.requests_card, stretch=3)

        # Right column (Last Wake + Wake Timers stacked)
        right_col = QVBoxLayout()
        right_col.setSpacing(14)

        # ── Last Wake Panel ────────────
        self.lastwake_card = Card("Last Wake Source", "⏰")
        self.lastwake_text = QLabel("Scanning…")
        self.lastwake_text.setFont(QFont("Cascadia Code", 10))
        self.lastwake_text.setWordWrap(True)
        self.lastwake_text.setStyleSheet(f"color: {Colors.TEXT_PRIMARY}; background: transparent; border: none;")
        self.lastwake_card.add_widget(self.lastwake_text)
        right_col.addWidget(self.lastwake_card)

        # ── Wake Timers Panel ──────────
        self.timers_card = Card("Wake Timers", "⏲️")
        self.timers_text = QLabel("Scanning…")
        self.timers_text.setFont(QFont("Cascadia Code", 10))
        self.timers_text.setWordWrap(True)
        self.timers_text.setStyleSheet(f"color: {Colors.TEXT_PRIMARY}; background: transparent; border: none;")
        self.timers_card.add_widget(self.timers_text)
        right_col.addWidget(self.timers_card)

        top_layout.addLayout(right_col, stretch=2)

        splitter.addWidget(top_widget)

        # ── Bottom: Log Window ────────────────────────────────────────────
        log_card = Card("Event Log", "📋")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Cascadia Code", 9))
        self.log_text.setStyleSheet(f"""
            QTextEdit {{
                background-color: {Colors.BG_SECONDARY};
                color: {Colors.TEXT_SECONDARY};
                border: 1px solid {Colors.BORDER};
                border-radius: 8px;
                padding: 10px;
            }}
        """)
        self.log_text.setMinimumHeight(120)
        log_card.add_widget(self.log_text)

        # Log controls
        log_controls = QHBoxLayout()
        clear_btn = QPushButton("🗑️ Clear Log")
        clear_btn.setCursor(Qt.PointingHandCursor)
        clear_btn.setFont(QFont("Segoe UI", 9))
        clear_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {Colors.TEXT_MUTED};
                border: 1px solid {Colors.BORDER};
                border-radius: 6px;
                padding: 4px 12px;
            }}
            QPushButton:hover {{
                color: {Colors.TEXT_PRIMARY};
                border-color: {Colors.TEXT_MUTED};
            }}
        """)
        clear_btn.clicked.connect(self.log_text.clear)
        log_controls.addStretch()
        log_controls.addWidget(clear_btn)
        log_card.add_layout(log_controls)

        splitter.addWidget(log_card)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)

        root_layout.addWidget(splitter)

        # Footer
        footer = QLabel("Auto-refresh every 5 seconds  •  Running as Administrator  •  powercfg monitoring active")
        footer.setFont(QFont("Segoe UI", 9))
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet(f"color: {Colors.TEXT_MUTED};")
        root_layout.addWidget(footer)

    # ── Helper: Create action button ──────────────────────────────────────

    def _make_action_btn(self, text, accent_color):
        btn = QPushButton(text)
        btn.setFont(QFont("Segoe UI", 10, QFont.DemiBold))
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {accent_color};
                border: 1px solid {accent_color};
                border-radius: 8px;
                padding: 8px 18px;
            }}
            QPushButton:hover {{
                background-color: {accent_color};
                color: #ffffff;
            }}
            QPushButton:pressed {{
                opacity: 0.8;
            }}
        """)
        return btn

    # ── Scanning ──────────────────────────────────────────────────────────

    def _run_scan(self):
        if self.worker and self.worker.isRunning():
            return  # Skip overlapping scans
        self.scan_btn.setEnabled(False)
        self.scan_btn.setText("⏳ Scanning…")
        self.worker = PowerCfgWorker()
        self.worker.results_ready.connect(self._on_results)
        self.worker.finished.connect(self._on_scan_done)
        self.worker.start()

    def _on_scan_done(self):
        self.scan_btn.setEnabled(True)
        self.scan_btn.setText("🔄 Scan Now")

    def _on_results(self, results):
        now = datetime.now().strftime("%H:%M:%S")

        # ── Parse /requests ───────────────────────────────────────────────
        requests_text = results.get("requests", "")
        categories = parse_requests(requests_text)
        blocked = has_active_blockers(categories)
        all_blockers = get_all_blockers(categories)

        # Update traffic light
        self.traffic_light.set_status(blocked)

        # Status text
        if blocked:
            blocker_count = len(all_blockers)
            self.status_label.setText(
                f"🔴 {blocker_count} active blocker{'s' if blocker_count != 1 else ''} detected  •  Last scan: {now}"
            )
            self.status_label.setStyleSheet(f"color: {Colors.ACCENT_RED}; background: transparent; border: none;")
        else:
            self.status_label.setText(f"🟢 System clear — no sleep blockers  •  Last scan: {now}")
            self.status_label.setStyleSheet(f"color: {Colors.ACCENT_GREEN}; background: transparent; border: none;")

        # Rebuild blockers list
        self._clear_layout(self.requests_layout)

        if all_blockers:
            for blocker in all_blockers:
                row = BlockerRow(blocker)
                row.kill_requested.connect(self._kill_process)
                row.override_requested.connect(self._override_process)
                self.requests_layout.addWidget(row)
        else:
            empty_label = QLabel("✅  No active sleep blockers found.")
            empty_label.setFont(QFont("Segoe UI", 11))
            empty_label.setAlignment(Qt.AlignCenter)
            empty_label.setStyleSheet(f"""
                color: {Colors.ACCENT_GREEN};
                background: {Colors.ACCENT_GREEN_DIM};
                border: 1px solid {Colors.ACCENT_GREEN};
                border-radius: 8px;
                padding: 20px;
            """)
            self.requests_layout.addWidget(empty_label)

        self.requests_layout.addStretch()

        # ── Last Wake ────────────────────────────────────────────────────
        lastwake = results.get("lastwake", "N/A")
        self.lastwake_text.setText(lastwake if lastwake else "No data available.")

        # ── Wake Timers ──────────────────────────────────────────────────
        timers = results.get("waketimers", "N/A")
        self.timers_text.setText(timers if timers else "No wake timers set.")

        # ── Log entry ────────────────────────────────────────────────────
        log_line = f"[{now}] "
        if blocked:
            names = [b["process"] or b["raw"][:40] for b in all_blockers]
            log_line += f"⚠ Blockers: {', '.join(names)}"
        else:
            log_line += "✓ System clear."
        self._log(log_line)

    # ── Actions ───────────────────────────────────────────────────────────

    def _kill_process(self, process_name):
        reply = QMessageBox.question(
            self,
            "Confirm Kill",
            f"Are you sure you want to terminate <b>{process_name}</b>?<br><br>"
            f"<code>taskkill /F /IM {process_name}</code>",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            try:
                result = subprocess.run(
                    f"taskkill /F /IM {process_name}",
                    capture_output=True, text=True, shell=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                output = result.stdout.strip() or result.stderr.strip()
                self._log(f"[KILL] {process_name}: {output}")
                QMessageBox.information(self, "Process Kill", output)
                self._run_scan()  # Refresh immediately
            except Exception as e:
                self._log(f"[KILL ERROR] {process_name}: {e}")
                QMessageBox.critical(self, "Error", str(e))

    def _override_process(self, process_name):
        cmd = f"powercfg /requestsoverride PROCESS {process_name} EXECUTION"
        reply = QMessageBox.question(
            self,
            "Confirm Override",
            f"Add a permanent sleep override for <b>{process_name}</b>?<br><br>"
            f"<code>{cmd}</code><br><br>"
            f"This means Windows will ignore EXECUTION requests from this process.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            try:
                result = subprocess.run(
                    cmd, capture_output=True, text=True, shell=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                output = result.stdout.strip() or result.stderr.strip() or "Override applied successfully."
                self._log(f"[OVERRIDE] {process_name}: {output}")
                QMessageBox.information(self, "Override Result", output)
                self._run_scan()
            except Exception as e:
                self._log(f"[OVERRIDE ERROR] {process_name}: {e}")
                QMessageBox.critical(self, "Error", str(e))

    def _open_device_manager(self):
        try:
            subprocess.Popen(
                "devmgmt.msc", shell=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            self._log("[ACTION] Opened Device Manager.")
        except Exception as e:
            self._log(f"[ERROR] Could not open Device Manager: {e}")

    def _open_power_options(self):
        try:
            subprocess.Popen(
                "control powercfg.cpl", shell=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            self._log("[ACTION] Opened Power Options.")
        except Exception as e:
            self._log(f"[ERROR] Could not open Power Options: {e}")

    # ── Utilities ─────────────────────────────────────────────────────────

    def _log(self, message):
        self.log_text.append(message)
        # Auto-scroll to bottom
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def closeEvent(self, event):
        self.poll_timer.stop()
        if self.worker and self.worker.isRunning():
            self.worker.quit()
            self.worker.wait(2000)
        event.accept()


# ─── Entry Point ─────────────────────────────────────────────────────────────

def main():
    # High-DPI support
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Dark Fusion palette (fallback)
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(Colors.BG_PRIMARY))
    palette.setColor(QPalette.WindowText, QColor(Colors.TEXT_PRIMARY))
    palette.setColor(QPalette.Base, QColor(Colors.BG_SECONDARY))
    palette.setColor(QPalette.AlternateBase, QColor(Colors.BG_CARD))
    palette.setColor(QPalette.Text, QColor(Colors.TEXT_PRIMARY))
    palette.setColor(QPalette.Button, QColor(Colors.BG_CARD))
    palette.setColor(QPalette.ButtonText, QColor(Colors.TEXT_PRIMARY))
    palette.setColor(QPalette.Highlight, QColor(Colors.ACCENT_BLUE))
    palette.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    app.setPalette(palette)

    # Admin check
    if not is_admin():
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Administrator Required")
        msg.setText(
            "This application requires <b>Administrator privileges</b> "
            "to run <code>powercfg</code> commands.\n\n"
            "Click OK to restart as Administrator, or Cancel to exit."
        )
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.setDefaultButton(QMessageBox.Ok)
        result = msg.exec_()
        if result == QMessageBox.Ok:
            request_admin()
        else:
            sys.exit(1)

    window = SleepManagerWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
