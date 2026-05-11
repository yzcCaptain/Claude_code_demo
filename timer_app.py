"""
Desktop Timer Application - 桌面定时器
- Countdown (倒计时) and Stopwatch (正计时)
- Always on top with acrylic blur background
- Drag to top of screen → compact mode
- Timer end → restore and pop to front
"""

import sys
import ctypes
from ctypes import wintypes

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QSpinBox, QGraphicsDropShadowEffect
)
from PySide6.QtCore import Qt, QTimer, QRect, QRectF, QPoint
from PySide6.QtGui import (
    QFont, QPainter, QColor, QPen, QPainterPath, QMouseEvent
)

# ==================== Windows Acrylic Blur ====================

class ACCENT_POLICY(ctypes.Structure):
    _fields_ = [
        ("AccentState", ctypes.c_uint),
        ("AccentFlags", ctypes.c_uint),
        ("GradientColor", ctypes.c_uint),
        ("AnimationId", ctypes.c_uint),
    ]

class WINCOMPATTRDATA(ctypes.Structure):
    _fields_ = [
        ("Attribute", ctypes.c_int),
        ("Data", ctypes.POINTER(ACCENT_POLICY)),
        ("SizeOfData", ctypes.c_size_t),
    ]

SetWindowCompositionAttribute = ctypes.windll.user32.SetWindowCompositionAttribute
SetWindowCompositionAttribute.argtypes = (wintypes.HWND, ctypes.POINTER(WINCOMPATTRDATA))
SetWindowCompositionAttribute.restype = ctypes.c_bool

def enable_acrylic_blur(hwnd):
    """Enable Windows 10/11 acrylic blur behind the window."""
    try:
        accent = ACCENT_POLICY()
        accent.AccentState = 4      # ACCENT_ENABLE_ACRYLICBLURBEHIND
        accent.AccentFlags = 2      # Draw with alpha channel
        accent.GradientColor = 0x01000000  # Very subtle tint
        data = WINCOMPATTRDATA()
        data.Attribute = 19         # WCA_ACCENT_POLICY
        data.Data = ctypes.pointer(accent)
        data.SizeOfData = ctypes.sizeof(accent)
        return SetWindowCompositionAttribute(hwnd, data)
    except Exception:
        return False

# ==================== Modern Styled Button ====================

class ModernButton(QPushButton):
    """Flat semi-transparent button with hover effect."""
    def __init__(self, text="", color="#FFFFFF", bg_hover="rgba(255,255,255,0.15)",
                 bg_pressed="rgba(255,255,255,0.1)", radius=6, padding="6px 16px", font_size="13px"):
        super().__init__(text)
        self.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {color};
                border: 1px solid rgba(255,255,255,0.15);
                border-radius: {radius}px;
                padding: {padding};
                font-size: {font_size};
            }}
            QPushButton:hover {{
                background: {bg_hover};
                border: 1px solid rgba(255,255,255,0.3);
            }}
            QPushButton:pressed {{
                background: {bg_pressed};
            }}
            QPushButton:disabled {{
                color: rgba(255,255,255,0.3);
                border: 1px solid rgba(255,255,255,0.05);
            }}
        """)
        self.setCursor(Qt.PointingHandCursor)


# ==================== Main Timer Application ====================

class TimerApp(QWidget):
    WINDOW_WIDTH = 340
    WINDOW_HEIGHT = 310
    COMPACT_HEIGHT = 42

    BG_COLOR_DARK = QColor(18, 20, 30, 200)
    BG_ALERT = QColor(160, 25, 35, 210)

    def __init__(self):
        super().__init__()
        # State
        self.compact_mode = False
        self.running = False
        self.is_countdown = True
        self.remaining_seconds = 0
        self.elapsed_seconds = 0
        self.normal_geometry = QRect(0, 0, self.WINDOW_WIDTH, self.WINDOW_HEIGHT)
        self.drag_pos = None
        self._blur_set = False
        self._total_countdown = 0
        self._alert_flash = False
        self._pulse_opacity = 0.0
        self._suppress_compact = False
        self._press_pos = QPoint()

        self.setup_window()
        self.setup_ui()
        self.setup_timer()

    # ────────── Window Setup ──────────

    def setup_window(self):
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Window
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(self.WINDOW_WIDTH, self.WINDOW_HEIGHT)
        self.setMouseTracking(True)

    def showEvent(self, event):
        super().showEvent(event)
        if not self._blur_set and self.winId():
            enable_acrylic_blur(int(self.winId()))
            self._blur_set = True

    # ────────── UI Setup ──────────

    def setup_ui(self):
        # ── Top bar ──
        self.top_bar = QHBoxLayout()
        self.top_bar.setContentsMargins(12, 6, 12, 0)

        self.icon_label = QLabel("⏱")
        self.icon_label.setStyleSheet("font-size: 16px; color: rgba(255,255,255,0.7);")

        self.title_label = QLabel("桌面定时器")
        self.title_label.setStyleSheet("font-size: 12px; color: rgba(255,255,255,0.5);")

        self.top_bar.addWidget(self.icon_label)
        self.top_bar.addWidget(self.title_label)
        self.top_bar.addStretch()

        # Compact toggle button
        self.compact_btn = ModernButton("—", padding="4px 10px", font_size="14px")
        self.compact_btn.setFixedSize(28, 24)
        self.compact_btn.clicked.connect(self.toggle_compact)
        self.compact_btn.setToolTip("折叠到顶部")

        # Close button
        self.close_btn = ModernButton("✕", bg_hover="rgba(200,40,40,0.7)", padding="4px 10px", font_size="14px")
        self.close_btn.setFixedSize(28, 24)
        self.close_btn.clicked.connect(self.close_app)
        self.close_btn.setToolTip("关闭")

        self.top_bar.addWidget(self.compact_btn)
        self.top_bar.addWidget(self.close_btn)

        # ── Time display ──
        self.time_label = QLabel("25:00")
        self.time_label.setAlignment(Qt.AlignCenter)
        self.time_label.setStyleSheet("color: white;")
        self.time_label.setGraphicsEffect(QGraphicsDropShadowEffect(
            blurRadius=8, offset=QPoint(0, 2), color=QColor(0, 0, 0, 80)))

        time_font = QFont("Consolas", 54, QFont.Weight.Light)
        time_font.setStyleHint(QFont.Monospace)
        self.time_label.setFont(time_font)
        self.time_label.setFixedHeight(80)

        # ── Mode selector ──
        self.mode_layout = QHBoxLayout()
        self.mode_layout.setContentsMargins(40, 0, 40, 0)
        self.mode_layout.setSpacing(0)

        self.countdown_btn = QPushButton("倒计时")
        self.stopwatch_btn = QPushButton("正计时")
        for btn in (self.countdown_btn, self.stopwatch_btn):
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFixedHeight(30)
            btn.setCheckable(True)

        self.countdown_btn.setChecked(True)
        self.countdown_btn.clicked.connect(lambda: self.set_mode(True))
        self.stopwatch_btn.clicked.connect(lambda: self.set_mode(False))
        self.update_mode_style()

        self.mode_layout.addWidget(self.countdown_btn)
        self.mode_layout.addWidget(self.stopwatch_btn)

        # ── Time input (hours : minutes : seconds) ──
        self.input_layout = QHBoxLayout()
        self.input_layout.setContentsMargins(60, 0, 60, 0)
        self.input_layout.setAlignment(Qt.AlignCenter)

        self.h_spin = QSpinBox()
        self.m_spin = QSpinBox()
        self.s_spin = QSpinBox()

        for spin, maxv, w in [
            (self.h_spin, 99, 55), (self.m_spin, 59, 45), (self.s_spin, 59, 45)
        ]:
            spin.setRange(0, maxv)
            spin.setFixedWidth(w)
            spin.setFixedHeight(32)
            spin.setAlignment(Qt.AlignCenter)
            spin.setButtonSymbols(QSpinBox.NoButtons)
            spin.setStyleSheet(f"""
                QSpinBox {{
                    background: transparent;
                    color: white;
                    border: none;
                    border-bottom: 1px solid rgba(255,255,255,0.2);
                    font-size: 20px;
                    font-weight: 300;
                    padding: 2px 0;
                }}
                QSpinBox:focus {{
                    border-bottom: 1px solid #7C3AED;
                }}
                QSpinBox::up-button, QSpinBox::down-button {{ width: 0px; }}
            """)

        self.h_spin.setValue(0)
        self.m_spin.setValue(25)
        self.s_spin.setValue(0)

        sep_style = "color: rgba(255,255,255,0.3); font-size: 20px; padding: 0 2px;"
        sep1 = QLabel(":")
        sep2 = QLabel(":")
        for s in (sep1, sep2):
            s.setStyleSheet(sep_style)

        self.input_layout.addWidget(self.h_spin)
        self.input_layout.addWidget(sep1)
        self.input_layout.addWidget(self.m_spin)
        self.input_layout.addWidget(sep2)
        self.input_layout.addWidget(self.s_spin)

        # ── Control buttons ──
        self.ctrl_layout = QHBoxLayout()
        self.ctrl_layout.setContentsMargins(40, 0, 40, 0)
        self.ctrl_layout.setSpacing(12)

        self.start_btn = ModernButton("▶  开始", padding="10px 28px", font_size="14px")
        self.start_btn.clicked.connect(self.toggle_start)
        self.start_btn.setToolTip("开始 / 暂停")

        self.reset_btn = ModernButton("⟳  重置", padding="10px 28px", font_size="14px")
        self.reset_btn.clicked.connect(self.reset)
        self.reset_btn.setToolTip("重置")

        self.ctrl_layout.addStretch()
        self.ctrl_layout.addWidget(self.start_btn)
        self.ctrl_layout.addWidget(self.reset_btn)
        self.ctrl_layout.addStretch()

        # ── Status label (hidden by default) ──
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: #F87171; font-size: 16px; font-weight: bold;")
        self.status_label.hide()

        # ── Main layout ──
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(16, 0, 16, 16)
        self.main_layout.setSpacing(2)

        self.main_layout.addLayout(self.top_bar)
        self.main_layout.addSpacing(4)
        self.main_layout.addWidget(self.time_label)
        self.main_layout.addSpacing(4)
        self.main_layout.addLayout(self.mode_layout)
        self.main_layout.addSpacing(6)
        self.main_layout.addLayout(self.input_layout)
        self.main_layout.addSpacing(10)
        self.main_layout.addLayout(self.ctrl_layout)
        self.main_layout.addWidget(self.status_label)

        self.update_display()
        self.update_mode_visibility()

    def setup_timer(self):
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.setInterval(1000)

    # ────────── Mode ──────────

    def set_mode(self, is_countdown):
        if self.running:
            self.stop()
        self.is_countdown = is_countdown
        self.countdown_btn.setChecked(is_countdown)
        self.stopwatch_btn.setChecked(not is_countdown)
        self.update_mode_style()
        self.update_mode_visibility()
        self.reset_state()

    def update_mode_style(self):
        active = """
            QPushButton {
                background: rgba(124, 58, 237, 0.35);
                color: white;
                border: 1px solid rgba(124, 58, 237, 0.6);
                border-radius: 6px;
                font-size: 13px;
            }
        """
        inactive = """
            QPushButton {
                background: transparent;
                color: rgba(255,255,255,0.5);
                border: 1px solid rgba(255,255,255,0.1);
                border-radius: 6px;
                font-size: 13px;
            }
            QPushButton:hover {
                background: rgba(255,255,255,0.08);
                border: 1px solid rgba(255,255,255,0.25);
            }
        """
        self.countdown_btn.setStyleSheet(active if self.countdown_btn.isChecked() else inactive)
        self.stopwatch_btn.setStyleSheet(active if self.stopwatch_btn.isChecked() else inactive)

    def update_mode_visibility(self):
        show_input = self.is_countdown and not self.running
        self.h_spin.setVisible(show_input)
        self.m_spin.setVisible(show_input)
        self.s_spin.setVisible(show_input)
        for i in range(self.input_layout.count()):
            w = self.input_layout.itemAt(i).widget()
            if w and isinstance(w, QLabel):
                w.setVisible(show_input)

    # ────────── Timer Logic ──────────

    def toggle_start(self):
        if self.running:
            self.pause()
        else:
            self.start()

    def start(self):
        if self.is_countdown:
            if self.remaining_seconds <= 0:
                # Read from input
                self.remaining_seconds = (
                    self.h_spin.value() * 3600 +
                    self.m_spin.value() * 60 +
                    self.s_spin.value()
                )
                if self.remaining_seconds <= 0:
                    self.status_label.setText("⚠ 请设置有效时间")
                    self.status_label.show()
                    return
                self._total_countdown = self.remaining_seconds
            self.status_label.hide()
            self.running = True
            self.timer.start()
        else:
            self.status_label.hide()
            self.running = True
            self.timer.start()
        self.update_mode_visibility()
        self.update_controls()
        self.update()

    def pause(self):
        self.running = False
        self.timer.stop()
        self.update_mode_visibility()
        self.update_controls()
        self.update()

    def stop(self):
        self.running = False
        self.timer.stop()
        self.update_controls()

    def tick(self):
        if self.is_countdown:
            if self.remaining_seconds > 0:
                self.remaining_seconds -= 1
                self.update_display()
            if self.remaining_seconds <= 0:
                self.on_timer_finished()
        else:
            self.elapsed_seconds += 1
            self.update_display()

    def on_timer_finished(self):
        self.running = False
        self.timer.stop()
        self.update_display()
        self.update_controls()
        self.update_mode_visibility()

        # Sound
        try:
            import winsound
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS | winsound.SND_ASYNC)
        except Exception:
            try:
                import winsound
                winsound.Beep(880, 200)
                winsound.Beep(660, 200)
                winsound.Beep(880, 400)
            except Exception:
                pass

        # Exit compact mode (requirement 4)
        if self.compact_mode:
            self.exit_compact_mode()

        # Raise to front
        self.raise_()
        self.activateWindow()

        # Flash alert
        self._alert_flash = True
        self._alert_count = 0
        self._alert_timer = QTimer(self)
        self._alert_timer.timeout.connect(self._alert_tick)
        self._alert_timer.start(500)

        # Show message
        self.status_label.setText("⏰ 时间到!")
        self.status_label.show()

        # Auto-stop alert after 4s
        QTimer.singleShot(4000, self._stop_alert)

    def _alert_tick(self):
        self._alert_count += 1
        if self._alert_count > 8:
            self._stop_alert()
            return
        self.update()

    def _stop_alert(self):
        self._alert_flash = False
        if hasattr(self, '_alert_timer') and self._alert_timer:
            self._alert_timer.stop()
            self._alert_timer = None
        self.update()

    def reset(self):
        was_running = self.running
        if self.running:
            self.timer.stop()
        self.running = False
        self._alert_flash = False
        self.status_label.hide()
        self.reset_state()
        self.update_mode_visibility()
        self.update_controls()
        if was_running:
            self.update_display()

    def reset_state(self):
        if self.is_countdown:
            self.remaining_seconds = 0
            self._total_countdown = 0
            self.time_label.setText(
                f"{self.h_spin.value():02d}:{self.m_spin.value():02d}:{self.s_spin.value():02d}")
        else:
            self.elapsed_seconds = 0
            self.time_label.setText("00:00:00")

    def update_display(self):
        if self.is_countdown:
            s = max(0, self.remaining_seconds)
            h = s // 3600
            m = (s % 3600) // 60
            sec = s % 60
            self.time_label.setText(f"{h:02d}:{m:02d}:{sec:02d}")
        else:
            s = self.elapsed_seconds
            h = s // 3600
            m = (s % 3600) // 60
            sec = s % 60
            self.time_label.setText(f"{h:02d}:{m:02d}:{sec:02d}")

    def update_controls(self):
        self.start_btn.setText("⏸  暂停" if self.running else "▶  开始")
        self.start_btn.setToolTip("暂停" if self.running else "开始")

    # ────────── Compact Mode ──────────

    def toggle_compact(self):
        if self.compact_mode:
            self.exit_compact_mode()
        else:
            self.enter_compact_mode()

    def enter_compact_mode(self):
        if self.compact_mode:
            return
        self.compact_mode = True
        self.normal_geometry = self.geometry()

        # Save layout state
        self._status_was_visible = self.status_label.isVisible()
        self._normal_margins = self.main_layout.contentsMargins()

        # Snap to top of screen, shrink height
        x = self.normal_geometry.x()
        self.setFixedHeight(self.COMPACT_HEIGHT)
        self.move(x, 0)

        # Tighten layout margins for compact mode
        self.main_layout.setContentsMargins(8, 0, 8, 0)

        # Hide all widgets except the time label
        self._set_compact_widgets_visible(False)

        # Remove drop shadow in compact mode (avoids clipping)
        self.time_label.setGraphicsEffect(None)

        # Compact time font
        self.time_label.setFont(QFont("Consolas", 26, QFont.Weight.Normal))
        self.time_label.setFixedHeight(self.COMPACT_HEIGHT)
        self.time_label.setAlignment(Qt.AlignCenter)

        # Update compact button text
        self.compact_btn.setText("↗")

    def exit_compact_mode(self):
        if not self.compact_mode:
            return
        self.compact_mode = False

        # Restore layout margins
        if hasattr(self, '_normal_margins'):
            self.main_layout.setContentsMargins(self._normal_margins)

        # Restore geometry
        self.setFixedHeight(self.WINDOW_HEIGHT)
        self.setFixedWidth(self.WINDOW_WIDTH)
        if self.normal_geometry:
            self.setGeometry(self.normal_geometry)

        # Show all widgets
        self._set_compact_widgets_visible(True)

        # Restore status label visibility
        if hasattr(self, '_status_was_visible') and not self._status_was_visible:
            self.status_label.hide()

        # Restore time font
        self.time_label.setFont(QFont("Consolas", 54, QFont.Weight.Light))
        self.time_label.setFixedHeight(80)
        self.time_label.setAlignment(Qt.AlignCenter)
        self.time_label.setGraphicsEffect(QGraphicsDropShadowEffect(
            blurRadius=8, offset=QPoint(0, 2), color=QColor(0, 0, 0, 80)))

        # Update compact button text
        self.compact_btn.setText("—")

        self.update_mode_visibility()

    def _set_compact_widgets_visible(self, visible):
        # Top bar widgets except compact button
        self.icon_label.setVisible(visible)
        self.title_label.setVisible(visible)
        self.close_btn.setVisible(visible)

        # Mode selector
        self.countdown_btn.setVisible(visible)
        self.stopwatch_btn.setVisible(visible)

        # Input
        for i in range(self.input_layout.count()):
            w = self.input_layout.itemAt(i).widget()
            if w:
                w.setVisible(visible)

        # Control buttons
        self.start_btn.setVisible(visible)
        self.reset_btn.setVisible(visible)
        self.status_label.setVisible(visible)

    # ────────── Mouse Events (Drag & Compact) ──────────

    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self.drag_pos = event.globalPosition().toPoint()
            # In compact mode: store press position to distinguish click vs drag
            if self.compact_mode:
                self._press_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event: QMouseEvent):
        if event.buttons() & Qt.LeftButton and self.drag_pos is not None:
            if self.compact_mode:
                # Dragging from compact mode — if moved enough, exit compact and follow mouse
                delta = event.globalPosition().toPoint() - self._press_pos
                if abs(delta.y()) > 15:
                    # Remember x position
                    old_x = self.normal_geometry.x() if self.normal_geometry else self.x()
                    was_visible = getattr(self, '_status_was_visible', False)
                    # Exit compact mode (restore size & widgets)
                    self.compact_mode = False
                    if hasattr(self, '_normal_margins'):
                        self.main_layout.setContentsMargins(self._normal_margins)
                    self.setFixedHeight(self.WINDOW_HEIGHT)
                    self.setFixedWidth(self.WINDOW_WIDTH)
                    self._set_compact_widgets_visible(True)
                    if not was_visible:
                        self.status_label.hide()
                    self.time_label.setFont(QFont("Consolas", 54, QFont.Weight.Light))
                    self.time_label.setFixedHeight(80)
                    self.time_label.setAlignment(Qt.AlignCenter)
                    self.time_label.setGraphicsEffect(QGraphicsDropShadowEffect(
                        blurRadius=8, offset=QPoint(0, 2), color=QColor(0, 0, 0, 80)))
                    self.compact_btn.setText("—")
                    self.update_mode_visibility()
                    # Move window so cursor grabs the title-bar area
                    cursor_y = int(event.globalPosition().y())
                    self.move(old_x, max(0, cursor_y - 40))
                    # Update normal_geometry to new position so click-restore uses correct position
                    self.normal_geometry = QRect(self.x(), self.y(), self.WINDOW_WIDTH, self.WINDOW_HEIGHT)
                    self.drag_pos = event.globalPosition().toPoint()
                    self._suppress_compact = True
                    QTimer.singleShot(400, lambda: setattr(self, '_suppress_compact', False))
                return

            # Normal drag (non-compact mode)
            delta = event.globalPosition().toPoint() - self.drag_pos
            new_pos = self.pos() + delta
            self.move(new_pos)
            self.drag_pos = event.globalPosition().toPoint()

    def mouseReleaseEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            if self.compact_mode:
                # Click on compact bar without dragging → restore at original position
                self.exit_compact_mode()
                self.drag_pos = None
                return
            self.drag_pos = None
            # Snap to compact mode if dragged to top of screen
            if not self._suppress_compact and self.pos().y() <= 5:
                self.enter_compact_mode()

    def mouseDoubleClickEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self.toggle_compact()

    # ────────── Painting (Frosted Glass Background) ──────────

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Rounded rect path
        path = QPainterPath()
        r = self.rect()
        if self.compact_mode:
            # In compact mode, only round bottom corners
            corner = 8
            path.moveTo(r.left(), r.top())
            path.lineTo(r.right(), r.top())
            path.lineTo(r.right(), r.bottom() - corner)
            path.quadTo(r.right(), r.bottom(), r.right() - corner, r.bottom())
            path.lineTo(r.left() + corner, r.bottom())
            path.quadTo(r.left(), r.bottom(), r.left(), r.bottom() - corner)
            path.closeSubpath()
        else:
            path.addRoundedRect(QRectF(r), 12, 12)

        # Background color
        if self._alert_flash:
            flash_on = (self._alert_count % 2 == 0)
            bg = self.BG_ALERT if flash_on else self.BG_COLOR_DARK
        else:
            bg = self.BG_COLOR_DARK

        painter.fillPath(path, bg)

        # Subtle border
        painter.setPen(QPen(QColor(255, 255, 255, 20), 1))
        painter.drawPath(path)

    # ────────── Close ──────────

    def close_app(self):
        self.close()


# ==================== Main Entry ====================

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setApplicationName("桌面定时器")

    # Set dark palette for controls
    palette = app.palette()
    palette.setColor(palette.ColorRole.WindowText, Qt.white)
    palette.setColor(palette.ColorRole.Text, Qt.white)
    app.setPalette(palette)

    window = TimerApp()
    # Center on screen
    screen = app.primaryScreen().availableGeometry()
    window.move(
        (screen.width() - window.width()) // 2,
        (screen.height() - window.height()) // 2
    )
    window.show()

    sys.exit(app.exec())
