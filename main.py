import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QColorDialog, QInputDialog, QLabel, QFrame,
    QMenuBar, QAction, QFileDialog, QMessageBox
)
from PyQt5.QtGui import QPainter, QColor, QPen, QFont, QLinearGradient
from PyQt5.QtCore import Qt, QRectF

import pandas as pd

import random


# Hilfsfunktion für harmonische Farben
def get_nice_color(idx):
    # Golden angle for color harmony
    golden_angle = 137.508
    hue = (idx * golden_angle) % 360
    import colorsys
    rgb = colorsys.hsv_to_rgb(hue/360, 0.5, 0.95)
    return QColor(int(rgb[0]*255), int(rgb[1]*255), int(rgb[2]*255))

class TimeSection:
    def __init__(self, start, duration, name, color):
        self.start = start  # in minutes
        self.duration = duration  # in minutes
        self.name = name
        self.color = color


class BarWidget(QWidget):
    def __init__(self, total_minutes=120, sections=None):
        super().__init__()
        self.total_minutes = total_minutes
        self.sections = sections if sections else []
        self.setMinimumWidth(1300)
        self.setMinimumHeight(340)
        self._drag_idx = None
        self._drag_offset = 0
        self._drag_x = None
        self._drag_y = None
        self._drag_section = None
        self.setMouseTracking(True)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        width = self.width()
        height = self.height()
        margin = 80
        bar_height = 120
        bar_y = height//2 - bar_height//2
        # Hintergrund
        painter.setBrush(QColor(250,252,255))
        painter.setPen(Qt.NoPen)
        painter.drawRect(0, 0, width, height)
        # Bar
        painter.setBrush(QColor(230,235,245))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(margin, bar_y, width-2*margin, bar_height, 16, 16)
        # Segmente (außer Dragged)
        for idx, section in enumerate(self.sections):
            if self._drag_idx == idx and self._drag_section is not None:
                continue
            x_start = margin + (section.start/self.total_minutes)*(width-2*margin)
            x_end = margin + ((section.start+section.duration)/self.total_minutes)*(width-2*margin)
            rect = QRectF(x_start, bar_y, x_end-x_start, bar_height)
            # Segment
            painter.setBrush(section.color)
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(rect, 12, 12)
            # Gestrichelte Linien am Start und Ende
            dash_pen = QPen(QColor(120,120,120,120), 1, Qt.DashLine)
            painter.setPen(dash_pen)
            painter.drawLine(int(x_start), int(bar_y-8), int(x_start), int(bar_y+bar_height+18))
            painter.drawLine(int(x_end), int(bar_y-8), int(x_end), int(bar_y+bar_height+18))
            # Text (Name + Dauer) UNTER dem Segment
            painter.setPen(QPen(QColor(30,30,30)))
            font = QFont()
            font.setPointSize(8)
            font.setBold(False)
            painter.setFont(font)
            text = f"{section.name} ({section.duration} min)"
            text_rect = QRectF(x_start, bar_y+bar_height+6, x_end-x_start, 18)
            painter.drawText(text_rect, Qt.AlignCenter, text)
        # Dragged Segment als Ghost am Cursor (folgt X und Y)
        if self._drag_idx is not None and self._drag_section is not None and self._drag_x is not None and self._drag_y is not None:
            drag = self._drag_section
            drag_width = (drag.duration/self.total_minutes)*(width-2*margin)
            rect = QRectF(self._drag_x - self._drag_offset, self._drag_y - bar_height//2, drag_width, bar_height)
            painter.setBrush(drag.color.lighter(120))
            painter.setPen(QPen(QColor(120,120,120,80), 2, Qt.DashLine))
            painter.drawRoundedRect(rect, 12, 12)
            # Gestrichelte Linien für Ghost
            dash_pen = QPen(QColor(120,120,120,120), 1, Qt.DashLine)
            painter.setPen(dash_pen)
            painter.drawLine(int(rect.left()), int(rect.top()-8), int(rect.left()), int(rect.bottom()+18))
            painter.drawLine(int(rect.right()), int(rect.top()-8), int(rect.right()), int(rect.bottom()+18))
            # Text (Name + Dauer) UNTER dem Ghost
            painter.setPen(QPen(QColor(30,30,30)))
            font = QFont()
            font.setPointSize(8)
            font.setBold(False)
            painter.setFont(font)
            text = f"{drag.name} ({drag.duration} min)"
            text_rect = QRectF(rect.left(), rect.bottom()+6, rect.width(), 18)
            painter.drawText(text_rect, Qt.AlignCenter, text)

    def mousePressEvent(self, event):
        if event.button() != Qt.LeftButton:
            return
        width = self.width()
        margin = 80  # match drawing
        bar_height = 120  # match drawing
        bar_width = width-2*margin
        x = event.x()
        y = event.y()
        bar_y = self.height()//2 - bar_height//2
        for idx, section in enumerate(self.sections):
            x_start = margin + (section.start/self.total_minutes)*bar_width
            x_end = margin + ((section.start+section.duration)/self.total_minutes)*bar_width
            rect = QRectF(x_start, bar_y, x_end-x_start, bar_height)
            if rect.contains(x, y):
                self._drag_idx = idx
                self._drag_offset = x - x_start
                self._drag_x = x
                self._drag_y = y
                self._drag_section = section
                self.setCursor(Qt.ClosedHandCursor)
                self.update()
                break

    def mouseMoveEvent(self, event):
        if self._drag_idx is not None and self._drag_section is not None:
            width = self.width()
            margin = 80
            bar_height = 120
            bar_width = width-2*margin
            x = event.x()
            y = event.y()
            self._drag_x = x
            self._drag_y = y
            self.setCursor(Qt.ClosedHandCursor)
            self.update()
        else:
            # Hover-Effekt
            width = self.width()
            margin = 80
            bar_height = 120
            bar_width = width-2*margin
            x = event.x()
            y = event.y()
            bar_y = self.height()//2 - bar_height//2
            hovered = None
            for idx, section in enumerate(self.sections):
                x_start = margin + (section.start/self.total_minutes)*bar_width
                x_end = margin + ((section.start+section.duration)/self.total_minutes)*bar_width
                rect = QRectF(x_start, bar_y, x_end-x_start, bar_height)
                if rect.contains(x, y):
                    hovered = idx
                    break
            if hovered is not None:
                self.setCursor(Qt.OpenHandCursor)
            else:
                self.setCursor(Qt.ArrowCursor)

    def mouseReleaseEvent(self, event):
        if self._drag_idx is not None and self._drag_section is not None:
            width = self.width()
            margin = 80
            bar_height = 120
            bar_width = width-2*margin
            x = event.x()
            # Bestimme neue Position im Array anhand Cursor-X
            rel_x = x - margin
            total = 0
            new_pos = 0
            for i, s in enumerate(self.sections):
                if i == self._drag_idx:
                    continue
                seg_len = (s.duration/self.total_minutes)*bar_width
                if rel_x < total + seg_len/2:
                    new_pos = i
                    break
                total += seg_len
            else:
                new_pos = len(self.sections)-1
            if new_pos != self._drag_idx:
                seg = self.sections.pop(self._drag_idx)
                self.sections.insert(new_pos, seg)
            # Startzeiten neu berechnen
            cur = 0
            for s in self.sections:
                s.start = cur
                cur += s.duration
            self._drag_idx = None
            self._drag_section = None
            self._drag_x = None
            self._drag_y = None
            self.setCursor(Qt.ArrowCursor)
            self.update()


class TimePlannerApp(QWidget):
    def __init__(self):
        super().__init__()
        name, ok1 = QInputDialog.getText(self, "Projektname", "Name des Zeitplans:", text="Mein Zeitplan")
        if not ok1 or not name.strip():
            name = "Mein Zeitplan"
        total, ok2 = QInputDialog.getInt(self, "Gesamtdauer", "Gesamtdauer (in Minuten):", 120, 10, 1440)
        if not ok2:
            total = 120
        self.plan_name = name.strip()
        self.total_minutes = total
        self.setWindowTitle(f"Zeitplaner - {self.plan_name}")
        self.setStyleSheet("""
            QWidget {
                background: #f4f7fa;
            }
            QLabel#titleLabel {
                font-size: 22px;
                font-weight: bold;
                color: #2a3a5a;
                margin-bottom: 10px;
            }
            QLabel#infoLabel {
                font-size: 14px;
                color: #444;
                margin-bottom: 10px;
            }
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6a8cff, stop:1 #3a4a7a);
                color: white;
                border-radius: 8px;
                font-size: 14px;
                padding: 8px 18px;
                margin: 8px 0;
            }
            QPushButton:hover {
                background: #4a6cff;
            }
        """)
        self.bar = BarWidget(total_minutes=self.total_minutes)
        self.title = QLabel(f"{self.plan_name}")
        self.title.setObjectName("titleLabel")
        self.info = QLabel(f"Gesamtdauer: {self.total_minutes//60}h {self.total_minutes%60}min ({self.total_minutes} Minuten)")
        self.info.setObjectName("infoLabel")
        self.add_btn = QPushButton("Abschnitt hinzufügen")
        self.add_btn.clicked.connect(self.add_section)
        self.scale_factor = 1.0
        # Reset Button
        self.reset_btn = QPushButton("Zurücksetzen")
        self.reset_btn.clicked.connect(self.reset_plan)
        # Skalierungs-Buttons
        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setToolTip("Größer (Strg +)")
        self.zoom_in_btn.clicked.connect(self.zoom_in)
        self.zoom_out_btn = QPushButton("–")
        self.zoom_out_btn.setToolTip("Kleiner (Strg -)")
        self.zoom_out_btn.clicked.connect(self.zoom_out)
        # Menüleiste
        self.menubar = QMenuBar(self)
        file_menu = self.menubar.addMenu("&Datei")
        import_excel_action = QAction("Excel importieren", self)
        import_excel_action.setShortcut("Ctrl+I")
        import_excel_action.triggered.connect(self.import_excel)
        file_menu.addAction(import_excel_action)
        export_excel_action = QAction("Als Excel exportieren", self)
        export_excel_action.setShortcut("Ctrl+E")
        export_excel_action.triggered.connect(self.export_excel)
        file_menu.addAction(export_excel_action)
        export_png_action = QAction("Als PNG exportieren", self)
        export_png_action.setShortcut("Ctrl+P")
        export_png_action.triggered.connect(self.export_png)
        file_menu.addAction(export_png_action)
        file_menu.addSeparator()
        reset_action = QAction("Zurücksetzen", self)
        reset_action.setShortcut("Ctrl+R")
        reset_action.triggered.connect(self.reset_plan)
        file_menu.addAction(reset_action)
        file_menu.addSeparator()
        scaler_action = QAction("Fenstergröße anpassen", self)
        scaler_action.setShortcut("F11")
        scaler_action.triggered.connect(self.toggle_window_scaler)
        file_menu.addAction(scaler_action)
        # Skalierung
        zoom_in_action = QAction("Größer", self)
        zoom_in_action.setShortcut("Ctrl++")
        zoom_in_action.triggered.connect(self.zoom_in)
        file_menu.addAction(zoom_in_action)
        zoom_out_action = QAction("Kleiner", self)
        zoom_out_action.setShortcut("Ctrl+-")
        zoom_out_action.triggered.connect(self.zoom_out)
        file_menu.addAction(zoom_out_action)
        # Toolbar row for buttons (nur Abschnitt hinzufügen)
        toolbar_row = QHBoxLayout()
        toolbar_row.addWidget(self.add_btn)
        toolbar_row.addStretch(1)
        layout = QVBoxLayout()
        layout.setContentsMargins(30, 20, 30, 20)
        layout.setSpacing(10)
        layout.setMenuBar(self.menubar)
        layout.addWidget(self.title, alignment=Qt.AlignCenter)
        layout.addWidget(self.info, alignment=Qt.AlignCenter)
        layout.addWidget(self.bar, alignment=Qt.AlignCenter)
        layout.addLayout(toolbar_row)
        self.setLayout(layout)
    # Entfernt: alles außerhalb von __init__ (Toolbar und Layout)
    def reset_plan(self):
        from PyQt5.QtWidgets import QMessageBox
        reply = QMessageBox.question(self, "Zurücksetzen", "Alle Abschnitte wirklich löschen?", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.bar.sections.clear()
            self.bar.update()

    def import_excel(self):
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        import pandas as pd
        path, _ = QFileDialog.getOpenFileName(self, "Excel importieren", "", "Excel Files (*.xlsx)")
        if not path:
            return
        try:
            df = pd.read_excel(path)
            sections = []
            cur_start = 0
            for _, row in df.iterrows():
                name = str(row.get('Abschnitt', ''))
                duration = int(row.get('Dauer (min)', 0))
                color_hex = row.get('Farbe', '#cccccc')
                color = QColor(color_hex)
                sections.append(TimeSection(cur_start, duration, name, color))
                cur_start += duration
            self.bar.sections = sections
            self.bar.update()
            QMessageBox.information(self, "Import erfolgreich", f"Excel importiert: {path}")
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Excel konnte nicht importiert werden.\n{e}")

    def zoom_in(self):
        self.scale_factor = min(self.scale_factor + 0.1, 2.5)
        self.apply_scale()

    def zoom_out(self):
        self.scale_factor = max(self.scale_factor - 0.1, 0.5)
        self.apply_scale()

    def apply_scale(self):
        # BarWidget skalieren
        w = int(1300 * self.scale_factor)
        h = int(340 * self.scale_factor)
        self.bar.setMinimumWidth(w)
        self.bar.setMinimumHeight(h)
        self.bar.resize(w, h)
        font = self.title.font()
        font.setPointSize(int(22 * self.scale_factor))
        self.title.setFont(font)
        font2 = self.info.font()
        font2.setPointSize(int(14 * self.scale_factor))
        self.info.setFont(font2)
        self.bar.update()

    def toggle_window_scaler(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def export_png(self):
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        path, _ = QFileDialog.getSaveFileName(self, "PNG exportieren", f"{self.plan_name}.png", "PNG Files (*.png)")
        if not path:
            return
        pixmap = self.bar.grab()
        if pixmap.save(path, "PNG"):
            QMessageBox.information(self, "Export erfolgreich", f"PNG gespeichert: {path}")
        else:
            QMessageBox.warning(self, "Fehler", "PNG konnte nicht gespeichert werden.")

    def export_excel(self):
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        import pandas as pd
        path, _ = QFileDialog.getSaveFileName(self, "Excel exportieren", f"{self.plan_name}.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return
        data = []
        for s in self.bar.sections:
            color_hex = '#{:02x}{:02x}{:02x}'.format(s.color.red(), s.color.green(), s.color.blue())
            data.append({
                'Abschnitt': s.name,
                'Start (min)': s.start,
                'Dauer (min)': s.duration,
                'Farbe': color_hex
            })
        df = pd.DataFrame(data)
        try:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Export erfolgreich", f"Excel gespeichert: {path}")
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Excel konnte nicht gespeichert werden.\n{e}")

    def add_section(self):
        # Ermittle das Ende des letzten Segments
        if self.bar.sections:
            last = self.bar.sections[-1]
            start = last.start + last.duration
        else:
            start = 0
        max_duration = self.bar.total_minutes - start
        if max_duration <= 0:
            return
        duration, ok2 = QInputDialog.getInt(self, "Dauer", f"Dauer (in Minuten, max {max_duration}):", 10, 1, max_duration)
        if not ok2:
            return
        name, ok3 = QInputDialog.getText(self, "Name", "Name des Abschnitts:")
        if not ok3 or not name.strip():
            return
        color = get_nice_color(len(self.bar.sections))
        self.bar.sections.append(TimeSection(start, duration, name.strip(), color))
        self.bar.update()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TimePlannerApp()
    window.resize(1500, 500)
    window.show()
    sys.exit(app.exec_())
