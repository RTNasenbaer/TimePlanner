import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QColorDialog, QInputDialog, QLabel, QFrame,
    QMenuBar, QAction, QFileDialog, QMessageBox, QDialog, QFormLayout, QLineEdit, QSpinBox, QTextEdit, QComboBox
)
from PyQt5.QtGui import QPainter, QColor, QPen, QFont, QLinearGradient
from PyQt5.QtCore import Qt, QRectF

import pandas as pd
import json
import os
from datetime import datetime

import random

try:
    from docx import Document
    from docx.shared import Inches
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


# Hilfsfunktion für harmonische Farben
def get_nice_color(idx):
    # Golden angle for color harmony
    golden_angle = 137.508
    hue = (idx * golden_angle) % 360
    import colorsys
    rgb = colorsys.hsv_to_rgb(hue/360, 0.5, 0.95)
    return QColor(int(rgb[0]*255), int(rgb[1]*255), int(rgb[2]*255))

class AppSettings:
    def __init__(self):
        self.user_name = "Trainer"
        self.player_number = 20
        self.requirements = "Grundfähigkeiten Ballkontrolle"
        self.team = "Jugendmannschaft"
        self.settings_file = "settings.json"
        self.load_settings()
    
    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.user_name = data.get('user_name', self.user_name)
                    self.player_number = data.get('player_number', self.player_number)
                    self.requirements = data.get('requirements', self.requirements)
                    self.team = data.get('team', self.team)
            except Exception:
                pass  # Use defaults if loading fails
    
    def save_settings(self):
        try:
            data = {
                'user_name': self.user_name,
                'player_number': self.player_number,
                'requirements': self.requirements,
                'team': self.team
            }
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass  # Silently fail if saving doesn't work


class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Einstellungen")
        self.setModal(True)
        self.setMinimumWidth(400)
        
        layout = QFormLayout()
        
        # User name
        self.name_edit = QLineEdit(settings.user_name)
        layout.addRow("Trainer Name:", self.name_edit)
        
        # Player number
        self.player_spin = QSpinBox()
        self.player_spin.setRange(1, 100)
        self.player_spin.setValue(settings.player_number)
        layout.addRow("Anzahl Spieler:", self.player_spin)
        
        # Requirements
        self.requirements_edit = QTextEdit(settings.requirements)
        self.requirements_edit.setMaximumHeight(80)
        layout.addRow("Anforderungen:", self.requirements_edit)
        
        # Team info
        self.team_edit = QTextEdit(settings.team)
        self.team_edit.setMaximumHeight(80)
        layout.addRow("Team Info:", self.team_edit)
        
        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Abbrechen")
        
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        
        layout.addRow(button_layout)
        self.setLayout(layout)
    
    def accept(self):
        # Save the values back to settings
        self.settings.user_name = self.name_edit.text().strip()
        self.settings.player_number = self.player_spin.value()
        self.settings.requirements = self.requirements_edit.toPlainText().strip()
        self.settings.team = self.team_edit.toPlainText().strip()
        self.settings.save_settings()
        super().accept()


class AddSectionDialog(QDialog):
    def __init__(self, max_duration, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Neuen Abschnitt hinzufügen")
        self.setModal(True)
        self.setMinimumWidth(400)
        
        layout = QFormLayout()
        
        # Name
        self.name_edit = QLineEdit()
        layout.addRow("Name des Abschnitts:", self.name_edit)
        
        # Duration
        self.duration_spin = QSpinBox()
        self.duration_spin.setRange(1, max_duration)
        self.duration_spin.setValue(10)
        self.duration_spin.setSuffix(" min")
        layout.addRow(f"Dauer (max {max_duration} min):", self.duration_spin)
        
        # Organisation
        self.organisation_combo = QComboBox()
        self.organisation_combo.addItems(["Übungsform", "Spielform"])
        layout.addRow("Organisationsform:", self.organisation_combo)
        
        # Explanation
        self.explanation_edit = QTextEdit()
        self.explanation_edit.setMaximumHeight(100)
        self.explanation_edit.setPlaceholderText("Beschreibung der Übung...")
        layout.addRow("Erklärung:", self.explanation_edit)
        
        # Tools
        self.tools_edit = QLineEdit()
        self.tools_edit.setPlaceholderText("z.B. Bälle, Hütchen, Tore...")
        layout.addRow("Hilfsmittel:", self.tools_edit)
        
        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Abbrechen")
        
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        
        layout.addRow(button_layout)
        self.setLayout(layout)
    
    def get_values(self):
        return {
            'name': self.name_edit.text().strip(),
            'duration': self.duration_spin.value(),
            'organisation': self.organisation_combo.currentText(),
            'explanation': self.explanation_edit.toPlainText().strip(),
            'tools': self.tools_edit.text().strip()
        }


class TimeSection:
    def __init__(self, start, duration, name, color, organisation="Übungsform", explanation="", tools=""):
        self.start = start  # in minutes
        self.duration = duration  # in minutes
        self.name = name
        self.color = color
        self.organisation = organisation  # "Spielform" or "Übungsform"
        self.explanation = explanation  # Detailed explanation of the exercise
        self.tools = tools  # List of tools needed


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
        # Initialize settings
        self.settings = AppSettings()
        
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
        export_docx_action = QAction("Als DOCX exportieren", self)
        export_docx_action.setShortcut("Ctrl+D")
        export_docx_action.triggered.connect(self.export_docx)
        file_menu.addAction(export_docx_action)
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
        
        # Settings menu
        settings_menu = self.menubar.addMenu("&Einstellungen")
        settings_action = QAction("Einstellungen...", self)
        settings_action.setShortcut("Ctrl+,")
        settings_action.triggered.connect(self.show_settings)
        settings_menu.addAction(settings_action)
        
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
    
    def show_settings(self):
        dialog = SettingsDialog(self.settings, self)
        dialog.exec_()
    
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
                organisation = str(row.get('Organisation', 'Übungsform'))
                explanation = str(row.get('Erklärung', ''))
                tools = str(row.get('Hilfsmittel', ''))
                sections.append(TimeSection(cur_start, duration, name, color, organisation, explanation, tools))
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
                'Farbe': color_hex,
                'Organisation': getattr(s, 'organisation', 'Übungsform'),
                'Erklärung': getattr(s, 'explanation', ''),
                'Hilfsmittel': getattr(s, 'tools', '')
            })
        df = pd.DataFrame(data)
        try:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Export erfolgreich", f"Excel gespeichert: {path}")
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Excel konnte nicht gespeichert werden.\n{e}")

    def export_docx(self):
        try:
            from docx import Document
            from datetime import datetime
        except ImportError:
            QMessageBox.warning(self, "Fehler", "python-docx ist nicht installiert. Bitte installieren Sie es mit 'pip install python-docx'")
            return
            
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        
        # Use the specific template file
        template_path = "Trainingseinheit_Name_StandardFile_Date.docx"
        if not os.path.exists(template_path):
            QMessageBox.warning(self, "Fehler", f"Template-Datei '{template_path}' nicht gefunden!")
            return
            
        # Ask for output file
        output_path, _ = QFileDialog.getSaveFileName(
            self, 
            "DOCX exportieren", 
            f"Trainingseinheit_{self.plan_name}_StandardFile_{datetime.now().strftime('%Y%m%d')}.docx", 
            "Word Documents (*.docx)"
        )
        if not output_path:
            return
            
        try:
            self._fill_docx_template(template_path, output_path)
            QMessageBox.information(self, "Export erfolgreich", f"DOCX gespeichert: {output_path}")
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"DOCX konnte nicht erstellt werden.\n{str(e)}")
    
    def _fill_docx_template(self, template_path, output_path):
        from docx import Document
        from datetime import datetime
        
        doc = Document(template_path)
        
        # Generate smart tools list from all sections
        tools_combined = self._combine_tools_intelligently()
        if not tools_combined:
            tools_combined = "Keine besonderen Werkzeuge"
        
        # Create replacements based on actual template placeholders
        replacements = {
            '{{Theme}}': self.plan_name,
            '{{Name}}': self.settings.user_name,
            '{{playNumber}}': str(self.settings.player_number),
            '{{tools}}': tools_combined,
            '{{Time}}': f"{self.total_minutes} Minuten",
            '{{empty}}': ""  # Clear empty placeholders
        }
        
        # Also need to handle the missing placeholders from the template analysis
        additional_replacements = {
            # The template seems to have some fields that need to be filled manually
            # We'll handle them in the table replacement logic
        }
        
        # Replace in all table cells (the template uses tables, not paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for placeholder, value in replacements.items():
                            if placeholder in paragraph.text:
                                paragraph.text = paragraph.text.replace(placeholder, value)
        
        # Handle the sections table (Table 2)
        if len(doc.tables) >= 2:
            sections_table = doc.tables[1]  # Second table for sections
            
            # Remove the existing template row (row 1, keep header row 0)
            if len(sections_table.rows) > 1:
                # Remove template row
                sections_table._element.remove(sections_table.rows[1]._element)
            
            # Add new rows for each section
            for i, section in enumerate(self.bar.sections):
                row_cells = sections_table.add_row().cells
                if len(row_cells) >= 5:  # Ensure we have enough columns
                    # Format duration instead of time slot
                    duration_text = f"{section.duration} min"
                    
                    # Fill the cells
                    row_cells[0].text = duration_text              # Zeit (Duration)
                    row_cells[1].text = section.name               # Absicht/Ziel
                    row_cells[2].text = section.organisation       # Organisation
                    row_cells[3].text = section.explanation        # Übungs-/Spielform
                    row_cells[4].text = section.tools             # Hilfsmittel
        
        doc.save(output_path)

    def _combine_tools_intelligently(self):
        """
        Intelligently combine tools from all sections, keeping the highest quantity for each tool type.
        Example: "3 Balls", "6 Balls", "12 Balls" -> "12 Balls"
        """
        import re
        
        tool_quantities = {}  # tool_name -> max_quantity
        
        for section in self.bar.sections:
            if not section.tools:
                continue
                
            # Split tools by comma and process each
            tools_list = [tool.strip() for tool in section.tools.split(',')]
            
            for tool in tools_list:
                if not tool:
                    continue
                    
                # Try to extract number and tool name using regex
                # Matches patterns like "3 Balls", "12 Hütchen", "1 Tor", etc.
                match = re.match(r'^(\d+)\s+(.+)$', tool.strip())
                
                if match:
                    quantity = int(match.group(1))
                    tool_name = match.group(2).strip()
                    
                    # Keep the highest quantity for this tool
                    if tool_name in tool_quantities:
                        tool_quantities[tool_name] = max(tool_quantities[tool_name], quantity)
                    else:
                        tool_quantities[tool_name] = quantity
                else:
                    # No quantity found, treat as single item
                    tool_name = tool.strip()
                    if tool_name in tool_quantities:
                        # If we already have a quantity for this tool, keep it
                        pass
                    else:
                        # Add as single item (quantity 1, but don't show the number)
                        tool_quantities[tool_name] = 0  # 0 means no quantity shown
        
        # Build the final tools list
        final_tools = []
        for tool_name, quantity in sorted(tool_quantities.items()):
            if quantity > 0:
                final_tools.append(f"{quantity} {tool_name}")
            else:
                final_tools.append(tool_name)
        
        return ', '.join(final_tools)

    def add_section(self):
        # Ermittle das Ende des letzten Segments
        if self.bar.sections:
            last = self.bar.sections[-1]
            start = last.start + last.duration
        else:
            start = 0
        max_duration = self.bar.total_minutes - start
        if max_duration <= 0:
            QMessageBox.warning(self, "Fehler", "Keine Zeit mehr verfügbar für weitere Abschnitte.")
            return
        
        dialog = AddSectionDialog(max_duration, self)
        if dialog.exec_() == QDialog.Accepted:
            values = dialog.get_values()
            if not values['name']:
                QMessageBox.warning(self, "Fehler", "Name des Abschnitts darf nicht leer sein.")
                return
            
            color = get_nice_color(len(self.bar.sections))
            section = TimeSection(
                start=start,
                duration=values['duration'],
                name=values['name'],
                color=color,
                organisation=values['organisation'],
                explanation=values['explanation'],
                tools=values['tools']
            )
            self.bar.sections.append(section)
            self.bar.update()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TimePlannerApp()
    window.resize(1500, 500)
    window.show()
    sys.exit(app.exec_())
