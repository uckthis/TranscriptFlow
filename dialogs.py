from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QFormLayout, QLineEdit, 
                             QComboBox, QDialogButtonBox, QLabel, QTabWidget, 
                             QWidget, QPushButton, QHBoxLayout, QListWidget, 
                             QListWidgetItem, QInputDialog, QMessageBox, QProgressDialog,
                             QColorDialog, QTextEdit, QTextBrowser, QDoubleSpinBox, QSpinBox, 
                             QCheckBox, QFileDialog, QFrame, QFontComboBox,
                             QGridLayout)
from PyQt6.QtCore import Qt, QSize, QTimer
from PyQt6.QtGui import QKeySequence, QKeyEvent, QColor, QFont, QFontDatabase, QTextCursor, QTextDocument
import subprocess
import re
import os
import shutil

# Configure enchant to use AppData dicts folder BEFORE importing enchant
from path_manager import get_dicts_dir
_dicts_path = get_dicts_dir()
os.environ["DICPATH"] = _dicts_path

import enchant
import webbrowser
import urllib.request
import threading
from PyQt6.QtGui import QMouseEvent, QPalette

# Premium Card Helper for Dialog Sections
class PremiumCard(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet("""
            PremiumCard {
                background-color: rgba(255, 255, 255, 0.05);
                border: 1px solid rgba(255, 255, 255, 0.1);
                border-radius: 15px;
                padding: 15px;
            }
            PremiumCard:hover {
                background-color: rgba(255, 255, 255, 0.08);
                border: 1px solid rgba(255, 255, 255, 0.2);
            }
        """)

# --- Original Media Dialog ---
class MediaSourceDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Media Source")
        self.resize(500, 250)
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        
        # --- File Tab ---
        self.tab_file = QWidget()
        file_layout = QVBoxLayout(self.tab_file)
        
        path_layout = QHBoxLayout()
        path_layout.addWidget(QLabel("File Path:"))
        self.inp_file = QLineEdit()
        path_layout.addWidget(self.inp_file)
        
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(self.browse_file)
        path_layout.addWidget(btn_browse)
        
        file_layout.addLayout(path_layout)
        file_layout.addStretch()
        
        # --- URL Tab ---
        self.tab_url = QWidget()
        form_url = QFormLayout()
        self.inp_url = QLineEdit()
        form_url.addRow("Media URL:", self.inp_url)
        self.tab_url.setLayout(form_url)
        
        # --- Offline Tab ---
        self.tab_offline = QWidget()
        form_off = QFormLayout()
        self.combo_mode = QComboBox()
        self.combo_mode.addItems(["Stopwatch Timer", "Real-time Clock"])
        form_off.addRow("Mode:", self.combo_mode)
        self.tab_offline.setLayout(form_off)
        
        self.tabs.addTab(self.tab_file, "File")
        self.tabs.addTab(self.tab_url, "URL")
        self.tabs.addTab(self.tab_offline, "Offline")
        layout.addWidget(self.tabs)
        
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def browse_file(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "Select Media File", "", 
            "Media Files (*.mp4 *.mp3 *.mov *.avi *.mkv *.wav);;All Files (*.*)"
        )
        if f:
            self.inp_file.setText(f)

    def get_data(self):
        idx = self.tabs.currentIndex()
        if idx == 0: return 'file', self.inp_file.text()
        if idx == 1: return 'url', self.inp_url.text()
        if idx == 2: return 'offline', self.combo_mode.currentText()
        return None, None

# --- Key Capture Dialog ---
class KeyCaptureDialog(QDialog):
    def __init__(self, parent=None, existing_triggers=None):
        super().__init__(parent)
        self.setWindowTitle("Define Trigger")
        self.setModal(True)
        self.resize(400, 200)
        self.captured_key = None
        self.existing_triggers = existing_triggers or []
        
        layout = QVBoxLayout(self)
        
        info = QLabel("Press the key combination you want to use:")
        info.setStyleSheet("font-size: 12px; padding: 5px;")
        layout.addWidget(info)
        
        self.display = QLabel("Waiting for key press...")
        self.display.setStyleSheet("""
            font-size: 16px; font-weight: bold; padding: 20px;
            background: #f0f0f0; border: 2px solid #ccc;
            border-radius: 5px; min-height: 60px;
        """)
        self.display.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.display)
        
        warning_label = QLabel("Press ESC to clear, Enter to accept")
        warning_label.setStyleSheet("color: #666; font-size: 10px; padding: 5px;")
        warning_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(warning_label)
        
        btn_layout = QHBoxLayout()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
        
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.clear_trigger)
        btn_layout.addWidget(clear_btn)
        
        self.ok_btn = QPushButton("OK")
        self.ok_btn.clicked.connect(self.accept)
        self.ok_btn.setEnabled(False)
        btn_layout.addWidget(self.ok_btn)
        
        layout.addLayout(btn_layout)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        
    def keyPressEvent(self, event: QKeyEvent):
        key = event.key()
        
        if key == Qt.Key.Key_Escape:
            self.clear_trigger()
            return
            
        if key == Qt.Key.Key_Return or key == Qt.Key.Key_Enter:
            if self.ok_btn.isEnabled():
                self.accept()
            return
        
        modifiers = event.modifiers()
        
        # Ignore modifier-only presses
        if key in [Qt.Key.Key_Control, Qt.Key.Key_Alt, Qt.Key.Key_Shift, 
                   Qt.Key.Key_Meta, Qt.Key.Key_AltGr]:
            return
        
        parts = []
        if modifiers & Qt.KeyboardModifier.ControlModifier:
            parts.append("Ctrl")
        if modifiers & Qt.KeyboardModifier.AltModifier:
            parts.append("Alt")
        if modifiers & Qt.KeyboardModifier.ShiftModifier:
            parts.append("Shift")
        
        key_text = QKeySequence(key).toString()
        if key_text:
            if key_text.startswith("+"):
                key_text = "+"
            parts.append(key_text)
        
        if parts:
            self.captured_key = "+".join(parts)
            self.display.setText(self.captured_key)
            
            # Check for duplicates
            if self.captured_key in self.existing_triggers:
                self.display.setStyleSheet("""
                    font-size: 16px; font-weight: bold; padding: 20px;
                    background: #ffcccc; border: 2px solid #ff0000; color: #cc0000;
                    border-radius: 5px; min-height: 60px;
                """)
                self.display.setText(f"{self.captured_key}\n(Already in use!)")
                self.ok_btn.setEnabled(False)
            else:
                self.display.setStyleSheet("""
                    font-size: 16px; font-weight: bold; padding: 20px;
                    background: #ccffcc; border: 2px solid #00cc00;
                    border-radius: 5px; min-height: 60px;
                """)
                self.ok_btn.setEnabled(True)
    
    def clear_trigger(self):
        self.captured_key = ""
        self.display.setText("Waiting for key press...")
        self.display.setStyleSheet("""
            font-size: 16px; font-weight: bold; padding: 20px;
            background: #f0f0f0; border: 2px solid #ccc;
            border-radius: 5px; min-height: 60px;
        """)
        self.ok_btn.setEnabled(False)

# --- Snippets Manager with Variable Dropdown ---
class SnippetsManagerDialog(QDialog):
    def __init__(self, parent=None, snippets=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Snippets")
        self.resize(900, 600)
        self.snippets = snippets if snippets else []
        self.current_snippet = None
        
        layout = QHBoxLayout(self)
        
        # Left Panel
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("Snippets"))
        
        self.list_widget = QListWidget()
        self.list_widget.currentItemChanged.connect(self.on_selected)
        left_layout.addWidget(self.list_widget)
        
        # Buttons
        btn_box = QHBoxLayout()
        add_btn = QPushButton("Add")
        add_btn.clicked.connect(self.add_snippet)
        rem_btn = QPushButton("Remove")
        rem_btn.clicked.connect(self.remove_snippet)
        reset_btn = QPushButton("Reset to Defaults")
        reset_btn.clicked.connect(self.reset_defaults)
        btn_box.addWidget(add_btn)
        btn_box.addWidget(rem_btn)
        btn_box.addWidget(reset_btn)
        left_layout.addLayout(btn_box)
        
        # Right Panel
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("Edit Snippet"))
        
        # Name
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Name:"))
        self.name_edit = QLineEdit()
        self.name_edit.textChanged.connect(self.update_current)
        name_layout.addWidget(self.name_edit)
        right_layout.addLayout(name_layout)
        
        hint_label = QLabel("Name shown in snippets list")
        hint_label.setStyleSheet("color: gray; font-size: 10px; padding-left: 5px;")
        right_layout.addWidget(hint_label)
        
        # Trigger
        trigger_layout = QHBoxLayout()
        trigger_layout.addWidget(QLabel("Trigger:"))
        self.trig_edit = QLineEdit()
        self.trig_edit.setReadOnly(True)
        trigger_layout.addWidget(self.trig_edit)
        right_layout.addLayout(trigger_layout)
        
        trigger_btn_layout = QHBoxLayout()
        def_btn = QPushButton("Define Trigger")
        def_btn.clicked.connect(self.define_trigger)
        trigger_btn_layout.addWidget(def_btn)
        clear_trig_btn = QPushButton("Clear Trigger")
        clear_trig_btn.clicked.connect(self.clear_trigger)
        trigger_btn_layout.addWidget(clear_trig_btn)
        right_layout.addLayout(trigger_btn_layout)
        
        # Formatting Toggles (Rich Snippets)
        fmt_layout = QHBoxLayout()
        btn_style = """
            QPushButton {
                background-color: #e0e0e0;
                color: #222;
                border: 1px solid #777;
                border-radius: 3px;
                font-weight: bold;
                font-size: 14px;
                min-width: 40px;
                min-height: 30px;
            }
            QPushButton:checked {
                background-color: #1976d2;
                color: white;
                border: 1px solid #1565c0;
            }
            QPushButton:hover {
                background-color: #d0d0d0;
            }
        """
        
        self.btn_bold = QPushButton("B")
        self.btn_bold.setFixedWidth(40)
        self.btn_bold.setCheckable(True)
        self.btn_bold.setFont(QFont("Tahoma", 12, QFont.Weight.Bold))
        self.btn_bold.setStyleSheet(btn_style)
        self.btn_bold.clicked.connect(self.update_current)
        
        self.btn_italic = QPushButton("I")
        self.btn_italic.setFixedWidth(40)
        self.btn_italic.setCheckable(True)
        self.btn_italic.setFont(QFont("Tahoma", 12, QFont.Weight.Normal, True))
        self.btn_italic.setStyleSheet(btn_style)
        self.btn_italic.clicked.connect(self.update_current)
        
        self.btn_under = QPushButton("U")
        self.btn_under.setFixedWidth(40)
        self.btn_under.setCheckable(True)
        self.btn_under.setFont(QFont("Tahoma", 12))
        self.btn_under.setStyleSheet(btn_style + " QPushButton { text-decoration: underline; }")
        self.btn_under.clicked.connect(self.update_current)
        
        fmt_layout.addWidget(self.btn_bold)
        fmt_layout.addWidget(self.btn_italic)
        fmt_layout.addWidget(self.btn_under)
        fmt_layout.addStretch()
        right_layout.addLayout(fmt_layout)
        
        # Color
        color_layout = QHBoxLayout()
        color_layout.addWidget(QLabel("Color:"))
        self.color_btn = QPushButton("Choose Color")
        self.color_btn.clicked.connect(self.choose_color)
        self.snippet_color = QColor("blue")
        self.update_color_btn()
        color_layout.addWidget(self.color_btn)
        color_layout.addStretch()
        right_layout.addLayout(color_layout)
        
        # Carry Formatting (Stylized Toggle Button for extreme prominence)
        carry_label_layout = QHBoxLayout()
        carry_label_layout.addWidget(QLabel("Carry formatting to speech:"))
        self.carry_toggle = QPushButton("OFF")
        self.carry_toggle.setCheckable(True)
        self.carry_toggle.setFixedWidth(100)
        self.carry_toggle.setFixedHeight(40)
        self.carry_toggle.clicked.connect(self.on_carry_toggled)
        self.update_carry_style()
        carry_label_layout.addWidget(self.carry_toggle)
        carry_label_layout.addStretch()
        right_layout.addLayout(carry_label_layout)
        
        hint_carry = QLabel("If ON, snippet styling applies to text you type after it.")
        hint_carry.setStyleSheet("color: #5d4037; font-size: 11px; font-style: italic; padding: 0 5px 10px 5px;")
        right_layout.addWidget(hint_carry)
        
        # Snippet text
        right_layout.addWidget(QLabel("Snippet Text:"))
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Enter snippet text...\nUse variables like {$time}, {$date_short}, etc.")
        self.text_edit.textChanged.connect(self.update_current)
        right_layout.addWidget(self.text_edit)
        
        # Variable insertion dropdown
        var_layout = QHBoxLayout()
        var_layout.addWidget(QLabel("Insert Variable:"))
        self.variable_combo = QComboBox()
        self.variable_combo.addItems([
            "-- Select Variable to Insert --",
            "{$time} - Current timecode",
            "{$time_raw} - Timecode without brackets",
            "{$time_hours} - Hours only",
            "{$time_minutes} - Minutes only",
            "{$time_seconds} - Seconds only",
            "{$time_frames} - Frames only",
            "{$time_offset(00:00:05.00)} - Offset timecode",
            "{$time_raw_offset(-00:00:02.00)} - Raw offset timecode",
            "{$date_short} - Short date (12/31/2024)",
            "{$date_long} - Long date (December 31, 2024)",
            "{$date_abbrev} - Abbreviated date (Dec 31, 2024)",
            "{$clock_short} - Short time (2:23 PM)",
            "{$clock_long} - Long time (2:23:50 PM)",
            "{$media_name} - Media filename",
            "{$media_path} - Media full path",
            "{$doc_name} - Document filename",
            "{$doc_path} - Document full path",
            "{$selection} - Selected text",
            "{$version} - Application version"
        ])
        self.variable_combo.currentIndexChanged.connect(self.insert_variable)
        var_layout.addWidget(self.variable_combo)
        right_layout.addLayout(var_layout)
        
        # Help button for variables
        help_btn = QPushButton("View All Variables Reference")
        help_btn.clicked.connect(self.show_variables_help)
        right_layout.addWidget(help_btn)
        
        # Done button
        done_btn = QPushButton("Done")
        done_btn.clicked.connect(self.accept)
        right_layout.addWidget(done_btn)
        
        layout.addLayout(left_layout, 1)
        layout.addLayout(right_layout, 2)
        
        self.load_list()

    def load_list(self):
        self.list_widget.clear()
        for s in self.snippets:
            # Truncate text for display
            snippet_text = s['text'].replace('\n', ' ')[:50]
            if len(snippet_text) == 50:
                snippet_text += "..."
            
            item_text = f"{s['name']}"
            if s['trigger']:
                item_text += f" [{s['trigger']}]"
            item_text += f"\n{snippet_text}"
            
            item = QListWidgetItem(item_text)
            item.setForeground(QColor(s['color']))
            self.list_widget.addItem(item)
        
        if self.snippets:
            self.list_widget.setCurrentRow(0)

    def on_selected(self, current, prev):
        if not current:
            return
        idx = self.list_widget.row(current)
        if idx < len(self.snippets):
            self.current_snippet = self.snippets[idx]
            
            # Block signals to prevent recursion
            self.name_edit.blockSignals(True)
            self.text_edit.blockSignals(True)
            
            self.name_edit.setText(self.current_snippet['name'])
            self.trig_edit.setText(self.current_snippet['trigger'])
            self.text_edit.setPlainText(self.current_snippet['text'])
            self.snippet_color = QColor(self.current_snippet['color'])
            
            # Update Toggle
            self.carry_toggle.blockSignals(True)
            self.carry_toggle.setChecked(self.current_snippet.get('carry_format', False))
            self.update_carry_style()
            self.carry_toggle.blockSignals(False)
            
            self.btn_bold.setChecked(self.current_snippet.get('bold', True))
            self.btn_italic.setChecked(self.current_snippet.get('italic', False))
            self.btn_under.setChecked(self.current_snippet.get('underline', False))
            self.update_color_btn()
            
            self.name_edit.blockSignals(False)
            self.text_edit.blockSignals(False)
            self.btn_bold.blockSignals(False)
            self.btn_italic.blockSignals(False)
            self.btn_under.blockSignals(False)

    def update_current(self):
        if self.current_snippet:
            self.current_snippet['name'] = self.name_edit.text()
            self.current_snippet['text'] = self.text_edit.toPlainText()
            self.current_snippet['color'] = self.snippet_color.name()
            self.current_snippet['carry_format'] = self.carry_toggle.isChecked()
            self.current_snippet['bold'] = self.btn_bold.isChecked()
            self.current_snippet['italic'] = self.btn_italic.isChecked()
            self.current_snippet['underline'] = self.btn_under.isChecked()
            
            # Update list display for the current item ONLY
            curr_row = self.list_widget.currentRow()
            if curr_row >= 0:
                item = self.list_widget.item(curr_row)
                snippet_text = self.current_snippet['text'].replace('\n', ' ')[:50]
                if len(snippet_text) == 50:
                    snippet_text += "..."
                
                item_text = f"{self.current_snippet['name']}"
                if self.current_snippet['trigger']:
                    item_text += f" [{self.current_snippet['trigger']}]"
                item_text += f"\n{snippet_text}"
                item.setText(item_text)
                item.setForeground(QColor(self.current_snippet['color']))

    def add_snippet(self):
        name, ok = QInputDialog.getText(self, "New Snippet", "Enter snippet name:")
        if ok and name:
            self.snippets.append({
                'name': name, 
                'trigger': '', 
                'text': f'{name}: ', 
                'color': 'blue',
                'carry_format': False,
                'bold': True,
                'italic': False,
                'underline': False
            })
            self.load_list()
            self.list_widget.setCurrentRow(len(self.snippets) - 1)

    def remove_snippet(self):
        row = self.list_widget.currentRow()
        if row >= 0 and row < len(self.snippets):
            del self.snippets[row]
            self.load_list()

    def reset_defaults(self):
        reply = QMessageBox.question(
            self, "Reset Snippets", 
            "Reset all snippets to defaults? This will delete all custom snippets.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.snippets.clear()
            self.snippets.extend(self.get_default_snippets())
            self.load_list()

    def define_trigger(self):
        # Get existing triggers
        existing_triggers = []
        for snippet in self.snippets:
            if snippet['trigger'] and snippet != self.current_snippet:
                existing_triggers.append(snippet['trigger'])
        
        dlg = KeyCaptureDialog(self, existing_triggers)
        if dlg.exec() and dlg.captured_key:
            self.trig_edit.setText(dlg.captured_key)
            if self.current_snippet:
                self.current_snippet['trigger'] = dlg.captured_key
                self.update_current()

    def clear_trigger(self):
        self.trig_edit.clear()
        if self.current_snippet:
            self.current_snippet['trigger'] = ''
            self.update_current()

    def on_carry_toggled(self):
        self.update_carry_style()
        self.update_current()

    def update_carry_style(self):
        if self.carry_toggle.isChecked():
            self.carry_toggle.setText("ON ✓")
            self.carry_toggle.setStyleSheet("""
                QPushButton {
                    background-color: #4caf50;
                    color: white;
                    border: 2px solid #2e7d32;
                    border-radius: 5px;
                    font-weight: bold;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #66bb6a;
                }
            """)
        else:
            self.carry_toggle.setText("OFF")
            self.carry_toggle.setStyleSheet("""
                QPushButton {
                    background-color: #f44336;
                    color: white;
                    border: 2px solid #c62828;
                    border-radius: 5px;
                    font-weight: bold;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #ef5350;
                }
            """)

    def choose_color(self):
        c = QColorDialog.getColor(self.snippet_color, self, "Choose Snippet Color")
        if c.isValid():
            self.snippet_color = c
            self.update_color_btn()
            if self.current_snippet:
                self.current_snippet['color'] = c.name()
                self.update_current()

    def update_color_btn(self):
        from utils import get_contrast_color
        text_color = get_contrast_color(self.snippet_color)
        
        self.color_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.snippet_color.name()};
                color: {text_color};
                border: 2px solid #555;
                font-weight: bold;
                padding: 10px 20px;
                border-radius: 5px;
            }}
            QPushButton:hover {{
                background-color: {self.snippet_color.lighter(115).name()};
            }}
        """)

    def insert_variable(self, index):
        """Insert variable at cursor position in text edit"""
        if index > 0:
            variable_text = self.variable_combo.currentText()
            # Extract just the variable part
            variable_match = re.search(r'\{[^}]+\}', variable_text)
            if variable_match:
                variable = variable_match.group(0)
                cursor = self.text_edit.textCursor()
                cursor.insertText(variable)
                self.text_edit.setTextCursor(cursor)
                self.text_edit.setFocus() # Focus fix
            # Reset dropdown
            self.variable_combo.setCurrentIndex(0)

    def show_variables_help(self):
        """Show comprehensive variables help"""
        help_text = """
        <h3>Snippet Variables Reference</h3>
        <p>These variables will be replaced with actual values when you insert a snippet:</p>
        
        <h4>Time Variables</h4>
        <ul>
        <li><b>{$time}</b> - Current timecode with formatting [00:01:23.05]</li>
        <li><b>{$time_raw}</b> - Timecode without brackets 00:01:23.05</li>
        <li><b>{$time_hours}</b> - Hours component (00)</li>
        <li><b>{$time_minutes}</b> - Minutes component (01)</li>
        <li><b>{$time_seconds}</b> - Seconds component (23)</li>
        <li><b>{$time_frames}</b> - Frames component (05)</li>
        <li><b>${time_offset(00:00:10.00)}</b> - Timecode offset by specified amount</li>
        <li><b>${time_raw_offset(-00:00:05.00)}</b> - Raw timecode with negative offset</li>
        </ul>
        
        <h4>Date & Clock Variables</h4>
        <ul>
        <li><b>{$date_short}</b> - 12/31/2024</li>
        <li><b>{$date_long}</b> - Wednesday, December 31, 2024</li>
        <li><b>{$date_abbrev}</b> - Wed, Dec 31, 2024</li>
        <li><b>{$clock_short}</b> - 2:23 PM</li>
        <li><b>{$clock_long}</b> - 2:23:50 PM</li>
        </ul>
        
        <h4>Media & Document Variables</h4>
        <ul>
        <li><b>{$media_name}</b> - Media filename (example.mov)</li>
        <li><b>{$media_path}</b> - Full path to media file</li>
        <li><b>{$doc_name}</b> - Document filename</li>
        <li><b>{$doc_path}</b> - Full path to document</li>
        </ul>
        
        <h4>Other Variables</h4>
        <ul>
        <li><b>{$selection}</b> - Currently selected text in editor</li>
        <li><b>{$version}</b> - Application version number</li>
        </ul>
        
        <h4>Examples</h4>
        <p><b>Speaker with time:</b> SPEAKER 1: {$time}</p>
        <p><b>Note with date:</b> [NOTE - {$date_short}]</p>
        <p><b>Wrap selection:</b> &lt;b&gt;{$selection}&lt;/b&gt;</p>
        """
        
        msg = QMessageBox(self)
        msg.setWindowTitle("Snippet Variables Help")
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setText(help_text)
        msg.exec()

    @staticmethod
    def get_default_snippets():
        return [
            {'name': 'Speaker 1', 'trigger': 'Ctrl+1', 'text': 'SPEAKER 1: ', 'color': 'blue', 'carry_format': False, 'bold': True, 'italic': False, 'underline': False},
            {'name': 'Speaker 2', 'trigger': 'Ctrl+2', 'text': 'SPEAKER 2: ', 'color': 'red', 'carry_format': False, 'bold': True, 'italic': False, 'underline': False},
            {'name': 'Interviewer', 'trigger': 'Ctrl+3', 'text': 'INTERVIEWER: ', 'color': 'darkgreen', 'carry_format': False, 'bold': True, 'italic': False, 'underline': False},
            {'name': 'Interviewee', 'trigger': 'Ctrl+4', 'text': 'INTERVIEWEE: ', 'color': 'purple', 'carry_format': False, 'bold': True, 'italic': False, 'underline': False},
            {'name': 'Note', 'trigger': 'Ctrl+N', 'text': '[NOTE] ', 'color': 'gray', 'carry_format': False, 'bold': True, 'italic': False, 'underline': False},
            {'name': 'Laughter', 'trigger': 'Ctrl+L', 'text': '[laughter] ', 'color': 'orange', 'carry_format': False, 'bold': False, 'italic': True, 'underline': False},
        ]

# --- Shortcuts Manager ---
class ShortcutsManagerDialog(QDialog):
    def __init__(self, parent=None, shortcuts=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Shortcuts")
        self.resize(800, 500)
        self.shortcuts = shortcuts if shortcuts else []
        self.current_sc = None
        
        layout = QHBoxLayout(self)
        
        # Left panel
        left = QVBoxLayout()
        left.addWidget(QLabel("Shortcuts"))
        self.list_w = QListWidget()
        self.list_w.currentItemChanged.connect(self.on_sel)
        left.addWidget(self.list_w)
        
        btn_row = QHBoxLayout()
        add_btn = QPushButton("Add")
        add_btn.clicked.connect(self.add_shortcut)
        rem_btn = QPushButton("Remove")
        rem_btn.clicked.connect(self.remove_shortcut)
        rst_btn = QPushButton("Reset Defaults")
        rst_btn.clicked.connect(self.reset_defaults)
        btn_row.addWidget(add_btn)
        btn_row.addWidget(rem_btn)
        btn_row.addWidget(rst_btn)
        left.addLayout(btn_row)
        
        # Right panel
        right = QVBoxLayout()
        right.addWidget(QLabel("Edit Shortcut"))
        
        form = QFormLayout()
        
        self.name_edit = QLineEdit()
        self.name_edit.textChanged.connect(self.update_val)
        form.addRow("Name:", self.name_edit)
        
        trig_box = QHBoxLayout()
        self.trig_edit = QLineEdit()
        self.trig_edit.setReadOnly(True)
        def_btn = QPushButton("Define")
        def_btn.clicked.connect(self.define)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.clear_trigger)
        trig_box.addWidget(self.trig_edit)
        trig_box.addWidget(def_btn)
        trig_box.addWidget(clear_btn)
        form.addRow("Trigger:", trig_box)
        
        self.cmd_combo = QComboBox()
        self.cmd_combo.addItems([
            "Toggle Pause and Play",
            "Skipback",
            "Stop",
            "Pause",
            "Insert Current Time",
            "Play",
            "Fast Forward",
            "Rewind",
            "Advance One Frame",
            "Rewind One Frame",
            "Jump To End",
            "Jump To Beginning",
            "Go To Next Time Code",
            "Go To Previous Time Code",
            "Increase Play Rate",
            "Decrease Play Rate",
            "Increase Volume",
            "Decrease Volume"
        ])
        self.cmd_combo.currentTextChanged.connect(self.on_command_changed)
        form.addRow("Command:", self.cmd_combo)
        
        self.skip_spin = QDoubleSpinBox()
        self.skip_spin.setRange(0, 10)
        self.skip_spin.setSuffix(" sec")
        self.skip_spin.valueChanged.connect(self.update_val)
        self.skip_label = QLabel("Skip:")
        self.skip_row = (self.skip_label, self.skip_spin)
        form.addRow(self.skip_label, self.skip_spin)
        
        self.value_spin = QDoubleSpinBox()
        self.value_spin.setRange(0.1, 100.0)  # Increased range for volume
        self.value_spin.setSingleStep(0.1)
        self.value_spin.valueChanged.connect(self.update_val)
        self.value_label = QLabel("Value:")
        self.value_row = (self.value_label, self.value_spin)
        form.addRow(self.value_label, self.value_spin)
        
        right.addLayout(form)
        right.addStretch()
        
        done = QPushButton("Done")
        done.clicked.connect(self.accept)
        right.addWidget(done)
        
        layout.addLayout(left, 1)
        layout.addLayout(right, 1)
        self.load()

    def load(self):
        self.list_w.clear()
        for sc in self.shortcuts:
            skip_text = f" ({sc['skip']} sec)" if sc.get('skip', 0) > 0 else ""
            value_text = f" ({sc.get('value', 1.0)}x)" if sc.get('value', 1.0) not in [0, 1.0] else ""
            item_text = f"{sc['name']}{skip_text}{value_text}"
            if sc['trigger']:
                item_text += f" [{sc['trigger']}]"
            item = QListWidgetItem(item_text)
            self.list_w.addItem(item)
        
        if self.shortcuts:
            self.list_w.setCurrentRow(0)

    def command_needs_skip(self, command):
        """Check if command needs skip back field"""
        skip_commands = [
            "Toggle Pause and Play",
            "Skipback",
            "Fast Forward",
            "Rewind"
        ]
        return command in skip_commands
    
    def command_needs_value(self, command):
        """Check if command needs value field"""
        value_commands = [
            "Increase Play Rate",
            "Decrease Play Rate",
            "Increase Volume",
            "Decrease Volume"
        ]
        return command in value_commands
    
    def update_field_visibility(self):
        """Show/hide skip and value fields based on current command"""
        if not self.current_sc:
            return
        
        command = self.current_sc['command']
        
        # Show/hide skip field
        needs_skip = self.command_needs_skip(command)
        self.skip_label.setVisible(needs_skip)
        self.skip_spin.setVisible(needs_skip)
        
        # Show/hide value field
        needs_value = self.command_needs_value(command)
        self.value_label.setVisible(needs_value)
        self.value_spin.setVisible(needs_value)
    
    def on_command_changed(self):
        """Handle command change - update visibility and value"""
        self.update_field_visibility()
        self.update_val()
    
    def on_sel(self, curr, prev):
        if not curr:
            return
        idx = self.list_w.row(curr)
        if idx < len(self.shortcuts):
            self.current_sc = self.shortcuts[idx]
            
            # Block signals to prevent recursion when setting values
            self.name_edit.blockSignals(True)
            self.cmd_combo.blockSignals(True)
            self.trig_edit.blockSignals(True)
            self.skip_spin.blockSignals(True)
            self.value_spin.blockSignals(True)
            
            self.name_edit.setText(self.current_sc['name'])
            self.cmd_combo.setCurrentText(self.current_sc['command'])
            self.trig_edit.setText(self.current_sc['trigger'])
            self.skip_spin.setValue(self.current_sc.get('skip', 0))
            self.value_spin.setValue(self.current_sc.get('value', 1.0))
            
            self.name_edit.blockSignals(False)
            self.cmd_combo.blockSignals(False)
            self.trig_edit.blockSignals(False)
            self.skip_spin.blockSignals(False)
            self.value_spin.blockSignals(False)
            
            # Update field visibility based on command
            self.update_field_visibility()

    def define(self):
        # Get existing triggers
        existing_triggers = []
        for sc in self.shortcuts:
            if sc['trigger'] and sc != self.current_sc:
                existing_triggers.append(sc['trigger'])
        
        dlg = KeyCaptureDialog(self, existing_triggers)
        if dlg.exec() and dlg.captured_key:
            self.trig_edit.setText(dlg.captured_key)
            if self.current_sc:
                self.current_sc['trigger'] = dlg.captured_key
                self.load()
                idx = self.shortcuts.index(self.current_sc)
                self.list_w.setCurrentRow(idx)

    def clear_trigger(self):
        self.trig_edit.clear()
        if self.current_sc:
            self.current_sc['trigger'] = ''
            self.load()
            idx = self.shortcuts.index(self.current_sc)
            self.list_w.setCurrentRow(idx)

    def update_val(self):
        if self.current_sc:
            self.current_sc['name'] = self.name_edit.text()
            command = self.cmd_combo.currentText()
            self.current_sc['command'] = command
            
            # Only save skip if command needs it
            if self.command_needs_skip(command):
                self.current_sc['skip'] = self.skip_spin.value()
            elif 'skip' in self.current_sc:
                del self.current_sc['skip']
            
            # Only save value if command needs it
            if self.command_needs_value(command):
                self.current_sc['value'] = self.value_spin.value()
            elif 'value' in self.current_sc:
                del self.current_sc['value']
            
            # Update specific item text instead of full load() to avoid selection change loop
            curr_row = self.list_w.currentRow()
            if curr_row >= 0:
                sc = self.current_sc
                skip_text = f" ({sc['skip']} sec)" if sc.get('skip', 0) > 0 else ""
                value_text = f" ({sc.get('value', 1.0)}x)" if sc.get('value', 1.0) not in [0, 1.0] else ""
                item_text = f"{sc['name']}{skip_text}{value_text}"
                if sc['trigger']:
                    item_text += f" [{sc['trigger']}]"
                self.list_w.item(curr_row).setText(item_text)

    def add_shortcut(self):
        name, ok = QInputDialog.getText(self, "New Shortcut", "Enter shortcut name:")
        if ok and name:
            self.shortcuts.append({
                'name': name,
                'trigger': '',
                'command': 'Toggle Pause and Play',
                'skip': 0.3  # Default command needs skip
            })
            self.load()
            self.list_w.setCurrentRow(len(self.shortcuts) - 1)

    def remove_shortcut(self):
        row = self.list_w.currentRow()
        if row >= 0 and row < len(self.shortcuts):
            del self.shortcuts[row]
            self.load()

    def reset_defaults(self):
        reply = QMessageBox.question(
            self, "Reset Shortcuts",
            "Reset all shortcuts to defaults? This will delete all custom shortcuts.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.shortcuts.clear()
            self.shortcuts.extend(self.get_default_shortcuts())
            self.load()

    @staticmethod
    def get_default_shortcuts():
        return [
            # Shortcuts with skip back (playback controls)
            {'name': 'Toggle Pause and Play', 'trigger': 'Tab', 'command': 'Toggle Pause and Play', 'skip': 0.5},
            {'name': 'Toggle Pause and Play (1 sec)', 'trigger': 'Ctrl+Space', 'command': 'Toggle Pause and Play', 'skip': 1.0},
            {'name': 'Skipback (1 sec)', 'trigger': 'Ctrl+0', 'command': 'Skipback', 'skip': 1},
            {'name': 'Skipback (2 sec)', 'trigger': 'Ctrl+9', 'command': 'Skipback', 'skip': 2},
            {'name': 'Fast Forward', 'trigger': 'Alt+Right', 'command': 'Fast Forward', 'skip': 0},
            {'name': 'Rewind', 'trigger': 'Alt+Left', 'command': 'Rewind', 'skip': 0},
            # Shortcuts without skip back
            {'name': 'Insert Current Time', 'trigger': 'Ctrl+;', 'command': 'Insert Current Time'},
            # Shortcuts with value (speed/volume controls)
            {'name': 'Increase Speed', 'trigger': 'Ctrl+Up', 'command': 'Increase Play Rate', 'value': 0.1},
            {'name': 'Decrease Speed', 'trigger': 'Ctrl+Down', 'command': 'Decrease Play Rate', 'value': 0.1},
            {'name': 'Increase Volume', 'trigger': 'Ctrl+Shift+Up', 'command': 'Increase Volume', 'value': 10},
            {'name': 'Decrease Volume', 'trigger': 'Ctrl+Shift+Down', 'command': 'Decrease Volume', 'value': 10},
        ]

# --- Settings Dialog ---
# --- Settings Dialog / Widget ---
class TranscriptSettingsWidget(QWidget):
    def __init__(self, parent=None, settings=None, defaults=None):
        super().__init__(parent)
        self.settings = settings if settings else {}
        self.defaults = defaults if defaults else {}
        self.timecode_color = self.settings.get('timecode_color', 'green')
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(0, 0, 0, 0)
        
        def add_separator():
            line = QFrame()
            line.setFrameShape(QFrame.Shape.HLine)
            line.setFrameShadow(QFrame.Shadow.Sunken)
            line.setStyleSheet("background-color: #ccc;")
            layout.addWidget(line)

        # --- Font Section ---
        font_layout = QFormLayout()
        self.font_combo = QFontComboBox()
        self.font_combo.setFontFilters(QFontComboBox.FontFilter.ScalableFonts) # Prevent terminal errors
        self.font_combo.setWritingSystem(QFontDatabase.WritingSystem.Latin)   # Exclude complex scripts causing warnings
        self.font_combo.setEditable(False)
        self.font_combo.setCurrentFont(QFont(self.settings.get('font', 'Tahoma')))
        font_layout.addRow("Font:", self.font_combo)
        
        self.size_spin = QSpinBox()
        self.size_spin.setRange(8, 72)
        self.size_spin.setValue(int(self.settings.get('size', 12)))
        font_layout.addRow("Size:", self.size_spin)
        layout.addLayout(font_layout)
        
        add_separator()
        
        # --- Print Margins Section ---
        units_layout = QHBoxLayout()
        units_layout.addStretch()
        units_layout.addWidget(QLabel("Units:"))
        self.units_combo = QComboBox()
        self.units_combo.addItems(["Inches", "cm"])
        self.units_combo.setCurrentText(self.settings.get('print_margins', {}).get('units', 'Inches'))
        units_layout.addWidget(self.units_combo)
        layout.addLayout(units_layout)
        
        margin_inputs = QHBoxLayout()
        left_m_form = QFormLayout()
        self.top_margin = QDoubleSpinBox()
        self.top_margin.setValue(self.settings.get('print_margins', {}).get('top', 1.0))
        left_m_form.addRow("Top:", self.top_margin)
        self.left_margin = QDoubleSpinBox()
        self.left_margin.setValue(self.settings.get('print_margins', {}).get('left', 1.0))
        left_m_form.addRow("Left:", self.left_margin)
        
        right_m_form = QFormLayout()
        self.bottom_margin = QDoubleSpinBox()
        self.bottom_margin.setValue(self.settings.get('print_margins', {}).get('bottom', 1.0))
        right_m_form.addRow("Bottom:", self.bottom_margin)
        self.right_margin = QDoubleSpinBox()
        self.right_margin.setValue(self.settings.get('print_margins', {}).get('right', 1.0))
        right_m_form.addRow("Right:", self.right_margin)
        
        margin_inputs.addLayout(left_m_form)
        margin_inputs.addLayout(right_m_form)
        
        margin_row_layout = QHBoxLayout()
        margin_row_layout.addWidget(QLabel("Print Margins:"))
        margin_row_layout.addLayout(margin_inputs)
        layout.addLayout(margin_row_layout)
        
        add_separator()
        
        # --- Frame Rate Section ---
        fps_form = QFormLayout()
        self.fps_combo = QComboBox()
        self.fps_combo.addItems(["23.976 fps (Film)", "24 fps (Film)", "25 fps (PAL)", 
                                  "29.97 fps (NTSC)", "30 fps (QuickTime Default)", 
                                  "50 fps", "59.94 fps", "60 fps"])
        
        current_fps = self.settings.get('fps', 23.976)
        best_index = 0
        min_diff = 999.0
        for i in range(self.fps_combo.count()):
            text = self.fps_combo.itemText(i)
            match = re.search(r'(\d+\.?\d*)', text)
            if match:
                val = float(match.group(1))
                diff = abs(val - current_fps)
                if diff < min_diff:
                    min_diff = diff
                    best_index = i
        self.fps_combo.setCurrentIndex(best_index)
        fps_form.addRow("Frame Rate:", self.fps_combo)
        layout.addLayout(fps_form)
        
        add_separator()
        
        # --- Timecode Layout ---
        tc_outer = QHBoxLayout()
        tc_outer.addWidget(QLabel("Inserted Time Code\nFormat:"), alignment=Qt.AlignmentFlag.AlignTop)
        
        tc_inner = QVBoxLayout()
        self.format_combo = QComboBox()
        self.format_combo.addItems(["[00:01:23:29]", "(00:01:23:29)", "{00:01:23:29}", "<00:01:23:29>", "00:01:23:29"])
        self.format_combo.setCurrentText(self.settings.get('timecode_format', '[00:01:23:29]'))
        tc_inner.addWidget(self.format_combo)
        
        
        self.omit_frames_check = QCheckBox("Omit Frames")
        self.omit_frames_check.setToolTip("Hides the frames segment [FF] from timecodes in the transcript.")
        self.omit_frames_check.setChecked(self.settings.get('omit_frames', False))
        tc_inner.addWidget(self.omit_frames_check)
        
        # Color row
        # Color rows
        color_grid = QGridLayout()
        
        # Timecode Color
        color_grid.addWidget(QLabel("Timecode Color:"), 0, 0)
        self.color_btn = QPushButton("Choose...")
        self.color_btn.clicked.connect(self.choose_timecode_color)
        self.update_color_button()
        color_grid.addWidget(self.color_btn, 0, 1)
        
        # Default Text Color
        color_grid.addWidget(QLabel("Default Text Color:"), 1, 0)
        self.text_color_btn = QPushButton("Choose...")
        self.text_color_btn.clicked.connect(self.choose_text_color)
        self.text_color = self.settings.get('font_color', 'black')
        self.update_text_color_button()
        color_grid.addWidget(self.text_color_btn, 1, 1)
        
        tc_inner.addLayout(color_grid)
        
        # Style row
        style_layout = QHBoxLayout()
        style_layout.addWidget(QLabel("Time Code Style:"))
        self.style_combo = QComboBox()
        self.style_combo.addItems(["Normal", "Bold", "Italic", "Underline"])
        self.style_combo.setCurrentText(self.settings.get('timecode_style', 'Bold'))
        style_layout.addWidget(self.style_combo)
        style_layout.addStretch()
        tc_inner.addLayout(style_layout)
        
        tc_outer.addLayout(tc_inner)
        layout.addLayout(tc_outer)
        
        add_separator()
        
        # --- Other Section ---
        other_layout = QHBoxLayout()
        other_layout.addWidget(QLabel("Other:"))
        self.unbracketed_check = QCheckBox("Recognize Unbracketed Time Codes")
        self.unbracketed_check.setChecked(self.settings.get('recognize_unbracketed', True))
        other_layout.addWidget(self.unbracketed_check)
        other_layout.addStretch()
        layout.addLayout(other_layout)
        
        layout.addStretch()

    def use_defaults(self):
        if not self.defaults: return
        self.font_combo.setCurrentFont(QFont(self.defaults.get('font', 'Tahoma')))
        self.size_spin.setValue(self.defaults.get('size', 12))
        df_fps = self.defaults.get('fps', 23.976)
        for i in range(self.fps_combo.count()):
            m = re.search(r'(\d+\.?\d*)', self.fps_combo.itemText(i))
            if m and abs(float(m.group(1)) - df_fps) < 0.1:
                self.fps_combo.setCurrentIndex(i); break
        self.format_combo.setCurrentText(self.defaults.get('timecode_format', '[00:01:23.29]'))
        self.omit_frames_check.setChecked(self.defaults.get('omit_frames', False))
        self.unbracketed_check.setChecked(self.defaults.get('recognize_unbracketed', True))
        margins = self.defaults.get('print_margins', {})
        self.units_combo.setCurrentText(margins.get('units', 'Inches'))
        self.top_margin.setValue(margins.get('top', 1.0))
        self.bottom_margin.setValue(margins.get('bottom', 1.0))
        self.left_margin.setValue(margins.get('left', 1.0))
        self.right_margin.setValue(margins.get('right', 1.0))
        self.timecode_color = self.defaults.get('timecode_color', 'green')
        self.update_color_button()
        self.text_color = self.defaults.get('font_color', 'black')
        self.update_text_color_button()
        self.style_combo.setCurrentText(self.defaults.get('timecode_style', 'Bold'))

    def sync_settings(self):
        fps_text = self.fps_combo.currentText()
        fps_match = re.search(r'(\d+\.?\d*)', fps_text)
        fps = float(fps_match.group(1)) if fps_match else 23.976
        
        self.settings.update({
            'font': self.font_combo.currentFont().family(),
            'size': self.size_spin.value(),
            'fps': fps,
            'timecode_format': self.format_combo.currentText(),
            'omit_frames': self.omit_frames_check.isChecked(),
            'timecode_color': self.timecode_color,
            'font_color': self.text_color,
            'timecode_style': self.style_combo.currentText(),
            'recognize_unbracketed': self.unbracketed_check.isChecked(),
            'print_margins': {
                'units': self.units_combo.currentText(),
                'top': self.top_margin.value(),
                'bottom': self.bottom_margin.value(),
                'left': self.left_margin.value(),
                'right': self.right_margin.value()
            }
        })

    def choose_timecode_color(self):
        color = QColorDialog.getColor(QColor(self.timecode_color), self, "Select Timecode Color")
        if color.isValid():
            self.timecode_color = color.name()
            self.update_color_button()

    def update_color_button(self):
        self.color_btn.setStyleSheet(f"background-color: {self.timecode_color}; color: {'white' if QColor(self.timecode_color).lightness() < 128 else 'black'}; font-weight: bold; border: 1px solid #555; border-radius: 3px;")

    def choose_text_color(self):
        color = QColorDialog.getColor(QColor(self.text_color), self, "Select Default Text Color")
        if color.isValid():
            self.text_color = color.name()
            self.update_text_color_button()

    def update_text_color_button(self):
        self.text_color_btn.setStyleSheet(f"background-color: {self.text_color}; color: {'white' if QColor(self.text_color).lightness() < 128 else 'black'}; font-weight: bold; border: 1px solid #555; border-radius: 3px;")

class TranscriptSettingsDialog(QDialog):
    def __init__(self, parent=None, settings=None, defaults=None):
        super().__init__(parent)
        self.setWindowTitle("Transcript Settings")
        self.resize(550, 650)
        layout = QVBoxLayout(self)
        self.widget = TranscriptSettingsWidget(self, settings, defaults)
        layout.addWidget(self.widget)
        
        btn_box = QHBoxLayout()
        self.btn_defaults = QPushButton("Use Defaults"); self.btn_defaults.clicked.connect(self.widget.use_defaults)
        btn_box.addWidget(self.btn_defaults); btn_box.addStretch()
        
        self.btn_cancel = QPushButton("Cancel"); self.btn_cancel.clicked.connect(self.reject); btn_box.addWidget(self.btn_cancel)
        self.btn_ok = QPushButton("OK"); self.btn_ok.setDefault(True); self.btn_ok.clicked.connect(self.accept); btn_box.addWidget(self.btn_ok)
        layout.addLayout(btn_box)

    def get_settings(self):
        self.widget.sync_settings()
        return self.widget.settings


    def choose_timecode_color(self):
        color = QColorDialog.getColor(QColor(self.timecode_color), self, "Select Timecode Color")
        if color.isValid():
            self.timecode_color = color.name()
            self.update_color_button()

    def update_color_button(self):
        self.color_btn.setStyleSheet(f"background-color: {self.timecode_color}; color: {'white' if QColor(self.timecode_color).lightness() < 128 else 'black'}; font-weight: bold;")

# --- Export Settings Dialog ---
class ExportSettingsDialog(QDialog):
    def __init__(self, parent=None, initial_format='html', current_file_path=None, initial_settings=None):
        super().__init__(parent)
        self.setWindowTitle("Export Settings")
        self.resize(750, 700)
        self.current_format = initial_format
        self.target_path = ""
        
        # Load persistent settings if available
        self.settings_cache = initial_settings if initial_settings else {}
        if 'format' in self.settings_cache:
            self.current_format = self.settings_cache['format']

        # Default filename based on current path
        import os
        if current_file_path:
            base = os.path.splitext(current_file_path)[0]
            self.target_path = f"{base}.{self._get_ext(self.current_format)}"
        else:
            self.target_path = f"export.{self._get_ext(self.current_format)}"

        # Access parent theme colors if available
        self.theme = {}
        if parent and hasattr(parent, 'active_theme_colors'):
            self.theme = parent.active_theme_colors
        
        bg = self.theme.get('main_bg', '#f8fafc')
        tc = self.theme.get('text', '#1e293b')
        accent = self.theme.get('accent', '#3b82f6')
        border = self.theme.get('border', '#cbd5e1')
        
        self.setStyleSheet(f"""
            QDialog {{ background-color: {bg}; color: {tc}; font-family: 'Segoe UI', system-ui; }}
            QLabel {{ color: {tc}; font-size: 13px; }}
            QCheckBox {{ color: {tc}; font-size: 13px; spacing: 8px; }}
            QCheckBox::indicator {{ width: 18px; height: 18px; }}
            QLineEdit {{ 
                background-color: {self.theme.get('edit_bg', 'white')}; 
                color: black; 
                border: 1px solid {border}; 
                border-radius: 6px;
                padding: 8px; 
            }}
            QComboBox {{ 
                background-color: {self.theme.get('edit_bg', 'white')}; 
                color: black; 
                border: 1px solid {border}; 
                border-radius: 6px;
                padding: 6px; 
            }}
            QPushButton#ExportBtn {{
                background-color: {accent};
                color: white;
                font-weight: bold;
                border-radius: 8px;
                padding: 10px 25px;
            }}
            QPushButton#SecondaryBtn {{
                background-color: transparent;
                color: {tc};
                border: 1px solid {border};
                border-radius: 8px;
                padding: 10px 25px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        # Header Section
        header = QLabel("Export Transcript")
        header.setStyleSheet(f"font-size: 24px; font-weight: 800; color: {accent}; margin-bottom: 5px;")
        layout.addWidget(header)

        # --- Section 1: Format Selection ---
        format_card = PremiumCard(self)
        format_layout = QVBoxLayout(format_card)
        
        fmt_title = QLabel("Choose Export Format")
        fmt_title.setStyleSheet("font-weight: bold; font-size: 14px; text-transform: uppercase; color: gray;")
        format_layout.addWidget(fmt_title)
        
        self.format_combo = QComboBox()
        self.formats = {
            'html': 'HTML Document',
            'odf': 'Open Document Text (.odt) - Microsoft Word / LibreOffice',
            'csv': 'Excel / CSV (.csv) - Spreadsheet data',
            'tab': 'Tab-delimited Text - For databases/coding',
            'txt': 'Plain Text (.txt) - Minimalist transcript'
        }
        for key, label in self.formats.items():
            self.format_combo.addItem(label, key)
        
        idx = self.format_combo.findData(self.current_format)
        self.format_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.format_combo.currentIndexChanged.connect(self.on_format_changed)
        format_layout.addWidget(self.format_combo)
        layout.addWidget(format_card)

        # --- Section 2: Content Options ---
        self.opts_card = PremiumCard(self)
        self.opts_layout = QVBoxLayout(self.opts_card)
        
        opts_title = QLabel("Content Options")
        opts_title.setStyleSheet("font-weight: bold; font-size: 14px; text-transform: uppercase; color: gray;")
        self.opts_layout.addWidget(opts_title)

        grid = QGridLayout()
        grid.setSpacing(15)

        # 1. Out Points
        self.chk_out_points = QCheckBox("Export Out Points")
        self.chk_out_points.setChecked(self.settings_cache.get('export_out_points', True))
        self.chk_out_points.stateChanged.connect(self.update_ui_state)
        grid.addWidget(self.chk_out_points, 0, 0)
        
        out_tip = QLabel("Infers end times based on next entry")
        out_tip.setStyleSheet("color: gray; font-size: 11px;")
        grid.addWidget(out_tip, 0, 1)

        # 2. Durations
        self.chk_durations = QCheckBox("Export Durations")
        self.chk_durations.setChecked(self.settings_cache.get('export_durations', False))
        grid.addWidget(self.chk_durations, 1, 0)
        
        dur_tip = QLabel("Include duration of each segment")
        dur_tip.setStyleSheet("color: gray; font-size: 11px;")
        grid.addWidget(dur_tip, 1, 1)

        # 3. Speakers
        self.chk_speaker_names = QCheckBox("Export Speaker Names")
        self.chk_speaker_names.setChecked(self.settings_cache.get('export_speakers', True))
        self.chk_speaker_names.stateChanged.connect(self.update_ui_state)
        grid.addWidget(self.chk_speaker_names, 2, 0)

        # Delimiter Row
        self.delim_widget = QWidget()
        delim_h = QHBoxLayout(self.delim_widget)
        delim_h.setContentsMargins(0,0,0,0)
        delim_h.addWidget(QLabel("Delimiter:"))
        self.delim_edit = QLineEdit(self.settings_cache.get('speaker_delimiter', ':'))
        self.delim_edit.setFixedWidth(50)
        self.delim_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        delim_h.addWidget(self.delim_edit)
        delim_h.addStretch()
        grid.addWidget(self.delim_widget, 2, 1)

        # 4. SFX
        self.chk_sfx = QCheckBox("Export SFX/Notes")
        self.chk_sfx.setChecked(self.settings_cache.get('export_sfx', True))
        grid.addWidget(self.chk_sfx, 3, 0)
        
        sfx_tip = QLabel("Treat [bracketed text] as separate field")
        sfx_tip.setStyleSheet("color: gray; font-size: 11px;")
        grid.addWidget(sfx_tip, 3, 1)

        self.opts_layout.addLayout(grid)
        layout.addWidget(self.opts_card)

        # --- Section 3: Destination ---
        dest_card = PremiumCard(self)
        dest_layout = QVBoxLayout(dest_card)
        
        dest_title = QLabel("Destination & Encoding")
        dest_title.setStyleSheet("font-weight: bold; font-size: 14px; text-transform: uppercase; color: gray;")
        dest_layout.addWidget(dest_title)

        target_row = QHBoxLayout()
        self.target_edit = QLineEdit(self.target_path)
        self.target_edit.setPlaceholderText("Select output file...")
        target_row.addWidget(self.target_edit, 1)
        
        self.btn_choose = QPushButton("Browse...")
        self.btn_choose.setObjectName("SecondaryBtn")
        self.btn_choose.clicked.connect(self.choose_target)
        target_row.addWidget(self.btn_choose)
        dest_layout.addLayout(target_row)

        enc_row = QHBoxLayout()
        enc_row.addWidget(QLabel("Encoding:"))
        self.enc_combo = QComboBox()
        self.enc_combo.addItems(["UTF-8", "UTF-16", "ISO-8859-1 (Latin-1)", "ASCII", "UTF-8 with BOM"])
        
        saved_enc = self.settings_cache.get('encoding', 'UTF-8')
        enc_idx = self.enc_combo.findText(saved_enc)
        if enc_idx >= 0: self.enc_combo.setCurrentIndex(enc_idx)
        
        enc_row.addWidget(self.enc_combo)
        enc_row.addStretch()
        
        self.chk_replace = QCheckBox("Overwite without asking")
        self.chk_replace.setChecked(self.settings_cache.get('replace_existing', False))
        self.chk_replace.setCursor(Qt.CursorShape.PointingHandCursor)
        enc_row.addWidget(self.chk_replace)
        dest_layout.addLayout(enc_row)
        
        layout.addWidget(dest_card)
        layout.addStretch()

        # Footer Buttons
        footer = QHBoxLayout()
        footer.setSpacing(15)
        
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setObjectName("SecondaryBtn")
        self.btn_cancel.clicked.connect(self.reject)
        
        self.btn_export = QPushButton("Start Export")
        self.btn_export.setObjectName("ExportBtn")
        self.btn_export.setDefault(True)
        self.btn_export.clicked.connect(self.accept)
        
        footer.addStretch()
        footer.addWidget(self.btn_cancel)
        footer.addWidget(self.btn_export)
        layout.addLayout(footer)

        self.update_ui_state()

    def _get_ext(self, fmt):
        ext_map = {
            'html': 'html', 'odf': 'odt',
            'csv': 'csv', 'tab': 'txt', 'srt': 'srt',
            'scc': 'scc', 'scc_export': 'scc'
        }
        return ext_map.get(fmt, 'txt')

    def on_format_changed(self):
        self.current_format = self.format_combo.currentData()
        
        # Update extension in target edit
        import os
        path = self.target_edit.text()
        if path:
            base = os.path.splitext(path)[0]
            ext = self._get_ext(self.current_format)
            self.target_edit.setText(f"{base}.{ext}")
        
        # Enable/Disable options based on format
        is_transcript = self.current_format in ['html', 'odf', 'csv', 'tab']
        self.opts_card.setEnabled(is_transcript)
        
        self.update_ui_state()

    def update_ui_state(self):
        self.chk_durations.setEnabled(self.chk_out_points.isChecked())
        self.delim_widget.setEnabled(self.chk_speaker_names.isChecked())

    def choose_target(self):
        ext = self._get_ext(self.current_format)
        f, _ = QFileDialog.getSaveFileName(self, "Export Target", self.target_edit.text(), f"{ext.upper()} Files (*.{ext})")
        if f:
            self.target_edit.setText(f)

    def get_settings(self):
        return {
            'format': self.current_format,
            'export_out_points': self.chk_out_points.isChecked(),
            'export_durations': self.chk_durations.isChecked(),
            'export_speakers': self.chk_speaker_names.isChecked(),
            'export_sfx': self.chk_sfx.isChecked(),
            'speaker_delimiter': self.delim_edit.text(),
            'target_path': self.target_edit.text(),
            'replace_existing': self.chk_replace.isChecked(),
            'encoding': self.enc_combo.currentText()
        }

    def on_apply(self):
        if self.parent() and hasattr(self.parent(), 'config'):
            self.parent().config['export'] = self.get_settings()
            self.parent().save_config()


# --- Adjust Time Codes Dialog ---
class AdjustTimecodesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Adjust Time Codes")
        self.resize(550, 350)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        # Row 1: Add/Subtract dropdown, Amount, Help, Label
        row1 = QHBoxLayout()
        self.op_combo = QComboBox()
        self.op_combo.addItems(["Add", "Subtract"])
        row1.addWidget(self.op_combo)
        
        self.amount_edit = QLineEdit("00:00:00.00")
        self.amount_edit.setFixedWidth(120)
        row1.addWidget(self.amount_edit)
        
        help_btn = QPushButton("?")
        help_btn.setFixedSize(24, 24)
        help_btn.setStyleSheet("border-radius: 12px; font-weight: bold; background: #ddd;")
        row1.addWidget(help_btn)
        
        row1.addWidget(QLabel("to each time code."))
        row1.addStretch()
        layout.addLayout(row1)
        
        # Hint text
        hint = QLabel("Enter an adjustment amount. This action is undoable.")
        hint.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(hint)
        
        # Checkbox: Adjust Selection Only
        self.chk_selection = QCheckBox("Adjust Selection Only")
        layout.addWidget(self.chk_selection)
        
        # Checkbox Hint
        chk_hint = QLabel("Normally all timecodes in the document will be adjusted. Check this box to adjust only those timecodes within the current selection.")
        chk_hint.setWordWrap(True)
        chk_hint.setStyleSheet("color: #666; font-size: 11px; margin-bottom: 20px;")
        layout.addWidget(chk_hint)
        
        layout.addStretch()
        
        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        self.btn_adjust = QPushButton("Adjust")
        self.btn_adjust.setDefault(True)
        self.btn_cancel = QPushButton("Cancel")
        
        self.btn_adjust.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addWidget(self.btn_adjust)
        btn_layout.addWidget(self.btn_cancel)
        layout.addLayout(btn_layout)

    def get_data(self):
        return {
            'operation': self.op_combo.currentText(), # "Add" or "Subtract"
            'amount': self.amount_edit.text(),
            'selection_only': self.chk_selection.isChecked()
        }

# --- Media Offset Dialog with Scanning ---
class MediaOffsetDialog(QDialog):
    def __init__(self, parent=None, current_offset_str="00:00:00.00", media_path=None):
        super().__init__(parent)
        self.setWindowTitle("Set Media Offset")
        self.resize(450, 200)
        self.media_path = media_path
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        info = QLabel("Set the starting timecode for this media. This will affect all new timecodes and the display.")
        info.setWordWrap(True)
        info.setStyleSheet("color: #64748b; margin-bottom: 5px;")
        layout.addWidget(info)
        
        form = QFormLayout()
        self.offset_edit = QLineEdit(current_offset_str)
        self.offset_edit.setPlaceholderText("HH:MM:SS.FF")
        self.offset_edit.setStyleSheet("font-size: 16px; font-weight: bold; padding: 5px;")
        form.addRow("Start Offset:", self.offset_edit)
        layout.addLayout(form)
        
        # Scan Button
        self.btn_scan = QPushButton("🔍 Scan Media Metadata for Timecode")
        self.btn_scan.setStyleSheet("""
            QPushButton { 
                padding: 8px; background: #f1f5f9; border: 1px solid #cbd5e1; 
                border-radius: 4px; font-weight: bold;
            }
            QPushButton:hover { background: #e2e8f0; }
        """)
        self.btn_scan.clicked.connect(self.scan_metadata)
        if not self.media_path:
            self.btn_scan.setEnabled(False)
            self.btn_scan.setToolTip("Load media first to scan metadata")
            
        layout.addWidget(self.btn_scan)
        
        # Buttons
        btns = QHBoxLayout()
        self.btn_ok = QPushButton("Set Offset")
        self.btn_ok.setDefault(True)
        self.btn_ok.clicked.connect(self.accept)
        self.btn_ok.setStyleSheet("padding: 8px 20px; background: #3b82f6; color: white; font-weight: bold; border-radius: 4px;")
        
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_cancel.setStyleSheet("padding: 8px 20px;")
        
        btns.addStretch()
        btns.addWidget(self.btn_cancel)
        btns.addWidget(self.btn_ok)
        layout.addLayout(btns)

    def scan_metadata(self):
        """Use ffprobe to find starting timecode metadata"""
        if not self.media_path or not os.path.exists(self.media_path):
            QMessageBox.warning(self, "No Media", "Please load a media file first.")
            return
            
        if not shutil.which("ffprobe"):
            QMessageBox.warning(self, "ffprobe Not Found", "ffmpeg/ffprobe (ffprobe.exe) must be in your PATH for metadata scanning.")
            return
            
        self.btn_scan.setText("⏳ Scanning Media...")
        self.btn_scan.setEnabled(False)
        self.repaint() # Force UI update
        
        try:
            # 1. Check format tags (Standard)
            cmd = ['ffprobe', '-v', 'error', '-show_entries', 'format_tags=timecode', '-of', 'default=noprint_wrappers=1:nokey=1', self.media_path]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
            tc = result.stdout.strip()
            
            # 2. Check stream tags if format tag empty
            if not tc:
                cmd = ['ffprobe', '-v', 'error', '-select_streams', 'v:0', '-show_entries', 'stream_tags=timecode', '-of', 'default=noprint_wrappers=1:nokey=1', self.media_path]
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
                tc = result.stdout.strip()
                
            if tc:
                # Clean up and validate
                tc = tc.replace('\r', '').strip()
                self.offset_edit.setText(tc)
                self.offset_edit.setStyleSheet("font-size: 16px; font-weight: bold; padding: 5px; background: #ccffcc; border: 2px solid #00cc00;")
                QTimer.singleShot(1500, lambda: self.offset_edit.setStyleSheet("font-size: 16px; font-weight: bold; padding: 5px;"))
            else:
                QMessageBox.information(self, "No Metadata", "No timecode metadata found in this specific file.")
        except Exception as e:
            QMessageBox.critical(self, "Scan Error", f"Failed to scan metadata: {str(e)}")
        finally:
            self.btn_scan.setText("🔍 Scan Media Metadata for Timecode")
            self.btn_scan.setEnabled(True)

    def get_offset(self):
        return self.offset_edit.text()



class PreferencesDialog(QDialog):
    def __init__(self, parent, config, defaults=None):
        super().__init__(parent)
        self.setWindowTitle("Preferences")
        self.resize(600, 550)
        self.config = config
        self.defaults = defaults
        
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        
        # --- TAB 1: General ---
        self.tab_general = QWidget()
        gen_layout = QVBoxLayout(self.tab_general)
        gen_layout.setSpacing(20)
        
        def add_sep(lay):
            line = QFrame(); line.setFrameShape(QFrame.Shape.HLine)
            line.setFrameShadow(QFrame.Shadow.Sunken); line.setStyleSheet("background-color: #ccc;")
            lay.addWidget(line)

        # Backups
        bk_layout = QHBoxLayout()
        bk_layout.addWidget(QLabel("Backups:"))
        self.save_backups_check = QCheckBox("Save backups every:")
        self.save_backups_check.setChecked(True)
        bk_layout.addWidget(self.save_backups_check)
        self.backup_interval = QSpinBox()
        self.backup_interval.setRange(1, 60)
        self.backup_interval.setValue(self.config.get('backup_interval', 10))
        bk_layout.addWidget(self.backup_interval)
        bk_layout.addWidget(QLabel("min"))
        bk_layout.addStretch()
        gen_layout.addLayout(bk_layout)
        
        add_sep(gen_layout)
        
        # Display
        disp_form = QFormLayout()
        disp_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        
        win_size_layout = QHBoxLayout()
        current_win_size = self.config.get('default_window_size', [1024, 800])
        self.win_w = QSpinBox(); self.win_w.setRange(800, 3840); self.win_w.setValue(current_win_size[0])
        win_size_layout.addWidget(self.win_w); win_size_layout.addWidget(QLabel("wide"))
        self.win_h = QSpinBox(); self.win_h.setRange(600, 2160); self.win_h.setValue(current_win_size[1])
        win_size_layout.addWidget(self.win_h); win_size_layout.addWidget(QLabel("high"))
        disp_form.addRow("New Window Size:", win_size_layout)
        
        self.color_tc_check = QCheckBox("Color Time Codes")
        self.color_tc_check.setChecked(True)
        self.tc_color_btn = QPushButton()
        self.tc_color_btn.setFixedSize(60, 25)
        self.tc_color_btn.clicked.connect(self.pick_tc_color)
        self.current_tc_color = self.config['settings'].get('timecode_color', 'green')
        self.update_tc_btn()
        
        tc_color_layout = QHBoxLayout()
        tc_color_layout.addWidget(self.tc_color_btn)
        tc_color_layout.addStretch()
        disp_form.addRow("Timecode Color", tc_color_layout)
        
        self.newline_tc_check = QCheckBox("Timecodes and Snippets always start on a new line")
        self.newline_tc_check.setChecked(self.config['settings'].get('timecode_new_line', True))
        disp_form.addRow("Timecodes in Lines", self.newline_tc_check)
        
        gen_layout.addLayout(disp_form)
        
        add_sep(gen_layout)
        
        # Media
        med_form = QFormLayout()
        med_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        
        self.pitch_lock_check = QCheckBox("Pitch Lock")
        self.pitch_lock_check.setToolTip("Keeps the voice sounding natural (not Chipmunks/Darth Vader) when changing playback speed.")
        self.pitch_lock_check.setChecked(self.config.get('playback', {}).get('pitch_lock', True))
        med_form.addRow("Media:", self.pitch_lock_check)
        
        self.autoplay_check = QCheckBox("Start playback automatically after loading media")
        self.autoplay_check.setChecked(self.config.get('autoplay_on_load', True))
        med_form.addRow("Behavior:", self.autoplay_check)
        
        self.auto_wf_check = QCheckBox("Automatically generate waveform on media load")
        self.auto_wf_check.setChecked(self.config.get('auto_generate_waveform', False))
        med_form.addRow("Waveform:", self.auto_wf_check)
        
        self.aspect_combo = QComboBox()
        self.aspect_combo.addItems(["Use Media Ratio", "16:9", "4:3", "2.35:1", "1:1"])
        med_form.addRow("Aspect Ratio:", self.aspect_combo)
        
        self.player_combo = QComboBox()
        self.player_combo.addItems(["MPV (Recommended)", "VLC", "FFmpeg", "QuickTime"])
        curr_p = self.config.get('playback', {}).get('preferred_player', 'mpv')
        p_idx = 0
        if 'vlc' in curr_p.lower(): p_idx = 1
        elif 'ffmpeg' in curr_p.lower(): p_idx = 2
        elif 'quicktime' in curr_p.lower(): p_idx = 3
        self.player_combo.setCurrentIndex(p_idx)
        med_form.addRow("Preferred Player:", self.player_combo)
        
        # MPV Specific - Version Selector
        from media_engine import MPVBackend
        self.detected_mpvs = MPVBackend.discover_dlls()
        if len(self.detected_mpvs) > 1:
            self.mpv_selector = QComboBox()
            curr_mpv_path = self.config.get('playback', {}).get('mpv_path')
            selected_idx = 0
            for i, md in enumerate(self.detected_mpvs):
                self.mpv_selector.addItem(f"{md['name']} ({os.path.basename(md['path'])})", md['path'])
                if md['path'] == curr_mpv_path:
                    selected_idx = i
            self.mpv_selector.setCurrentIndex(selected_idx)
            med_form.addRow("MPV Engine Version:", self.mpv_selector)
        else:
            self.mpv_selector = None
            
        gen_layout.addLayout(med_form)
        
        add_sep(gen_layout)
        
        # Waveform Cache
        wf_layout = QHBoxLayout()
        wf_layout.addWidget(QLabel("Waveform Cache:"))
        
        self.btn_clear_wf = QPushButton("Clear Waveform Folder")
        self.btn_clear_wf.clicked.connect(self.clear_waveform_cache)
        self.btn_clear_wf.setStyleSheet("color: #d32f2f;")
        wf_layout.addWidget(self.btn_clear_wf)
        
        wf_layout.addSpacing(20)
        wf_layout.addWidget(QLabel("Delete after:"))
        self.wf_retention = QComboBox()
        self.wf_retention.addItems(["NEVER", "1 month", "3 months", "6 months", "12 months"])
        
        # Map months to index
        wf_ret_map = {0: 0, 1: 1, 3: 2, 6: 3, 12: 4}
        current_wf_ret = self.config.get('waveform_retention_months', 3)
        self.wf_retention.setCurrentIndex(wf_ret_map.get(current_wf_ret, 2))
        wf_layout.addWidget(self.wf_retention)
        
        wf_layout.addStretch()
        gen_layout.addLayout(wf_layout)
        
        add_sep(gen_layout)
        
        # Player Scaling Behavior
        scaling_layout = QHBoxLayout()
        scaling_layout.addWidget(QLabel("Player Resize Behavior:"))
        
        self.scaling_combo = QComboBox()
        self.scaling_behaviors = [
            ("Stretch with Window (Classic)", "proportional"),
            ("Limit Control Width (Wide Prevention)", "cap"),
            ("Center Controls (Modern Island)", "island")
        ]
        for name, key in self.scaling_behaviors:
            self.scaling_combo.addItem(name, key)
            
        curr_behavior = self.config.get('ui', {}).get('player_scaling_behavior', 'proportional')
        for i, (name, key) in enumerate(self.scaling_behaviors):
            if key == curr_behavior:
                self.scaling_combo.setCurrentIndex(i)
                break
                
        scaling_layout.addWidget(self.scaling_combo)
        scaling_layout.addStretch()
        gen_layout.addLayout(scaling_layout)
        
        add_sep(gen_layout)
        
        gen_layout.addStretch()
        
        # --- TAB 2: New Document ---
        self.tab_newdoc = QWidget()
        nd_layout = QVBoxLayout(self.tab_newdoc)
        self.ts_widget = TranscriptSettingsWidget(self, self.config['settings'], self.defaults)
        nd_layout.addWidget(self.ts_widget)
        
        self.tabs.addTab(self.tab_general, "General")
        self.tabs.addTab(self.tab_newdoc, "New Document")
        layout.addWidget(self.tabs)
        
        # Bottom Buttons
        bot_layout = QHBoxLayout()
        self.btn_edit_sc = QPushButton("Edit Shortcuts...")
        self.btn_edit_sc.clicked.connect(self.on_edit_shortcuts)
        bot_layout.addWidget(self.btn_edit_sc)
        bot_layout.addStretch()
        
        self.btn_ok = QPushButton("OK"); self.btn_ok.setDefault(True); self.btn_ok.clicked.connect(self.accept)
        self.btn_can = QPushButton("Cancel"); self.btn_can.clicked.connect(self.reject)
        bot_layout.addWidget(self.btn_ok); bot_layout.addWidget(self.btn_can)
        layout.addLayout(bot_layout)

    def pick_tc_color(self):
        c = QColorDialog.getColor(QColor(self.current_tc_color), self, "Select Timecode Color")
        if c.isValid():
            self.current_tc_color = c.name()
            self.update_tc_btn()

    def update_tc_btn(self):
        self.tc_color_btn.setStyleSheet(f"background-color: {self.current_tc_color}; border: 1px solid #555;")

    def on_edit_shortcuts(self):
        if self.parent():
            self.parent().open_shortcuts_dialog()

    def clear_waveform_cache(self):
        reply = QMessageBox.question(
            self, "Clear Waveform Cache", 
            "Are you sure you want to delete all cached waveform files? This will cause waveforms to be regenerated next time you open media.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            from waveform import WaveformWorker
            if WaveformWorker.clear_cache():
                QMessageBox.information(self, "Clear Cache", "Waveform cache cleared successfully.")
            else:
                QMessageBox.warning(self, "Clear Cache", "Failed to clear some cache files.")

    def get_config(self):
        self.ts_widget.sync_settings()
        self.config['backup_interval'] = self.backup_interval.value()
        self.config['default_window_size'] = [self.win_w.value(), self.win_h.value()]
        self.config['settings']['timecode_color'] = self.current_tc_color
        self.config['settings']['timecode_new_line'] = self.newline_tc_check.isChecked()
        if 'playback' not in self.config: self.config['playback'] = {}
        self.config['playback']['pitch_lock'] = self.pitch_lock_check.isChecked()
        players = ['mpv', 'vlc', 'ffmpeg', 'quicktime']
        self.config['playback']['preferred_player'] = players[self.player_combo.currentIndex()]
        
        self.config['ui']['player_scaling_behavior'] = self.scaling_combo.currentData()
        self.config['autoplay_on_load'] = self.autoplay_check.isChecked()
        self.config['auto_generate_waveform'] = self.auto_wf_check.isChecked()
        
        if hasattr(self, 'mpv_selector') and self.mpv_selector:
            self.config['playback']['mpv_path'] = self.mpv_selector.currentData()
            
        # Waveform retention
        wf_ret_map = {0: 0, 1: 1, 2: 3, 3: 6, 4: 12}
        self.config['waveform_retention_months'] = wf_ret_map.get(self.wf_retention.currentIndex(), 3)
        
        return self.config

from hardware import USBManager

class ManageUSBDevicesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage USB Devices")
        self.resize(700, 400)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        layout.addWidget(QLabel("<b>Connected HID Class Devices:</b>"))
        
        self.card = PremiumCard()
        card_layout = QVBoxLayout(self.card)
        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet("border: none; background: transparent;")
        card_layout.addWidget(self.list_widget)
        layout.addWidget(self.card)
        
        btn_refresh = QPushButton("Refresh Device List")
        btn_refresh.clicked.connect(self.refresh_list)
        layout.addWidget(btn_refresh)
        
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        
        self.refresh_list()

    def refresh_list(self):
        self.list_widget.clear()
        devices = USBManager.list_hid_devices()
        for d in devices:
            name = d.get('FriendlyName') or "Unknown Device"
            inst_id = d.get('InstanceId') or "No ID"
            status = d.get('Status') or "Unknown"
            
            item_text = f"{name}\nID: {inst_id} ({status})"
            item = QListWidgetItem(item_text)
            self.list_widget.addItem(item)

class FootPedalSetupDialog(QDialog):
    def __init__(self, parent, config):
        super().__init__(parent)
        self.setWindowTitle("Setup Foot Pedal")
        self.resize(600, 450)
        self.config = config
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)
        
        # Device Selection Card
        self.card_dev = PremiumCard()
        dev_card_layout = QVBoxLayout(self.card_dev)
        dev_card_layout.addWidget(QLabel("<b>1. Select your Foot Pedal device:</b>"))
        self.device_combo = QComboBox()
        self.refresh_devices()
        dev_card_layout.addWidget(self.device_combo)
        
        btn_refresh = QPushButton("Refresh Device List")
        btn_refresh.clicked.connect(self.refresh_devices)
        dev_card_layout.addWidget(btn_refresh)
        layout.addWidget(self.card_dev)
        
        # Calibration Section Card
        self.card_cal = PremiumCard()
        cal_card_layout = QVBoxLayout(self.card_cal)
        cal_card_layout.addWidget(QLabel("<b>2. Pedal Mapping (Calibration):</b>"))
        self.calib_info = QLabel("Click a pedal below, then press that pedal on your hardware.")
        self.calib_info.setStyleSheet("font-style: italic; color: #5dade2;")
        cal_card_layout.addWidget(self.calib_info)
        
        btn_layout = QHBoxLayout()
        self.btn_left = QPushButton("Left Pedal")
        self.btn_mid = QPushButton("Middle Pedal")
        self.btn_right = QPushButton("Right Pedal")
        
        for b in [self.btn_left, self.btn_mid, self.btn_right]:
            btn_layout.addWidget(b)
        cal_card_layout.addLayout(btn_layout)
        layout.addWidget(self.card_cal)
        
        layout.addStretch()
        
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.on_save)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def refresh_devices(self):
        self.device_combo.clear()
        devices = USBManager.list_hid_devices()
        for d in devices:
            name = d.get('FriendlyName') or "HID Device"
            id_ = d.get('InstanceId')
            self.device_combo.addItem(name, id_)

    def on_save(self):
        device_id = self.device_combo.currentData()
        if not device_id:
            QMessageBox.warning(self, "No Device", "Please select a device first.")
            return

        # Prepare config struct
        if 'hardware' not in self.config: self.config['hardware'] = {}
        self.config['hardware']['pedal_id'] = device_id
        
        # In a real scenario, we'd save buttons here from calibration
        # self.config['hardware']['pedal_mapping'] = self.mappings
        
        QMessageBox.information(self, "Hardware Saved", 
            f"Foot pedal settings for '{self.device_combo.currentText()}' saved.")
        self.accept()


class ThemePropertyEditor(QDialog):
    def __init__(self, parent, theme_data, name):
        super().__init__(parent)
        self.setWindowTitle(f"Editing Theme: {name}")
        self.resize(500, 600)
        self.theme_data = theme_data.copy()
        
        layout = QVBoxLayout(self)
        form = QFormLayout()
        
        self.widgets = {}
        
        # Define editable fields
        fields = [
            ('main_bg', 'Main Background'),
            ('text', 'Main Text Color'),
            ('edit_bg', 'Editor Background'),
            ('top_bar', 'Ribbon/Top Bar'),
            ('accent', 'Primary Accent'),
            ('border', 'Border Color'),
            ('glow', 'Glow/Hover Effect'),
        ]
        
        for key, label in fields:
            card = QFrame()
            card_layout = QHBoxLayout(card)
            
            color = self.theme_data.get(key, "#000000")
            btn = QPushButton()
            btn.setFixedSize(60, 30)
            btn.setStyleSheet(f"background-color: {color}; border: 1px solid #666; border-radius: 4px;")
            btn.clicked.connect(lambda checked, k=key, b=btn: self.pick_color(k, b))
            
            card_layout.addWidget(QLabel(label))
            card_layout.addStretch()
            card_layout.addWidget(btn)
            
            self.widgets[key] = btn
            form.addRow(card)

        # Gradient Support
        grad_frame = QFrame()
        grad_layout = QVBoxLayout(grad_frame)
        grad_layout.addWidget(QLabel("Button Gradients (Primary, Mid, Bottom):"))
        
        grad_btns_layout = QHBoxLayout()
        self.grad_btns = []
        grads = self.theme_data.get('btn_grad', ["#ffffff", "#f1f5f9", "#e2e8f0"])
        for i in range(3):
            gcol = grads[i]
            btn = QPushButton()
            btn.setFixedSize(60, 30)
            btn.setStyleSheet(f"background-color: {gcol}; border: 1px solid #666; border-radius: 4px;")
            btn.clicked.connect(lambda checked, idx=i, b=btn: self.pick_grad_color(idx, b))
            grad_btns_layout.addWidget(btn)
            self.grad_btns.append(btn)
        
        grad_layout.addLayout(grad_btns_layout)
        form.addRow(grad_frame)

        layout.addLayout(form)
        
        # Buttons
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def pick_color(self, key, btn):
        initial = QColor(self.theme_data[key])
        color = QColorDialog.getColor(initial, self, f"Select {key}")
        if color.isValid():
            self.theme_data[key] = color.name()
            btn.setStyleSheet(f"background-color: {color.name()}; border: 1px solid #666; border-radius: 4px;")

    def pick_grad_color(self, idx, btn):
        grads = list(self.theme_data.get('btn_grad', ["#ffffff", "#f1f5f9", "#e2e8f0"]))
        initial = QColor(grads[idx])
        color = QColorDialog.getColor(initial, self, f"Select Gradient Step {idx+1}")
        if color.isValid():
            grads[idx] = color.name()
            self.theme_data['btn_grad'] = grads
            btn.setStyleSheet(f"background-color: {color.name()}; border: 1px solid #666; border-radius: 4px;")

    def get_data(self):
        return self.theme_data


class ThemeBuilderDialog(QDialog):
    def __init__(self, parent, builtin_themes, custom_themes, hidden_themes):
        super().__init__(parent)
        self.setWindowTitle("Theme Builder & Manager")
        self.resize(700, 500)
        
        self.builtin = builtin_themes
        self.custom_themes = custom_themes.copy()
        self.hidden_themes = hidden_themes.copy()
        
        layout = QHBoxLayout(self)
        
        # Left Panel: List
        left = QVBoxLayout()
        left.addWidget(QLabel("All Themes:"))
        self.list = QListWidget()
        self.list.itemSelectionChanged.connect(self.update_ui)
        left.addWidget(self.list)
        layout.addLayout(left, 2)
        
        # Right Panel: Actions
        right = QVBoxLayout()
        right.addSpacing(20)
        
        self.btn_new = QPushButton("Create Copy/New...")
        self.btn_new.clicked.connect(self.create_new)
        right.addWidget(self.btn_new)
        
        self.btn_edit = QPushButton("Edit Properties...")
        self.btn_edit.clicked.connect(self.edit_theme)
        right.addWidget(self.btn_edit)
        
        self.btn_toggle = QPushButton("Hide/Show")
        self.btn_toggle.clicked.connect(self.toggle_visibility)
        right.addWidget(self.btn_toggle)
        
        self.btn_delete = QPushButton("Delete Custom")
        self.btn_delete.clicked.connect(self.delete_custom)
        right.addWidget(self.btn_delete)
        
        right.addStretch()
        
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        right.addWidget(btns)
        
        layout.addLayout(right, 1)
        
        self.refresh_list()

    def refresh_list(self):
        self.list.clear()
        
        # Built-in
        for name in self.builtin.keys():
            item = QListWidgetItem(name)
            if name in self.hidden_themes:
                item.setText(f"{name} (Hidden)")
                item.setForeground(QColor("gray"))
            self.list.addItem(item)
            
        # Custom
        for name in self.custom_themes.keys():
            item = QListWidgetItem(f"[Custom] {name}")
            if name in self.hidden_themes:
                item.setText(f"[Custom] {name} (Hidden)")
                item.setForeground(QColor("gray"))
            item.setData(Qt.ItemDataRole.UserRole, name)
            self.list.addItem(item)

    def update_ui(self):
        item = self.list.currentItem()
        if not item:
            self.btn_edit.setEnabled(False)
            self.btn_delete.setEnabled(False)
            self.btn_toggle.setEnabled(False)
            return
            
        is_custom = "[Custom]" in item.text()
        self.btn_edit.setEnabled(is_custom)
        self.btn_delete.setEnabled(is_custom)
        self.btn_toggle.setEnabled(True)

    def create_new(self):
        item = self.list.currentItem()
        base_name = item.text().replace("[Custom] ", "").replace(" (Hidden)", "") if item else "Light"
        
        new_name, ok = QInputDialog.getText(self, "New Theme", "Enter theme name:", text=f"{base_name} Copy")
        if ok and new_name:
            # Copy from current or base
            themes = self.builtin.copy()
            themes.update(self.custom_themes)
            base_data = themes.get(base_name, themes['Light'])
            
            self.custom_themes[new_name] = base_data.copy()
            self.refresh_list()

    def edit_theme(self):
        item = self.list.currentItem()
        if not item: return
        name = item.data(Qt.ItemDataRole.UserRole)
        if not name: return # Not a custom theme
        
        dlg = ThemePropertyEditor(self, self.custom_themes[name], name)
        if dlg.exec():
            self.custom_themes[name] = dlg.get_data()
            # Show live preview by signal or just wait for Save
            if hasattr(self.parent(), 'apply_theme'):
                self.parent().apply_theme(name)

    def toggle_visibility(self):
        item = self.list.currentItem()
        if not item: return
        
        name = item.text().replace("[Custom] ", "").replace(" (Hidden)", "")
        # Resolve real name for custom
        real_name = item.data(Qt.ItemDataRole.UserRole) or name
        
        if real_name in self.hidden_themes:
            self.hidden_themes.remove(real_name)
        else:
            self.hidden_themes.append(real_name)
            
        self.refresh_list()

    def delete_custom(self):
        item = self.list.currentItem()
        if not item: return
        name = item.data(Qt.ItemDataRole.UserRole)
        if not name: return
        
        if QMessageBox.question(self, "Delete?", f"Delete custom theme '{name}'?") == QMessageBox.StandardButton.Yes:
            del self.custom_themes[name]
            if name in self.hidden_themes: self.hidden_themes.remove(name)
            self.refresh_list()



# --- Sleek Floating Spell Check Assistant ---
class SpellCheckDialog(QDialog):
    def __init__(self, parent=None, editor=None):
        super().__init__(parent)
        self.editor = editor
        self.setWindowTitle("Spell Check Assistant")
        # Removed WindowStaysOnTopHint so dialog doesn't stay on top of other apps
        self.resize(380, 780) # Sleek, vertical profile
        
        self.current_error = None
        self.current_index = 0
        self.history = [] # Action history for Undo

        self.init_ui()
        self.load_languages()
        
        # Start immediately from the very first character
        self.next_error(0)

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(12)

        # Header: Word Not Found (Big and Bold)
        top_v = QVBoxLayout()
        header = QLabel("WORD NOT FOUND")
        header.setStyleSheet("font-size: 10px; font-weight: bold; color: #64748b; letter-spacing: 1px;")
        top_v.addWidget(header)
        
        self.word_edit = QLineEdit()
        self.word_edit.setStyleSheet("""
            font-size: 20px; font-weight: bold; padding: 12px; 
            color: #e11d48; background: #fff1f2; border: 2px solid #fda4af;
            border-radius: 8px;
        """)
        top_v.addWidget(self.word_edit)
        main_layout.addLayout(top_v)

        # Suggestions Region (The Centerpiece)
        mid_v = QVBoxLayout()
        header_s = QLabel("SUGGESTIONS")
        header_s.setStyleSheet("font-size: 10px; font-weight: bold; color: #64748b; letter-spacing: 1px;")
        mid_v.addWidget(header_s)
        
        self.sugg_list = QListWidget()
        self.sugg_list.setStyleSheet("""
            QListWidget { font-size: 15px; padding: 5px; border-radius: 8px; border: 1px solid #cbd5e1; }
            QListWidget::item { padding: 10px; border-bottom: 1px solid #f1f5f9; }
            QListWidget::item:selected { background: #3b82f6; color: white; border-radius: 4px; }
        """)
        self.sugg_list.itemDoubleClicked.connect(self.on_use)
        mid_v.addWidget(self.sugg_list, 1) # Expandable
        
        btn_h = QHBoxLayout()
        self.btn_use = QPushButton("Use Selection")
        self.btn_use.setStyleSheet("font-weight: bold; background: #3b82f6; color: white; height: 40px;")
        self.btn_use.clicked.connect(self.on_use)
        btn_h.addWidget(self.btn_use)
        mid_v.addLayout(btn_h)
        main_layout.addLayout(mid_v)

        # Action Buttons (Clean Grid)
        grid_v = QVBoxLayout()
        grid = QGridLayout()
        grid.setSpacing(8)
        
        def create_act_btn(text, slot, color=None):
            btn = QPushButton(text)
            btn.setFixedHeight(34)
            btn.clicked.connect(slot)
            if color: btn.setStyleSheet(f"color: {color};")
            return btn

        self.btn_change_all = create_act_btn("Change All Instances", self.on_use_always)
        self.btn_skip = create_act_btn("Skip Once", self.on_skip)
        self.btn_skip_all = create_act_btn("Skip All (Session)", self.on_skip_all)
        self.btn_add_dict = create_act_btn("Add to Dictionary", self.on_add_to_dict)
        self.btn_google = create_act_btn("🔎 Search Google", self.on_google, "#3b82f6")
        
        grid.addWidget(self.btn_skip, 0, 0)
        grid.addWidget(self.btn_skip_all, 0, 1)
        grid.addWidget(self.btn_change_all, 1, 0, 1, 2)
        grid.addWidget(self.btn_add_dict, 2, 0, 1, 2)
        grid.addWidget(self.btn_google, 3, 0, 1, 2)
        grid_v.addLayout(grid)
        main_layout.addLayout(grid_v)

        # Language Control
        lang_h = QHBoxLayout()
        lang_h.addWidget(QLabel("Dictionary:"))
        self.lang_combo = QComboBox()
        self.lang_combo.currentIndexChanged.connect(self.on_lang_changed)
        # Give combo a bit less stretch to make room for 'Add'
        lang_h.addWidget(self.lang_combo, 4)
        
        # DOWNLOAD BUTTON - Wider to avoid 'd' cut-off
        self.btn_download = QPushButton("Add")
        self.btn_download.setToolTip("Download more dictionaries...")
        self.btn_download.setFixedWidth(75)
        self.btn_download.setStyleSheet("font-weight: bold; background: #f1f5f9; border: 1px solid #cbd5e1;")
        self.btn_download.clicked.connect(self.open_downloader)
        lang_h.addWidget(self.btn_download)
        
        main_layout.addLayout(lang_h)

        # Footer with Undo
        main_layout.addSpacing(10)
        self.btn_undo = QPushButton("Undo last action")
        self.btn_undo.setEnabled(False)
        self.btn_undo.setStyleSheet("font-style: italic; font-size: 12px; color: #64748b; border: 1px dashed #cbd5e1; height: 36px;")
        self.btn_undo.clicked.connect(self.on_undo)
        main_layout.addWidget(self.btn_undo)
        
        bottom_h = QHBoxLayout()
        self.status_lbl = QLabel("Checking...")
        self.status_lbl.setStyleSheet("color: #94a3b8; font-size: 11px;")
        bottom_h.addWidget(self.status_lbl)
        bottom_h.addStretch()
        self.btn_abort = QPushButton("Close")
        self.btn_abort.clicked.connect(self.reject)
        bottom_h.addWidget(self.btn_abort)
        main_layout.addLayout(bottom_h)

    def load_languages(self):
        self.lang_combo.blockSignals(True)
        self.lang_combo.clear()
        
        # Human-readable language names
        lang_names = {
            'en_US': 'English (US)',
            'en_GB': 'English (GB)',
            'en_AU': 'English (AU)',
            'en_CA': 'English (CA)',
            'en_NZ': 'English (NZ)',
            'en_ZA': 'English (ZA)',
            'en_IN': 'English (IN)',
            'es_ES': 'Spanish',
            'fr_FR': 'French',
            'de_DE': 'German',
            'it_IT': 'Italian',
            'pt_BR': 'Portuguese (Brazil)',
            'pt_PT': 'Portuguese (Portugal)',
            'ru_RU': 'Russian',
            'ar': 'Arabic',
            'tr_TR': 'Turkish',
            'nl_NL': 'Dutch',
            'pl_PL': 'Polish',
            'cs_CZ': 'Czech',
            'da_DK': 'Danish',
            'sv_SE': 'Swedish',
            'nb_NO': 'Norwegian (Bokmål)',
            'el_GR': 'Greek',
            'hu_HU': 'Hungarian',
            'ro': 'Romanian',
            'hr_HR': 'Croatian',
            'sk_SK': 'Slovak',
            'sl_SI': 'Slovenian',
            'uk_UA': 'Ukrainian',
        }
        
        # Scan ONLY local dicts folder for available dictionaries
        import glob
        base_dir = os.path.dirname(os.path.abspath(__file__))
        dicts_dir = os.path.join(base_dir, "dicts")
        
        available_langs = []
        if os.path.exists(dicts_dir):
            # Find all .dic files
            dic_files = glob.glob(os.path.join(dicts_dir, "*.dic"))
            for dic_file in dic_files:
                # Extract language code from filename (e.g., en_US from en_US.dic)
                lang_code = os.path.splitext(os.path.basename(dic_file))[0]
                # Check if corresponding .aff file exists
                aff_file = os.path.join(dicts_dir, f"{lang_code}.aff")
                if os.path.exists(aff_file):
                    available_langs.append(lang_code)
        
        # Sort: English variants first, then others alphabetically
        english_langs = [l for l in available_langs if l.startswith('en_')]
        other_langs = [l for l in available_langs if not l.startswith('en_')]
        
        # Add English variants
        for lang_code in sorted(english_langs):
            display_name = lang_names.get(lang_code, lang_code)
            self.lang_combo.addItem(display_name, lang_code)
        
        # Add other languages
        for lang_code in sorted(other_langs):
            display_name = lang_names.get(lang_code, lang_code)
            self.lang_combo.addItem(display_name, lang_code)
        
        # Sync with editor's current lang
        curr_lang = self.editor.highlighter.lang
        idx = self.lang_combo.findData(curr_lang)
        if idx >= 0: self.lang_combo.setCurrentIndex(idx)
        else:
            # Fallback to US if found
            idx_us = self.lang_combo.findData("en_US")
            if idx_us >= 0: self.lang_combo.setCurrentIndex(idx_us)
            
        self.lang_combo.blockSignals(False)

    def open_downloader(self):
        dlg = DictionaryDownloaderDialog(self)
        if dlg.exec():
            self.load_languages() # Refresh list

    def next_error(self, start_from=None):
        if start_from is not None:
            self.current_index = start_from
            
        error = self.editor.find_next_error(self.current_index)
        if error:
            self.current_error = error
            self.current_index = error['end']
            self.update_error_ui()
        else:
            self.status_lbl.setText("Finished!")
            QMessageBox.information(self, "Spell Check", "No more misspellings found in document.")
            self.accept()

    def update_error_ui(self):
        word = self.current_error['word']
        self.word_edit.setText(word)
        
        # Suggestions Refresh
        self.sugg_list.clear()
        suggestions = self.editor.highlighter.dict.suggest(word)
        if suggestions:
            self.sugg_list.addItems(suggestions[:25])
            self.sugg_list.setCurrentRow(0)
            
        # Context Handling: FOCUS REAL EDITOR
        self.editor.highlight_word(self.current_error['start'], self.current_error['end'])
        
        self.status_lbl.setText(f"Found '{word}'")

    def add_history(self, action_type, word, details=""):
        desc = f"Undo: {action_type} '{word}'"
        if details: desc = f"Undo: {details}"
        self.history.append({'type': action_type, 'word': word, 'details': details, 'index': self.current_error['start']})
        self.btn_undo.setEnabled(True)
        self.btn_undo.setText(desc)

    def on_undo(self):
        if not self.history: return
        last = self.history.pop()
        
        if last['type'] == 'Skip All':
            self.editor.remove_from_skip_list(last['word'])
        elif last['type'] == 'Add to Dictionary':
            self.editor.remove_from_dictionary(last['word'])
        elif last['type'] == 'Change':
            pass # Standard internal undo chain would handle this
            
        self.btn_undo.setEnabled(len(self.history) > 0)
        self.btn_undo.setText(f"Undo: {self.history[-1]['details']}" if self.history else "Undo last action")
        
        # Jump back
        self.next_error(last['index'])

    def on_lang_changed(self, index):
        lang = self.lang_combo.currentData()
        self.editor.set_spell_language(lang)
        self.next_error(0)

    def on_use(self):
        old_word = self.current_error['word']
        new_word = self.word_edit.text()
        if self.sugg_list.currentItem() and new_word == old_word:
            new_word = self.sugg_list.currentItem().text()
        
        self.add_history("Change", old_word, f"{old_word} -> {new_word}")
        self.editor.replace_selection(self.current_error['start'], self.current_error['end'], new_word)
        # Advance
        self.current_index = self.current_error['start'] + len(new_word)
        self.next_error()

    def on_use_always(self):
        # Global replace logic
        old_word = self.current_error['word']
        new_word = self.word_edit.text()
        if self.sugg_list.currentItem() and new_word == old_word:
             new_word = self.sugg_list.currentItem().text()
             
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        doc = self.editor.document()
        search_cursor = QTextCursor(doc)
        count = 0
        while True:
            found_cursor = doc.find(old_word, search_cursor, QTextDocument.FindFlag.FindWholeWords)
            if found_cursor.isNull(): break
            found_cursor.insertText(new_word)
            search_cursor = found_cursor
            count += 1
        cursor.endEditBlock()
        
        self.add_history("Change All", old_word, f"Replaced {count} instances of '{old_word}'")
        self.next_error()

    def on_skip(self):
        self.next_error()

    def on_skip_all(self):
        word = self.current_error['word']
        self.add_history("Skip All", word)
        self.editor.add_to_skip_list(word)
        self.next_error()

    def on_add_to_dict(self):
        word = self.current_error['word']
        self.add_history("Add to Dictionary", word)
        self.editor.highlighter.dict.add(word)
        self.editor.highlighter.rehighlight()
        self.next_error()

    def on_google(self):
        word = self.current_error['word']
        webbrowser.open(f"https://www.google.com/search?q={word}")


# --- One-Click Dictionary Downloader ---
class DictionaryDownloaderDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Download Dictionaries")
        self.resize(500, 600)
        
        # Only show dictionaries available in LibreOffice repository
        # Based on https://github.com/LibreOffice/dictionaries
        self.langs = {
            "English (United States)": "en_US",
            "English (United Kingdom)": "en_GB",
            "English (Australia)": "en_AU",
            "English (Canada)": "en_CA",
            "Spanish": "es_ES",
            "French": "fr_FR",
            "German": "de_DE",
            "Italian": "it_IT",
            "Portuguese (Brazil)": "pt_BR",
            "Portuguese (Portugal)": "pt_PT",
            "Russian": "ru_RU",
            "Arabic": "ar",
            "Turkish": "tr_TR",
            "Dutch": "nl_NL",
            "Polish": "pl_PL",
            "Czech": "cs_CZ",
            "Danish": "da_DK",
            "Swedish": "sv_SE",
            "Norwegian (Bokmål)": "nb_NO",
            "Greek": "el_GR",
            "Hungarian": "hu_HU",
            "Romanian": "ro",
            "Croatian": "hr_HR",
            "Slovak": "sk_SK",
            "Slovenian": "sl_SI",
            "Ukrainian": "uk_UA",
        }
        
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        layout.addWidget(QLabel("Available Languages:"))
        self.list = QListWidget()
        self.list.setStyleSheet("""
            QListWidget { font-size: 14px; border: 1px solid #cbd5e1; outline: none; }
            QListWidget::item { padding: 8px; border-bottom: 1px solid #f1f5f9; }
            QListWidget::item:selected { background: #3b82f6; color: white; }
        """)
        for name in sorted(self.langs.keys()):
            item = QListWidgetItem(name)
            code = self.langs[name]
            # Check if already installed in LOCAL dicts folder
            base_dir = os.path.dirname(os.path.abspath(__file__))
            dicts_dir = os.path.join(base_dir, "dicts")
            dic_file = os.path.join(dicts_dir, f"{code}.dic")
            aff_file = os.path.join(dicts_dir, f"{code}.aff")
            
            if os.path.exists(dic_file) and os.path.exists(aff_file):
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEnabled)
                item.setText(f"{name} (Installed)")
            self.list.addItem(item)
        layout.addWidget(self.list)
        
        # Progress dialog will be created only when needed during download
        self.progress = None
        
        btn_h = QHBoxLayout()
        self.btn_install = QPushButton("Install Selected")
        self.btn_install.setFixedHeight(40)
        self.btn_install.setStyleSheet("background: #10b981; color: white; font-weight: bold;")
        self.btn_install.clicked.connect(self.start_download)
        btn_h.addWidget(self.btn_install)
        
        self.btn_close = QPushButton("Cancel")
        self.btn_close.clicked.connect(self.reject)
        btn_h.addWidget(self.btn_close)
        layout.addLayout(btn_h)

    def start_download(self):
        item = self.list.currentItem()
        if not item or "Installed" in item.text(): return
        
        lang_name = item.text()
        lang_code = self.langs[lang_name]
        
        # Create progress dialog only when download starts
        self.progress = QProgressDialog(f"Downloading {lang_name}...", "Cancel", 0, 100, self)
        self.progress.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress.setValue(0)
        self.progress.show()
        
        # Start in thread
        threading.Thread(target=self.do_download, args=(lang_code,), daemon=True).start()

    def do_download(self, code):
        try:
            # Use LibreOffice dictionaries repository
            # Structure: https://github.com/LibreOffice/dictionaries/tree/master/[folder]/[code].dic
            
            # Map language codes to their folder names in LibreOffice repo
            folder_map = {
                'en_US': 'en',
                'en_GB': 'en',
                'en_AU': 'en',
                'en_CA': 'en',
                'es_ES': 'es',
                'fr_FR': 'fr_FR',
                'de_DE': 'de',
                'it_IT': 'it_IT',
                'pt_BR': 'pt_BR',
                'pt_PT': 'pt_PT',
                'ru_RU': 'ru_RU',
                'ar': 'ar',
                'tr_TR': 'tr_TR',
                'nl_NL': 'nl_NL',
                'pl_PL': 'pl_PL',
                'cs_CZ': 'cs_CZ',
                'da_DK': 'da_DK',
                'sv_SE': 'sv_SE',
                'nb_NO': 'no',  # Norwegian Bokmål is in 'no' folder
                'el_GR': 'el_GR',
                'hu_HU': 'hu_HU',
                'ro': 'ro',
                'hr_HR': 'hr_HR',
                'sk_SK': 'sk_SK',
                'sl_SI': 'sl_SI',
                'uk_UA': 'uk_UA',
            }
            
            folder_name = folder_map.get(code, code.split('_')[0] if '_' in code else code)
            
            # Some dictionaries have different filenames than their language codes
            file_name_map = {
                'fr_FR': 'fr',  # French uses fr.dic not fr_FR.dic
                'nb_NO': 'nb_NO',  # Norwegian Bokmål
                'el_GR': 'el_GR',  # Greek
            }
            file_name = file_name_map.get(code, code)
            
            # Destination
            base = os.path.dirname(os.path.abspath(__file__))
            target_dir = os.path.join(base, "dicts")
            os.makedirs(target_dir, exist_ok=True)
            
            for i, ext in enumerate(['dic', 'aff']):
                # Construct URL for LibreOffice dictionaries repository
                url = f"https://raw.githubusercontent.com/LibreOffice/dictionaries/master/{folder_name}/{file_name}.{ext}"
                path = os.path.join(target_dir, f"{code}.{ext}")
                
                # Use a proper User-Agent to avoid being blocked
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req) as response:
                    data = response.read()
                    with open(path, 'wb') as out_file:
                        out_file.write(data)
                
                if self.progress:
                    self.progress.setValue(50 if i == 0 else 100)
            
            QTimer.singleShot(0, lambda: self.on_finished(True, code))
            
        except Exception as e:
            print(f"Download Error: {e}")
            QTimer.singleShot(0, lambda: self.on_finished(False, str(e)))

    def on_finished(self, success, msg):
        # Hide progress dialog safely
        if self.progress:
            try:
                self.progress.hide()
                self.progress.close()
                self.progress = None
            except:
                pass
        
        if success:
            # Reload enchant broker to pick up new dictionaries
            try:
                import importlib
                importlib.reload(enchant)
            except:
                pass
            
            QMessageBox.information(self, "Success", 
                f"Dictionary installed successfully! It should now be available in the language dropdown.")
            self.accept()
        else:
            QMessageBox.critical(self, "Error", f"Failed to download dictionary:\n{msg}")

class SyncTranscriptDialog(QDialog):
    def __init__(self, current_time_str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Sync Transcript With Current Time")
        self.setFixedSize(450, 300)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Header Info
        header = QLabel("📌 Synchronize Timecodes")
        header.setStyleSheet("font-size: 16px; font-weight: bold; color: #3b82f6;")
        layout.addWidget(header)
        
        info = QLabel(f"Aligning transcript to: <b>{current_time_str}</b><br><br>"
                     "The first timecode found in your chosen scope will be updated to "
                     "match this moment. All subsequent timecodes will then <b>ripple</b> "
                     "by the same offset to maintain their relative positions.<br><br>"
                     "<i>Your original formatting (colors, bolding) will be preserved.</i>")
        info.setWordWrap(True)
        layout.addWidget(info)
        
        # Options Group
        self.card = PremiumCard()
        card_layout = QVBoxLayout(self.card)
        
        self.opt_all = QCheckBox("Ripple Entire Transcript")
        self.opt_all.setChecked(True)
        self.opt_cursor = QCheckBox("Ripple From Cursor Onward")
        
        # Exclusive check
        self.opt_all.toggled.connect(lambda chk: self.opt_cursor.setChecked(not chk) if chk else None)
        self.opt_cursor.toggled.connect(lambda chk: self.opt_all.setChecked(not chk) if chk else None)
        
        card_layout.addWidget(self.opt_all)
        card_layout.addWidget(self.opt_cursor)
        layout.addWidget(self.card)
        
        # Buttons
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        
    def get_options(self):
        return {
            'scope': 'all' if self.opt_all.isChecked() else 'cursor'
        }
