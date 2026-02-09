from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                             QListWidget, QListWidgetItem, QPushButton, 
                             QSpinBox, QFileDialog, QMessageBox, QWidget, QMenu)
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QDesktopServices
import os

class BackupSettingsDialog(QDialog):
    def __init__(self, parent, backup_manager):
        super().__init__(parent)
        self.setWindowTitle("Backup Manager")
        self.resize(600, 500)
        self.bm = backup_manager
        self.selected_backup = None
        
        layout = QVBoxLayout(self)
        
        # --- Settings Section ---
        settings_group = QWidget()
        settings_layout = QVBoxLayout(settings_group)
        settings_layout.setContentsMargins(0, 0, 0, 10)
        
        # Row 1: Location
        loc_layout = QHBoxLayout()
        loc_layout.addWidget(QLabel("Backup Location:"))
        self.lbl_path = QLabel(self.bm.get_backup_dir())
        self.lbl_path.setStyleSheet("font-style: italic; color: gray;")
        loc_layout.addWidget(self.lbl_path, stretch=1)
        
        btn_change = QPushButton("Change...")
        btn_change.clicked.connect(self.change_location)
        loc_layout.addWidget(btn_change)
        settings_layout.addLayout(loc_layout)
        
        # Row 2: Interval & Retention
        int_layout = QHBoxLayout()
        int_layout.addWidget(QLabel("Backup every"))
        self.spin_interval = QSpinBox()
        self.spin_interval.setRange(1, 1440) 
        current_interval = 5
        if self.parent() and hasattr(self.parent(), 'config'):
            current_interval = self.parent().config.get('backup_interval', 5)
        self.spin_interval.setValue(current_interval)
        self.spin_interval.valueChanged.connect(self.on_interval_changed)
        int_layout.addWidget(self.spin_interval)
        int_layout.addWidget(QLabel("min."))
        
        int_layout.addSpacing(20)
        int_layout.addWidget(QLabel("Keep"))
        self.spin_per_file = QSpinBox()
        self.spin_per_file.setRange(1, 100)
        current_per_file = 5
        if self.parent() and hasattr(self.parent(), 'config'):
            current_per_file = self.parent().config.get('backups_per_file', 5)
        self.spin_per_file.setValue(current_per_file)
        self.spin_per_file.valueChanged.connect(self.on_per_file_changed)
        int_layout.addWidget(self.spin_per_file)
        int_layout.addWidget(QLabel("per file"))
        
        int_layout.addSpacing(20)
        int_layout.addWidget(QLabel("Delete after"))
        from PyQt6.QtWidgets import QComboBox
        self.combo_retention = QComboBox()
        self.combo_retention.addItems(["NEVER", "1 month", "3 months", "6 months", "12 months", "24 months"])
        
        # Map months to index
        ret_map = {0: 0, 1: 1, 3: 2, 6: 3, 12: 4, 24: 5}
        current_ret = 6
        if self.parent() and hasattr(self.parent(), 'config'):
            current_ret = self.parent().config.get('backup_retention_months', 6)
        self.combo_retention.setCurrentIndex(ret_map.get(current_ret, 3))
        self.combo_retention.currentIndexChanged.connect(self.on_retention_changed)
        int_layout.addWidget(self.combo_retention)
        
        int_layout.addStretch()
        settings_layout.addLayout(int_layout)
        
        layout.addWidget(settings_group)
        
        # --- List Section ---
        layout.addWidget(QLabel("Available Backups (Double-click or Right-click to restore):"))
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_widget.customContextMenuRequested.connect(self.show_context_menu)
        self.list_widget.itemSelectionChanged.connect(self.on_selection_changed)
        self.list_widget.itemDoubleClicked.connect(self.accept) 
        layout.addWidget(self.list_widget)
        
        # --- Controls Section ---
        controls_layout = QHBoxLayout()
        
        self.btn_restore = QPushButton("Restore")
        self.btn_restore.setEnabled(False)
        self.btn_restore.clicked.connect(self.on_restore_clicked)
        controls_layout.addWidget(self.btn_restore)
        
        self.btn_delete_sel = QPushButton("Delete Selected")
        self.btn_delete_sel.setEnabled(False)
        self.btn_delete_sel.clicked.connect(self.delete_selected)
        self.btn_delete_sel.setStyleSheet("color: #d32f2f;")
        controls_layout.addWidget(self.btn_delete_sel)

        self.btn_open_folder = QPushButton("📂 Open Folder")
        self.btn_open_folder.clicked.connect(self.open_backup_folder)
        controls_layout.addWidget(self.btn_open_folder)

        self.btn_clear = QPushButton("Clear All")
        self.btn_clear.clicked.connect(self.clear_backups)
        self.btn_clear.setStyleSheet("color: #d32f2f;") 
        controls_layout.addWidget(self.btn_clear)

        self.btn_backup_now = QPushButton("Backup Now")
        self.btn_backup_now.clicked.connect(self.do_backup_now)
        controls_layout.addWidget(self.btn_backup_now)
        
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.reject)
        controls_layout.addWidget(btn_close)
        
        layout.addLayout(controls_layout)
        
        self.refresh_list()

    def refresh_list(self):
        self.list_widget.clear()
        backups = self.bm.get_backups()
        
        for b in backups:
            size_kb = b['size'] / 1024
            date_str = b['date'].strftime("%Y-%m-%d %H:%M:%S")
            item_text = f"{date_str}  -  {b['filename']}  ({size_kb:.1f} KB)"
            
            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, b)
            self.list_widget.addItem(item)

    def on_selection_changed(self):
        items = self.list_widget.selectedItems()
        count = len(items)
        
        # Restore only for single selection
        self.btn_restore.setEnabled(count == 1)
        if count == 1:
            self.selected_backup = items[0].data(Qt.ItemDataRole.UserRole)
        else:
            self.selected_backup = None
            
        # Delete enabled for 1 or more
        self.btn_delete_sel.setEnabled(count >= 1)

    def show_context_menu(self, pos):
        item = self.list_widget.itemAt(pos)
        selected_count = len(self.list_widget.selectedItems())
        if selected_count == 0:
            return
            
        menu = QMenu(self) # Pass self to inherit stylesheet
        
        if selected_count == 1:
            act_restore = menu.addAction("Restore This Backup")
            act_restore.triggered.connect(self.accept)
            menu.addSeparator()
            
        act_delete = menu.addAction(f"Delete Selected ({selected_count})")
        act_delete.triggered.connect(self.delete_selected)
        
        menu.addSeparator()
        act_open = menu.addAction("Open Backup Folder")
        act_open.triggered.connect(self.open_backup_folder)
        
        menu.exec(self.list_widget.mapToGlobal(pos))

    def delete_selected(self):
        items = self.list_widget.selectedItems()
        if not items:
            return
            
        reply = QMessageBox.question(
            self, "Delete Backups", 
            f"Are you sure you want to delete {len(items)} selected backup(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
            
        deleted_count = 0
        for item in items:
            b = item.data(Qt.ItemDataRole.UserRole)
            try:
                if os.path.exists(b['path']):
                    os.remove(b['path'])
                    deleted_count += 1
            except Exception as e:
                print(f"Error deleting {b['path']}: {e}")
                
        self.refresh_list()
        if deleted_count > 0:
            self.on_selection_changed() # Reset buttons

    def open_backup_folder(self):
        path = self.bm.get_backup_dir()
        if os.path.exists(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def on_selection(self, current, prev):
        # Deprecated by on_selection_changed but kept for compatibility if needed
        pass

    def change_location(self):
        d = QFileDialog.getExistingDirectory(self, "Select Backup Directory", self.bm.get_backup_dir())
        if d:
            self.bm.set_backup_dir(d)
            self.lbl_path.setText(d)
            self.refresh_list()

    def do_backup_now(self):
        # Trigger parent's save functionality but as backup
        if self.parent():
            self.parent().perform_auto_backup(force=True)
            self.refresh_list()
            QMessageBox.information(self, "Backup", "Backup created successfully!")

    def on_interval_changed(self, value):
        if self.parent() and hasattr(self.parent(), 'config'):
            self.parent().config['backup_interval'] = value
            self.parent().save_config()

    def on_per_file_changed(self, value):
        if self.parent() and hasattr(self.parent(), 'config'):
            self.parent().config['backups_per_file'] = value
            self.parent().save_config()

    def on_retention_changed(self, index):
        if self.parent() and hasattr(self.parent(), 'config'):
            # Map index to months
            ret_map = {0: 0, 1: 1, 2: 3, 3: 6, 4: 12, 5: 24}
            self.parent().config['backup_retention_months'] = ret_map.get(index, 6)
            self.parent().save_config()

    def clear_backups(self):
        reply = QMessageBox.question(
            self, "Clear All Backups", 
            "Are you sure you want to delete ALL available backups? This cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            if self.bm.clear_all_backups():
                self.refresh_list()
                QMessageBox.information(self, "Clear Backups", "All backups deleted.")
            else:
                QMessageBox.warning(self, "Clear Backups", "Failed to delete some backups.")

    def on_restore_clicked(self):
        # We wrap accept to ensure the parent knows this was a "Restore" action
        # The main logic is in main.py:show_backups after dlg.exec()
        self.accept()

    def get_selected_backup_path(self):
        if self.selected_backup:
            return self.selected_backup['path']
        return None
