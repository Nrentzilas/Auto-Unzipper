import os
import time
import subprocess
import sys
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, 
                            QWidget, QLabel, QFileDialog, QCheckBox, QListWidget,
                            QSpinBox, QLineEdit, QGroupBox, QFormLayout, QSystemTrayIcon, QMenu,
                            QGridLayout, QTabWidget, QTextBrowser)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer, QEvent
from PyQt6.QtGui import QIcon, QAction, QFont, QColor, QPalette, QDesktopServices
from PyQt6.QtCore import QUrl

class ExtractorThread(QThread):
    extract_complete = pyqtSignal(str, bool)
    log_message = pyqtSignal(str)
    
    def __init__(self, zip_file, extract_path, delete_after=False):
        super().__init__()
        self.zip_file = zip_file
        self.extract_path = extract_path
        self.seven_zip_path = r"C:\Program Files\7-Zip\7z.exe"
        self.delete_after = delete_after
        
    def run(self):
        try:
            self.log_message.emit(f"Extracting: {self.zip_file} -> {self.extract_path}")
            
            process = subprocess.run(
                [self.seven_zip_path, "x", self.zip_file, f"-o{self.extract_path}", "-y"], 
                shell=True, 
                capture_output=True,
                text=True
            )
            
            if process.returncode == 0:
                if self.delete_after:
                    os.remove(self.zip_file)
                    self.log_message.emit(f"Deleted original file: {self.zip_file}")
                self.extract_complete.emit(self.zip_file, True)
            else:
                self.log_message.emit(f"Error extracting {self.zip_file}: {process.stderr}")
                self.extract_complete.emit(self.zip_file, False)
        except Exception as e:
            self.log_message.emit(f"Exception during extraction: {str(e)}")
            self.extract_complete.emit(self.zip_file, False)

class MonitorThread(QThread):
    new_file_found = pyqtSignal(str)
    log_message = pyqtSignal(str)
    
    def __init__(self, folder_to_monitor, file_extensions, interval):
        super().__init__()
        self.folder_to_monitor = folder_to_monitor
        self.file_extensions = file_extensions
        self.interval = interval
        self.running = True
        self.processed_files = set()
        
    def run(self):
        self.log_message.emit(f"Monitoring started for {self.folder_to_monitor}")
        while self.running:
            try:
                if not os.path.exists(self.folder_to_monitor):
                    self.log_message.emit(f"Warning: Folder {self.folder_to_monitor} does not exist!")
                    time.sleep(self.interval)
                    continue
                    
                files = [f for f in os.listdir(self.folder_to_monitor) 
                        if any(f.lower().endswith(ext.lower()) for ext in self.file_extensions)]
                
                for file in files:
                    file_path = os.path.join(self.folder_to_monitor, file)
                    if file_path not in self.processed_files:
                        self.new_file_found.emit(file_path)
                        self.processed_files.add(file_path)
                
                time.sleep(self.interval)
            except Exception as e:
                self.log_message.emit(f"Error in monitoring thread: {str(e)}")
                time.sleep(self.interval)
    
    def stop(self):
        self.running = False
        self.log_message.emit("Monitoring stopped")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.settings = QSettings("AutoUnzip", "Settings")
        
        self.downloads_folder = self.settings.value("downloads_folder", os.path.expanduser("~\\Downloads"))
        self.extract_folder = self.settings.value("extract_folder", os.path.join(self.downloads_folder, "Extracted"))
        self.auto_delete = self.settings.value("auto_delete", "false") == "true"
        self.monitor_interval = int(self.settings.value("monitor_interval", 10))
        self.auto_start_monitoring = self.settings.value("auto_start_monitoring", "false") == "true"
        self.dark_mode = self.settings.value("dark_mode", "false") == "true"
        
        self.supported_extensions = [".zip", ".rar", ".7z"]
        
        # Set application and window icons
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "icon.ico")
        self.setWindowIcon(QIcon(icon_path))
        app = QApplication.instance()
        app.setWindowIcon(QIcon(icon_path))
        
        self.setup_ui()
        self.setup_system_tray()
        
        self.monitor_thread = None
        self.extractor_threads = []
        
        self.apply_theme()
        self.add_log("Application started")
        
        if self.auto_start_monitoring:
            QTimer.singleShot(1000, self.start_monitoring)
    
    def apply_theme(self):
        app = QApplication.instance()
        app.setStyle("Fusion")
        
        if self.dark_mode:
            dark_palette = QPalette()
            dark_color = QColor(45, 45, 45)
            disabled_color = QColor(127, 127, 127)
            text_color = QColor(210, 210, 210)
            highlight_color = QColor(42, 130, 218)
            
            dark_palette.setColor(QPalette.ColorRole.Window, dark_color)
            dark_palette.setColor(QPalette.ColorRole.WindowText, text_color)
            dark_palette.setColor(QPalette.ColorRole.Base, QColor(18, 18, 18))
            dark_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ColorRole.ToolTipBase, dark_color)
            dark_palette.setColor(QPalette.ColorRole.ToolTipText, text_color)
            dark_palette.setColor(QPalette.ColorRole.Text, text_color)
            dark_palette.setColor(QPalette.ColorRole.Button, dark_color)
            dark_palette.setColor(QPalette.ColorRole.ButtonText, text_color)
            dark_palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
            dark_palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
            dark_palette.setColor(QPalette.ColorRole.Highlight, highlight_color)
            dark_palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
            dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, disabled_color)
            dark_palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, disabled_color)
            
            app.setPalette(dark_palette)
            
            button_style = """
                QPushButton {
                    border: 1px solid #555555;
                    border-radius: 4px;
                    padding: 6px 12px;
                    background-color: #444444;
                    color: #E0E0E0;
                }
                QPushButton:hover {
                    background-color: #555555;
                }
                QPushButton:pressed {
                    background-color: #666666;
                }
                QPushButton:disabled {
                    background-color: #333333;
                    color: #777777;
                }
            """
            
            self.start_button.setStyleSheet("""
                QPushButton {
                    border: 1px solid #2e7d32;
                    border-radius: 4px;
                    padding: 6px 12px;
                    background-color: #43a047;
                    color: white;
                }
                QPushButton:hover {
                    background-color: #388e3c;
                }
                QPushButton:pressed {
                    background-color: #2e7d32;
                }
                QPushButton:disabled {
                    background-color: #388e3c;
                    color: #dddddd;
                }
            """)
            
            self.stop_button.setStyleSheet("""
                QPushButton {
                    border: 1px solid #c62828;
                    border-radius: 4px;
                    padding: 6px 12px;
                    background-color: #d32f2f;
                    color: white;
                }
                QPushButton:hover {
                    background-color: #c62828;
                }
                QPushButton:pressed {
                    background-color: #b71c1c;
                }
                QPushButton:disabled {
                    background-color: #c62828;
                    color: #dddddd;
                }
            """)
            
            groupbox_style = """
                QGroupBox {
                    border: 1px solid #555555;
                    border-radius: 4px;
                    margin-top: 1.5ex;
                    font-weight: bold;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    padding: 0 5px;
                    color: #E0E0E0;
                }
            """
        else:
            palette = QPalette()
            base_color = QColor(245, 245, 245) 
            text_color = QColor(50, 50, 50)    
            accent_color = QColor(70, 130, 180) 
            
            palette.setColor(QPalette.ColorRole.Window, base_color)
            palette.setColor(QPalette.ColorRole.WindowText, text_color)
            palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(235, 235, 235))
            palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(255, 255, 255))
            palette.setColor(QPalette.ColorRole.ToolTipText, text_color)
            palette.setColor(QPalette.ColorRole.Text, text_color)
            palette.setColor(QPalette.ColorRole.Button, base_color)
            palette.setColor(QPalette.ColorRole.ButtonText, text_color)
            palette.setColor(QPalette.ColorRole.BrightText, QColor(0, 0, 0))
            palette.setColor(QPalette.ColorRole.Link, accent_color)
            palette.setColor(QPalette.ColorRole.Highlight, accent_color)
            palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
            
            app.setPalette(palette)
            
            button_style = """
                QPushButton {
                    border: 1px solid #C0C0C0;
                    border-radius: 4px;
                    padding: 6px 12px;
                    background-color: #F5F5F5;
                    color: #333333;
                }
                QPushButton:hover {
                    background-color: #E0E0E0;
                }
                QPushButton:pressed {
                    background-color: #D0D0D0;
                }
                QPushButton:disabled {
                    background-color: #F0F0F0;
                    color: #A0A0A0;
                }
            """
            
            self.start_button.setStyleSheet("""
                QPushButton {
                    border: 1px solid #4CAF50;
                    border-radius: 4px;
                    padding: 6px 12px;
                    background-color: #4CAF50;
                    color: white;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
                QPushButton:pressed {
                    background-color: #3d8b40;
                }
                QPushButton:disabled {
                    background-color: #a5d6a7;
                    color: #f5f5f5;
                }
            """)
            
            self.stop_button.setStyleSheet("""
                QPushButton {
                    border: 1px solid #f44336;
                    border-radius: 4px;
                    padding: 6px 12px;
                    background-color: #f44336;
                    color: white;
                }
                QPushButton:hover {
                    background-color: #e53935;
                }
                QPushButton:pressed {
                    background-color: #d32f2f;
                }
                QPushButton:disabled {
                    background-color: #ef9a9a;
                    color: #f5f5f5;
                }
            """)
            
            groupbox_style = """
                QGroupBox {
                    border: 1px solid #C0C0C0;
                    border-radius: 4px;
                    margin-top: 1.5ex;
                    font-weight: bold;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    padding: 0 5px;
                }
            """
        
        for widget in self.findChildren(QPushButton):
            if widget not in [self.start_button, self.stop_button, self.theme_toggle_button]:
                widget.setStyleSheet(button_style)
        
        for widget in self.findChildren(QGroupBox):
            widget.setStyleSheet(groupbox_style)
            
        if hasattr(self, 'theme_toggle_button'):
            self.theme_toggle_button.setText("Light Theme" if self.dark_mode else "Dark Theme")

    def setup_ui(self):
        self.setWindowTitle("Auto Unzip v.1.1.0")
        self.setMinimumSize(650, 550)
        
        self.tab_widget = QTabWidget()
        
        main_tab = QWidget()
        main_layout = QVBoxLayout()
        
        theme_layout = QHBoxLayout()
        self.theme_toggle_button = QPushButton("Dark Theme") 
        self.theme_toggle_button.clicked.connect(self.toggle_theme)
        theme_layout.addStretch()
        theme_layout.addWidget(self.theme_toggle_button)
        main_layout.addLayout(theme_layout)
        
        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout()
        
        self.status_label = QLabel("Monitoring: Stopped")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("font-size: 14px; font-weight: bold; padding: 5px;")
        status_layout.addWidget(self.status_label)
        
        control_layout = QHBoxLayout()
        self.start_button = QPushButton("Start Monitoring")
        self.stop_button = QPushButton("Stop Monitoring")
        self.stop_button.setEnabled(False)
        control_layout.addWidget(self.start_button)
        control_layout.addWidget(self.stop_button)
        status_layout.addLayout(control_layout)
        
        status_group.setLayout(status_layout)
        main_layout.addWidget(status_group)
        
        settings_group = QGroupBox("Settings")
        settings_layout = QFormLayout()
        
        self.downloads_path_input = QLineEdit(self.downloads_folder)
        browse_downloads_button = QPushButton("Browse...")
        browse_downloads_button.clicked.connect(lambda: self.browse_folder(self.downloads_path_input))
        downloads_path_layout = QHBoxLayout()
        downloads_path_layout.addWidget(self.downloads_path_input)
        downloads_path_layout.addWidget(browse_downloads_button)
        settings_layout.addRow("Monitor folder:", downloads_path_layout)
        
        self.extract_path_input = QLineEdit(self.extract_folder)
        browse_extract_button = QPushButton("Browse...")
        browse_extract_button.clicked.connect(lambda: self.browse_folder(self.extract_path_input))
        extract_path_layout = QHBoxLayout()
        extract_path_layout.addWidget(self.extract_path_input)
        extract_path_layout.addWidget(browse_extract_button)
        settings_layout.addRow("Extract to:", extract_path_layout)
        
        self.monitor_interval_spinner = QSpinBox()
        self.monitor_interval_spinner.setRange(1, 3600)
        self.monitor_interval_spinner.setValue(self.monitor_interval)
        self.monitor_interval_spinner.setSuffix(" seconds")
        settings_layout.addRow("Check interval:", self.monitor_interval_spinner)
        
        checkbox_container = QWidget()
        checkbox_layout = QGridLayout(checkbox_container)
        checkbox_layout.setColumnStretch(0, 1)
        checkbox_layout.setColumnStretch(1, 1)
        
        self.auto_delete_checkbox = QCheckBox("Delete archives after extraction")
        self.auto_delete_checkbox.setChecked(self.auto_delete)
        checkbox_layout.addWidget(self.auto_delete_checkbox, 0, 0)
        
        self.auto_start_checkbox = QCheckBox("Auto-start monitoring on launch")
        self.auto_start_checkbox.setChecked(self.auto_start_monitoring)
        checkbox_layout.addWidget(self.auto_start_checkbox, 0, 1)
        
        self.minimize_to_tray_checkbox = QCheckBox("Minimize to system tray")
        self.minimize_to_tray_checkbox.setChecked(self.settings.value("minimize_to_tray", "true") == "true")
        checkbox_layout.addWidget(self.minimize_to_tray_checkbox, 1, 0)
        
        settings_layout.addRow("Options:", checkbox_container)
        
        save_settings_button = QPushButton("Save Settings")
        save_settings_button.clicked.connect(self.save_settings)
        settings_layout.addRow("", save_settings_button)
        
        settings_group.setLayout(settings_layout)
        main_layout.addWidget(settings_group)
        
        log_group = QGroupBox("Activity Log")
        log_layout = QVBoxLayout()
        
        self.log_list = QListWidget()
        self.log_list.setAlternatingRowColors(True)
        log_layout.addWidget(self.log_list)
        
        clear_log_button = QPushButton("Clear Log")
        clear_log_button.clicked.connect(self.log_list.clear)
        log_layout.addWidget(clear_log_button)
        
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        main_tab.setLayout(main_layout)
        
        about_tab = QWidget()
        about_layout = QVBoxLayout()
        
        about_text = QTextBrowser()
        about_text.setOpenExternalLinks(True)
        about_text.setHtml(f"""
        <div style="text-align: center;">
            <h1>Auto Unzip</h1>
            <h2>Version 1.1.0</h2>
            <p>Automatically extracts and organizes archive files from your downloads folder.</p>
            <p>Supports ZIP, RAR, and 7Z formats.</p>
            <br>
            <h2>Created by <b>Nrentzilas</b></h2>
            <br>
            <p>This tool monitors your downloads folder for new archive files 
            and automatically extracts them to a designated folder.Feel free to star the repo as it help continue developing</p>
        </div>
        """)
        
        about_layout.addWidget(about_text)
        
        github_button = QPushButton("Visit GitHub Repository")
        github_button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl("https://github.com/Nrentzilas/Auto-Unzipper")))
        
        about_layout.addWidget(github_button)
        about_tab.setLayout(about_layout)
        
        self.tab_widget.addTab(main_tab, "Main")
        self.tab_widget.addTab(about_tab, "About")
        
        self.setCentralWidget(self.tab_widget)
        
        self.start_button.clicked.connect(self.start_monitoring)
        self.stop_button.clicked.connect(self.stop_monitoring)
    
    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.settings.setValue("dark_mode", str(self.dark_mode).lower())
        self.apply_theme()
        self.add_log(f"Switched to {'dark' if self.dark_mode else 'light'} theme")
    
    def setup_system_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "icon.ico")
        self.tray_icon.setIcon(QIcon(icon_path))
        self.tray_icon.setToolTip("Auto Unzip")
        
        tray_menu = QMenu()
        
        show_action = QAction("Show", self)
        show_action.triggered.connect(self.show_from_tray)
        
        hide_action = QAction("Hide", self)
        hide_action.triggered.connect(self.hide)
        
        start_action = QAction("Start Monitoring", self)
        start_action.triggered.connect(self.start_monitoring)
        
        stop_action = QAction("Stop Monitoring", self)
        stop_action.triggered.connect(self.stop_monitoring)
        
        toggle_theme_action = QAction("Toggle Theme", self)
        toggle_theme_action.triggered.connect(self.toggle_theme)
        
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close_application)
        
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addSeparator()
        tray_menu.addAction(start_action)
        tray_menu.addAction(stop_action)
        tray_menu.addSeparator()
        tray_menu.addAction(toggle_theme_action)
        tray_menu.addSeparator()
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()
    
    def show_from_tray(self):
        self.showNormal()  
        self.activateWindow()
    
    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show_from_tray()
    
    def add_log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_list.addItem(f"[{timestamp}] {message}")
        self.log_list.scrollToBottom()
    
    def browse_folder(self, line_edit):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder", line_edit.text())
        if folder:
            line_edit.setText(folder)
    
    def save_settings(self):
        self.downloads_folder = self.downloads_path_input.text()
        self.extract_folder = self.extract_path_input.text()
        self.auto_delete = self.auto_delete_checkbox.isChecked()
        self.monitor_interval = self.monitor_interval_spinner.value()
        self.auto_start_monitoring = self.auto_start_checkbox.isChecked()
        minimize_to_tray = self.minimize_to_tray_checkbox.isChecked()
        
        self.settings.setValue("downloads_folder", self.downloads_folder)
        self.settings.setValue("extract_folder", self.extract_folder)
        self.settings.setValue("auto_delete", str(self.auto_delete).lower())
        self.settings.setValue("monitor_interval", self.monitor_interval)
        self.settings.setValue("auto_start_monitoring", str(self.auto_start_monitoring).lower())
        self.settings.setValue("minimize_to_tray", str(minimize_to_tray).lower())
        
        self.add_log("Settings saved")
    
    def start_monitoring(self):
        os.makedirs(self.extract_folder, exist_ok=True)
        
        self.monitor_thread = MonitorThread(self.downloads_folder, self.supported_extensions, self.monitor_interval)
        self.monitor_thread.new_file_found.connect(self.handle_new_file)
        self.monitor_thread.log_message.connect(self.add_log)
        self.monitor_thread.start()
        
        self.status_label.setText("Monitoring: Active")
        self.status_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #4CAF50; padding: 5px;")
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.add_log(f"Started monitoring {self.downloads_folder} for archives")
    
    def stop_monitoring(self):
        if self.monitor_thread and self.monitor_thread.isRunning():
            self.monitor_thread.stop()
            self.monitor_thread.wait()
            
        self.status_label.setText("Monitoring: Stopped")
        self.status_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #f44336; padding: 5px;")
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.add_log("Stopped monitoring")
    
    def handle_new_file(self, file_path):
        self.add_log(f"New file detected: {file_path}")
        
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        extract_path = os.path.join(self.extract_folder, file_name)
        os.makedirs(extract_path, exist_ok=True)
        
        extractor = ExtractorThread(file_path, extract_path, self.auto_delete)
        extractor.extract_complete.connect(self.extraction_finished)
        extractor.log_message.connect(self.add_log)
        extractor.start()
        
        self.extractor_threads.append(extractor)
    
    def extraction_finished(self, file_path, success):
        self.add_log(f"Extraction {'completed' if success else 'failed'} for {file_path}")
        
        self.extractor_threads = [t for t in self.extractor_threads if t.isRunning()]
    
    def changeEvent(self, event):
        if event.type() == QEvent.Type.WindowStateChange:
            minimize_to_tray = self.settings.value("minimize_to_tray", "true") == "true"
            if self.isMinimized() and minimize_to_tray:
                QTimer.singleShot(0, self.hide)
                self.tray_icon.show() 
        super().changeEvent(event)
    
    def close_application(self):
        self.close()
    
    def closeEvent(self, event):
        self.stop_monitoring()
        for thread in self.extractor_threads:
            if thread.isRunning():
                thread.wait(1000)
        event.accept()

def main():
    app = QApplication(sys.argv)
    
    if not QSystemTrayIcon.isSystemTrayAvailable():
        print("System tray not available on this system")
        return 1
    
    QApplication.setQuitOnLastWindowClosed(False)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()