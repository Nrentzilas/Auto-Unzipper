import os
import time
import subprocess
import sys
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, 
                            QWidget, QLabel, QFileDialog, QCheckBox, QListWidget,
                            QSpinBox, QLineEdit, QGroupBox, QFormLayout)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer

class ExtractorThread(QThread):
    extract_complete = pyqtSignal(str, bool)
    log_message = pyqtSignal(str)
    
    def __init__(self, zip_file, extract_path, seven_zip_path, delete_after=False):
        super().__init__()
        self.zip_file = zip_file
        self.extract_path = extract_path
        self.seven_zip_path = seven_zip_path
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
        self.seven_zip_path = self.settings.value("seven_zip_path", r"C:\Program Files\7-Zip\7z.exe")
        self.auto_delete = self.settings.value("auto_delete", "false") == "true"
        self.monitor_interval = int(self.settings.value("monitor_interval", 10))
        self.auto_start_monitoring = self.settings.value("auto_start_monitoring", "false") == "true"
        
        self.supported_extensions = [".zip", ".rar", ".7z"]
        
        self.setWindowTitle("Auto Unzip")
        self.setMinimumSize(600, 500)
        
        main_layout = QVBoxLayout()
        
        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout()
        
        self.status_label = QLabel("Monitoring: Stopped")
        status_layout.addWidget(self.status_label)
        
        control_layout = QHBoxLayout()
        self.start_button = QPushButton("Start Monitoring")
        self.start_button.clicked.connect(self.start_monitoring)
        self.stop_button = QPushButton("Stop Monitoring")
        self.stop_button.clicked.connect(self.stop_monitoring)
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
        
        self.seven_zip_input = QLineEdit(self.seven_zip_path)
        browse_7zip_button = QPushButton("Browse...")
        browse_7zip_button.clicked.connect(lambda: self.browse_file(self.seven_zip_input))
        seven_zip_layout = QHBoxLayout()
        seven_zip_layout.addWidget(self.seven_zip_input)
        seven_zip_layout.addWidget(browse_7zip_button)
        settings_layout.addRow("7-Zip executable:", seven_zip_layout)
        
        self.monitor_interval_spinner = QSpinBox()
        self.monitor_interval_spinner.setRange(1, 3600)
        self.monitor_interval_spinner.setValue(self.monitor_interval)
        self.monitor_interval_spinner.setSuffix(" seconds")
        settings_layout.addRow("Check interval:", self.monitor_interval_spinner)
        
        self.auto_delete_checkbox = QCheckBox()
        self.auto_delete_checkbox.setChecked(self.auto_delete)
        settings_layout.addRow("Delete archives after extraction:", self.auto_delete_checkbox)
        
        self.auto_start_checkbox = QCheckBox()
        self.auto_start_checkbox.setChecked(self.auto_start_monitoring)
        settings_layout.addRow("Auto-start monitoring on launch:", self.auto_start_checkbox)
        
        save_settings_button = QPushButton("Save Settings")
        save_settings_button.clicked.connect(self.save_settings)
        settings_layout.addRow("", save_settings_button)
        
        settings_group.setLayout(settings_layout)
        main_layout.addWidget(settings_group)
        
        log_group = QGroupBox("Activity Log")
        log_layout = QVBoxLayout()
        
        self.log_list = QListWidget()
        log_layout.addWidget(self.log_list)
        
        clear_log_button = QPushButton("Clear Log")
        clear_log_button.clicked.connect(self.log_list.clear)
        log_layout.addWidget(clear_log_button)
        
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        main_widget = QWidget()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        self.monitor_thread = None
        self.extractor_threads = []
        
        self.add_log("Application started")
        
        if self.auto_start_monitoring:
            QTimer.singleShot(1000, self.start_monitoring)
    
    def add_log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_list.addItem(f"[{timestamp}] {message}")
        self.log_list.scrollToBottom()
    
    def browse_folder(self, line_edit):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder", line_edit.text())
        if folder:
            line_edit.setText(folder)
    
    def browse_file(self, line_edit):
        file, _ = QFileDialog.getOpenFileName(self, "Select File", line_edit.text())
        if file:
            line_edit.setText(file)
    
    def save_settings(self):
        self.downloads_folder = self.downloads_path_input.text()
        self.extract_folder = self.extract_path_input.text()
        self.seven_zip_path = self.seven_zip_input.text()
        self.auto_delete = self.auto_delete_checkbox.isChecked()
        self.monitor_interval = self.monitor_interval_spinner.value()
        self.auto_start_monitoring = self.auto_start_checkbox.isChecked()
        
        self.settings.setValue("downloads_folder", self.downloads_folder)
        self.settings.setValue("extract_folder", self.extract_folder)
        self.settings.setValue("seven_zip_path", self.seven_zip_path)
        self.settings.setValue("auto_delete", str(self.auto_delete).lower())
        self.settings.setValue("monitor_interval", self.monitor_interval)
        self.settings.setValue("auto_start_monitoring", str(self.auto_start_monitoring).lower())
        
        self.add_log("Settings saved")
    
    def start_monitoring(self):
        os.makedirs(self.extract_folder, exist_ok=True)
        
        self.monitor_thread = MonitorThread(self.downloads_folder, self.supported_extensions, self.monitor_interval)
        self.monitor_thread.new_file_found.connect(self.handle_new_file)
        self.monitor_thread.log_message.connect(self.add_log)
        self.monitor_thread.start()
        
        self.status_label.setText("Monitoring: Active")
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.add_log(f"Started monitoring {self.downloads_folder} for archives")
    
    def stop_monitoring(self):
        if self.monitor_thread and self.monitor_thread.isRunning():
            self.monitor_thread.stop()
            self.monitor_thread.wait()
            
        self.status_label.setText("Monitoring: Stopped")
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.add_log("Stopped monitoring")
    
    def handle_new_file(self, file_path):
        self.add_log(f"New file detected: {file_path}")
        
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        extract_path = os.path.join(self.extract_folder, file_name)
        os.makedirs(extract_path, exist_ok=True)
        
        extractor = ExtractorThread(file_path, extract_path, self.seven_zip_path, self.auto_delete)
        extractor.extract_complete.connect(self.extraction_finished)
        extractor.log_message.connect(self.add_log)
        extractor.start()
        
        self.extractor_threads.append(extractor)
    
    def extraction_finished(self, file_path, success):
        self.add_log(f"Extraction {'completed' if success else 'failed'} for {file_path}")
        
        self.extractor_threads = [t for t in self.extractor_threads if t.isRunning()]
    
    def closeEvent(self, event):
        self.stop_monitoring()
        for thread in self.extractor_threads:
            if thread.isRunning():
                thread.wait(1000)
        super().closeEvent(event)

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()