import os
import time
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PyQt5.QtWidgets import (QApplication, QMainWindow, QDialog, QVBoxLayout, QLineEdit, QLabel, QHBoxLayout, QPushButton,
                             QMessageBox, QSpinBox, QFileDialog, QCheckBox, QAction, QMenu)
from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QStyle, QSystemTrayIcon
from configparser import ConfigParser
from plyer import notification


class SettingsDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Settings")
        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        self.folder_path_input = QLineEdit(self)
        self.log_file_path_input = QLineEdit(self)
        self.move_folder_path_input = QLineEdit(self)
        self.move_log_file_path_input = QLineEdit(self)
        self.monitor_interval_input = QSpinBox(self)
        self.monitor_interval_input.setRange(1, 86400)
        self.move_delay_input = QSpinBox(self)
        self.move_delay_input.setRange(0, 1440)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Folder Path:"))
        layout.addWidget(self.folder_path_input)
        browse_folder_btn = QPushButton("Browse...")
        browse_folder_btn.clicked.connect(self.browse_folder)
        layout.addWidget(browse_folder_btn)

        layout.addWidget(QLabel("Log File Path:"))
        layout.addWidget(self.log_file_path_input)
        browse_log_file_btn = QPushButton("Browse...")
        browse_log_file_btn.clicked.connect(self.browse_log_file)
        layout.addWidget(browse_log_file_btn)

        layout.addWidget(QLabel("Move Folder Path:"))
        layout.addWidget(self.move_folder_path_input)
        browse_move_folder_btn = QPushButton("Browse...")
        browse_move_folder_btn.clicked.connect(self.browse_move_folder)
        layout.addWidget(browse_move_folder_btn)

        layout.addWidget(QLabel("Monitor Interval (seconds):"))
        layout.addWidget(self.monitor_interval_input)

        layout.addWidget(QLabel("Move Delay (seconds):"))
        layout.addWidget(self.move_delay_input)

        buttons_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_settings)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        layout.addLayout(buttons_layout)

        self.setLayout(layout)

    def browse_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.folder_path_input.setText(folder_path)

    def browse_log_file(self):
        log_file_path, _ = QFileDialog.getSaveFileName(self, "Select Log File", "", "Log Files (*.log);;All Files (*)")
        if log_file_path:
            self.log_file_path_input.setText(log_file_path)

    def browse_move_folder(self):
        move_folder_path = QFileDialog.getExistingDirectory(self, "Select Move Folder")
        if move_folder_path:
            self.move_folder_path_input.setText(move_folder_path)

    def load_settings(self):
        config = ConfigParser()
        config.read('settings.ini')
        if 'SETTINGS' in config:
            self.folder_path_input.setText(config['SETTINGS'].get('folder_path', ''))
            self.log_file_path_input.setText(config['SETTINGS'].get('log_file_path', ''))
            self.move_folder_path_input.setText(config['SETTINGS'].get('move_folder_path', ''))
            self.monitor_interval_input.setValue(config['SETTINGS'].getint('monitor_interval', 10))
            self.move_delay_input.setValue(config['SETTINGS'].getint('move_delay', 5))

    def save_settings(self):
        config = ConfigParser()
        config['SETTINGS'] = {
            'folder_path': self.folder_path_input.text(),
            'log_file_path': self.log_file_path_input.text(),
            'move_folder_path': self.move_folder_path_input.text(),
            'monitor_interval': self.monitor_interval_input.value(),
            'move_delay': self.move_delay_input.value()
        }
        with open('settings.ini', 'w') as configfile:
            config.write(configfile)
        self.accept()


class FolderMonitor(QMainWindow):
    file_dropped_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.initUI()
        self.load_settings()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_for_files)
        self.observer = None
        self.last_event_time = time.time()

        self.file_dropped_signal.connect(self.on_file_dropped)

    def initUI(self):
        self.setWindowTitle("Folder Monitor")
        self.setGeometry(200, 200, 400, 300)
        self.settings_action = QAction("Settings", self)
        self.settings_action.triggered.connect(self.show_settings)

        self.start_action = QAction("Start Monitoring", self)
        self.start_action.triggered.connect(self.start_monitoring)

        self.stop_action = QAction("Stop Monitoring", self)
        self.stop_action.triggered.connect(self.stop_monitoring)
        self.stop_action.setDisabled(True)

        menubar = self.menuBar()
        file_menu = menubar.addMenu("File")
        file_menu.addAction(self.settings_action)
        file_menu.addAction(self.start_action)
        file_menu.addAction(self.stop_action)

    def load_settings(self):
        config = ConfigParser()
        config.read('settings.ini')
        if 'SETTINGS' in config:
            self.folder_path = config['SETTINGS'].get('folder_path', '')
            self.log_file_path = config['SETTINGS'].get('log_file_path', '')
            self.move_folder_path = config['SETTINGS'].get('move_folder_path', '')
            self.monitor_interval = config['SETTINGS'].getint('monitor_interval', 10)
            self.move_delay = config['SETTINGS'].getint('move_delay', 5)

    def show_settings(self):
        dialog = SettingsDialog()
        if dialog.exec_() == QDialog.Accepted:
            self.load_settings()

    def start_monitoring(self):
        if not os.path.exists(self.folder_path):
            QMessageBox.warning(self, "Error", "Folder path does not exist!")
            return

        self.observer = Observer()
        event_handler = FileSystemEventHandler()
        event_handler.on_created = self.on_created
        self.observer.schedule(event_handler, self.folder_path, recursive=False)
        self.observer.start()

        self.start_action.setDisabled(True)
        self.stop_action.setDisabled(False)
        self.timer.start(self.monitor_interval * 1000)

    def stop_monitoring(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()

        self.timer.stop()
        self.start_action.setDisabled(False)
        self.stop_action.setDisabled(True)

    def on_created(self, event):
        if event.is_directory:
            return

        file_path = event.src_path
        self.file_dropped_signal.emit(file_path)

    def on_file_dropped(self, file_path):
        self.log_event(f"File dropped: {file_path}")
        time.sleep(self.move_delay)
        if self.move_folder_path:
            try:
                shutil.move(file_path, os.path.join(self.move_folder_path, os.path.basename(file_path)))
                self.log_event(f"File moved: {file_path}")
            except Exception as e:
                self.log_event(f"Error moving file: {str(e)}")

    def check_for_files(self):
        if time.time() - self.last_event_time > self.monitor_interval:
            notification.notify(
                title="No Files Detected",
                message=f"No files were dropped in the last {self.monitor_interval} seconds.",
                timeout=5
            )

    def log_event(self, message):
        print(message)
        if self.log_file_path:
            with open(self.log_file_path, "a") as log_file:
                log_file.write(f"{time.ctime()} - {message}\n")


if __name__ == "__main__":
    app = QApplication([])
    window = FolderMonitor()
    window.show()
    app.exec_()
