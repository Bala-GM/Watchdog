import os
import sys
import time
import shutil
import win32com.client as win32
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PyQt5.QtWidgets import (QApplication, QMainWindow, QDialog, QVBoxLayout, QLineEdit, QLabel, QHBoxLayout, QPushButton,
                             QMessageBox, QSpinBox, QFileDialog, QCheckBox, QAction, QMenu)
from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QStyle, QSystemTrayIcon
from configparser import ConfigParser
from plyer import notification  # Import plyer for notifications

# Version V-1.0.5 Jan|23|2025

def send_email(subject, body, to_recipients, cc_recipients=None, attachment_paths=None):
    """Sends an email with the specified subject, body, and attachments."""
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)  # 0: olMailItem
    mail.Subject = subject
    mail.Body = body

    # Handle multiple recipients in 'To' field
    if isinstance(to_recipients, list):
        mail.To = '; '.join(to_recipients)
    else:
        mail.To = to_recipients

    # Handle multiple recipients in 'CC' field
    if cc_recipients:
        if isinstance(cc_recipients, list):
            mail.CC = '; '.join(cc_recipients)
        else:
            mail.CC = cc_recipients

    # Add attachments if any
    if attachment_paths:
        for attachment in attachment_paths:
            mail.Attachments.Add(attachment)

    try:
        mail.Send()
        print("Email sent successfully!")
    except Exception as e:
        print("Failed to send email:", str(e))

class SettingsDialog(QDialog):
    """Dialog for configuring application settings."""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Settings")
        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        """Sets up the UI for the settings dialog."""
        self.folder_path_input = QLineEdit(self)
        self.log_file_path_input = QLineEdit(self)
        self.move_folder_path_input = QLineEdit(self)  # Folder Move Path input
        self.move_log_file_path_input = QLineEdit(self)  # Move Log File Path input
        self.monitor_interval_input = QSpinBox(self)
        self.monitor_interval_input.setRange(1, 86400)
        self.extended_monitor_interval_input = QSpinBox(self)
        self.extended_monitor_interval_input.setRange(1, 86400)
        self.notification_duration_input = QSpinBox(self)  # Add notification duration input
        self.notification_duration_input.setRange(1, 3600)
        self.auto_start_monitoring_checkbox = QCheckBox("Start monitoring on application start", self)
        self.move_delay_input = QSpinBox(self)  # Delay for file movement (in minutes)
        self.move_delay_input.setRange(0, 1440)  # Limit from 0 to 1440 minutes (24 hours)

        self.email_subject_input = QLineEdit(self)
        self.email_body_input = QLineEdit(self)
        self.email_to_input = QLineEdit(self)
        self.email_cc_input = QLineEdit(self)
        self.notification_message_input = QLineEdit(self)

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

        layout.addWidget(QLabel("Folder Move Path:"))  # New label for move folder path
        layout.addWidget(self.move_folder_path_input)
        browse_move_folder_btn = QPushButton("Browse...")
        browse_move_folder_btn.clicked.connect(self.browse_move_folder)
        layout.addWidget(browse_move_folder_btn)

        layout.addWidget(QLabel("Move Log File Path:"))  # New label for move log file path
        layout.addWidget(self.move_log_file_path_input)
        browse_move_log_file_btn = QPushButton("Browse...")
        browse_move_log_file_btn.clicked.connect(self.browse_move_log_file)
        layout.addWidget(browse_move_log_file_btn)

        layout.addWidget(QLabel("Monitor Interval (seconds):"))
        layout.addWidget(self.monitor_interval_input)

        layout.addWidget(QLabel("Extended Monitor Interval (seconds):"))
        layout.addWidget(self.extended_monitor_interval_input)

        layout.addWidget(QLabel("Notification Duration (seconds):"))  # Add label for notification duration
        layout.addWidget(self.notification_duration_input)  # Add notification duration input
        
        layout.addWidget(QLabel("File Move Delay (Seconds):")) #File Move Delay (minutes)
        layout.addWidget(self.move_delay_input)

        layout.addWidget(self.auto_start_monitoring_checkbox)

        layout.addWidget(QLabel("Email Subject:"))
        layout.addWidget(self.email_subject_input)

        layout.addWidget(QLabel("Email Body:"))
        layout.addWidget(self.email_body_input)

        layout.addWidget(QLabel("To Recipients:"))
        layout.addWidget(self.email_to_input)

        layout.addWidget(QLabel("CC Recipients:"))
        layout.addWidget(self.email_cc_input)

        layout.addWidget(QLabel("Notification Message:"))
        layout.addWidget(self.notification_message_input)

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
        """Opens a dialog to select a folder."""
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.folder_path_input.setText(folder_path)

    def browse_log_file(self):
        """Opens a dialog to select a log file."""
        log_file_path, _ = QFileDialog.getSaveFileName(self, "Select Log File", "", "Log Files (*.log);;All Files (*)")
        if log_file_path:
            self.log_file_path_input.setText(log_file_path)

    def browse_move_folder(self):
        """Opens a dialog to select the folder where files will be moved."""
        move_folder_path = QFileDialog.getExistingDirectory(self, "Select Move Folder")
        if move_folder_path:
            self.move_folder_path_input.setText(move_folder_path)

    def browse_move_log_file(self):
        """Opens a dialog to select a log file for move logs."""
        move_log_file_path, _ = QFileDialog.getSaveFileName(self, "Select Move Log File", "", "Text Files (*.txt);;All Files (*)")
        if move_log_file_path:
            self.move_log_file_path_input.setText(move_log_file_path)

    def load_settings(self):
        """Loads settings from the configuration file."""
        config = ConfigParser()
        config.read('settings.ini')
        if 'SETTINGS' in config:
            self.folder_path_input.setText(config['SETTINGS'].get('folder_path', ''))
            self.log_file_path_input.setText(config['SETTINGS'].get('log_file_path', ''))
            self.move_folder_path_input.setText(config['SETTINGS'].get('move_folder_path', ''))  # Load Move Folder Path
            self.move_log_file_path_input.setText(config['SETTINGS'].get('move_log_file_path', ''))  # Load Move Log File Path
            self.monitor_interval_input.setValue(config['SETTINGS'].getint('monitor_interval', 120))
            self.extended_monitor_interval_input.setValue(config['SETTINGS'].getint('extended_monitor_interval', 2700))
            self.notification_duration_input.setValue(config['SETTINGS'].getint('notification_duration', 10))  # Load notification duration
            self.move_delay_input.setValue(config['SETTINGS'].getint('move_delay', 30))  # Default delay to 10 minutes to 30 seconds
            self.auto_start_monitoring_checkbox.setChecked(config['SETTINGS'].getboolean('auto_start_monitoring', False))
        if 'EMAIL' in config:
            self.email_subject_input.setText(config['EMAIL'].get('subject', 'No file drop alert'))
            self.email_body_input.setText(config['EMAIL'].get('body', 'No file has been dropped in the monitored folder within the specified interval.'))
            self.email_to_input.setText(config['EMAIL'].get('to', ''))
            self.email_cc_input.setText(config['EMAIL'].get('cc', ''))
        if 'NOTIFICATION' in config:
            self.notification_message_input.setText(config['NOTIFICATION'].get('message', 'No file dropped within the specified interval.'))

    def save_settings(self):
        """Saves the current settings to the configuration file."""
        config = ConfigParser()
        config['SETTINGS'] = {
            'folder_path': self.folder_path_input.text(),
            'log_file_path': self.log_file_path_input.text(),
            'move_folder_path': self.move_folder_path_input.text(),  # Save Move Folder Path
            'move_log_file_path': self.move_log_file_path_input.text(),  # Save Move Log File Path
            'monitor_interval': self.monitor_interval_input.value(),
            'extended_monitor_interval': self.extended_monitor_interval_input.value(),
            'notification_duration': self.notification_duration_input.value(),  # Save notification duration
            'move_delay': self.move_delay_input.value(),  # Save move delay setting
            'auto_start_monitoring': self.auto_start_monitoring_checkbox.isChecked()
        }
        config['EMAIL'] = {
            'subject': self.email_subject_input.text(),
            'body': self.email_body_input.text(),
            'to': self.email_to_input.text(),
            'cc': self.email_cc_input.text()
        }
        config['NOTIFICATION'] = {
            'message': self.notification_message_input.text()
        }
        with open('settings.ini', 'w') as configfile:
            config.write(configfile)
        self.accept()

class MonitorApp(QMainWindow):
    """Main application window for monitoring folder and sending notifications."""
    file_dropped_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.initUI()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_file_drop)
        self.observer = None
        self.notification_count = 0  # To track the number of notifications

        self.file_dropped_signal.connect(self.on_file_dropped)

        self.load_settings()

        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))  # Correct icon
        self.tray_icon.setVisible(True)

        # Add minimize to system tray functionality
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_menu = QMenu(self)
        self.restore_action = QAction("Restore", self)
        self.restore_action.triggered.connect(self.show)
        self.quit_action = QAction("Quit", self)
        self.quit_action.triggered.connect(QApplication.instance().quit)
        self.tray_menu.addAction(self.restore_action)
        self.tray_menu.addAction(self.quit_action)
        self.tray_icon.setContextMenu(self.tray_menu)

    def initUI(self):
        """Initializes the UI components and settings."""
        self.setWindowTitle('Folder Monitor')
        self.setGeometry(100, 100, 600, 400)
        self.statusBar().showMessage('Ready')

        self.settings_action = QAction('Settings', self)
        self.settings_action.triggered.connect(self.show_settings)

        self.start_action = QAction('Start Monitoring', self)
        self.start_action.triggered.connect(self.start_monitoring)

        self.stop_action = QAction('Stop Monitoring', self)
        self.stop_action.triggered.connect(self.stop_monitoring)
        self.stop_action.setDisabled(True)

        menubar = self.menuBar()
        file_menu = menubar.addMenu('File')
        file_menu.addAction(self.settings_action)
        file_menu.addAction(self.start_action)
        file_menu.addAction(self.stop_action)

        self.folder_path = ''
        self.log_file_path = ''
        self.move_folder_path = ''
        self.move_log_file_path = ''
        self.monitor_interval_ns = 120 * 1e9  # Default to 120 seconds in nanoseconds
        self.extended_monitor_interval_ns = 2700 * 1e9  # Default to 2700 seconds in nanoseconds
        self.notification_duration = 600  # Default notification duration in seconds
        self.move_delay_sec = 30  # Default delay for file movement in seconds
        self.auto_start_monitoring = False
        self.email_subject = 'No file drop alert'
        self.email_body = 'No file has been dropped in the monitored folder within the specified interval.'
        self.email_to = ''
        self.email_cc = ''
        self.notification_message = 'No file dropped within the specified interval.'

        self.load_settings()

        if self.auto_start_monitoring:
            self.start_monitoring()

    def load_settings(self):
        """Loads settings from the configuration file."""
        config = ConfigParser()
        config.read('settings.ini')
        if 'SETTINGS' in config:
            self.folder_path = config['SETTINGS'].get('folder_path', '')
            self.log_file_path = config['SETTINGS'].get('log_file_path', '')
            self.move_folder_path = config['SETTINGS'].get('move_folder_path', '')
            self.move_log_file_path = config['SETTINGS'].get('move_log_file_path', '')
            self.monitor_interval_ns = config['SETTINGS'].getint('monitor_interval', 120) * 1e9
            self.extended_monitor_interval_ns = config['SETTINGS'].getint('extended_monitor_interval', 2700) * 1e9
            self.notification_duration = config['SETTINGS'].getint('notification_duration', 600)  # Load notification duration
            self.move_delay_sec = config['SETTINGS'].getint('move_delay', 30)  # Load file move delay
            self.auto_start_monitoring = config['SETTINGS'].getboolean('auto_start_monitoring', False)
        if 'EMAIL' in config:
            self.email_subject = config['EMAIL'].get('subject', 'No file drop alert')
            self.email_body = config['EMAIL'].get('body', 'No file has been dropped in the monitored folder within the specified interval.')
            self.email_to = config['EMAIL'].get('to', '')
            self.email_cc = config['EMAIL'].get('cc', '')
        if 'NOTIFICATION' in config:
            self.notification_message = config['NOTIFICATION'].get('message', 'No file dropped within the specified interval.')

    def show_settings(self):
        """Displays the settings dialog."""
        dialog = SettingsDialog()
        if dialog.exec_() == QDialog.Accepted:
            self.load_settings()

    def start_monitoring(self):
        """Starts the folder monitoring process."""
        if not os.path.exists(self.folder_path):
            QMessageBox.warning(self, 'Error', 'The folder path does not exist. Please configure the folder path in settings.')
            return

        self.observer = Observer()
        event_handler = FileSystemEventHandler()
        event_handler.on_created = self.on_created
        self.observer.schedule(event_handler, self.folder_path, recursive=False)
        self.observer.start()

        self.start_action.setDisabled(True)
        self.stop_action.setDisabled(False)
        self.statusBar().showMessage('Monitoring started')

        self.timer.start(int(self.monitor_interval_ns / 1e6))  # Convert ns to ms for QTimer

    def stop_monitoring(self):
        """Stops the folder monitoring process."""
        if self.observer:
            self.observer.stop()
            self.observer.join()

        self.start_action.setDisabled(False)
        self.stop_action.setDisabled(True)
        self.statusBar().showMessage('Monitoring stopped')

        self.timer.stop()

    def on_created(self, event):
        """Handles file creation events in the monitored folder."""
        try:
            if os.path.exists(event.src_path):
                file_path = event.src_path
                if time.time() - os.path.getmtime(file_path) < self.monitor_interval_ns:
                    self.file_dropped_signal.emit(file_path)
        except FileNotFoundError:
            self.log_event(f"File not found during event handling: {event.src_path}")
        except Exception as e:
            self.log_event(f"Unexpected error during file event handling: {str(e)}")


    def on_file_dropped(self, file_path):
        """Handles actions when a file is dropped in the monitored folder."""
        self.notification_count = 0  # Reset the notification count on file drop
        self.timer.start(int(self.monitor_interval_ns / 1e6))  # Restart the timer

        self.log_event(f'File dropped: {file_path}')  # Log in the regular log file

        # Add a delay before moving the file
        config = ConfigParser()
        config.read('settings.ini')
        delay_sec = config['SETTINGS'].getint('move_delay', 30)
        time.sleep(delay_sec / 60)  # Convert delay to seconds

        # Move file to the specified folder
        if self.move_folder_path:
            try:
                new_file_path = os.path.join(self.move_folder_path, os.path.basename(file_path))
                shutil.move(file_path, new_file_path)
                self.log_event_move(f"File moved successfully after {delay_sec} Second(s).", new_file_path)
            except Exception as e:
                self.log_event(f"Failed to move file: {str(e)}")

    def check_file_drop(self):
        """Checks for file drops within the specified interval."""
        self.notification_count += 1

        if self.notification_count == 1:
            self.log_event('First alert: No file dropped within the specified interval.')
        elif self.notification_count == 2:
            self.log_event('Second alert: No file dropped within the specified interval.')
        elif self.notification_count == 3:
            self.log_event('Third alert: No file dropped within the specified interval. Sending email notification.')
            self.send_notification()
            self.send_email_notification()
            self.notification_count = 0  # Reset the notification count after sending the email

    def send_notification(self):
        """Sends a desktop notification."""
        notification.notify(
            title='Folder Monitor Alert',
            message=self.notification_message,
            timeout=self.notification_duration  # Use customizable notification duration
        )

    def send_email_notification(self):
        """Sends an email notification with the log file as an attachment."""
        to_list = [email.strip() for email in self.email_to.split(';') if email.strip()]
        cc_list = [email.strip() for email in self.email_cc.split(';') if email.strip()] if self.email_cc else None
        
        send_email(
            subject=self.email_subject,
            body=self.email_body,
            to_recipients=to_list,
            cc_recipients=cc_list,
            attachment_paths=[self.log_file_path] if self.log_file_path else None
        )
        
    def log_event(self, message):
        """Logs general events to the specified log file."""
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        log_message = f'[{timestamp}] {message}\n'
        with open(self.log_file_path, 'a') as log_file:
            log_file.write(log_message)

    def log_event_move(self, message, destination_path):
        """Logs the file move events to a separate log file."""
        if self.move_log_file_path:
            try:
                with open(self.move_log_file_path, 'a') as log_file:
                    log_file.write(f'{time.strftime("%Y-%m-%d %H:%M:%S")} - {message} - Moved to: {destination_path}\n')
            except Exception as e:
                self.log_event(f"Failed to log move event: {str(e)}")

            
    def closeEvent(self, event):
        """Handles the close event to minimize the application to the system tray."""
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "Folder Monitor",
            "Application minimized to tray",
            QSystemTrayIcon.Information,
            2000
        )
    
    def on_tray_icon_activated(self, reason):
        """Handles system tray icon activation to restore the application."""
        if reason == QSystemTrayIcon.Trigger:
            self.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    monitor_app = MonitorApp()
    monitor_app.show()
    sys.exit(app.exec_())

    #pyinstaller -F -i "icons8-briefcase-512.ico" --noconsole Watchdog.py  & pyinstaller -F -i "icons8-briefcase-512.ico" --onefile Watchdog.py
    #'''Try adding --hidden-import plyer.platforms.win.notification in the pyinstaller command For example : pyinstaller --onefile --windowed --hidden-import plyer.platforms.win.notification example.py'''
    #pyinstaller -F -i "icons8-briefcase-512.ico" --noconsole --onefile --windowed --hidden-import plyer.platforms.win.notification Watchdog.py https://stackoverflow.com/questions/60884169/why-shows-modulenotfounderror-after-converting-py-to-exe