from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QTabWidget,
    QMessageBox,
)
from services.font_service import register_fonts

from datetime import datetime
import os
import sys

from ui.generate_letters_tab import GenerateLettersTab
from ui.email_letters_tab import EmailLettersTab


class MainWindow(QMainWindow):
    def __init__(self, username):
        super().__init__()

        self.username = username
        self.setWindowTitle(f"Claims Letters System - {username}")
        self.resize(900, 600)

        self.output_folder = None

        self.init_ui()

    # --------------------------------------------------
    # UI
    # --------------------------------------------------
    def init_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()

        # Create Folder Button
        self.create_folder_btn = QPushButton("Create Month Folder")
        self.create_folder_btn.clicked.connect(self.create_month_folder)

        layout.addWidget(self.create_folder_btn)

        # Tabs
        self.tabs = QTabWidget()

        self.generate_tab = GenerateLettersTab(self)
        self.email_tab = EmailLettersTab()

        self.tabs.addTab(self.generate_tab, "Generate Letters")
        self.tabs.addTab(self.email_tab, "Email Letters")

        layout.addWidget(self.tabs)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    # --------------------------------------------------
    # CREATE MONTH FOLDER
    # --------------------------------------------------
    def create_month_folder(self):
        base_path = os.path.expanduser("~/Documents")

        folder_name = datetime.now().strftime("%B_%Y")
        full_path = os.path.join(base_path, folder_name)

        os.makedirs(full_path, exist_ok=True)

        self.output_folder = full_path

        QMessageBox.information(
            self,
            "Folder Created",
            f"Folder created at:\n{full_path}",
        )
