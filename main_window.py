# main_window.py

from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QTabWidget,
    QMessageBox,
    QComboBox,
    QFileDialog,
    QLabel,
    QHBoxLayout,
)
from datetime import datetime
import os

from ui.generate_letters_tab import GenerateLettersTab
from ui.email_letters_tab import EmailLettersTab


class MainWindow(QMainWindow):
    def __init__(self, username):
        super().__init__()

        self.username = username
        self.setWindowTitle(f"Claims Letters System - {username}")
        self.resize(1000, 700)

        self.output_folder = None
        self.base_path = None

        self.init_ui()

    # --------------------------------------------------
    # UI
    # --------------------------------------------------
    def init_ui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()

        # ---------------------------------------------
        # MONTH + YEAR + LOCATION SECTION
        # ---------------------------------------------
        folder_layout = QHBoxLayout()

        self.month_dropdown = QComboBox()
        self.month_dropdown.addItems(
            [
                "January",
                "February",
                "March",
                "April",
                "May",
                "June",
                "July",
                "August",
                "September",
                "October",
                "November",
                "December",
            ]
        )

        self.year_dropdown = QComboBox()
        current_year = datetime.now().year
        for year in range(current_year - 2, current_year + 3):
            self.year_dropdown.addItem(str(year))

        self.choose_location_btn = QPushButton("Choose Location")
        self.choose_location_btn.clicked.connect(self.choose_location)

        self.create_folder_btn = QPushButton("Create Folder")
        self.create_folder_btn.clicked.connect(self.create_month_folder)

        folder_layout.addWidget(QLabel("Month:"))
        folder_layout.addWidget(self.month_dropdown)
        folder_layout.addWidget(QLabel("Year:"))
        folder_layout.addWidget(self.year_dropdown)
        folder_layout.addWidget(self.choose_location_btn)
        folder_layout.addWidget(self.create_folder_btn)

        layout.addLayout(folder_layout)

        # ---------------------------------------------
        # TABS
        # ---------------------------------------------
        self.tabs = QTabWidget()

        self.generate_tab = GenerateLettersTab(self)
        self.email_tab = EmailLettersTab()

        self.tabs.addTab(self.generate_tab, "Generate Letters")
        self.tabs.addTab(self.email_tab, "Email Letters")

        layout.addWidget(self.tabs)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    # --------------------------------------------------
    # Choose Folder Location
    # --------------------------------------------------
    def choose_location(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Base Folder Location")

        if folder:
            self.base_path = folder
            QMessageBox.information(
                self, "Location Selected", f"Base location set to:\n{folder}"
            )

    # --------------------------------------------------
    # Create Month Folder
    # --------------------------------------------------
    def create_month_folder(self):
        if not self.base_path:
            QMessageBox.warning(self, "Error", "Please choose a base location first.")
            return

        month = self.month_dropdown.currentText()
        year = self.year_dropdown.currentText()

        folder_name = f"{month}_{year}"
        full_path = os.path.join(self.base_path, folder_name)

        os.makedirs(full_path, exist_ok=True)

        self.output_folder = full_path

        QMessageBox.information(
            self, "Folder Created", f"Folder created at:\n{full_path}"
        )
