from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QMessageBox,
    QProgressBar,
    QTableWidget,
    QTableWidgetItem,
    QLineEdit,
    QTextEdit,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import QApplication

import pandas as pd
import os

from services.data_processing_service import process_data
from generator_router import route_letter_generator


# ==========================================================
# BACKGROUND WORKER (THREAD)
# ==========================================================
class LetterGenerationWorker(QThread):
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(int)
    error_signal = pyqtSignal(str)
    log_signal = pyqtSignal(str)

    def __init__(self, df, output_folder, header_path, footer_path):
        super().__init__()
        self.df = df
        self.output_folder = output_folder
        self.header_path = header_path
        self.footer_path = footer_path
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True
        self.log_signal.emit("⚠ Generation cancelled by user.")

    def run(self):
        try:
            self.log_signal.emit("🔄 Processing data...")

            service_data = process_data(
                df=self.df,
                output_folder=self.output_folder,
            )

            total_letters = sum(len(df) for df in service_data.values())
            current_progress = 0

            self.log_signal.emit(f"📦 Total letters to generate: {total_letters}")

            for service, service_df in service_data.items():
                if self._is_cancelled:
                    break

                if service_df.empty:
                    continue

                self.log_signal.emit(f"\n🗂 Service: {service}")

                service_folder = os.path.join(self.output_folder, service)
                os.makedirs(service_folder, exist_ok=True)

                for _, row in service_df.iterrows():
                    if self._is_cancelled:
                        break

                    policy = row.get("Policy_Number", "Unknown")

                    self.log_signal.emit(f"   ➜ Generating policy {policy}")

                    route_letter_generator(
                        service=service,
                        row=row,
                        output_folder=service_folder,
                        header_image_path=self.header_path,
                        footer_image_path=self.footer_path,
                    )

                    current_progress += 1
                    self.progress_signal.emit(current_progress)

            self.log_signal.emit("\n✅ Generation finished.")
            self.finished_signal.emit(current_progress)

        except Exception as e:
            self.error_signal.emit(str(e))


# ==========================================================
# MAIN TAB
# ==========================================================
class GenerateLettersTab(QWidget):
    def __init__(self, main_window):
        super().__init__()

        self.main_window = main_window
        self.df = None
        self.worker = None

        self.init_ui()

    # ------------------------------------------------------
    # UI SETUP
    # ------------------------------------------------------
    def init_ui(self):
        layout = QVBoxLayout()

        self.browse_btn = QPushButton("Browse CSV File")
        self.browse_btn.clicked.connect(self.load_csv)

        self.file_label = QLabel("No file selected")

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search preview...")
        self.search_input.textChanged.connect(self.filter_preview)

        self.generate_btn = QPushButton("Process Data & Generate Letters")
        self.generate_btn.clicked.connect(self.start_generation)

        self.cancel_btn = QPushButton("Cancel Generation")
        self.cancel_btn.clicked.connect(self.cancel_generation)
        self.cancel_btn.setVisible(False)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)

        self.preview_table = QTableWidget()
        self.preview_table.setVisible(False)
        self.preview_table.setAlternatingRowColors(True)

        self.log_panel = QTextEdit()
        self.log_panel.setReadOnly(True)
        self.log_panel.setPlaceholderText("Real-time logs will appear here...")

        layout.addWidget(self.browse_btn)
        layout.addWidget(self.file_label)
        layout.addWidget(self.search_input)
        layout.addWidget(self.generate_btn)
        layout.addWidget(self.cancel_btn)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.preview_table)
        layout.addWidget(self.log_panel)

        self.setLayout(layout)

    # ------------------------------------------------------
    # LOAD CSV
    # ------------------------------------------------------
    def load_csv(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select CSV File", "", "CSV Files (*.csv)"
        )

        if file_path:
            try:
                self.df = pd.read_csv(file_path)
                self.file_label.setText(file_path)
                self.populate_preview(self.df)
                QMessageBox.information(self, "Success", "CSV loaded successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    # ------------------------------------------------------
    # PREVIEW TABLE
    # ------------------------------------------------------
    def populate_preview(self, df):
        self.preview_table.setVisible(True)

        self.preview_table.setRowCount(len(df))
        self.preview_table.setColumnCount(len(df.columns))
        self.preview_table.setHorizontalHeaderLabels(df.columns)

        for row in range(len(df)):
            for col in range(len(df.columns)):
                self.preview_table.setItem(
                    row, col, QTableWidgetItem(str(df.iat[row, col]))
                )

        self.preview_table.resizeColumnsToContents()

    # ------------------------------------------------------
    # SEARCH FILTER
    # ------------------------------------------------------
    def filter_preview(self):
        if self.df is None:
            return

        text = self.search_input.text().lower()

        if text == "":
            filtered_df = self.df
        else:
            filtered_df = self.df[
                self.df.apply(
                    lambda row: row.astype(str).str.lower().str.contains(text).any(),
                    axis=1,
                )
            ]

        self.populate_preview(filtered_df)

    # ------------------------------------------------------
    # VALIDATION
    # ------------------------------------------------------
    def validate_columns(self):
        required = [
            "Insurance_ref",
            "client_name",
            "cover_start_date",
            "product_description",
        ]

        missing = [col for col in required if col not in self.df.columns]

        if missing:
            QMessageBox.critical(
                self,
                "Missing Columns",
                f"Missing required columns:\n{', '.join(missing)}",
            )
            return False

        return True

    # ------------------------------------------------------
    # START GENERATION (THREAD)
    # ------------------------------------------------------
    def start_generation(self):
        if self.df is None:
            QMessageBox.warning(self, "Error", "Load CSV first.")
            return

        if not self.main_window.output_folder:
            QMessageBox.warning(self, "Error", "Create month folder first.")
            return

        if not self.validate_columns():
            return

        self.generate_btn.setEnabled(False)
        self.cancel_btn.setVisible(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_panel.clear()

        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        header_path = os.path.join(base_dir, "resources", "Header.png")
        footer_path = os.path.join(base_dir, "resources", "Footer.jpg")

        self.worker = LetterGenerationWorker(
            self.df,
            self.main_window.output_folder,
            header_path,
            footer_path,
        )

        self.worker.progress_signal.connect(self.progress_bar.setValue)
        self.worker.finished_signal.connect(self.generation_finished)
        self.worker.error_signal.connect(self.generation_error)
        self.worker.log_signal.connect(self.append_log)

        self.worker.start()

    # ------------------------------------------------------
    # CANCEL
    # ------------------------------------------------------
    def cancel_generation(self):
        if self.worker:
            self.worker.cancel()

    # ------------------------------------------------------
    # LOG PANEL
    # ------------------------------------------------------
    def append_log(self, message):
        self.log_panel.append(message)
        self.log_panel.verticalScrollBar().setValue(
            self.log_panel.verticalScrollBar().maximum()
        )

    # ------------------------------------------------------
    # FINISHED
    # ------------------------------------------------------
    def generation_finished(self, total):
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setVisible(False)
        self.progress_bar.setVisible(False)

        QMessageBox.information(
            self,
            "Success",
            f"Processing complete.\n\nTotal letters generated: {total}",
        )

    # ------------------------------------------------------
    # ERROR
    # ------------------------------------------------------
    def generation_error(self, message):
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setVisible(False)
        self.progress_bar.setVisible(False)

        QMessageBox.critical(
            self,
            "Error",
            f"Processing failed:\n{message}",
        )
