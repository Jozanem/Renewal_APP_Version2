import os
import time
import pandas as pd
import pythoncom
import win32com.client as win32

from PyQt6.QtCore import Qt, QThread, pyqtSignal

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QComboBox,
    QLineEdit,
    QTextEdit,
    QListWidget,
    QLabel,
    QProgressBar,
    QFileDialog,
    QMessageBox,
    QSpinBox,
    QSizePolicy,
    QCheckBox,
)
# =====================================================
# SAFE TEXT FORMATTING
# =====================================================

def safe_format(value, row):
    """Safely replace {ColumnName} placeholders using values from a pandas row."""
    if value is None:
        return ""

    text = str(value)
    data = {}

    for key, val in row.items():
        if val is None:
            data[key] = ""
        else:
            try:
                if pd.isna(val):
                    data[key] = ""
                else:
                    data[key] = str(val)
            except (TypeError, ValueError):
                data[key] = str(val)

    try:
        return text.format(**data)
    except (KeyError, ValueError, IndexError):
        # If a placeholder is missing or invalid, keep the original text
        # rather than stopping the email process.
        return text


# =====================================================
# EMAIL WORKER
# =====================================================

class EmailWorker(QThread):

    progress_signal = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(list)

    def __init__(self, df, settings):
        super().__init__()

        self.df = df
        self.settings = settings
        self._cancel = False
        self.report_data = []

    # -------------------------------------------------
    # CANCEL
    # -------------------------------------------------

    def cancel(self):
        self._cancel = True

    # -------------------------------------------------
    # FIND ATTACHMENTS
    # -------------------------------------------------

    def find_attachments(self, username):

        folder_path = self.settings["attachment_folder"]

        if not folder_path or not os.path.exists(folder_path):
            return []

        # Number of characters to match
        match_length = self.settings.get(
            "attachment_match_length",
            18
        )

        username_str = str(username).strip().lower()

        matches = []

        for filename in os.listdir(folder_path):

            file_prefix = filename[:match_length].strip().lower()

            if file_prefix == username_str[:match_length]:

                full_path = os.path.join(
                    folder_path,
                    filename
                )

                if os.path.isfile(full_path):
                    matches.append(
                        os.path.abspath(full_path)
                    )

        return matches

    # -------------------------------------------------
    # THREAD EXECUTION
    # -------------------------------------------------

    def run(self):

        pythoncom.CoInitialize()

        try:

            # Use DispatchEx for stability
            outlook = win32.DispatchEx(
                "Outlook.Application"
            )

            sent_count = 0

            for _, row in self.df.iterrows():

                if self._cancel:
                    break

                recipient = str(
                    row[self.settings["email_column"]]
                ).strip()

                username = row[
                    self.settings["match_column"]
                ]

                subject = safe_format(
                    self.settings["subject_template"],
                    row
                )

                body = safe_format(
                    self.settings["body_template"],
                    row
                )

                # Convert newlines to HTML
                body = body.replace(
                    "\n",
                    "<br>"
                )

                cc_value = safe_format(
                    self.settings["cc"],
                    row
                )

                reply_to_value = safe_format(
                    self.settings["reply_to"],
                    row
                )

                # -------------------------------------------------
                # COLLECT ATTACHMENTS
                # -------------------------------------------------

                attachments = self.find_attachments(
                    username
                )

                # Add shared attachments
                attachments.extend(
                    self.settings["shared_attachments"]
                )

                retry = 0
                success = False

                while retry < 3 and not success:

                    try:

                        mail = outlook.CreateItem(0)

                        mail.To = recipient
                        mail.Subject = subject

                        if cc_value:
                            mail.CC = cc_value

                        if reply_to_value:
                            mail.ReplyRecipients.Add(
                                reply_to_value
                            )

                        mail.HTMLBody = body

                        # -------------------------------------------------
                        # ATTACH FILES SAFELY
                        # -------------------------------------------------

                        for file in attachments:

                            try:

                                if os.path.exists(file):

                                    mail.Attachments.Add(
                                        file
                                    )

                            except Exception as e:

                                self.log_signal.emit(
                                    f"⚠ Attachment failed: "
                                    f"{file} - {str(e)}"
                                )

                        # -------------------------------------------------
                        # SEND
                        # -------------------------------------------------

                        mail.Send()

                        time.sleep(0.5)

                        success = True

                        sent_count += 1

                        self.log_signal.emit(
                            f"✅ Sent to {recipient}"
                        )

                        self.report_data.append(
                            [
                                recipient,
                                "Sent",
                                retry
                            ]
                        )

                    except Exception as e:

                        retry += 1

                        self.log_signal.emit(
                            f"⚠ Retry {retry} for "
                            f"{recipient} - {str(e)}"
                        )

                if not success:

                    self.report_data.append(
                        [
                            recipient,
                            "Failed",
                            retry
                        ]
                    )

                self.progress_signal.emit(
                    sent_count
                )

        finally:

            pythoncom.CoUninitialize()

            self.finished_signal.emit(
                self.report_data
            )


# =====================================================
# EMAIL TAB UI
# =====================================================

class EmailLettersTab(QWidget):

    def __init__(self):

        super().__init__()

        self.df = None

        self.shared_attachments = []

        self.attachment_folder = ""

        self.worker = None

        # Preview position
        self.preview_index = 0
        # Date column checkboxes
        self.date_checkboxes = []
        
        self.init_ui()
        
    
    # -------------------------------------------------
    # UI SETUP
    # -------------------------------------------------

    def init_ui(self):

        layout = QVBoxLayout()

        # -------------------------------------------------
        # UPLOAD EXCEL
        # -------------------------------------------------

        self.upload_btn = QPushButton(
            "Upload Excel"
        )

        self.upload_btn.clicked.connect(
            self.load_excel
        )
        # -------------------------------------------------
        # EMAIL COLUMN
        # -------------------------------------------------

        self.email_column_dropdown = QComboBox()

        # -------------------------------------------------
        # MATCH COLUMN
        # -------------------------------------------------

        self.match_column_dropdown = QComboBox()
    
        # -------------------------------------------------
        # DATE COLUMNS
        # -------------------------------------------------
        layout.addWidget(QLabel("Date Columns"))
        self.date_column_container = QVBoxLayout()
        layout.addLayout(self.date_column_container)
        
        
        # -------------------------------------------------
        # PLACEHOLDERS
        # -------------------------------------------------

        self.placeholder_list = QListWidget()
        self.placeholder_list.setMaximumHeight(100)

        self.placeholder_list.itemClicked.connect(
            self.insert_placeholder)

        # -------------------------------------------------
        # CC
        # -------------------------------------------------

        self.cc_input = QLineEdit()
        self.cc_input.setPlaceholderText(
            "CC emails (comma separated or use {ColumnName})"
        )

        # -------------------------------------------------
        # REPLY TO
        # -------------------------------------------------

        self.reply_to_input = QLineEdit()

        self.reply_to_input.setPlaceholderText(
            "Reply-To email (or use {ColumnName})"
        )

        # -------------------------------------------------
        # SUBJECT
        # -------------------------------------------------

        self.subject_input = QLineEdit()

        self.subject_input.setPlaceholderText(
            "Subject (use {ColumnName})"
        )

        # -------------------------------------------------
        # BODY
        # -------------------------------------------------

        self.body_input = QTextEdit()

        self.body_input.setPlaceholderText(
            "Email Body (use {ColumnName})"
        )

        # Make email body much larger
        self.body_input.setMinimumHeight(250)

        # Allow the body box to expand
        self.body_input.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding
)

        # -------------------------------------------------
        # ATTACHMENT FOLDER
        # -------------------------------------------------

        self.attachment_btn = QPushButton(
            "Select Attachment Folder"
        )

        self.attachment_btn.clicked.connect(
            self.select_attachment_folder
        )

        # -------------------------------------------------
        # ATTACHMENT FOLDER PATH
        # -------------------------------------------------

        self.attachment_folder_label = QLabel(
            "Attachment Folder: Not selected"
        )

        self.attachment_folder_label.setWordWrap(
            True
        )

        # Give the path a visible border
        self.attachment_folder_label.setStyleSheet(
            """
            QLabel {
                border: 1px solid #cccccc;
                padding: 6px;
                background-color: #f5f5f5;
            }
            """
        )

        # -------------------------------------------------
        # MATCH LENGTH
        # -------------------------------------------------

        self.match_length_label = QLabel(
            "Characters to match:"
        )

        self.match_length_spin = QSpinBox()

        # Minimum number of characters
        self.match_length_spin.setMinimum(1)

        # Maximum number of characters
        self.match_length_spin.setMaximum(255)

        # Default = 18
        self.match_length_spin.setValue(18)
        self.match_length_spin.setToolTip(
            "Number of characters from the username/file name "
            "to use when matching attachments."
        )

        # -------------------------------------------------
        # SHARED ATTACHMENTS
        # -------------------------------------------------

        self.add_shared_btn = QPushButton(
            "Add Shared Attachment"
        )

        self.add_shared_btn.clicked.connect(
            self.add_shared_attachment
        )

        self.shared_list = QListWidget()

        # -------------------------------------------------
        # PREVIEW BUTTONS
        # -------------------------------------------------

        self.preview_btn = QPushButton(
            "Preview Email"
        )

        self.preview_btn.clicked.connect(
            self.preview_email
        )

        # Previous button
        self.previous_preview_btn = QPushButton(
            "← Previous"
        )

        self.previous_preview_btn.clicked.connect(
            self.previous_preview
        )

        # Next button
        self.next_preview_btn = QPushButton(
            "Next →"
        )

        self.next_preview_btn.clicked.connect(
            self.next_preview
        )

        # Preview counter
        self.preview_counter = QLabel(
            "No email loaded"
        )

        self.preview_counter.setAlignment(
            Qt.AlignmentFlag.AlignCenter
        )

        # -------------------------------------------------
        # SEND / CANCEL
        # -------------------------------------------------

        self.send_btn = QPushButton(
            "Send Emails"
        )

        self.send_btn.clicked.connect(
            self.send_emails
        )

        self.cancel_btn = QPushButton(
            "Cancel Sending"
        )

        self.cancel_btn.clicked.connect(
            self.cancel_sending
        )

        self.cancel_btn.setVisible(False)

        # -------------------------------------------------
        # PROGRESS
        # -------------------------------------------------

        self.progress_bar = QProgressBar()

        self.progress_bar.setVisible(False)

        # -------------------------------------------------
        # LOG
        # -------------------------------------------------

        self.log_panel = QTextEdit()

        self.log_panel.setReadOnly(
            True
        )

        self.log_panel.setMinimumHeight(
            120
        )

        # -------------------------------------------------
        # ADD WIDGETS TO LAYOUT
        # -------------------------------------------------

        layout.addWidget(self.upload_btn)

        layout.addWidget(QLabel("Email Column"))
        layout.addWidget(self.email_column_dropdown)

        layout.addWidget(QLabel("Match Column"))
        layout.addWidget(self.match_column_dropdown)
        
        layout.addWidget(
            QLabel("Available Placeholders")
        ) 
        layout.addWidget(self.placeholder_list)
        
        layout.addWidget(self.cc_input)

        layout.addWidget(self.reply_to_input)

        layout.addWidget(
            self.subject_input
        )

        layout.addWidget(
            
            QLabel("Email Body")
        )

        layout.addWidget(
            self.body_input
        )
        


        layout.addWidget(
            self.placeholder_list
        )

        # -------------------------------------------------
        # ATTACHMENT SECTION
        # -------------------------------------------------

        layout.addWidget(
            self.attachment_btn
        )

        layout.addWidget(
            self.attachment_folder_label
        )

        # Match length
        match_layout = QHBoxLayout()

        match_layout.addWidget(
            self.match_length_label
        )

        match_layout.addWidget(
            self.match_length_spin
        )

        match_layout.addStretch()

        layout.addLayout(
            match_layout
        )

        # -------------------------------------------------
        # SHARED ATTACHMENTS
        # -------------------------------------------------

        layout.addWidget(
            self.add_shared_btn
        )

        layout.addWidget(
            QLabel("Shared Attachments")
        )

        layout.addWidget(
            self.shared_list
        )

        # -------------------------------------------------
        # PREVIEW CONTROLS
        # -------------------------------------------------

        preview_layout = QHBoxLayout()

        preview_layout.addWidget(
            self.previous_preview_btn
        )

        preview_layout.addWidget(
            self.preview_btn
        )

        preview_layout.addWidget(
            self.next_preview_btn
        )

        layout.addLayout(
            preview_layout
        )

        layout.addWidget(
            self.preview_counter
        )

        # -------------------------------------------------
        # SEND
        # -------------------------------------------------

        layout.addWidget(
            self.send_btn
        )

        layout.addWidget(
            self.cancel_btn
        )

        layout.addWidget(
            self.progress_bar
        )

        # -------------------------------------------------
        # LOG
        # -------------------------------------------------

        layout.addWidget(
            QLabel("Log Panel")
        )

        layout.addWidget(
            self.log_panel
        )

        self.setLayout(
            layout
        )

        # Disable preview navigation initially
        self.previous_preview_btn.setEnabled(
            False
        )

        self.next_preview_btn.setEnabled(
            False
        )
    
    # -------------------------------------------------
    # LOAD EXCEL
    # -------------------------------------------------
    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )

        if file_path:
            self.df = pd.read_excel(file_path)

            self.email_column_dropdown.clear()
            self.match_column_dropdown.clear()
            self.placeholder_list.clear()

            self.email_column_dropdown.addItems(self.df.columns)
            self.match_column_dropdown.addItems(self.df.columns)

            # placeholders
            for col in self.df.columns:
                self.placeholder_list.addItem(f"{{{col}}}")

            # clear old date checkboxes
            for cb in self.date_checkboxes:
                cb.deleteLater()

            self.date_checkboxes = []

            for col in self.df.columns:
                checkbox = QCheckBox(col)

                detected = False

                if pd.api.types.is_datetime64_any_dtype(self.df[col]):
                    detected = True

                else:
                    try:
                        pd.to_datetime(self.df[col].dropna().iloc[0])
                        detected = True
                    except Exception:
                        detected = False

                if detected:
                    checkbox.setChecked(True)

                self.date_column_container.addWidget(checkbox)
                self.date_checkboxes.append(checkbox)

            QMessageBox.information(self, "Loaded", "Excel Loaded Successfully")

        else:
            QMessageBox.warning(self, "No File", "No file was selected.")
                                
        # -------------------------------------------------
        # Placeholder SETUP
        # -------------------------------------------------
                
    def insert_placeholder(self, item):
        placeholder = item.text()

        widget = self.focusWidget()

        if isinstance(widget, QLineEdit):
            widget.insert(placeholder)

        elif isinstance(widget, QTextEdit):
            widget.insertPlainText(placeholder)
            
    # -------------------------------------------------
    # DATE FORMATTER
    # -------------------------------------------------
    def format_dates(self):
        selected_date_cols = [
            cb.text() for cb in self.date_checkboxes if cb.isChecked()
        ]

        df_copy = self.df.copy()

        for col in selected_date_cols:
            try:
                df_copy[col] = pd.to_datetime(df_copy[col]).dt.strftime("%d %B %Y")
            except Exception:
                pass

        return df_copy

    # -------------------------------------------------
    # ATTACHMENT FOLDER
    # -------------------------------------------------

    def select_attachment_folder(self):

        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Attachment Folder"
        )

        if folder:

            self.attachment_folder = folder

            # Display the selected shared folder
            self.attachment_folder_label.setText(
                f"Attachment Folder:\n{folder}"
            )
            
            # 2. Force high-contrast colors and raise the widget to the top layer
            self.attachment_folder_label.setStyleSheet("""
                QLabel {
                    color: #000000 !important; 
                    background-color: #FFFFE0 !important; /* Bright light-yellow background to test visibility */
                    font-weight: bold;
                }
            """)
            self.attachment_folder_label.raise_() # Forces it to render on top of overlapping widgets
            self.attachment_folder_label.update()  # Triggers an instant visual redraw

    # -------------------------------------------------
    # SHARED ATTACHMENTS
    # -------------------------------------------------

    def add_shared_attachment(self):

        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Shared Attachments"
        )

        for file in files:

            if file not in self.shared_attachments:

                self.shared_attachments.append(
                    file
                )

                self.shared_list.addItem(
                    file
                )

    # -------------------------------------------------
    # PREVIEW
    # -------------------------------------------------

    def preview_email(self):

        if self.df is None or self.df.empty:

            QMessageBox.warning(
                self,
                "Preview",
                "Please load an Excel file first."
            )

            return

        # Always preview current index
        self.show_preview()
        
    # -------------------------------------------------
    # SHOW CURRENT PREVIEW
    # -------------------------------------------------
    def show_preview(self):

        if self.df is None or self.df.empty:
            return
        df_preview = self.format_dates()

        # Make sure index is valid
        if self.preview_index < 0:

            self.preview_index = 0

        if self.preview_index >= len(self.df):

            self.preview_index = len(self.df) - 1

        row = df_preview.iloc[
            self.preview_index
        ]

        # Generate subject
                
        subject = safe_format(
            self.subject_input.text(),
            row
        )

        # Generate body
        body = safe_format(
            self.body_input.toPlainText(),
            row
        )

        # Find recipient
        recipient = ""

        email_column = (
            self.email_column_dropdown.currentText()
        )

        if email_column:

            recipient = str(
                row[email_column]
            ).strip()

        # Find username
        username = ""

        match_column = (
            self.match_column_dropdown.currentText()
        )

        if match_column:

            username = str(
                row[match_column]
            ).strip()

        # Find matching attachments
        attachments = []

        if self.attachment_folder:

            match_length = (
                self.match_length_spin.value()
            )

            username_prefix = (
                username[:match_length]
                .lower()
            )

            for filename in os.listdir(
                self.attachment_folder
            ):

                file_prefix = (
                    filename[:match_length]
                    .strip()
                    .lower()
                )

                if file_prefix == username_prefix:

                    full_path = os.path.join(
                        self.attachment_folder,
                        filename
                    )

                    if os.path.isfile(full_path):

                        attachments.append(
                            full_path
                        )

        # Add shared attachments
        attachments.extend(
            self.shared_attachments
        )

        # Build preview message
        attachment_text = ""

        if attachments:

            attachment_text = (
                "\n\nAttachments:\n"
                + "\n".join(
                    [
                        os.path.basename(file)
                        for file in attachments
                    ]
                )
            )

        else:

            attachment_text = (
                "\n\nAttachments:\nNone"
            )

        preview_text = (
            f"Email {self.preview_index + 1} "
            f"of {len(self.df)}\n\n"
            f"To:\n{recipient}\n\n"
            f"Subject:\n{subject}\n\n"
            f"Body:\n{body}"
            f"{attachment_text}"
        )

        QMessageBox.information(
            self,
            "Email Preview",
            preview_text
        )

        self.update_preview_buttons()

    # -------------------------------------------------
    # PREVIOUS PREVIEW
    # -------------------------------------------------

    def previous_preview(self):

        if self.df is None or self.df.empty:
            return

        if self.preview_index > 0:

            self.preview_index -= 1

            self.show_preview()

    # -------------------------------------------------
    # NEXT PREVIEW
    # -------------------------------------------------

    def next_preview(self):

        if self.df is None or self.df.empty:
            return

        if self.preview_index < len(self.df) - 1:

            self.preview_index += 1

            self.show_preview()

    # -------------------------------------------------
    # UPDATE PREVIEW BUTTONS
    # -------------------------------------------------

    def update_preview_buttons(self):

        if self.df is None or self.df.empty:

            self.preview_counter.setText(
                "No email loaded"
            )

            self.previous_preview_btn.setEnabled(
                False
            )

            self.next_preview_btn.setEnabled(
                False
            )

            return

        total = len(self.df)

        self.preview_counter.setText(
            f"Previewing email "
            f"{self.preview_index + 1} of {total}"
        )

        self.previous_preview_btn.setEnabled(
            self.preview_index > 0
        )

        self.next_preview_btn.setEnabled(
            self.preview_index < total - 1
        )

    # -------------------------------------------------
    # SEND EMAILS
    # -------------------------------------------------

    def send_emails(self):

        if self.df is None or self.df.empty:

            QMessageBox.warning(
                self,
                "Send Emails",
                "Please load an Excel file first."
            )

            return
        
        # Format selected date columns before sending
        df_send = self.format_dates()

        # Get attachment matching length
        match_length = (
            self.match_length_spin.value()
        )

        settings = {

            "email_column":
                self.email_column_dropdown.currentText(),

            "match_column":
                self.match_column_dropdown.currentText(),

            "subject_template":
                self.subject_input.text(),

            "body_template":
                self.body_input.toPlainText(),

            "cc":
                self.cc_input.text(),

            "reply_to":
                self.reply_to_input.text(),

            "attachment_folder":
                self.attachment_folder,

            "shared_attachments":
                self.shared_attachments,

            # NEW
            "attachment_match_length":
                match_length,
        }

        self.progress_bar.setVisible(
            True
        )

        self.progress_bar.setMaximum(
            len(self.df)
        )

        self.progress_bar.setValue(
            0
        )

        self.cancel_btn.setVisible(
            True
        )

        self.send_btn.setEnabled(
            False
        )

        self.worker = EmailWorker(
            df_send,
            settings
        )

        self.worker.progress_signal.connect(
            self.progress_bar.setValue
        )

        self.worker.log_signal.connect(
            self.log_panel.append
        )

        self.worker.finished_signal.connect(
            self.save_report
        )

        self.worker.start()

    # -------------------------------------------------
    # CANCEL
    # -------------------------------------------------

    def cancel_sending(self):

        if self.worker:

            self.worker.cancel()

            self.log_panel.append(
                "⛔ Sending Cancelled"
            )

    # -------------------------------------------------
    # SAVE REPORT
    # -------------------------------------------------

    def save_report(
        self,
        report_data
    ):

        report_df = pd.DataFrame(
            report_data,
            columns=[
                "Recipient",
                "Status",
                "Retries"
            ]
        )

        report_path = os.path.join(
            os.getcwd(),
            "Email_Report.xlsx"
        )

        report_df.to_excel(
            report_path,
            index=False
        )

        QMessageBox.information(
            self,
            "Done",
            f"Emails processed.\n\n"
            f"Report saved to:\n"
            f"{report_path}"
        )

        self.progress_bar.setVisible(
            False
        )

        self.cancel_btn.setVisible(
            False
        )

        self.send_btn.setEnabled(
            True
        )