"""
Data Source Page: File Selection and Database Export
"""

import sys
import os
from pathlib import Path
import pandas as pd
import shutil

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QLineEdit,
        QPushButton, QFileDialog, QComboBox, QTableWidget, QTableWidgetItem,
        QProgressBar, QMessageBox
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

from edm_wizard.workers.threads import AccessExportThread, SQLiteExportThread



class DataSourcePage(QWizardPage):
    """Step 1: Choose between Access DB, SQLite DB, or existing Excel file"""

    def __init__(self):
        super().__init__()
        self.setTitle("Step 1: Select Data Source")
        self.setSubTitle("Choose your data source for EDM library processing")

        layout = QVBoxLayout()

        # Single file selection with auto-detection
        file_group = QGroupBox("📂 Data File Selection")
        file_layout = QVBoxLayout()

        # File browser
        browser_layout = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("Select Access DB (.mdb/.accdb), SQLite DB (.db/.sqlite/.sqlite3), Excel (.xlsx/.xls), or CSV (.csv)...")
        self.file_path.textChanged.connect(self.on_file_selected)

        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self.browse_file)

        browser_layout.addWidget(self.file_path)
        browser_layout.addWidget(browse_button)
        file_layout.addLayout(browser_layout)

        # File type detection display
        detection_layout = QHBoxLayout()
        detection_layout.addWidget(QLabel("Detected Type:"))
        self.file_type_label = QLabel("No file selected")
        self.file_type_label.setStyleSheet("font-weight: bold; color: #666;")
        detection_layout.addWidget(self.file_type_label)
        detection_layout.addStretch()
        file_layout.addLayout(detection_layout)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Action button (Load or Convert)
        self.action_button = QPushButton("Load Data")
        self.action_button.clicked.connect(self.process_file)
        self.action_button.setEnabled(False)

        # Preview section
        preview_group = QGroupBox("Data Preview")
        preview_layout = QVBoxLayout()

        # Sheet selector
        sheet_selector_layout = QHBoxLayout()
        sheet_selector_layout.addWidget(QLabel("Sheet:"))
        self.sheet_selector = QComboBox()
        self.sheet_selector.currentTextChanged.connect(self.on_sheet_changed)
        sheet_selector_layout.addWidget(self.sheet_selector)
        sheet_selector_layout.addStretch()

        self.preview_label = QLabel("No data loaded")
        self.preview_table = QTableWidget()
        self.preview_table.setSortingEnabled(True)  # Enable sorting

        preview_layout.addLayout(sheet_selector_layout)
        preview_layout.addWidget(self.preview_label)
        preview_layout.addWidget(self.preview_table)
        preview_group.setLayout(preview_layout)

        # Progress
        self.progress_label = QLabel("")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)

        # Add widgets
        layout.addWidget(self.action_button)
        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress_bar)
        layout.addWidget(preview_group, stretch=1)  # Preview fills available space

        self.setLayout(layout)

        # Store data
        self.exported_excel_path = None
        self.dataframes = {}
        self.detected_file_type = None  # Will be set by auto-detection

    def browse_file(self):
        """Browse for any supported file type"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Data File",
            "",
            "All Supported Files (*.mdb *.accdb *.db *.sqlite *.sqlite3 *.xlsx *.xls *.csv);;"
            "Access Database (*.mdb *.accdb);;"
            "SQLite Database (*.db *.sqlite *.sqlite3);;"
            "Excel Files (*.xlsx *.xls);;"
            "CSV Files (*.csv);;"
            "All Files (*.*)"
        )
        if file_path:
            self.file_path.setText(file_path)

    def on_file_selected(self, file_path):
        """Auto-detect file type when file is selected"""
        if not file_path or not os.path.exists(file_path):
            self.file_type_label.setText("No file selected")
            self.file_type_label.setStyleSheet("font-weight: bold; color: #666;")
            self.action_button.setEnabled(False)
            self.detected_file_type = None
            return

        # Auto-detect by file extension
        file_ext = Path(file_path).suffix.lower()

        if file_ext in ['.mdb', '.accdb']:
            self.detected_file_type = 'access'
            self.file_type_label.setText("🗄️ Access Database")
            self.file_type_label.setStyleSheet("font-weight: bold; color: #2196F3;")
            self.action_button.setText("Convert to Excel")
            self.action_button.setEnabled(True)

        elif file_ext in ['.db', '.sqlite', '.sqlite3']:
            self.detected_file_type = 'sqlite'
            self.file_type_label.setText("🗄️ SQLite Database")
            self.file_type_label.setStyleSheet("font-weight: bold; color: #4CAF50;")
            self.action_button.setText("Convert to Excel")
            self.action_button.setEnabled(True)

        elif file_ext in ['.xlsx', '.xls']:
            self.detected_file_type = 'excel'
            self.file_type_label.setText("📊 Excel Workbook")
            self.file_type_label.setStyleSheet("font-weight: bold; color: #FF9800;")
            self.action_button.setText("Load Excel")
            self.action_button.setEnabled(True)

        elif file_ext == '.csv':
            self.detected_file_type = 'csv'
            self.file_type_label.setText("📋 CSV File")
            self.file_type_label.setStyleSheet("font-weight: bold; color: #9C27B0;")
            self.action_button.setText("Load CSV")
            self.action_button.setEnabled(True)

        else:
            self.detected_file_type = None
            self.file_type_label.setText("❌ Unsupported File Type")
            self.file_type_label.setStyleSheet("font-weight: bold; color: #F44336;")
            self.action_button.setEnabled(False)

    def process_file(self):
        """Process the selected file based on detected type"""
        file_path = self.file_path.text()

        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Invalid File", "Please select a valid file.")
            return

        if self.detected_file_type in ['access', 'sqlite']:
            self.export_database()
        elif self.detected_file_type == 'excel':
            self.load_excel_preview(file_path)
        elif self.detected_file_type == 'csv':
            self.load_csv_preview(file_path)

    def export_database(self):
        """Export database (Access or SQLite) to Excel using detected file type"""
        db_file = self.file_path.text()

        # Validate file exists
        if not db_file or not os.path.exists(db_file):
            QMessageBox.warning(self, "Invalid File", "Please select a valid database file.")
            return

        # Get output folder from StartPage
        start_page = self.wizard().page(0)
        output_folder = start_page.output_folder_input.text() if hasattr(start_page, 'output_folder_input') else None

        if not output_folder or not os.path.exists(output_folder):
            QMessageBox.warning(self, "No Output Folder",
                               "Output folder not set. Please go back to the Welcome page and select an output folder.")
            return

        # Select thread class based on detected type
        if self.detected_file_type == 'sqlite':
            thread_class = SQLiteExportThread
            db_type = "SQLite"
        elif self.detected_file_type == 'access':
            thread_class = AccessExportThread
            db_type = "Access"
        else:
            QMessageBox.warning(self, "Invalid Type", "Unsupported database type.")
            return

        # Generate output filename in the output folder
        output_file = os.path.join(output_folder, f"{Path(db_file).stem}.xlsx")

        # Start export in background thread
        self.action_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(0)  # Indeterminate

        self.export_thread = thread_class(db_file, output_file)
        self.export_thread.progress.connect(self.update_progress)
        self.export_thread.finished.connect(self.export_finished)
        self.export_thread.error.connect(self.export_error)
        self.export_thread.start()

    def update_progress(self, message):
        """Update progress label"""
        self.progress_label.setText(message)

    def export_finished(self, excel_path, dataframes):
        """Handle export completion"""
        self.progress_bar.setVisible(False)
        self.action_button.setEnabled(True)
        self.exported_excel_path = excel_path
        self.dataframes = dataframes

        # Show preview
        self.show_preview(dataframes)

        QMessageBox.information(self, "Export Complete",
                               f"Database exported successfully to:\n{excel_path}")

        self.completeChanged.emit()

    def export_error(self, error_msg):
        """Handle export error"""
        self.progress_bar.setVisible(False)
        self.action_button.setEnabled(True)
        QMessageBox.critical(self, "Export Error", error_msg)

    def load_excel_preview(self, excel_path):
        """Load and preview Excel file, copying it to output folder"""
        try:
            # Get output folder from StartPage
            start_page = self.wizard().page(0)
            output_folder = start_page.output_folder_input.text() if hasattr(start_page, 'output_folder_input') else None

            if not output_folder or not os.path.exists(output_folder):
                QMessageBox.warning(self, "No Output Folder",
                                   "Output folder not set. Please go back to the Welcome page and select an output folder.")
                return

            # Load the Excel file
            xl_file = pd.ExcelFile(excel_path)
            self.dataframes = {sheet: pd.read_excel(excel_path, sheet_name=sheet)
                             for sheet in xl_file.sheet_names}

            # Copy Excel file to output folder
            import shutil
            base_name = Path(excel_path).name
            output_excel = os.path.join(output_folder, base_name)

            # Copy the file
            shutil.copy2(excel_path, output_excel)

            # Store the output path (not the original path)
            self.exported_excel_path = output_excel

            self.show_preview(self.dataframes)
            self.completeChanged.emit()

            QMessageBox.information(self, "Excel Loaded",
                                   f"Excel file copied to output folder:\n{output_excel}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load Excel file: {str(e)}")

    def load_csv_preview(self, csv_path):
        """Load and preview CSV file, converting it to Excel in output folder"""
        try:
            # Get output folder from StartPage
            start_page = self.wizard().page(0)
            output_folder = start_page.output_folder_input.text() if hasattr(start_page, 'output_folder_input') else None

            if not output_folder or not os.path.exists(output_folder):
                QMessageBox.warning(self, "No Output Folder",
                                   "Output folder not set. Please go back to the Welcome page and select an output folder.")
                return

            # Load the CSV file
            # Try different encodings if utf-8 fails
            try:
                df = pd.read_csv(csv_path, encoding='utf-8')
            except UnicodeDecodeError:
                try:
                    df = pd.read_csv(csv_path, encoding='latin-1')
                except UnicodeDecodeError:
                    df = pd.read_csv(csv_path, encoding='cp1252')

            # Use filename without extension as sheet name
            sheet_name = Path(csv_path).stem
            # Clean sheet name for Excel compatibility
            sheet_name = sheet_name[:31]  # Excel max sheet name length
            for char in ['\\', '/', '*', '?', ':', '[', ']']:
                sheet_name = sheet_name.replace(char, '_')

            self.dataframes = {sheet_name: df}

            # Convert to Excel in output folder
            base_name = Path(csv_path).stem
            output_excel = os.path.join(output_folder, f"{base_name}.xlsx")

            # Write to Excel
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Store the output path
            self.exported_excel_path = output_excel

            self.show_preview(self.dataframes)
            self.completeChanged.emit()

            QMessageBox.information(self, "CSV Loaded",
                                   f"CSV file converted to Excel in output folder:\n{output_excel}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load CSV file: {str(e)}")

    def show_preview(self, dataframes):
        """Show preview of first 100 rows and populate sheet selector"""
        if not dataframes:
            return

        # Populate sheet selector
        self.sheet_selector.blockSignals(True)  # Prevent triggering on_sheet_changed during population
        self.sheet_selector.clear()
        self.sheet_selector.addItems(list(dataframes.keys()))
        self.sheet_selector.blockSignals(False)

        # Show first sheet
        first_sheet = list(dataframes.keys())[0]
        self.display_sheet_preview(first_sheet)

    def on_sheet_changed(self, sheet_name):
        """Handle sheet selection change"""
        if sheet_name and self.dataframes:
            self.display_sheet_preview(sheet_name)

    def display_sheet_preview(self, sheet_name):
        """Display preview of the selected sheet"""
        if sheet_name not in self.dataframes:
            return

        df = self.dataframes[sheet_name]

        # Limit to first 100 rows
        preview_df = df.head(100)

        self.preview_label.setText(
            f"Preview: {sheet_name} ({len(df)} total rows, showing first {len(preview_df)})"
        )

        # Populate table
        self.preview_table.setRowCount(len(preview_df))
        self.preview_table.setColumnCount(len(preview_df.columns))
        self.preview_table.setHorizontalHeaderLabels(preview_df.columns.tolist())

        for i in range(len(preview_df)):
            for j in range(len(preview_df.columns)):
                value = preview_df.iloc[i, j]
                item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                self.preview_table.setItem(i, j, item)

        self.preview_table.resizeColumnsToContents()

    def isComplete(self):
        """Check if page is complete"""
        if self.detected_file_type in ['access', 'sqlite']:
            # Database files need to be exported first
            return self.exported_excel_path is not None
        elif self.detected_file_type in ['excel', 'csv']:
            # Excel and CSV files just need to be loaded
            return len(self.dataframes) > 0
        return False

    def get_excel_path(self):
        """Get the Excel file path"""
        if self.detected_file_type in ['access', 'sqlite', 'csv']:
            return self.exported_excel_path
        elif self.detected_file_type == 'excel':
            return self.exported_excel_path  # Return the copied file in output folder
        return None

    def get_dataframes(self):
        """Get the loaded dataframes"""
        return self.dataframes


class NoScrollComboBox(QComboBox):
    """ComboBox that ignores mouse wheel events"""
    def wheelEvent(self, event):
        event.ignore()

