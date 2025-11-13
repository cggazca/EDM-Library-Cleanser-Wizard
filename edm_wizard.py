#!/usr/bin/env python3
"""
EDM Library Wizard
A comprehensive wizard for converting Access databases to Excel and generating XML files for EDM Library Creator
"""

import sys
import os
from pathlib import Path
from datetime import datetime
import pandas as pd
import sqlalchemy as sa
import urllib
from sqlalchemy import inspect
import xml.etree.ElementTree as ET
from xml.dom import minidom

try:
    from PyQt5.QtWidgets import (
        QApplication, QWizard, QWizardPage, QVBoxLayout, QHBoxLayout,
        QRadioButton, QPushButton, QLabel, QLineEdit, QFileDialog,
        QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QComboBox,
        QGroupBox, QMessageBox, QTextEdit, QProgressBar, QSpacerItem,
        QSizePolicy, QGridLayout, QWidget, QSplitter
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings
    from PyQt5.QtGui import QFont, QIcon
except ImportError:
    print("Error: PyQt5 is required. Install it with: pip install PyQt5")
    sys.exit(1)

import json
import shutil

try:
    from anthropic import Anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False


class StartPage(QWizardPage):
    """Start Page: Claude AI API Key Configuration"""

    def __init__(self):
        super().__init__()
        self.setTitle("Welcome to EDM Library Wizard")
        self.setSubTitle("Configure Claude AI for intelligent column mapping assistance")

        layout = QVBoxLayout()

        # AI Info section
        info_group = QGroupBox("🤖 AI-Powered Column Mapping")
        info_layout = QVBoxLayout()

        info_text = QLabel(
            "This wizard can use Claude AI to automatically detect and map your Excel columns.\n"
            "Enter your Claude API key below to enable AI features, or skip to continue manually."
        )
        info_text.setWordWrap(True)
        info_layout.addWidget(info_text)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # API Key input section
        api_group = QGroupBox("API Configuration")
        api_layout = QVBoxLayout()

        # API Key input
        key_layout = QHBoxLayout()
        key_layout.addWidget(QLabel("Claude API Key:"))
        self.api_key_input = QLineEdit()
        self.api_key_input.setPlaceholderText("sk-ant-api03-...")
        self.api_key_input.setEchoMode(QLineEdit.Password)
        self.api_key_input.textChanged.connect(self.on_api_key_changed)
        key_layout.addWidget(self.api_key_input)

        # Show/Hide button
        self.show_key_btn = QPushButton("Show")
        self.show_key_btn.setMaximumWidth(60)
        self.show_key_btn.clicked.connect(self.toggle_key_visibility)
        key_layout.addWidget(self.show_key_btn)

        api_layout.addLayout(key_layout)

        # Save API key checkbox
        self.save_key_checkbox = QCheckBox("Remember API key for future sessions")
        self.save_key_checkbox.setChecked(True)
        api_layout.addWidget(self.save_key_checkbox)

        # Test connection button
        test_layout = QHBoxLayout()
        self.test_btn = QPushButton("Test Connection")
        self.test_btn.clicked.connect(self.test_api_key)
        self.test_btn.setEnabled(False)
        test_layout.addWidget(self.test_btn)

        self.test_status = QLabel("")
        test_layout.addWidget(self.test_status)
        test_layout.addStretch()

        api_layout.addLayout(test_layout)

        # Get API key link
        link_label = QLabel('<a href="https://console.anthropic.com/settings/keys">Get your API key from Anthropic Console</a>')
        link_label.setOpenExternalLinks(True)
        api_layout.addWidget(link_label)

        api_group.setLayout(api_layout)
        layout.addWidget(api_group)

        # Skip AI section
        skip_layout = QHBoxLayout()
        skip_layout.addStretch()
        self.skip_ai_btn = QPushButton("Continue without AI")
        self.skip_ai_btn.clicked.connect(self.skip_ai)
        skip_layout.addWidget(self.skip_ai_btn)
        layout.addLayout(skip_layout)

        layout.addStretch()
        self.setLayout(layout)

        # Load saved API key if available
        self.load_saved_api_key()

        # Store whether API is validated
        self.api_validated = False
        self.skip_ai_mode = False

    def load_saved_api_key(self):
        """Load API key from config file if it exists"""
        config_file = Path.home() / ".edm_wizard_config.json"
        if config_file.exists():
            try:
                with open(config_file, 'r') as f:
                    config = json.load(f)
                    if 'api_key' in config:
                        self.api_key_input.setText(config['api_key'])
                        self.test_status.setText("✓ Loaded saved API key")
                        self.test_status.setStyleSheet("color: green;")
            except Exception as e:
                pass

    def save_api_key(self):
        """Save API key to config file if checkbox is checked"""
        if self.save_key_checkbox.isChecked():
            config_file = Path.home() / ".edm_wizard_config.json"
            try:
                config = {'api_key': self.api_key_input.text()}
                with open(config_file, 'w') as f:
                    json.dump(config, f)
            except Exception as e:
                QMessageBox.warning(self, "Save Error", f"Could not save API key: {str(e)}")

    def clear_saved_api_key(self):
        """Clear saved API key from config file"""
        config_file = Path.home() / ".edm_wizard_config.json"
        if config_file.exists():
            try:
                config_file.unlink()
            except Exception as e:
                pass

    def on_api_key_changed(self):
        """Enable test button when API key is entered"""
        self.test_btn.setEnabled(len(self.api_key_input.text().strip()) > 0)
        self.api_validated = False
        self.test_status.setText("")

    def toggle_key_visibility(self):
        """Toggle API key visibility"""
        if self.api_key_input.echoMode() == QLineEdit.Password:
            self.api_key_input.setEchoMode(QLineEdit.Normal)
            self.show_key_btn.setText("Hide")
        else:
            self.api_key_input.setEchoMode(QLineEdit.Password)
            self.show_key_btn.setText("Show")

    def test_api_key(self):
        """Test the Claude API connection"""
        if not ANTHROPIC_AVAILABLE:
            QMessageBox.warning(
                self,
                "Anthropic Package Not Installed",
                "The 'anthropic' package is not installed.\n\n"
                "Please install it with: pip install anthropic"
            )
            return

        api_key = self.api_key_input.text().strip()
        if not api_key:
            self.test_status.setText("⚠ Please enter an API key")
            self.test_status.setStyleSheet("color: orange;")
            return

        self.test_status.setText("Testing connection...")
        self.test_status.setStyleSheet("color: blue;")
        self.test_btn.setEnabled(False)
        QApplication.processEvents()

        try:
            client = Anthropic(api_key=api_key)
            # Simple test message - use Claude Haiku 4.5 (fast and cost-effective)
            response = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=10,
                messages=[{"role": "user", "content": "test"}]
            )

            self.api_validated = True
            self.test_status.setText("✓ Connection successful!")
            self.test_status.setStyleSheet("color: green;")

            # Save API key if checkbox is checked
            self.save_api_key()

        except Exception as e:
            self.api_validated = False
            error_msg = str(e)
            # Show more detailed error message
            self.test_status.setText(f"✗ Failed: {error_msg[:50]}...")
            self.test_status.setStyleSheet("color: red;")

            # Show full error in a message box
            QMessageBox.critical(
                self,
                "Connection Test Failed",
                f"Failed to connect to Claude API:\n\n{error_msg}\n\n"
                "Please check:\n"
                "1. Your API key is correct\n"
                "2. Your API key has sufficient credits\n"
                "3. You have internet connectivity"
            )

        self.test_btn.setEnabled(True)

    def skip_ai(self):
        """Skip AI features and continue without API key"""
        self.skip_ai_mode = True
        self.wizard().next()

    def validatePage(self):
        """Validate before proceeding to next page"""
        # If skipping AI, always allow
        if self.skip_ai_mode:
            # Clear saved key if not saving
            if not self.save_key_checkbox.isChecked():
                self.clear_saved_api_key()
            return True

        # If API key is entered but not tested
        if self.api_key_input.text().strip() and not self.api_validated:
            reply = QMessageBox.question(
                self,
                "API Key Not Tested",
                "You entered an API key but haven't tested it.\n\n"
                "Do you want to continue without testing?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return False

        # Save or clear API key based on checkbox
        if self.save_key_checkbox.isChecked() and self.api_key_input.text().strip():
            self.save_api_key()
        else:
            self.clear_saved_api_key()

        return True

    def get_api_key(self):
        """Get the entered API key"""
        if self.skip_ai_mode:
            return None
        return self.api_key_input.text().strip() if self.api_key_input.text().strip() else None


class AccessExportThread(QThread):
    """Background thread for exporting Access database"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(str, object)  # excel_path, dataframes_dict
    error = pyqtSignal(str)

    def __init__(self, mdb_file, output_file):
        super().__init__()
        self.mdb_file = mdb_file
        self.output_file = output_file

    def run(self):
        try:
            self.progress.emit("Connecting to Access database...")

            # Create connection string
            conn_str = (
                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                r"DBQ=" + self.mdb_file
            )
            quoted_conn_str = urllib.parse.quote_plus(conn_str)
            engine = sa.create_engine(f"access+pyodbc:///?odbc_connect={quoted_conn_str}")

            # Get table names
            inspector = inspect(engine)
            tables = inspector.get_table_names()

            self.progress.emit(f"Found {len(tables)} tables. Exporting...")

            # Export all tables
            dataframes = {}
            with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
                for idx, table in enumerate(tables, 1):
                    self.progress.emit(f"Exporting table {idx}/{len(tables)}: {table}")
                    df = pd.read_sql(f"SELECT * FROM [{table}]", engine)

                    # Clean sheet name
                    sheet_name = self.clean_sheet_name(table)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    dataframes[sheet_name] = df

            self.progress.emit("Export completed successfully!")
            self.finished.emit(self.output_file, dataframes)

        except Exception as e:
            self.error.emit(f"Error exporting Access database: {str(e)}")

    @staticmethod
    def clean_sheet_name(name):
        """Clean Excel sheet names"""
        invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
        for char in invalid_chars:
            name = name.replace(char, '')
        return name[:31]


class AIDetectionThread(QThread):
    """Background thread for AI column detection"""
    progress = pyqtSignal(str, int, int)  # message, current, total
    finished = pyqtSignal(dict)  # mappings
    error = pyqtSignal(str)

    def __init__(self, api_key, dataframes):
        super().__init__()
        self.api_key = api_key
        self.dataframes = dataframes

    def run(self):
        try:
            self.progress.emit("🔄 Preparing data...", 0, len(self.dataframes))

            client = Anthropic(api_key=self.api_key)

            # Process sheets in chunks (3 sheets at a time to avoid token limits)
            chunk_size = 3
            sheet_names = list(self.dataframes.keys())
            all_mappings = {}

            for chunk_idx in range(0, len(sheet_names), chunk_size):
                chunk_sheets = sheet_names[chunk_idx:chunk_idx + chunk_size]

                self.progress.emit(
                    f"🤖 Analyzing sheets {chunk_idx + 1}-{min(chunk_idx + chunk_size, len(sheet_names))} of {len(sheet_names)}...",
                    chunk_idx,
                    len(sheet_names)
                )

                # Prepare column information for this chunk
                sheets_info = []
                for sheet_name in chunk_sheets:
                    df = self.dataframes[sheet_name]
                    columns = df.columns.tolist()

                    # Filter out rows that are mostly empty (less than 30% of columns have data)
                    min_fields_threshold = max(2, len(columns) * 0.3)
                    non_empty_counts = df.notna().sum(axis=1)
                    df_filtered = df[non_empty_counts >= min_fields_threshold].copy()

                    if len(df_filtered) == 0:
                        df_filtered = df.copy()

                    # Increase sample size to 50 rows for better detection
                    sample_rows = []

                    # First 20 rows
                    if len(df_filtered) > 0:
                        sample_rows.extend(df_filtered.head(20).to_dict('records'))

                    # Random sample from middle (if we have more than 40 rows)
                    if len(df_filtered) > 40:
                        middle_sample = df_filtered.iloc[20:-10].sample(n=min(20, len(df_filtered) - 30), random_state=42)
                        sample_rows.extend(middle_sample.to_dict('records'))

                    # Last 10 rows (if we have more than 30 rows total)
                    if len(df_filtered) > 30:
                        sample_rows.extend(df_filtered.tail(10).to_dict('records'))

                    # Get basic statistics
                    stats = {
                        'total_rows': len(df),
                        'rows_with_data': len(df_filtered),
                        'non_empty_counts': {}
                    }

                    for col in columns:
                        non_empty = df_filtered[col].notna().sum()
                        stats['non_empty_counts'][col] = non_empty

                    sheets_info.append({
                        'sheet_name': sheet_name,
                        'columns': columns,
                        'sample_data': sample_rows,
                        'statistics': stats
                    })

                # Create prompt for Claude
                prompt = f"""Analyze the following Excel sheets and their columns. For each sheet, identify which columns correspond to:
1. MFG (Manufacturer name) - Look for manufacturer names like "Siemens", "ABB", "Schneider", etc.
2. MFG_PN (Manufacturer Part Number) - The primary part number from the manufacturer
3. MFG_PN_2 (Secondary/alternative Manufacturer Part Number) - An alternative or backup part number
4. Part_Number (Internal part number) - Internal reference numbers
5. Description (Part description) - Text description of the part

Here are the sheets with their columns, sample data (up to 50 rows), and statistics:

{json.dumps(sheets_info, indent=2, default=str)}

Note: Rows with little to no information (less than 30% of columns filled) have been filtered out.

Analyze the sample data carefully. Look at:
- Column names (they might have hints like "Mfg", "Manufacturer", "PN", "Part", "Description", etc.)
- Data patterns (manufacturer names vs part numbers vs descriptions)
- Data completeness (statistics show total_rows, rows_with_data after filtering, and non_empty_counts per column)
- Data consistency across the sample rows

For each sheet, return a JSON object with the mapping and confidence scores (0-100). Base confidence on:
- How well the column name matches the expected field
- How consistent the data pattern is with the expected field type
- How complete the data is (columns with mostly empty values should have lower confidence)

Format:
{{
  "sheet_name": {{
    "MFG": {{"column": "column_name or null", "confidence": 0-100}},
    "MFG_PN": {{"column": "column_name or null", "confidence": 0-100}},
    "MFG_PN_2": {{"column": "column_name or null", "confidence": 0-100}},
    "Part_Number": {{"column": "column_name or null", "confidence": 0-100}},
    "Description": {{"column": "column_name or null", "confidence": 0-100}}
  }}
}}

Only return the JSON, no other text."""

                # Call Claude API - use Claude Haiku 4.5 (fast and cost-effective)
                response = client.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=4096,
                    messages=[{"role": "user", "content": prompt}]
                )

                # Parse response
                response_text = response.content[0].text.strip()
                if response_text.startswith('```'):
                    response_text = response_text.split('```')[1]
                    if response_text.startswith('json'):
                        response_text = response_text[4:]
                    response_text = response_text.strip()

                chunk_mappings = json.loads(response_text)
                all_mappings.update(chunk_mappings)

            self.progress.emit("✅ Applying mappings...", len(sheet_names), len(sheet_names))
            self.finished.emit(all_mappings)

        except Exception as e:
            self.error.emit(str(e))


class DataSourcePage(QWizardPage):
    """Step 1: Choose between Access DB or existing Excel file"""

    def __init__(self):
        super().__init__()
        self.setTitle("Step 1: Select Data Source")
        self.setSubTitle("Choose your data source for EDM library processing")

        layout = QVBoxLayout()

        # Source selection
        source_group = QGroupBox("Data Source")
        source_layout = QVBoxLayout()

        self.access_radio = QRadioButton("Convert Access Database to Excel")
        self.excel_radio = QRadioButton("Use existing Excel file")
        self.access_radio.setChecked(True)

        source_layout.addWidget(self.access_radio)
        source_layout.addWidget(self.excel_radio)
        source_group.setLayout(source_layout)

        # Access DB file selection
        self.access_group = QGroupBox("Access Database File")
        access_layout = QHBoxLayout()
        self.access_path = QLineEdit()
        self.access_path.setPlaceholderText("Select .mdb or .accdb file...")
        access_browse = QPushButton("Browse...")
        access_browse.clicked.connect(self.browse_access)
        access_layout.addWidget(self.access_path)
        access_layout.addWidget(access_browse)
        self.access_group.setLayout(access_layout)

        # Excel file selection
        self.excel_group = QGroupBox("Excel File")
        excel_layout = QHBoxLayout()
        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("Select .xlsx file...")
        excel_browse = QPushButton("Browse...")
        excel_browse.clicked.connect(self.browse_excel)
        excel_layout.addWidget(self.excel_path)
        excel_layout.addWidget(excel_browse)
        self.excel_group.setLayout(excel_layout)
        self.excel_group.setEnabled(False)

        # Preview section
        preview_group = QGroupBox("Data Preview")
        preview_layout = QVBoxLayout()
        self.preview_label = QLabel("No data loaded")
        self.preview_table = QTableWidget()
        preview_layout.addWidget(self.preview_label)
        preview_layout.addWidget(self.preview_table)
        preview_group.setLayout(preview_layout)

        # Export button (only for Access DB)
        self.export_button = QPushButton("Export Access Database")
        self.export_button.clicked.connect(self.export_access)
        self.export_button.setEnabled(False)

        # Progress
        self.progress_label = QLabel("")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)

        # Connect radio buttons
        self.access_radio.toggled.connect(self.update_ui)
        self.excel_radio.toggled.connect(self.update_ui)
        self.access_path.textChanged.connect(self.validate_page)
        self.excel_path.textChanged.connect(self.validate_page)

        # Add widgets
        layout.addWidget(source_group)
        layout.addWidget(self.access_group)
        layout.addWidget(self.excel_group)
        layout.addWidget(self.export_button)
        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress_bar)
        layout.addWidget(preview_group, stretch=1)  # Preview fills available space

        self.setLayout(layout)

        # Store exported data
        self.exported_excel_path = None
        self.dataframes = {}

    def update_ui(self):
        """Update UI based on radio selection"""
        is_access = self.access_radio.isChecked()
        self.access_group.setEnabled(is_access)
        self.excel_group.setEnabled(not is_access)
        self.export_button.setVisible(is_access)
        self.validate_page()

    def browse_access(self):
        """Browse for Access database file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Access Database",
            "", "Access Database (*.mdb *.accdb);;All Files (*.*)"
        )
        if file_path:
            self.access_path.setText(file_path)

    def browse_excel(self):
        """Browse for Excel file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File",
            "", "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )
        if file_path:
            self.excel_path.setText(file_path)
            self.load_excel_preview(file_path)

    def export_access(self):
        """Export Access database to Excel"""
        access_file = self.access_path.text()
        if not access_file or not os.path.exists(access_file):
            QMessageBox.warning(self, "Invalid File", "Please select a valid Access database file.")
            return

        # Generate output filename
        output_file = str(Path(access_file).parent / f"{Path(access_file).stem}.xlsx")

        # Start export in background thread
        self.export_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(0)  # Indeterminate

        self.export_thread = AccessExportThread(access_file, output_file)
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
        self.export_button.setEnabled(True)
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
        self.export_button.setEnabled(True)
        QMessageBox.critical(self, "Export Error", error_msg)

    def load_excel_preview(self, excel_path):
        """Load and preview Excel file"""
        try:
            xl_file = pd.ExcelFile(excel_path)
            self.dataframes = {sheet: pd.read_excel(excel_path, sheet_name=sheet)
                             for sheet in xl_file.sheet_names}
            self.show_preview(self.dataframes)
            self.exported_excel_path = excel_path
            self.completeChanged.emit()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load Excel file: {str(e)}")

    def show_preview(self, dataframes):
        """Show preview of first 100 rows from first sheet"""
        if not dataframes:
            return

        first_sheet = list(dataframes.keys())[0]
        df = dataframes[first_sheet]

        # Limit to first 100 rows
        preview_df = df.head(100)

        self.preview_label.setText(
            f"Preview: {first_sheet} ({len(df)} total rows, showing first {len(preview_df)})"
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

    def validate_page(self):
        """Validate page completion"""
        self.export_button.setEnabled(
            self.access_radio.isChecked() and
            bool(self.access_path.text()) and
            os.path.exists(self.access_path.text())
        )

    def isComplete(self):
        """Check if page is complete"""
        if self.access_radio.isChecked():
            return self.exported_excel_path is not None
        else:
            return bool(self.excel_path.text()) and os.path.exists(self.excel_path.text())

    def get_excel_path(self):
        """Get the Excel file path"""
        if self.access_radio.isChecked():
            return self.exported_excel_path
        else:
            return self.excel_path.text()

    def get_dataframes(self):
        """Get the loaded dataframes"""
        return self.dataframes


class NoScrollComboBox(QComboBox):
    """ComboBox that ignores mouse wheel events"""
    def wheelEvent(self, event):
        event.ignore()


class ColumnMappingPage(QWizardPage):
    """Step 2: Map columns and configure combine options"""

    def __init__(self):
        super().__init__()
        self.setTitle("Step 2: Column Mapping & Combine Options")
        self.setSubTitle("Map columns for each sheet and configure combination settings")

        # Main layout with splitter for resizable sections
        main_layout = QHBoxLayout()

        # Create splitter for resizable panels
        splitter = QSplitter(Qt.Horizontal)

        # Left panel widget
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        # Bulk assign section
        bulk_group = QGroupBox("Bulk Assign Columns")
        bulk_layout = QHBoxLayout()

        bulk_layout.addWidget(QLabel("Column Type:"))
        self.bulk_column_type = NoScrollComboBox()
        self.bulk_column_type.addItems(["MFG", "MFG PN", "MFG PN 2", "Part Number", "Description"])
        bulk_layout.addWidget(self.bulk_column_type)

        bulk_layout.addWidget(QLabel("Assign To:"))
        self.bulk_column_name = NoScrollComboBox()
        bulk_layout.addWidget(self.bulk_column_name)

        self.bulk_apply_btn = QPushButton("Apply to All Sheets")
        self.bulk_apply_btn.clicked.connect(self.apply_bulk_assignment)
        bulk_layout.addWidget(self.bulk_apply_btn)

        bulk_group.setLayout(bulk_layout)
        left_layout.addWidget(bulk_group)

        # AI Auto-detect section
        ai_group = QGroupBox("🤖 AI-Powered Auto-Detection")
        ai_layout = QHBoxLayout()

        ai_info = QLabel("Let Claude AI automatically detect column mappings")
        ai_layout.addWidget(ai_info)

        self.ai_detect_btn = QPushButton("🤖 Auto-Detect Column Mappings")
        self.ai_detect_btn.clicked.connect(self.auto_detect_with_ai)
        self.ai_detect_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        ai_layout.addWidget(self.ai_detect_btn)

        self.ai_status = QLabel("")
        ai_layout.addWidget(self.ai_status)
        ai_layout.addStretch()

        ai_group.setLayout(ai_layout)
        left_layout.addWidget(ai_group)

        # Mapping table
        mapping_group = QGroupBox("Column Mapping")
        mapping_layout = QVBoxLayout()

        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(7)
        self.mapping_table.setHorizontalHeaderLabels([
            "Include", "Sheet Name", "MFG Column", "MFG PN Column", "MFG PN Column 2", "Part Number Column", "Description Column"
        ])
        self.mapping_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.mapping_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.mapping_table.setSelectionMode(QTableWidget.SingleSelection)
        self.mapping_table.itemSelectionChanged.connect(self.on_sheet_selected)

        # Save/Load configuration buttons
        config_layout = QHBoxLayout()
        self.save_config_btn = QPushButton("Save Mapping Config")
        self.load_config_btn = QPushButton("Load Mapping Config")
        self.save_config_btn.clicked.connect(self.save_configuration)
        self.load_config_btn.clicked.connect(self.load_configuration)
        config_layout.addWidget(self.save_config_btn)
        config_layout.addWidget(self.load_config_btn)
        config_layout.addStretch()
        mapping_layout.addLayout(config_layout)

        mapping_layout.addWidget(self.mapping_table)
        mapping_group.setLayout(mapping_layout)
        left_layout.addWidget(mapping_group, stretch=1)  # Mapping fills available space

        # Combine options
        combine_group = QGroupBox("Combine Options")
        combine_layout = QVBoxLayout()

        self.combine_checkbox = QCheckBox("Combine selected sheets into single 'Combined' sheet")
        self.combine_checkbox.toggled.connect(self.toggle_combine_options)

        self.filter_group = QGroupBox("Filter Conditions (rows must meet ALL checked conditions)")
        filter_layout = QVBoxLayout()

        self.filter_mfg = QCheckBox("MFG must not be empty")
        self.filter_mfg_pn = QCheckBox("MFG PN must not be empty")
        self.filter_part_number = QCheckBox("Part Number must not be empty")
        self.filter_description = QCheckBox("Description must not be empty")

        # TBD fill option
        self.fill_tbd_checkbox = QCheckBox("Fill empty MFG values with 'TBD'")
        self.fill_tbd_checkbox.setToolTip("If MFG PN is not empty but MFG is empty, set MFG to 'TBD'")

        filter_layout.addWidget(self.filter_mfg)
        filter_layout.addWidget(self.filter_mfg_pn)
        filter_layout.addWidget(self.filter_part_number)
        filter_layout.addWidget(self.filter_description)
        filter_layout.addWidget(self.fill_tbd_checkbox)

        self.filter_group.setLayout(filter_layout)
        self.filter_group.setEnabled(False)

        combine_layout.addWidget(self.combine_checkbox)
        combine_layout.addWidget(self.filter_group)
        combine_group.setLayout(combine_layout)
        left_layout.addWidget(combine_group)  # Combine stays at bottom, no stretch

        # Right panel widget - Preview
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        preview_group = QGroupBox("Sheet Preview")
        preview_layout = QVBoxLayout()

        self.preview_label = QLabel("Select a sheet to preview")
        self.preview_label.setStyleSheet("font-weight: bold;")
        preview_layout.addWidget(self.preview_label)

        self.preview_table = QTableWidget()
        preview_layout.addWidget(self.preview_table)

        preview_group.setLayout(preview_layout)
        right_layout.addWidget(preview_group)

        # Add widgets to splitter
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([700, 500])  # Initial sizes

        # Make splitter expand to fill available space
        splitter.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Add splitter to main layout
        main_layout.addWidget(splitter, stretch=1)  # Splitter fills available space

        self.setLayout(main_layout)

        self.sheet_mappings = {}
        self.dataframes = {}

    def initializePage(self):
        """Initialize page with data from previous step"""
        # Get API key from start page
        start_page = self.wizard().page(0)
        self.api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None

        # Enable/disable AI button based on API key availability
        if self.api_key and ANTHROPIC_AVAILABLE:
            self.ai_detect_btn.setEnabled(True)
            self.ai_status.setText("")
        else:
            self.ai_detect_btn.setEnabled(False)
            if not ANTHROPIC_AVAILABLE:
                self.ai_status.setText("⚠ Anthropic package not installed")
                self.ai_status.setStyleSheet("color: orange;")
            elif not self.api_key:
                self.ai_status.setText("ℹ No API key provided")
                self.ai_status.setStyleSheet("color: gray;")

        prev_page = self.wizard().page(1)  # DataSourcePage is now page 1
        dataframes = prev_page.get_dataframes()

        if not dataframes:
            excel_path = prev_page.get_excel_path()
            if excel_path:
                xl_file = pd.ExcelFile(excel_path)
                dataframes = {sheet: pd.read_excel(excel_path, sheet_name=sheet)
                            for sheet in xl_file.sheet_names}

        self.dataframes = dataframes
        self.populate_mapping_table(dataframes)
        self.populate_bulk_column_names()

    def populate_bulk_column_names(self):
        """Populate bulk assign dropdown with all available columns"""
        all_columns = set()
        for df in self.dataframes.values():
            all_columns.update(df.columns.tolist())

        self.bulk_column_name.clear()
        self.bulk_column_name.addItem("")
        self.bulk_column_name.addItems(sorted(all_columns))

    def apply_bulk_assignment(self):
        """Apply bulk column assignment to all sheets"""
        column_type = self.bulk_column_type.currentText()
        column_name = self.bulk_column_name.currentText()

        if not column_name:
            QMessageBox.warning(self, "No Selection", "Please select a column name to assign.")
            return

        # Map column type to table column index
        # Columns: Include(0), Sheet Name(1), MFG(2), MFG PN(3), MFG PN 2(4), Part Number(5), Description(6)
        type_map = {
            "MFG": 2,
            "MFG PN": 3,
            "MFG PN 2": 4,
            "Part Number": 5,
            "Description": 6
        }
        col_idx = type_map.get(column_type)

        if col_idx is None:
            return

        # Apply to all rows
        for row in range(self.mapping_table.rowCount()):
            combo = self.mapping_table.cellWidget(row, col_idx)
            if combo:
                # Check if this column exists in this sheet
                index = combo.findText(column_name)
                if index >= 0:
                    combo.setCurrentIndex(index)

        QMessageBox.information(self, "Bulk Assign Complete",
                               f"Assigned '{column_name}' to {column_type} for all applicable sheets.")

    def on_sheet_selected(self):
        """Handle sheet selection to show preview"""
        selected_rows = self.mapping_table.selectedIndexes()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        sheet_item = self.mapping_table.item(row, 1)
        if not sheet_item:
            return

        sheet_name = sheet_item.text()
        if sheet_name in self.dataframes:
            self.show_sheet_preview(sheet_name, self.dataframes[sheet_name])

    def show_sheet_preview(self, sheet_name, df):
        """Show preview of selected sheet"""
        preview_df = df.head(100)

        self.preview_label.setText(
            f"Preview: {sheet_name} ({len(df)} total rows, showing first {len(preview_df)})"
        )

        # Populate preview table
        self.preview_table.setRowCount(len(preview_df))
        self.preview_table.setColumnCount(len(preview_df.columns))
        self.preview_table.setHorizontalHeaderLabels(preview_df.columns.tolist())

        for i in range(len(preview_df)):
            for j in range(len(preview_df.columns)):
                value = preview_df.iloc[i, j]
                item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.preview_table.setItem(i, j, item)

        self.preview_table.resizeColumnsToContents()

    def populate_mapping_table(self, dataframes):
        """Populate the mapping table with sheets and column dropdowns"""
        self.mapping_table.setRowCount(len(dataframes))

        for row, (sheet_name, df) in enumerate(dataframes.items()):
            # Include checkbox
            include_checkbox = QCheckBox()
            include_checkbox.setChecked(True)
            include_widget = QWidget()
            include_layout = QHBoxLayout(include_widget)
            include_layout.addWidget(include_checkbox)
            include_layout.setAlignment(Qt.AlignCenter)
            include_layout.setContentsMargins(0, 0, 0, 0)
            self.mapping_table.setCellWidget(row, 0, include_widget)

            # Sheet name
            sheet_item = QTableWidgetItem(sheet_name)
            sheet_item.setFlags(sheet_item.flags() & ~Qt.ItemIsEditable)
            self.mapping_table.setItem(row, 1, sheet_item)

            columns = [""] + df.columns.tolist()

            # Create dropdowns for each mapping type
            for col_idx, mapping_type in enumerate(["MFG", "MFG_PN", "MFG_PN_2", "Part_Number", "Description"], 2):
                combo = NoScrollComboBox()
                combo.addItems(columns)
                combo.setProperty("sheet_name", sheet_name)
                combo.setProperty("mapping_type", mapping_type)
                self.mapping_table.setCellWidget(row, col_idx, combo)

    def toggle_combine_options(self, checked):
        """Enable/disable combine filter options"""
        self.filter_group.setEnabled(checked)

    def get_included_sheets(self):
        """Get list of sheets that are checked for inclusion"""
        included = []
        for row in range(self.mapping_table.rowCount()):
            include_widget = self.mapping_table.cellWidget(row, 0)
            checkbox = include_widget.findChild(QCheckBox)
            if checkbox and checkbox.isChecked():
                sheet_name = self.mapping_table.item(row, 1).text()
                included.append(sheet_name)
        return included

    def save_configuration(self):
        """Save current column mappings to a JSON file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Mapping Configuration",
            "mapping_config.json", "JSON Files (*.json);;All Files (*.*)"
        )

        if not file_path:
            return

        config = {
            'mappings': self.get_mappings(),
            'version': '1.0'
        }

        try:
            with open(file_path, 'w') as f:
                json.dump(config, f, indent=2)
            QMessageBox.information(self, "Success", f"Configuration saved to:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save configuration:\n{str(e)}")

    def load_configuration(self):
        """Load column mappings from a JSON file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Load Mapping Configuration",
            "", "JSON Files (*.json);;All Files (*.*)"
        )

        if not file_path:
            return

        try:
            with open(file_path, 'r') as f:
                config = json.load(f)

            mappings = config.get('mappings', {})

            # Apply loaded mappings to table
            for row in range(self.mapping_table.rowCount()):
                sheet_name = self.mapping_table.item(row, 1).text()

                if sheet_name in mappings:
                    sheet_config = mappings[sheet_name]

                    # Set each dropdown
                    for col_idx, key in enumerate(['MFG', 'MFG_PN', 'MFG_PN_2', 'Part_Number', 'Description'], 2):
                        combo = self.mapping_table.cellWidget(row, col_idx)
                        if combo and key in sheet_config:
                            value = sheet_config[key]
                            index = combo.findText(value)
                            if index >= 0:
                                combo.setCurrentIndex(index)

            QMessageBox.information(self, "Success", "Configuration loaded successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load configuration:\n{str(e)}")

    def auto_detect_with_ai(self):
        """Use Claude AI to automatically detect column mappings"""
        if not self.api_key or not ANTHROPIC_AVAILABLE:
            QMessageBox.warning(
                self,
                "AI Not Available",
                "Claude AI is not available. Please provide an API key in the Start page."
            )
            return

        # Create progress bar
        self.ai_progress = QProgressBar(self)
        self.ai_progress.setMinimum(0)
        self.ai_progress.setMaximum(len(self.dataframes))
        self.ai_progress.setValue(0)

        # Add progress bar to AI section temporarily
        ai_group = self.ai_detect_btn.parent()
        ai_layout = ai_group.layout()
        ai_layout.addWidget(self.ai_progress)

        # Disable controls but keep UI responsive
        self.ai_detect_btn.setEnabled(False)
        self.bulk_apply_btn.setEnabled(False)
        self.save_config_btn.setEnabled(False)
        self.load_config_btn.setEnabled(False)

        # Disable all dropdowns in the mapping table
        for row in range(self.mapping_table.rowCount()):
            for col in range(2, 7):  # Columns 2-6 are the dropdowns
                combo = self.mapping_table.cellWidget(row, col)
                if combo:
                    combo.setEnabled(False)

        self.ai_status.setText("🔄 Starting AI analysis...")
        self.ai_status.setStyleSheet("color: blue;")

        # Create and start AI detection thread
        self.ai_thread = AIDetectionThread(self.api_key, self.dataframes)
        self.ai_thread.progress.connect(self.on_ai_progress)
        self.ai_thread.finished.connect(self.on_ai_finished)
        self.ai_thread.error.connect(self.on_ai_error)
        self.ai_thread.start()

    def on_ai_progress(self, message, current, total):
        """Update progress during AI detection"""
        self.ai_status.setText(message)
        self.ai_progress.setValue(current)

    def on_ai_finished(self, all_mappings):
        """Handle AI detection completion"""
        # Apply mappings to table with confidence indicators
        self.ai_status.setText("✅ Applying mappings...")

        for row in range(self.mapping_table.rowCount()):
            sheet_name = self.mapping_table.item(row, 1).text()

            if sheet_name in all_mappings:
                sheet_mapping = all_mappings[sheet_name]

                # Column index mapping
                col_map = {
                    'MFG': 2,
                    'MFG_PN': 3,
                    'MFG_PN_2': 4,
                    'Part_Number': 5,
                    'Description': 6
                }

                for field, col_idx in col_map.items():
                    if field in sheet_mapping:
                        mapping_info = sheet_mapping[field]
                        column_name = mapping_info.get('column')
                        confidence = mapping_info.get('confidence', 0)

                        combo = self.mapping_table.cellWidget(row, col_idx)
                        if combo and column_name:
                            # Find and set the column
                            index = combo.findText(column_name)
                            if index >= 0:
                                combo.setCurrentIndex(index)

                                # Apply color coding based on confidence
                                if confidence >= 80:
                                    # High confidence - green
                                    combo.setStyleSheet("background-color: #c8e6c9;")
                                elif confidence >= 50:
                                    # Medium confidence - yellow
                                    combo.setStyleSheet("background-color: #fff9c4;")
                                else:
                                    # Low confidence - orange
                                    combo.setStyleSheet("background-color: #ffe0b2;")

                                # Add tooltip with confidence score
                                combo.setToolTip(f"AI Confidence: {confidence}%")

        self.ai_status.setText("✓ Auto-detection complete!")
        self.ai_status.setStyleSheet("color: green;")

        # Re-enable controls
        self.ai_detect_btn.setEnabled(True)
        self.bulk_apply_btn.setEnabled(True)
        self.save_config_btn.setEnabled(True)
        self.load_config_btn.setEnabled(True)

        # Re-enable all dropdowns
        for row in range(self.mapping_table.rowCount()):
            for col in range(2, 7):
                combo = self.mapping_table.cellWidget(row, col)
                if combo:
                    combo.setEnabled(True)

        # Remove progress bar
        ai_group = self.ai_detect_btn.parent()
        ai_layout = ai_group.layout()
        ai_layout.removeWidget(self.ai_progress)
        self.ai_progress.deleteLater()

        # Show legend
        QMessageBox.information(
            self,
            "AI Detection Complete",
            f"Column mappings detected for {len(all_mappings)} sheets using Claude Haiku 4.5!\n\n"
            "Color coding:\n"
            "🟢 Green: High confidence (80%+)\n"
            "🟡 Yellow: Medium confidence (50-79%)\n"
            "🟠 Orange: Low confidence (<50%)\n\n"
            "Please review and adjust as needed. "
            "Hover over dropdowns to see confidence scores."
        )

    def on_ai_error(self, error_msg):
        """Handle AI detection error"""
        self.ai_status.setText(f"✗ Error: {error_msg[:30]}")
        self.ai_status.setStyleSheet("color: red;")

        # Re-enable controls
        self.ai_detect_btn.setEnabled(True)
        self.bulk_apply_btn.setEnabled(True)
        self.save_config_btn.setEnabled(True)
        self.load_config_btn.setEnabled(True)

        # Re-enable all dropdowns
        for row in range(self.mapping_table.rowCount()):
            for col in range(2, 7):
                combo = self.mapping_table.cellWidget(row, col)
                if combo:
                    combo.setEnabled(True)

        # Remove progress bar
        ai_group = self.ai_detect_btn.parent()
        ai_layout = ai_group.layout()
        ai_layout.removeWidget(self.ai_progress)
        self.ai_progress.deleteLater()

        QMessageBox.critical(
            self,
            "AI Detection Failed",
            f"Failed to auto-detect columns:\n{error_msg}"
        )

    def get_mappings(self):
        """Get all column mappings"""
        mappings = {}

        for row in range(self.mapping_table.rowCount()):
            sheet_name = self.mapping_table.item(row, 1).text()

            mappings[sheet_name] = {
                'MFG': self.mapping_table.cellWidget(row, 2).currentText(),
                'MFG_PN': self.mapping_table.cellWidget(row, 3).currentText(),
                'MFG_PN_2': self.mapping_table.cellWidget(row, 4).currentText(),
                'Part_Number': self.mapping_table.cellWidget(row, 5).currentText(),
                'Description': self.mapping_table.cellWidget(row, 6).currentText()
            }

        return mappings

    def should_combine(self):
        """Check if sheets should be combined"""
        return self.combine_checkbox.isChecked()

    def get_filter_conditions(self):
        """Get filter conditions for combining"""
        return {
            'MFG': self.filter_mfg.isChecked(),
            'MFG_PN': self.filter_mfg_pn.isChecked(),
            'Part_Number': self.filter_part_number.isChecked(),
            'Description': self.filter_description.isChecked(),
            'Fill_TBD': self.fill_tbd_checkbox.isChecked()
        }

    def validatePage(self):
        """Validate mappings and perform combine if requested"""
        mappings = self.get_mappings()
        included_sheets = self.get_included_sheets()

        if not included_sheets:
            QMessageBox.warning(self, "No Sheets Selected",
                              "Please select at least one sheet to include.")
            return False

        # Check if at least one included sheet has MFG and MFG_PN mapped
        has_valid_mapping = False
        for sheet_name in included_sheets:
            if sheet_name in mappings:
                sheet_mappings = mappings[sheet_name]
                if sheet_mappings['MFG'] and sheet_mappings['MFG_PN']:
                    has_valid_mapping = True
                    break

        if not has_valid_mapping:
            reply = QMessageBox.warning(
                self, "Missing Mappings",
                "No selected sheets have both MFG and MFG PN columns mapped.\n"
                "XML generation may not work properly.\n\n"
                "Do you want to continue anyway?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return False

        # Perform combine if requested
        if self.should_combine():
            try:
                self.combine_sheets()
            except Exception as e:
                QMessageBox.critical(self, "Combine Error", f"Failed to combine sheets: {str(e)}")
                return False

        return True

    def combine_sheets(self):
        """Combine sheets based on mappings and filters"""
        prev_page = self.wizard().page(0)
        excel_path = prev_page.get_excel_path()

        mappings = self.get_mappings()
        filters = self.get_filter_conditions()
        included_sheets = self.get_included_sheets()

        combined_data = []

        for sheet_name in included_sheets:
            if sheet_name not in self.dataframes:
                continue

            df = self.dataframes[sheet_name]
            df_copy = df.copy()
            df_copy['Source_Sheet'] = sheet_name

            # Get mapped columns
            sheet_mapping = mappings[sheet_name]

            # Rename columns to standard names
            rename_dict = {}
            for key, col_name in sheet_mapping.items():
                if col_name and key != 'MFG_PN_2':  # MFG_PN_2 is handled separately
                    rename_dict[col_name] = key

            if rename_dict:
                df_copy = df_copy.rename(columns=rename_dict)

            # Handle MFG PN fallback: if MFG_PN is empty, use MFG_PN_2
            if 'MFG_PN' in df_copy.columns and sheet_mapping.get('MFG_PN_2'):
                mfg_pn_2_col = sheet_mapping['MFG_PN_2']
                if mfg_pn_2_col in df.columns:
                    # Fill empty MFG_PN with values from MFG_PN_2
                    empty_mask = df_copy['MFG_PN'].isna() | (df_copy['MFG_PN'].astype(str).str.strip() == '')
                    df_copy.loc[empty_mask, 'MFG_PN'] = df[mfg_pn_2_col]

            # Handle TBD fill: if MFG_PN is not empty but MFG is empty, set MFG to 'TBD'
            if filters.get('Fill_TBD') and 'MFG' in df_copy.columns and 'MFG_PN' in df_copy.columns:
                mfg_pn_filled = df_copy['MFG_PN'].notna() & (df_copy['MFG_PN'].astype(str).str.strip() != '')
                mfg_empty = df_copy['MFG'].isna() | (df_copy['MFG'].astype(str).str.strip() == '')
                df_copy.loc[mfg_pn_filled & mfg_empty, 'MFG'] = 'TBD'

            # Apply filters
            mask = pd.Series([True] * len(df_copy))

            if filters['MFG'] and 'MFG' in df_copy.columns:
                mask &= df_copy['MFG'].notna() & (df_copy['MFG'].astype(str).str.strip() != '')

            if filters['MFG_PN'] and 'MFG_PN' in df_copy.columns:
                mask &= df_copy['MFG_PN'].notna() & (df_copy['MFG_PN'].astype(str).str.strip() != '')

            if filters['Part_Number'] and 'Part_Number' in df_copy.columns:
                mask &= df_copy['Part_Number'].notna() & (df_copy['Part_Number'].astype(str).str.strip() != '')

            if filters['Description'] and 'Description' in df_copy.columns:
                mask &= df_copy['Description'].notna() & (df_copy['Description'].astype(str).str.strip() != '')

            df_filtered = df_copy[mask]

            if len(df_filtered) > 0:
                combined_data.append(df_filtered)

        if combined_data:
            combined_df = pd.concat(combined_data, ignore_index=True)

            # Write combined sheet to original Excel file
            with pd.ExcelFile(excel_path) as xls:
                existing_sheets = {sheet: pd.read_excel(excel_path, sheet_name=sheet)
                                 for sheet in xls.sheet_names}

            existing_sheets['Combined'] = combined_df

            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                for sheet_name, df in existing_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            QMessageBox.information(
                self, "Combine Complete",
                f"Successfully combined {len(included_sheets)} sheets into 'Combined' sheet.\n"
                f"Total rows: {len(combined_df)}"
            )


class XMLGenerationPage(QWizardPage):
    """Step 3: Generate XML files"""

    def __init__(self):
        super().__init__()
        self.setTitle("Step 3: XML Generation")
        self.setSubTitle("Configure and generate XML files for EDM Library Creator")

        layout = QVBoxLayout()

        # Project settings
        settings_group = QGroupBox("Project Settings")
        settings_layout = QGridLayout()

        settings_layout.addWidget(QLabel("Project Name:"), 0, 0)
        self.project_name = QLineEdit("VarTrainingLab")
        settings_layout.addWidget(self.project_name, 0, 1)

        settings_layout.addWidget(QLabel("Catalog:"), 1, 0)
        self.catalog = QLineEdit("VV")
        settings_layout.addWidget(self.catalog, 1, 1)

        settings_group.setLayout(settings_layout)

        # Output settings
        output_group = QGroupBox("Output Settings")
        output_layout = QVBoxLayout()

        location_layout = QHBoxLayout()
        location_layout.addWidget(QLabel("Output Location:"))
        self.output_path = QLineEdit()
        self.output_path.setReadOnly(True)
        location_layout.addWidget(self.output_path)

        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_output)
        location_layout.addWidget(browse_btn)

        output_layout.addLayout(location_layout)
        output_group.setLayout(output_layout)

        # TBD option
        self.tbd_checkbox = QCheckBox("If MFG PN exists but MFG is empty, set MFG to 'TBD' in XML")
        self.tbd_checkbox.setChecked(True)

        # Generate button
        self.generate_button = QPushButton("Generate XML Files")
        self.generate_button.clicked.connect(self.generate_xml)

        # Status
        self.status_label = QLabel("")

        # Summary
        summary_group = QGroupBox("Generation Summary")
        summary_layout = QVBoxLayout()
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        summary_layout.addWidget(self.summary_text)
        summary_group.setLayout(summary_layout)

        layout.addWidget(settings_group)
        layout.addWidget(output_group)
        layout.addWidget(self.tbd_checkbox)
        layout.addWidget(self.generate_button)
        layout.addWidget(self.status_label)
        layout.addWidget(summary_group, stretch=1)  # Summary fills available space

        self.setLayout(layout)

        self.xml_generated = False

    def initializePage(self):
        """Initialize with default output path"""
        prev_page = self.wizard().page(0)
        excel_path = prev_page.get_excel_path()

        if excel_path:
            # Create timestamped folder for this run
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_dir = Path(excel_path).parent
            output_folder = base_dir / f"EDM_Output_{timestamp}"
            output_folder.mkdir(exist_ok=True)

            self.output_path.setText(str(output_folder))
            self.timestamp = timestamp  # Store for later use

    def browse_output(self):
        """Browse for output directory"""
        directory = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if directory:
            self.output_path.setText(directory)

    def generate_xml(self):
        """Generate MFG and MFGPN XML files and copy all files to output folder"""
        try:
            import shutil

            prev_page_0 = self.wizard().page(0)
            prev_page_1 = self.wizard().page(1)

            excel_path = prev_page_0.get_excel_path()
            dataframes = prev_page_0.get_dataframes()
            mappings = prev_page_1.get_mappings()
            output_dir = Path(self.output_path.text())

            # Check if Combined sheet should be used
            if prev_page_1.should_combine():
                # Reload to get Combined sheet
                xl_file = pd.ExcelFile(excel_path)
                if 'Combined' in xl_file.sheet_names:
                    combined_df = pd.read_excel(excel_path, sheet_name='Combined')
                    # Combined sheet already has standardized column names
                    self.generate_xml_from_df(combined_df, excel_path,
                                             {'MFG': 'MFG', 'MFG_PN': 'MFG_PN',
                                              'Description': 'Description'})
                else:
                    QMessageBox.warning(self, "Warning", "Combined sheet not found. Using individual sheets.")
                    self.generate_xml_from_sheets(dataframes, excel_path, mappings)
            else:
                self.generate_xml_from_sheets(dataframes, excel_path, mappings)

            # Copy Excel file to output folder
            excel_filename = Path(excel_path).name
            dest_excel = output_dir / excel_filename
            if Path(excel_path) != dest_excel:
                shutil.copy2(excel_path, dest_excel)

            # Save configuration file to output folder
            config_file = output_dir / "column_mapping_config.json"
            config = {
                'mappings': mappings,
                'timestamp': self.timestamp,
                'version': '1.0'
            }
            with open(config_file, 'w') as f:
                json.dump(config, f, indent=2)

            self.xml_generated = True
            self.completeChanged.emit()

        except Exception as e:
            QMessageBox.critical(self, "Generation Error", f"Failed to generate XML files: {str(e)}")

    def generate_xml_from_sheets(self, dataframes, excel_path, mappings):
        """Generate XML from multiple sheets"""
        prev_page_1 = self.wizard().page(1)
        included_sheets = prev_page_1.get_included_sheets()

        all_mfg = set()
        all_mfgpn = []

        for sheet_name, df in dataframes.items():
            # Skip sheets that are not included
            if sheet_name not in included_sheets:
                continue

            mapping = mappings[sheet_name]

            if not mapping['MFG'] or not mapping['MFG_PN']:
                continue

            # Extract data
            mfg_col = mapping['MFG']
            mfgpn_col = mapping['MFG_PN']
            desc_col = mapping.get('Description', '')

            df_filtered = df[[mfg_col, mfgpn_col]].copy()
            if desc_col:
                df_filtered['Description'] = df[desc_col]
            else:
                df_filtered['Description'] = "This is the PN description."

            df_filtered.columns = ['MFG', 'MFG_PN', 'Description']

            # Handle TBD option
            if self.tbd_checkbox.isChecked():
                mask = (df_filtered['MFG_PN'].notna()) & (df_filtered['MFG_PN'].astype(str).str.strip() != '')
                df_filtered.loc[mask & (df_filtered['MFG'].isna() | (df_filtered['MFG'].astype(str).str.strip() == '')), 'MFG'] = 'TBD'

            # Collect unique MFG
            mfg_values = df_filtered['MFG'].dropna()
            all_mfg.update(mfg_values.astype(str).str.strip().unique())

            # Collect MFG/MFGPN pairs
            df_pairs = df_filtered[['MFG', 'MFG_PN', 'Description']].dropna(subset=['MFG', 'MFG_PN'])
            for _, row in df_pairs.iterrows():
                all_mfgpn.append({
                    'MFG': str(row['MFG']).strip(),
                    'MFG_PN': str(row['MFG_PN']).strip(),
                    'Description': str(row['Description']) if pd.notna(row['Description']) else "This is the PN description."
                })

        # Generate XML files
        self.create_xml_files(all_mfg, all_mfgpn, excel_path)

    def generate_xml_from_df(self, df, excel_path, mapping):
        """Generate XML from a single dataframe"""
        all_mfg = set()
        all_mfgpn = []

        mfg_col = mapping['MFG']
        mfgpn_col = mapping['MFG_PN']
        desc_col = mapping.get('Description', '')

        df_copy = df.copy()

        # Handle TBD option
        if self.tbd_checkbox.isChecked():
            mask = (df_copy[mfgpn_col].notna()) & (df_copy[mfgpn_col].astype(str).str.strip() != '')
            df_copy.loc[mask & (df_copy[mfg_col].isna() | (df_copy[mfg_col].astype(str).str.strip() == '')), mfg_col] = 'TBD'

        # Collect unique MFG
        mfg_values = df_copy[mfg_col].dropna()
        all_mfg.update(mfg_values.astype(str).str.strip().unique())

        # Collect MFG/MFGPN pairs
        for _, row in df_copy.iterrows():
            if pd.notna(row[mfg_col]) and pd.notna(row[mfgpn_col]):
                desc = row[desc_col] if desc_col and pd.notna(row[desc_col]) else "This is the PN description."
                all_mfgpn.append({
                    'MFG': str(row[mfg_col]).strip(),
                    'MFG_PN': str(row[mfgpn_col]).strip(),
                    'Description': str(desc)
                })

        # Generate XML files
        self.create_xml_files(all_mfg, all_mfgpn, excel_path)

    def create_xml_files(self, manufacturers, mfgpn_data, excel_path):
        """Create MFG and MFGPN XML files"""
        output_dir = Path(self.output_path.text())
        base_name = Path(excel_path).stem
        project_name = self.project_name.text()
        catalog = self.catalog.text()

        mfg_xml_path = output_dir / f"{base_name}_MFG.xml"
        mfgpn_xml_path = output_dir / f"{base_name}_MFGPN.xml"

        # Create MFG XML
        mfg_count = self.create_mfg_xml(manufacturers, mfg_xml_path, project_name, catalog)

        # Create MFGPN XML
        mfgpn_count = self.create_mfgpn_xml(mfgpn_data, mfgpn_xml_path, project_name, catalog)

        # Build comprehensive summary
        summary = f"✓ All Files Generated Successfully!\n\n"
        summary += f"Output Folder: {output_dir}\n"
        summary += f"{'-' * 60}\n\n"

        # List all files in output folder
        summary += "Files Created:\n"
        summary += f"  1. {Path(excel_path).name}\n"
        summary += f"      - Excel workbook with all data\n"
        summary += f"  2. column_mapping_config.json\n"
        summary += f"      - Column mapping configuration (reusable)\n"
        summary += f"  3. {mfg_xml_path.name}\n"
        summary += f"      - Manufacturers ({mfg_count} entries)\n"
        summary += f"  4. {mfgpn_xml_path.name}\n"
        summary += f"      - Manufacturer Part Numbers ({mfgpn_count} entries)\n\n"

        summary += f"All files are saved in:\n{output_dir}"

        self.summary_text.setText(summary)
        self.status_label.setText("✓ All files generated and saved successfully")
        self.status_label.setStyleSheet("color: green; font-weight: bold;")

        QMessageBox.information(self, "Success",
                               f"All files generated successfully!\n\n"
                               f"Output folder:\n{output_dir}\n\n"
                               f"Files created:\n"
                               f"- Excel workbook\n"
                               f"- Config file\n"
                               f"- MFG XML ({mfg_count} manufacturers)\n"
                               f"- MFGPN XML ({mfgpn_count} part numbers)")

    def escape_xml(self, text):
        """Escape special XML characters"""
        if pd.isna(text):
            return ""
        text = str(text)
        text = text.replace("&", "&amp;")
        text = text.replace("<", "&lt;")
        text = text.replace(">", "&gt;")
        text = text.replace('"', "&quot;")
        text = text.replace("'", "&apos;")
        return text

    def create_mfg_xml(self, manufacturers, output_file, project_name, catalog):
        """Create MFG XML file"""
        manufacturers = sorted([m for m in manufacturers if m])

        root = ET.Element('data')

        for mfg in manufacturers:
            obj = ET.SubElement(root, 'object')
            obj.set('objectid', self.escape_xml(mfg))
            obj.set('catalog', catalog)
            obj.set('class', '090')

            field1 = ET.SubElement(obj, 'field')
            field1.set('id', '090obj_skn')
            field1.text = catalog

            field2 = ET.SubElement(obj, 'field')
            field2.set('id', '090obj_id')
            field2.text = self.escape_xml(mfg)

            field3 = ET.SubElement(obj, 'field')
            field3.set('id', '090her_name')
            field3.text = self.escape_xml(mfg)

        self.save_xml(root, output_file, project_name)
        return len(manufacturers)

    def create_mfgpn_xml(self, mfgpn_data, output_file, project_name, catalog):
        """Create MFGPN XML file"""
        # Remove duplicates
        unique_pairs = {}
        for item in mfgpn_data:
            key = (item['MFG'], item['MFG_PN'])
            if key not in unique_pairs:
                unique_pairs[key] = item['Description']

        root = ET.Element('data')

        for (mfg, mfg_pn), description in unique_pairs.items():
            objectid = f"{mfg}:{mfg_pn}"

            obj = ET.SubElement(root, 'object')
            obj.set('objectid', self.escape_xml(objectid))
            obj.set('class', '060')

            field1 = ET.SubElement(obj, 'field')
            field1.set('id', '060partnumber')
            field1.text = self.escape_xml(mfg_pn)

            field2 = ET.SubElement(obj, 'field')
            field2.set('id', '060mfgref')
            field2.text = self.escape_xml(mfg)

            field3 = ET.SubElement(obj, 'field')
            field3.set('id', '060komp_name')
            field3.text = self.escape_xml(description)

        self.save_xml(root, output_file, project_name)
        return len(unique_pairs)

    def save_xml(self, root, output_file, project_name):
        """Format and save XML file"""
        xml_str = ET.tostring(root, encoding='utf-8', method='xml')
        dom = minidom.parseString(xml_str)

        comment_lines = [
            f'Created By: EDM Library Creator v1.7.000.0130',
            f'DDP Project: {project_name}',
            f'Date: {datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")}'
        ]

        xml_lines = ['<?xml version="1.0" encoding="utf-8" standalone="yes"?>']
        for comment in comment_lines:
            xml_lines.append(f'<!--{comment}-->')

        formatted = dom.toprettyxml(indent='  ', encoding='utf-8').decode('utf-8')
        xml_content = '\n'.join(formatted.split('\n')[1:])

        final_xml = '\n'.join(xml_lines) + '\n' + xml_content

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(final_xml)

    def isComplete(self):
        """Check if page is complete"""
        return self.xml_generated


class EDMWizard(QWizard):
    """Main wizard window"""

    def __init__(self):
        super().__init__()

        self.setWindowTitle("EDM Library Wizard")
        self.setWizardStyle(QWizard.ModernStyle)

        # Set window flags to enable minimize, maximize, and close buttons
        self.setWindowFlags(
            Qt.Window |
            Qt.WindowCloseButtonHint |
            Qt.WindowMinimizeButtonHint |
            Qt.WindowMaximizeButtonHint
        )

        # Add pages
        self.start_page = StartPage()
        self.data_source_page = DataSourcePage()
        self.column_mapping_page = ColumnMappingPage()
        self.xml_generation_page = XMLGenerationPage()

        self.addPage(self.start_page)
        self.addPage(self.data_source_page)
        self.addPage(self.column_mapping_page)
        self.addPage(self.xml_generation_page)

        # Customize buttons
        self.setButtonText(QWizard.FinishButton, "Finish")

        # Set size constraints and enable resizing
        self.setMinimumSize(1000, 750)
        self.resize(1200, 800)  # Default size
        self.setMaximumSize(16777215, 16777215)  # Remove maximum size constraint

        # Set size policy to allow expansion
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Apply styling
        self.setStyleSheet("""
            QWizard {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QPushButton {
                padding: 5px 15px;
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 3px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #cccccc;
                border-radius: 3px;
            }
            QTableWidget {
                border: 1px solid #cccccc;
                gridline-color: #e0e0e0;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                padding: 5px;
                border: 1px solid #cccccc;
                font-weight: bold;
            }
        """)


def main():
    """Main entry point"""
    app = QApplication(sys.argv)

    # Set application style
    app.setStyle('Fusion')

    wizard = EDMWizard()
    wizard.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
