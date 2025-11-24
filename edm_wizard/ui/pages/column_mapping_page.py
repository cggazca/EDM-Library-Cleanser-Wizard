"""
Column Mapping Page: AI-Assisted Column Detection
"""

import sys
import os
from pathlib import Path
import json
import pandas as pd

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QLineEdit,
        QPushButton, QFileDialog, QComboBox, QCheckBox, QTableWidget,
        QTableWidgetItem, QHeaderView, QProgressBar, QMessageBox, QWidget,
        QSplitter, QScrollArea, QSpinBox, QSizePolicy
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

try:
    from anthropic import Anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

from edm_wizard.workers.threads import AIDetectionThread, SheetDetectionWorker
from edm_wizard.ui.components.custom_widgets import NoScrollComboBox



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

        # Style the splitter handle to make it more noticeable
        splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #d0d0d0;
                border: 1px solid #a0a0a0;
                width: 4px;
            }
            QSplitter::handle:horizontal {
                width: 4px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #e0e0e0, stop:0.5 #c0c0c0, stop:1 #e0e0e0);
                border-left: 1px solid #a0a0a0;
                border-right: 1px solid #a0a0a0;
            }
            QSplitter::handle:hover {
                background-color: #0078d7;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #4da6ff, stop:0.5 #0078d7, stop:1 #4da6ff);
            }
        """)

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
        self.mapping_table.setColumnCount(8)
        self.mapping_table.setHorizontalHeaderLabels([
            "Include", "Sheet Name", "MFG Column", "MFG PN Column", "MFG PN Column 2", "Part Number Column", "Description Column", "Actions"
        ])
        self.mapping_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.mapping_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.mapping_table.setSelectionMode(QTableWidget.SingleSelection)
        self.mapping_table.setSortingEnabled(True)  # Enable sorting
        self.mapping_table.itemSelectionChanged.connect(self.on_sheet_selected)

        # Save/Load configuration and Toggle Select All button
        config_layout = QHBoxLayout()

        # Toggle Select All button
        self.toggle_select_btn = QPushButton("✓ Select All")
        self.toggle_select_btn.setCheckable(True)
        self.toggle_select_btn.clicked.connect(self.toggle_all_sheets)
        config_layout.addWidget(self.toggle_select_btn)

        config_layout.addSpacing(20)  # Add spacing between button groups

        # Save/Load configuration buttons
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

        # Combine options (mandatory - always enabled)
        combine_group = QGroupBox("Combine Options")
        combine_layout = QVBoxLayout()

        # Info label explaining that combining is mandatory
        combine_info = QLabel("ℹ️ Sheets will be automatically combined for PAS search")
        combine_info.setStyleSheet("color: #0066cc; font-weight: bold;")
        combine_layout.addWidget(combine_info)

        # Explanatory text about what gets combined
        explanation = QLabel(
            "The combined data will include:\n"
            "  • MFG = Manufacturer Name (e.g., 'Texas Instruments')\n"
            "  • MFG PN = Manufacturer Part Number (e.g., 'TPS54360DDAR') ← Used for PAS search\n"
            "  • Part Number = Your internal/company part number (not used for PAS search)\n"
            "  • Description = Part description\n\n"
            "Use filters below to exclude rows with missing data:"
        )
        explanation.setStyleSheet("font-size: 10pt; color: #555; padding: 10px; background-color: #f5f5f5; border-radius: 5px;")
        explanation.setWordWrap(True)
        combine_layout.addWidget(explanation)

        self.filter_group = QGroupBox("Data Quality Filters (exclude rows that don't meet ALL checked conditions)")
        filter_layout = QVBoxLayout()

        self.filter_mfg = QCheckBox("Require MFG (Manufacturer Name)")
        self.filter_mfg.setToolTip("Exclude rows where Manufacturer Name is empty or missing")

        self.filter_mfg_pn = QCheckBox("Require MFG PN (Manufacturer Part Number)")
        self.filter_mfg_pn.setToolTip("Exclude rows where Manufacturer Part Number is empty or missing.\nRECOMMENDED: Check this to avoid PAS search errors.")

        self.filter_part_number = QCheckBox("Require Part Number (Internal/Company Part Number)")
        self.filter_part_number.setToolTip("Exclude rows where your internal Part Number is empty or missing")

        self.filter_description = QCheckBox("Require Description")
        self.filter_description.setToolTip("Exclude rows where Description is empty or missing")

        # TBD fill option
        self.fill_tbd_checkbox = QCheckBox("Auto-fill empty MFG with 'TBD' when MFG PN exists")
        self.fill_tbd_checkbox.setToolTip("If Manufacturer Part Number exists but Manufacturer Name is missing, automatically set MFG to 'TBD'")

        filter_layout.addWidget(self.filter_mfg)
        filter_layout.addWidget(self.filter_mfg_pn)
        filter_layout.addWidget(self.filter_part_number)
        filter_layout.addWidget(self.filter_description)
        filter_layout.addWidget(self.fill_tbd_checkbox)

        self.filter_group.setLayout(filter_layout)
        # Enable filter group by default since combining is mandatory
        self.filter_group.setEnabled(True)

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
        self.preview_table.setSortingEnabled(True)  # Enable sorting
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
        self.combined_data = None  # Will store combined dataframe for PAS Search

        # Set recommended defaults for filters
        self.filter_mfg.setChecked(False)  # Require MFG by default
        self.filter_mfg_pn.setChecked(True)  # Require MFG PN by default (CRITICAL for PAS search)

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

        # Enable/disable per-row action buttons based on API key availability
        self.update_action_buttons_state()

    def update_action_buttons_state(self):
        """Enable or disable per-row action buttons based on API key availability"""
        enabled = self.api_key and ANTHROPIC_AVAILABLE

        for row in range(self.mapping_table.rowCount()):
            action_btn = self.mapping_table.cellWidget(row, 7)
            if action_btn:
                action_btn.setEnabled(enabled)
                if not enabled:
                    if not ANTHROPIC_AVAILABLE:
                        action_btn.setToolTip("Anthropic package not installed")
                    elif not self.api_key:
                        action_btn.setToolTip("No API key provided. Please configure in the Start page.")
                else:
                    action_btn.setToolTip("Auto-detect column mappings for this sheet using AI")

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

    def toggle_all_sheets(self):
        """Toggle all sheets (select all or unselect all based on button state)"""
        is_checked = self.toggle_select_btn.isChecked()

        for row in range(self.mapping_table.rowCount()):
            include_widget = self.mapping_table.cellWidget(row, 0)
            if include_widget:
                checkbox = include_widget.findChild(QCheckBox)
                if checkbox:
                    checkbox.setChecked(is_checked)

        # Update button text based on state
        if is_checked:
            self.toggle_select_btn.setText("✗ Unselect All")
        else:
            self.toggle_select_btn.setText("✓ Select All")

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

            # Add auto-detect action button
            action_btn = QPushButton("🤖 Auto-Detect")
            action_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    font-weight: bold;
                    padding: 5px 10px;
                    border-radius: 3px;
                    font-size: 10pt;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
                QPushButton:disabled {
                    background-color: #cccccc;
                    color: #666666;
                }
            """)
            action_btn.setProperty("sheet_name", sheet_name)
            action_btn.setProperty("row_index", row)
            action_btn.clicked.connect(lambda checked, r=row: self.auto_detect_single_row(r))
            action_btn.setToolTip("Auto-detect column mappings for this sheet using AI")
            self.mapping_table.setCellWidget(row, 7, action_btn)

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

    def auto_detect_single_row(self, row):
        """Auto-detect column mappings for a single row using AI"""
        if not self.api_key or not ANTHROPIC_AVAILABLE:
            QMessageBox.warning(
                self,
                "AI Not Available",
                "Claude AI is not available. Please provide an API key in the Start page."
            )
            return

        # Get sheet name for this row
        sheet_item = self.mapping_table.item(row, 1)
        if not sheet_item:
            return

        sheet_name = sheet_item.text()

        # Get the dataframe for this sheet
        if sheet_name not in self.dataframes:
            QMessageBox.warning(
                self,
                "Sheet Not Found",
                f"Could not find data for sheet: {sheet_name}"
            )
            return

        # Get the action button for this row
        action_btn = self.mapping_table.cellWidget(row, 7)
        if action_btn:
            action_btn.setEnabled(False)
            action_btn.setText("⏳ Detecting...")

        # Get selected model from StartPage
        start_page = self.wizard().page(0)
        model = start_page.get_selected_model() if hasattr(start_page, 'get_selected_model') else "claude-sonnet-4-5-20250929"

        # Create and start single sheet detection worker
        self.single_sheet_worker = SheetDetectionWorker(
            self.api_key,
            sheet_name,
            self.dataframes[sheet_name],
            model
        )

        # Connect signals with row information
        self.single_sheet_worker.finished.connect(
            lambda sname, mapping, r=row: self.on_single_sheet_finished(r, sname, mapping)
        )
        self.single_sheet_worker.error.connect(
            lambda sname, error, r=row: self.on_single_sheet_error(r, sname, error)
        )

        self.single_sheet_worker.start()

    def on_single_sheet_finished(self, row, sheet_name, mapping):
        """Handle completion of single sheet auto-detection"""
        # Column index mapping
        col_map = {
            'MFG': 2,
            'MFG_PN': 3,
            'MFG_PN_2': 4,
            'Part_Number': 5,
            'Description': 6
        }

        # Apply mappings to this row
        for field, col_idx in col_map.items():
            if field in mapping:
                mapping_info = mapping[field]
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

        # Re-enable the action button
        action_btn = self.mapping_table.cellWidget(row, 7)
        if action_btn:
            action_btn.setEnabled(True)
            action_btn.setText("🤖 Auto-Detect")

        # Show success message with confidence info
        QMessageBox.information(
            self,
            "Detection Complete",
            f"Column mappings detected for '{sheet_name}'!\n\n"
            "Color coding:\n"
            "🟢 Green: High confidence (80%+)\n"
            "🟡 Yellow: Medium confidence (50-79%)\n"
            "🟠 Orange: Low confidence (<50%)\n\n"
            "Hover over dropdowns to see confidence scores."
        )

    def on_single_sheet_error(self, row, sheet_name, error_msg):
        """Handle error from single sheet auto-detection"""
        # Re-enable the action button
        action_btn = self.mapping_table.cellWidget(row, 7)
        if action_btn:
            action_btn.setEnabled(True)
            action_btn.setText("🤖 Auto-Detect")

        QMessageBox.critical(
            self,
            "Detection Failed",
            f"Failed to auto-detect columns for '{sheet_name}':\n{error_msg}"
        )

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

        # Disable all dropdowns and action buttons in the mapping table
        for row in range(self.mapping_table.rowCount()):
            for col in range(2, 7):  # Columns 2-6 are the dropdowns
                combo = self.mapping_table.cellWidget(row, col)
                if combo:
                    combo.setEnabled(False)
            # Disable per-row action button
            action_btn = self.mapping_table.cellWidget(row, 7)
            if action_btn:
                action_btn.setEnabled(False)

        self.ai_status.setText("🔄 Starting AI analysis...")
        self.ai_status.setStyleSheet("color: blue;")

        # Get selected model from StartPage
        start_page = self.wizard().page(0)
        model = start_page.get_selected_model() if hasattr(start_page, 'get_selected_model') else "claude-sonnet-4-5-20250929"

        # Create and start AI detection thread
        self.ai_thread = AIDetectionThread(self.api_key, self.dataframes, model)
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

        # Re-enable all dropdowns and action buttons
        for row in range(self.mapping_table.rowCount()):
            for col in range(2, 7):
                combo = self.mapping_table.cellWidget(row, col)
                if combo:
                    combo.setEnabled(True)
            # Re-enable per-row action button
            action_btn = self.mapping_table.cellWidget(row, 7)
            if action_btn:
                action_btn.setEnabled(True)

        # Remove progress bar
        ai_group = self.ai_detect_btn.parent()
        ai_layout = ai_group.layout()
        ai_layout.removeWidget(self.ai_progress)
        self.ai_progress.deleteLater()

        # Show legend with failed sheets info
        failed_sheets = getattr(self.ai_thread, 'failed_sheets', [])
        success_count = len(all_mappings)
        total_count = len(self.dataframes)

        message = f"Column mappings detected for {success_count} of {total_count} sheets!\n\n"

        if failed_sheets:
            message += f"⚠️ {len(failed_sheets)} sheet(s) failed:\n"
            for item in failed_sheets[:5]:
                error_short = item['error'][:60]
                if 'rate_limit' in error_short.lower() or '429' in error_short:
                    message += f"  • {item['sheet']}: Rate limit error (429)\n"
                else:
                    message += f"  • {item['sheet']}: {error_short}...\n"
            if len(failed_sheets) > 5:
                message += f"  ... and {len(failed_sheets) - 5} more\n"
            message += "\n"

        message += (
            "Color coding:\n"
            "🟢 Green: High confidence (80%+)\n"
            "🟡 Yellow: Medium confidence (50-79%)\n"
            "🟠 Orange: Low confidence (<50%)\n\n"
            "Please review and adjust as needed. "
            "Hover over dropdowns to see confidence scores."
        )

        QMessageBox.information(self, "AI Detection Complete", message)

    def on_ai_error(self, error_msg):
        """Handle AI detection error"""
        self.ai_status.setText(f"✗ Error: {error_msg[:30]}")
        self.ai_status.setStyleSheet("color: red;")

        # Re-enable controls
        self.ai_detect_btn.setEnabled(True)
        self.bulk_apply_btn.setEnabled(True)
        self.save_config_btn.setEnabled(True)
        self.load_config_btn.setEnabled(True)

        # Re-enable all dropdowns and action buttons
        for row in range(self.mapping_table.rowCount()):
            for col in range(2, 7):
                combo = self.mapping_table.cellWidget(row, col)
                if combo:
                    combo.setEnabled(True)
            # Re-enable per-row action button
            action_btn = self.mapping_table.cellWidget(row, 7)
            if action_btn:
                action_btn.setEnabled(True)

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
        """Check if sheets should be combined - always True (mandatory)"""
        return True

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

        # Auto-save configuration
        self.auto_save_configuration()

        return True

    def auto_save_configuration(self):
        """Automatically save mapping configuration to a default file"""
        try:
            # Save to a fixed file in the current directory or a .gemini folder if preferred
            # For now, saving to 'mapping_config_autosave.json' in the current working directory
            file_path = "mapping_config_autosave.json"
            
            config = {
                'mappings': self.get_mappings(),
                'version': '1.0',
                'timestamp': pd.Timestamp.now().isoformat()
            }

            with open(file_path, 'w') as f:
                json.dump(config, f, indent=2)
            print(f"Auto-saved mapping configuration to {file_path}")
        except Exception as e:
            print(f"Failed to auto-save configuration: {str(e)}")

    def combine_sheets(self):
        """Combine sheets based on mappings and filters"""
        prev_page = self.wizard().page(1)  # DataSourcePage is page 1
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

            # Store combined data for PAS Search page to access
            self.combined_data = combined_df

            # The Excel file is already in the output folder (from Step 1)
            # We just need to update it by adding the Combined sheet
            if not excel_path or not os.path.exists(excel_path):
                raise Exception("Excel file not found. Please go back to Step 1.")

            # Read existing sheets from the Excel file
            with pd.ExcelFile(excel_path) as xls:
                existing_sheets = {sheet: pd.read_excel(excel_path, sheet_name=sheet)
                                 for sheet in xls.sheet_names}

            # Add/update the Combined sheet
            existing_sheets['Combined'] = combined_df

            # Write back all sheets to the same Excel file
            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                for sheet_name, df in existing_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Store the Excel path for later use (same file, just updated)
            self.output_excel_path = excel_path

            QMessageBox.information(
                self, "Combine Complete",
                f"Successfully combined {len(included_sheets)} sheets into 'Combined' sheet.\n"
                f"Total rows: {len(combined_df)}\n\n"
                f"Updated file:\n{excel_path}"
            )
        else:
            # No data after filtering - set empty dataframe
            self.combined_data = pd.DataFrame()
            raise Exception("No data remained after applying filters. Please adjust your filter settings or column mappings.")

