"""
Comparison Page - Step 6 of EDM Wizard

Displays a Beyond Compare style side-by-side comparison of the original
and modified datasets, with export capabilities to CSV and Excel.
"""

import os
import csv
from pathlib import Path
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("Error: pandas is required. Install it with: pip install pandas")

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel,
        QTableWidget, QTableWidgetItem, QPushButton, QRadioButton,
        QMessageBox, QCheckBox, QApplication, QDialog, QDialogButtonBox,
        QLineEdit, QFormLayout
    )
    from PyQt5.QtCore import Qt
    from PyQt5.QtGui import QColor
except ImportError:
    print("Error: PyQt5 is required. Install it with: pip install PyQt5")

try:
    from edm_wizard.utils.xml_generation import create_mfg_xml, create_mfgpn_xml
    from edm_wizard.utils.constants import DEFAULT_PROJECT_NAME, DEFAULT_CATALOG
    XML_AVAILABLE = True
except ImportError:
    XML_AVAILABLE = False
    DEFAULT_PROJECT_NAME = "VarTrainingLab"
    DEFAULT_CATALOG = "VV"


class ComparisonPage(QWizardPage):
    """
    Step 6: Side-by-Side Comparison - Beyond Compare Style

    Displays original vs modified data with:
    - Side-by-side table comparison
    - Color-coded changes (red=old, green=new)
    - Filter controls (all rows or changes only)
    - Export to CSV and Excel options
    """

    def __init__(self):
        super().__init__()
        self.setTitle("Step 6: Review Changes - Side-by-Side Comparison")
        self.setSubTitle("Compare Combined (original) vs Combined_New (with normalization)")

        layout = QVBoxLayout()

        # Summary section
        summary_group = QGroupBox("Summary")
        summary_layout = QHBoxLayout()

        self.summary_label = QLabel("No changes to display")
        self.summary_label.setWordWrap(True)
        summary_layout.addWidget(self.summary_label)

        # Filter controls
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Show:"))

        self.show_all_radio = QRadioButton("All Rows")
        self.show_changes_radio = QRadioButton("Changes Only")
        self.show_all_radio.setChecked(True)
        self.show_all_radio.toggled.connect(self.apply_filter)
        self.show_changes_radio.toggled.connect(self.apply_filter)

        filter_layout.addWidget(self.show_all_radio)
        filter_layout.addWidget(self.show_changes_radio)
        filter_layout.addStretch()

        summary_layout.addLayout(filter_layout)
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)

        # Side-by-side comparison tables
        comparison_layout = QHBoxLayout()

        # Left table: Combined (Original)
        left_group = QGroupBox("Combined (Original)")
        left_layout = QVBoxLayout()

        self.left_table = QTableWidget()
        self.left_table.setSortingEnabled(True)
        self.left_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.left_table.setSelectionMode(QTableWidget.SingleSelection)
        self.left_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.left_table.verticalScrollBar().valueChanged.connect(self.sync_scroll_right)

        left_layout.addWidget(self.left_table)
        left_group.setLayout(left_layout)
        comparison_layout.addWidget(left_group)

        # Right table: Combined_New (After Changes)
        right_group = QGroupBox("Combined_New (After Changes)")
        right_layout = QVBoxLayout()

        self.right_table = QTableWidget()
        self.right_table.setSortingEnabled(True)
        self.right_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.right_table.setSelectionMode(QTableWidget.SingleSelection)
        self.right_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.right_table.verticalScrollBar().valueChanged.connect(self.sync_scroll_left)

        right_layout.addWidget(self.right_table)
        right_group.setLayout(right_layout)
        comparison_layout.addWidget(right_group)

        layout.addLayout(comparison_layout, stretch=1)

        # Data Writeback Options
        writeback_group = QGroupBox("🔄 Data Writeback Options")
        writeback_layout = QVBoxLayout()

        # Option 1: Column update strategy
        strategy_label = QLabel("<b>MFG/MFG PN Column Update Strategy:</b>")
        writeback_layout.addWidget(strategy_label)

        self.overwrite_radio = QRadioButton("Overwrite existing MFG/MFG PN columns with normalized values")
        self.overwrite_radio.setChecked(True)
        self.overwrite_radio.setToolTip("Replace original MFG and MFG PN values with normalized values from PAS")
        writeback_layout.addWidget(self.overwrite_radio)

        self.new_columns_radio = QRadioButton("Create new columns (MFG_Normalized, MFG_PN_Normalized, External_Content_ID)")
        self.new_columns_radio.setToolTip("Keep original columns and add new normalized columns alongside")
        writeback_layout.addWidget(self.new_columns_radio)

        # Option 2: External Content ID
        self.include_external_id_cb = QCheckBox("Include External Content ID from PAS search results")
        self.include_external_id_cb.setChecked(True)
        self.include_external_id_cb.setToolTip("Add External Content ID column with unique identifiers from PAS database")
        writeback_layout.addWidget(self.include_external_id_cb)

        # Option 3: Writeback to source
        writeback_info = QLabel(
            "<i>Note: Changes will be written back to the original source Excel sheets based on Source_Sheet tracking.</i>"
        )
        writeback_info.setWordWrap(True)
        writeback_info.setStyleSheet("color: #666; font-size: 9pt; padding: 5px;")
        writeback_layout.addWidget(writeback_info)

        # Writeback button
        writeback_btn_layout = QHBoxLayout()
        self.writeback_btn = QPushButton("📝 Write Back to Source Excel File")
        self.writeback_btn.clicked.connect(self.writeback_to_source)
        self.writeback_btn.setToolTip("Apply changes back to the original source Excel file")
        self.writeback_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 8px; }")
        writeback_btn_layout.addWidget(self.writeback_btn)
        writeback_btn_layout.addStretch()
        writeback_layout.addLayout(writeback_btn_layout)

        self.writeback_status = QLabel("")
        writeback_layout.addWidget(self.writeback_status)

        writeback_group.setLayout(writeback_layout)
        layout.addWidget(writeback_group)

        # Export options
        export_group = QGroupBox("💾 Export Options")
        export_layout = QVBoxLayout()

        export_btn_layout = QHBoxLayout()

        self.export_csv_btn = QPushButton("Export to CSV")
        self.export_csv_btn.clicked.connect(self.export_to_csv)
        export_btn_layout.addWidget(self.export_csv_btn)

        self.export_excel_btn = QPushButton("Export to Excel")
        self.export_excel_btn.clicked.connect(self.export_to_excel)
        export_btn_layout.addWidget(self.export_excel_btn)

        export_btn_layout.addStretch()
        export_layout.addLayout(export_btn_layout)

        self.export_status = QLabel("")
        export_layout.addWidget(self.export_status)

        export_group.setLayout(export_layout)
        layout.addWidget(export_group)

        self.setLayout(layout)

        # Store data
        self.original_df = None
        self.new_df = None
        self.all_rows = []
        self.syncing_scroll = False  # Prevent scroll recursion

    def sync_scroll_right(self, value):
        """Sync right table scroll with left table"""
        if not self.syncing_scroll:
            self.syncing_scroll = True
            self.right_table.verticalScrollBar().setValue(value)
            self.syncing_scroll = False

    def sync_scroll_left(self, value):
        """Sync left table scroll with right table"""
        if not self.syncing_scroll:
            self.syncing_scroll = True
            self.left_table.verticalScrollBar().setValue(value)
            self.syncing_scroll = False

    def initializePage(self):
        """Initialize by loading Combined from source and Combined_New from normalized output"""
        # Get the column mapping page for file paths
        column_mapping_page = self.wizard().page(2)

        # Get original (source) Excel file path
        if not hasattr(column_mapping_page, 'output_excel_path') or not column_mapping_page.output_excel_path:
            self.summary_label.setText("Source Excel file not found. Please go back and complete Step 2.")
            return

        source_excel_path = column_mapping_page.output_excel_path

        if not os.path.exists(source_excel_path):
            self.summary_label.setText(f"Source Excel file not found: {source_excel_path}")
            return

        # Get normalized (updated) Excel file path
        updated_excel_path = getattr(column_mapping_page, 'updated_excel_path', None)

        try:
            # Load Combined (original) sheet from source file
            if 'Combined' in pd.ExcelFile(source_excel_path).sheet_names:
                self.original_df = pd.read_excel(source_excel_path, sheet_name='Combined')
            else:
                self.summary_label.setText("'Combined' sheet not found in source Excel file")
                return

            # Load Combined_New (after changes) sheet from normalized output file
            if updated_excel_path and os.path.exists(updated_excel_path):
                if 'Combined_New' in pd.ExcelFile(updated_excel_path).sheet_names:
                    self.new_df = pd.read_excel(updated_excel_path, sheet_name='Combined_New')
                else:
                    # Fallback: try Combined sheet from updated file
                    if 'Combined' in pd.ExcelFile(updated_excel_path).sheet_names:
                        self.new_df = pd.read_excel(updated_excel_path, sheet_name='Combined')
                    else:
                        self.new_df = self.original_df.copy()
                        self.summary_label.setText("'Combined_New' sheet not found in normalized file - showing original data only")
            else:
                # If no updated file exists, use original as placeholder
                self.new_df = self.original_df.copy()
                self.summary_label.setText("Normalized output file not found - showing original data only")

            # Store paths for writeback
            self.source_excel_path = source_excel_path
            self.updated_excel_path = updated_excel_path

            # Build comparison
            self.build_comparison()

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.summary_label.setText(f"Error loading data: {str(e)}")

    def get_mapped_columns(self):
        """Get only the mapped columns from Column Mapping step"""
        # These are the standard column names after combination
        mapped_columns = ['MFG', 'MFG_PN', 'Part_Number', 'Description', 'Source_Sheet',
                         'External_Content_ID', 'MFG_Normalized', 'MFG_PN_Normalized']

        # Only include columns that exist in the dataframes
        available_columns = []
        all_cols = set(self.original_df.columns) | set(self.new_df.columns)
        for col in mapped_columns:
            if col in all_cols:
                available_columns.append(col)

        return available_columns

    def get_display_column_name(self, col):
        """Convert internal column name to user-friendly display name"""
        display_names = {
            'MFG': 'MFG',
            'MFG_PN': 'MFG PN',
            'Part_Number': 'Part Number',
            'Description': 'Description',
            'Source_Sheet': 'Source Sheet',
            'External_Content_ID': 'External Content ID',
            'MFG_Normalized': 'MFG (Normalized)',
            'MFG_PN_Normalized': 'MFG PN (Normalized)'
        }
        return display_names.get(col, col)

    def build_comparison(self):
        """Build side-by-side comparison with Beyond Compare styling"""
        try:
            if self.original_df is None or self.new_df is None:
                return

            # Get only the mapped columns to display
            mapped_columns = self.get_mapped_columns()

            # Ensure both DataFrames have the same columns (only mapped ones)
            for col in mapped_columns:
                if col not in self.original_df.columns:
                    self.original_df[col] = ""
                if col not in self.new_df.columns:
                    self.new_df[col] = ""

            # Filter to only show mapped columns
            self.original_df = self.original_df[mapped_columns]
            self.new_df = self.new_df[mapped_columns]

            # Build row comparison data
            self.all_rows = []
            max_rows = max(len(self.original_df), len(self.new_df))

            changed_count = 0
            for i in range(max_rows):
                row_changed = False
                if i < len(self.original_df) and i < len(self.new_df):
                    # Compare each cell (only mapped columns)
                    for col in mapped_columns:
                        old_val = str(self.original_df.iloc[i][col]) if pd.notna(self.original_df.iloc[i][col]) else ""
                        new_val = str(self.new_df.iloc[i][col]) if pd.notna(self.new_df.iloc[i][col]) else ""
                        if old_val != new_val:
                            row_changed = True
                            break
                else:
                    row_changed = True  # Row exists in one but not the other

                if row_changed:
                    changed_count += 1

                self.all_rows.append({
                    'index': i,
                    'changed': row_changed
                })

            # Update summary
            total = len(self.all_rows)
            unchanged = total - changed_count
            if total > 0:
                self.summary_label.setText(
                    f"<b>Total Rows:</b> {total} | "
                    f"<b>Changed:</b> {changed_count} ({changed_count/total*100:.1f}%) | "
                    f"<b>Unchanged:</b> {unchanged} ({unchanged/total*100:.1f}%)"
                )
            else:
                self.summary_label.setText("<b>No data to compare</b>")

            # Populate tables
            self.populate_tables()

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.summary_label.setText(f"Error building comparison: {str(e)}")

    def populate_tables(self):
        """Populate both tables with data and Beyond Compare style formatting"""
        if self.original_df is None or self.new_df is None:
            return

        # Filter rows based on radio selection
        if self.show_changes_radio.isChecked():
            display_rows = [r for r in self.all_rows if r['changed']]
        else:
            display_rows = self.all_rows

        # Set up columns (use only mapped columns)
        columns = list(self.original_df.columns)
        display_headers = [self.get_display_column_name(col) for col in columns]

        self.left_table.setColumnCount(len(columns))
        self.left_table.setHorizontalHeaderLabels(display_headers)
        self.right_table.setColumnCount(len(columns))
        self.right_table.setHorizontalHeaderLabels(display_headers)

        # Set row counts
        self.left_table.setRowCount(len(display_rows))
        self.right_table.setRowCount(len(display_rows))

        # Populate rows with Beyond Compare styling
        for display_idx, row_info in enumerate(display_rows):
            i = row_info['index']
            row_changed = row_info['changed']

            # Populate left table (original)
            if i < len(self.original_df):
                for col_idx, col in enumerate(columns):
                    old_val = str(self.original_df.iloc[i][col]) if pd.notna(self.original_df.iloc[i][col]) else ""
                    item = QTableWidgetItem(old_val)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)

                    # Compare with new value for cell-level highlighting
                    if i < len(self.new_df):
                        new_val = str(self.new_df.iloc[i][col]) if pd.notna(self.new_df.iloc[i][col]) else ""
                        if old_val != new_val:
                            # Cell changed - light red background, bold font
                            item.setBackground(QColor(255, 200, 200))  # Light red
                            font = item.font()
                            font.setBold(True)
                            item.setFont(font)

                    self.left_table.setItem(display_idx, col_idx, item)

            # Populate right table (new)
            if i < len(self.new_df):
                for col_idx, col in enumerate(columns):
                    new_val = str(self.new_df.iloc[i][col]) if pd.notna(self.new_df.iloc[i][col]) else ""
                    item = QTableWidgetItem(new_val)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)

                    # Compare with old value for cell-level highlighting
                    if i < len(self.original_df):
                        old_val = str(self.original_df.iloc[i][col]) if pd.notna(self.original_df.iloc[i][col]) else ""
                        if old_val != new_val:
                            # Cell changed - light green background, bold font
                            item.setBackground(QColor(200, 255, 200))  # Light green
                            font = item.font()
                            font.setBold(True)
                            item.setFont(font)

                    self.right_table.setItem(display_idx, col_idx, item)

        # Resize columns to fit content
        self.left_table.resizeColumnsToContents()
        self.right_table.resizeColumnsToContents()

    def apply_filter(self):
        """Re-populate tables based on filter selection"""
        self.populate_tables()

    def export_to_csv(self):
        """Export comparison to CSV"""
        try:
            start_page = self.wizard().page(0)
            output_folder = start_page.get_output_folder() if hasattr(start_page, 'get_output_folder') else None

            if not output_folder:
                QMessageBox.warning(self, "Error", "Output folder not configured")
                return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_path = Path(output_folder) / f"Comparison_{timestamp}.csv"

            # Export comparison data
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)

                # Write headers
                columns = list(self.original_df.columns)
                header = []
                for col in columns:
                    header.append(f"Original {self.get_display_column_name(col)}")
                    header.append(f"New {self.get_display_column_name(col)}")
                writer.writerow(header)

                # Write rows
                max_rows = max(len(self.original_df), len(self.new_df))
                for i in range(max_rows):
                    row = []
                    for col in columns:
                        old_val = str(self.original_df.iloc[i][col]) if i < len(self.original_df) and pd.notna(self.original_df.iloc[i][col]) else ""
                        new_val = str(self.new_df.iloc[i][col]) if i < len(self.new_df) and pd.notna(self.new_df.iloc[i][col]) else ""
                        row.append(old_val)
                        row.append(new_val)
                    writer.writerow(row)

            self.export_status.setText(f"Exported to: {csv_path.name}")
            self.export_status.setStyleSheet("color: green;")

        except Exception as e:
            self.export_status.setText(f"Export failed: {str(e)}")
            self.export_status.setStyleSheet("color: red;")

    def export_to_excel(self):
        """Export comparison to Excel"""
        try:
            start_page = self.wizard().page(0)
            output_folder = start_page.get_output_folder() if hasattr(start_page, 'get_output_folder') else None

            if not output_folder:
                QMessageBox.warning(self, "Error", "Output folder not configured")
                return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_path = Path(output_folder) / f"Comparison_{timestamp}.xlsx"

            # Create comparison DataFrame
            columns = list(self.original_df.columns)
            export_data = []

            max_rows = max(len(self.original_df), len(self.new_df))
            for i in range(max_rows):
                row = {}
                for col in columns:
                    old_val = str(self.original_df.iloc[i][col]) if i < len(self.original_df) and pd.notna(self.original_df.iloc[i][col]) else ""
                    new_val = str(self.new_df.iloc[i][col]) if i < len(self.new_df) and pd.notna(self.new_df.iloc[i][col]) else ""
                    row[f"Original {self.get_display_column_name(col)}"] = old_val
                    row[f"New {self.get_display_column_name(col)}"] = new_val
                export_data.append(row)

            df = pd.DataFrame(export_data)

            # Write to Excel
            df.to_excel(excel_path, index=False, engine='xlsxwriter')

            self.export_status.setText(f"Exported to: {excel_path.name}")
            self.export_status.setStyleSheet("color: green;")

        except Exception as e:
            self.export_status.setText(f"Export failed: {str(e)}")
            self.export_status.setStyleSheet("color: red;")

    def writeback_to_source(self):
        """Write normalized data and External Content IDs back to source Excel file"""
        try:
            # Get review page to access search results
            review_page = self.wizard().page(4)  # SupplyFrameReviewPage is page 4
            if not review_page or not hasattr(review_page, 'search_results'):
                QMessageBox.warning(self, "Error", "Search results not found. Please complete the review step first.")
                return

            # Use the source file path stored during initializePage
            if not hasattr(self, 'source_excel_path') or not self.source_excel_path:
                # Fallback to column_mapping_page
                column_mapping_page = self.wizard().page(2)
                if not hasattr(column_mapping_page, 'output_excel_path') or not column_mapping_page.output_excel_path:
                    QMessageBox.warning(self, "Error", "Source Excel file not found. Cannot write back changes.")
                    return
                source_file = Path(column_mapping_page.output_excel_path)
            else:
                source_file = Path(self.source_excel_path)
            if not source_file.exists():
                QMessageBox.warning(self, "Error", f"Source file not found:\n{source_file}")
                return

            # Confirm action
            overwrite_mode = self.overwrite_radio.isChecked()
            include_external_id = self.include_external_id_cb.isChecked()

            strategy_msg = "overwrite existing MFG/MFG PN columns" if overwrite_mode else "create new columns (MFG_Normalized, MFG_PN_Normalized)"
            external_id_msg = "and add External_Content_ID column" if include_external_id else "without External_Content_ID"

            confirm = QMessageBox.question(
                self,
                "Confirm Writeback",
                f"This will update the source Excel file:\n{source_file.name}\n\n"
                f"Strategy: {strategy_msg}\n{external_id_msg}\n\n"
                f"Do you want to proceed?",
                QMessageBox.Yes | QMessageBox.No
            )

            if confirm != QMessageBox.Yes:
                return

            # Show progress
            self.writeback_status.setText("⏳ Writing back to source file...")
            self.writeback_status.setStyleSheet("color: blue;")
            QApplication.processEvents()

            # Build lookup dictionary for updates
            # Key: (original_pn, original_mfg, source_sheet)
            # Value: (normalized_mfg, normalized_pn, external_id)
            updates = {}

            for result in review_page.search_results:
                original_pn = result.get('PartNumber', '')
                original_mfg = result.get('ManufacturerName', '')

                # Determine the source sheet
                # First check if we have combined_data with Source_Sheet
                pas_page = self.wizard().page(3)
                source_sheet = ''
                if hasattr(pas_page, 'combined_data') and pas_page.combined_data is not None:
                    # Find matching row in combined_data
                    matching_rows = pas_page.combined_data[
                        (pas_page.combined_data['MFG_PN'] == original_pn) &
                        (pas_page.combined_data['MFG'] == original_mfg)
                    ]
                    if not matching_rows.empty:
                        source_sheet = matching_rows.iloc[0].get('Source_Sheet', '')

                # Get selected match (normalized values)
                selected_match = result.get('selected_match')
                if not selected_match:
                    # For "Found" status, use the first match
                    if result.get('MatchStatus') == 'Found' and result.get('matches'):
                        selected_match = result['matches'][0]

                if selected_match:
                    # Extract normalized values from selected match
                    if isinstance(selected_match, dict):
                        normalized_mfg = selected_match.get('mfg', original_mfg)
                        normalized_pn = selected_match.get('mpn', original_pn)
                        external_id = selected_match.get('external_id', '')
                    elif isinstance(selected_match, str) and '@' in selected_match:
                        # Old format: "PartNumber@Manufacturer"
                        normalized_pn, normalized_mfg = selected_match.split('@', 1)
                        external_id = ''
                    else:
                        normalized_pn = original_pn
                        normalized_mfg = original_mfg
                        external_id = ''

                    # Check if manufacturer was normalized
                    if hasattr(review_page, 'manufacturer_normalizations'):
                        if original_mfg in review_page.manufacturer_normalizations:
                            normalized_mfg = review_page.manufacturer_normalizations[original_mfg]

                    # Store update
                    key = (original_pn, original_mfg, source_sheet)
                    updates[key] = (normalized_mfg, normalized_pn, external_id)

            if not updates:
                self.writeback_status.setText("⚠ No updates to write back")
                self.writeback_status.setStyleSheet("color: orange;")
                return

            # Load the source Excel file
            excel_file = pd.ExcelFile(source_file)

            # Process each sheet
            updates_count = 0
            sheets_updated = []

            with pd.ExcelWriter(source_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)

                    # Find MFG and MFG PN columns (may have different names in original sheets)
                    mfg_col = None
                    mfg_pn_col = None

                    # Try to find columns from column_mapping_page
                    if hasattr(column_mapping_page, 'sheet_mappings') and sheet_name in column_mapping_page.sheet_mappings:
                        mfg_col = column_mapping_page.sheet_mappings[sheet_name].get('mfg_col')
                        mfg_pn_col = column_mapping_page.sheet_mappings[sheet_name].get('mfg_pn_col')

                    if not mfg_col or not mfg_pn_col:
                        # Try to find by common names
                        for col in df.columns:
                            col_lower = col.lower()
                            if 'mfg' in col_lower and 'pn' not in col_lower and not mfg_col:
                                mfg_col = col
                            elif ('pn' in col_lower or 'part' in col_lower) and not mfg_pn_col:
                                mfg_pn_col = col

                    if not mfg_col or not mfg_pn_col:
                        continue  # Skip sheets without MFG/MFG PN columns

                    # Add new columns if in "create new columns" mode
                    if not overwrite_mode:
                        if 'MFG_Normalized' not in df.columns:
                            df['MFG_Normalized'] = ''
                        if 'MFG_PN_Normalized' not in df.columns:
                            df['MFG_PN_Normalized'] = ''

                    if include_external_id and 'External_Content_ID' not in df.columns:
                        df['External_Content_ID'] = ''

                    # Apply updates row by row
                    sheet_updates = 0
                    for idx, row in df.iterrows():
                        original_pn = str(row[mfg_pn_col]) if pd.notna(row[mfg_pn_col]) else ''
                        original_mfg = str(row[mfg_col]) if pd.notna(row[mfg_col]) else ''

                        key = (original_pn, original_mfg, sheet_name)
                        if key in updates:
                            normalized_mfg, normalized_pn, external_id = updates[key]

                            # Track if this row has any actual changes
                            row_has_changes = False

                            if overwrite_mode:
                                # Overwrite existing columns - only if values changed
                                if normalized_mfg and normalized_mfg != original_mfg:
                                    df.at[idx, mfg_col] = normalized_mfg
                                    sheet_updates += 1
                                    row_has_changes = True
                                if normalized_pn and normalized_pn != original_pn:
                                    df.at[idx, mfg_pn_col] = normalized_pn
                                    sheet_updates += 1
                                    row_has_changes = True
                            else:
                                # Create new columns - only if values changed or external_id exists
                                has_mfg_change = normalized_mfg and normalized_mfg != original_mfg
                                has_pn_change = normalized_pn and normalized_pn != original_pn
                                has_external_id = include_external_id and external_id

                                # Only write if there's an actual change or external ID
                                if has_mfg_change or has_pn_change or has_external_id:
                                    df.at[idx, 'MFG_Normalized'] = normalized_mfg if normalized_mfg else original_mfg
                                    df.at[idx, 'MFG_PN_Normalized'] = normalized_pn if normalized_pn else original_pn
                                    sheet_updates += 1
                                    row_has_changes = True

                            # Add External Content ID if enabled and exists
                            if include_external_id and external_id:
                                df.at[idx, 'External_Content_ID'] = external_id
                                if not row_has_changes:  # Only count if not already counted
                                    sheet_updates += 1
                                    row_has_changes = True

                    # Write sheet back
                    if sheet_updates > 0:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        sheets_updated.append(sheet_name)
                        updates_count += sheet_updates

            # Show success message
            self.writeback_status.setText(
                f"✓ Successfully updated {updates_count} values across {len(sheets_updated)} sheet(s):\n"
                f"{', '.join(sheets_updated[:3])}{' and more...' if len(sheets_updated) > 3 else ''}"
            )
            self.writeback_status.setStyleSheet("color: green;")

            QMessageBox.information(
                self,
                "Writeback Complete",
                f"Successfully updated source file:\n{source_file.name}\n\n"
                f"• Updated {updates_count} values\n"
                f"• Modified {len(sheets_updated)} sheet(s)\n\n"
                f"Strategy: {strategy_msg}\n{external_id_msg}"
            )

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.writeback_status.setText(f"✗ Writeback failed: {str(e)}")
            self.writeback_status.setStyleSheet("color: red;")
            QMessageBox.critical(
                self,
                "Writeback Error",
                f"Failed to write back to source file:\n\n{str(e)}\n\nDetails:\n{error_details}"
            )

    def validatePage(self):
        """Override to show export dialog instead of closing the wizard"""
        # Create export dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Export Normalized Data")
        dialog.setMinimumWidth(500)

        layout = QVBoxLayout(dialog)

        # Info label
        info_label = QLabel(
            "<b>Export your normalized data:</b><br>"
            "Choose export format(s) for your finished data with all normalizations applied."
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        # XML Settings
        xml_group = QGroupBox("XML Export Settings")
        xml_form = QFormLayout()

        self.project_name_input = QLineEdit(DEFAULT_PROJECT_NAME)
        xml_form.addRow("Project Name:", self.project_name_input)

        self.catalog_input = QLineEdit(DEFAULT_CATALOG)
        xml_form.addRow("Catalog:", self.catalog_input)

        xml_group.setLayout(xml_form)
        layout.addWidget(xml_group)

        # Export buttons - Data formats
        data_group = QGroupBox("📊 Data Export")
        data_layout = QHBoxLayout()

        excel_btn = QPushButton("Export Excel\n(.xlsx)")
        excel_btn.setStyleSheet("QPushButton { padding: 15px; font-size: 11pt; }")
        excel_btn.setToolTip("Export normalized data as Excel workbook")
        excel_btn.clicked.connect(lambda: self._do_export_normalized_excel(dialog))
        data_layout.addWidget(excel_btn)

        csv_btn = QPushButton("Export CSV\n(.csv)")
        csv_btn.setStyleSheet("QPushButton { padding: 15px; font-size: 11pt; }")
        csv_btn.setToolTip("Export normalized data as CSV (UTF-8)")
        csv_btn.clicked.connect(lambda: self._do_export_normalized_csv(dialog))
        data_layout.addWidget(csv_btn)

        data_group.setLayout(data_layout)
        layout.addWidget(data_group)

        # Export buttons - XML formats
        xml_btn_group = QGroupBox("📝 XML Export (for EDM Library Creator)")
        xml_layout = QHBoxLayout()

        mfg_xml_btn = QPushButton("Export MFG XML\n(Manufacturers)")
        mfg_xml_btn.setStyleSheet("QPushButton { padding: 15px; font-size: 11pt; }")
        mfg_xml_btn.setToolTip("Export unique manufacturers as XML (Class 090)")
        mfg_xml_btn.clicked.connect(lambda: self._do_export_mfg_xml(dialog))
        if not XML_AVAILABLE:
            mfg_xml_btn.setEnabled(False)
            mfg_xml_btn.setToolTip("XML export not available")
        xml_layout.addWidget(mfg_xml_btn)

        mfgpn_xml_btn = QPushButton("Export MFG PN XML\n(Part Numbers)")
        mfgpn_xml_btn.setStyleSheet("QPushButton { padding: 15px; font-size: 11pt; }")
        mfgpn_xml_btn.setToolTip("Export MFG:PN pairs as XML (Class 060)")
        mfgpn_xml_btn.clicked.connect(lambda: self._do_export_mfgpn_xml(dialog))
        if not XML_AVAILABLE:
            mfgpn_xml_btn.setEnabled(False)
            mfgpn_xml_btn.setToolTip("XML export not available")
        xml_layout.addWidget(mfgpn_xml_btn)

        both_xml_btn = QPushButton("Export Both XMLs")
        both_xml_btn.setStyleSheet("QPushButton { padding: 15px; font-size: 11pt; background-color: #4CAF50; color: white; }")
        both_xml_btn.setToolTip("Export both MFG and MFG PN XML files")
        both_xml_btn.clicked.connect(lambda: self._do_export_both_xml(dialog))
        if not XML_AVAILABLE:
            both_xml_btn.setEnabled(False)
            both_xml_btn.setToolTip("XML export not available")
        xml_layout.addWidget(both_xml_btn)

        xml_btn_group.setLayout(xml_layout)
        layout.addWidget(xml_btn_group)

        # Status label
        self.dialog_status = QLabel("")
        self.dialog_status.setWordWrap(True)
        self.dialog_status.setStyleSheet("padding: 10px; background-color: #f0f0f0; border-radius: 5px;")
        layout.addWidget(self.dialog_status)

        # Close buttons
        close_layout = QHBoxLayout()
        close_layout.addStretch()

        cancel_btn = QPushButton("Continue Editing")
        cancel_btn.clicked.connect(dialog.reject)
        close_layout.addWidget(cancel_btn)

        close_btn = QPushButton("Close Wizard")
        close_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; padding: 8px 16px; }")
        close_btn.clicked.connect(dialog.accept)
        close_layout.addWidget(close_btn)

        layout.addLayout(close_layout)

        # Show dialog
        result = dialog.exec_()

        # Only close wizard if user clicked "Close Wizard"
        return result == QDialog.Accepted

    def _get_output_folder(self):
        """Get output folder path"""
        start_page = self.wizard().page(0)
        return start_page.get_output_folder() if hasattr(start_page, 'get_output_folder') else None

    def _do_export_normalized_excel(self, dialog):
        """Export normalized data as Excel"""
        try:
            output_folder = self._get_output_folder()
            if not output_folder:
                self.dialog_status.setText("❌ Output folder not configured")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            if self.new_df is None or self.new_df.empty:
                self.dialog_status.setText("❌ No normalized data to export")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_path = Path(output_folder) / f"Normalized_Export_{timestamp}.xlsx"

            self.new_df.to_excel(excel_path, index=False, sheet_name='Normalized_Data', engine='xlsxwriter')

            self.dialog_status.setText(f"✓ Exported: {excel_path.name}")
            self.dialog_status.setStyleSheet("color: green; padding: 10px; background-color: #f0f0f0;")

        except Exception as e:
            self.dialog_status.setText(f"❌ Export failed: {str(e)}")
            self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")

    def _do_export_normalized_csv(self, dialog):
        """Export normalized data as CSV (UTF-8)"""
        try:
            output_folder = self._get_output_folder()
            if not output_folder:
                self.dialog_status.setText("❌ Output folder not configured")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            if self.new_df is None or self.new_df.empty:
                self.dialog_status.setText("❌ No normalized data to export")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_path = Path(output_folder) / f"Normalized_Export_{timestamp}.csv"

            self.new_df.to_csv(csv_path, index=False, encoding='utf-8')

            self.dialog_status.setText(f"✓ Exported: {csv_path.name}")
            self.dialog_status.setStyleSheet("color: green; padding: 10px; background-color: #f0f0f0;")

        except Exception as e:
            self.dialog_status.setText(f"❌ Export failed: {str(e)}")
            self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")

    def _do_export_mfg_xml(self, dialog):
        """Export MFG XML (manufacturers)"""
        try:
            output_folder = self._get_output_folder()
            if not output_folder:
                self.dialog_status.setText("❌ Output folder not configured")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            if self.new_df is None or 'MFG' not in self.new_df.columns:
                self.dialog_status.setText("❌ No MFG data to export")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            project_name = self.project_name_input.text().strip() or DEFAULT_PROJECT_NAME
            catalog = self.catalog_input.text().strip() or DEFAULT_CATALOG

            # Get unique manufacturers
            manufacturers = self.new_df['MFG'].dropna().unique().tolist()
            manufacturers = [m for m in manufacturers if str(m).strip()]

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            xml_path = Path(output_folder) / f"MFG_{timestamp}.xml"

            count = create_mfg_xml(manufacturers, str(xml_path), project_name, catalog)

            self.dialog_status.setText(f"✓ Exported {count} manufacturers: {xml_path.name}")
            self.dialog_status.setStyleSheet("color: green; padding: 10px; background-color: #f0f0f0;")

        except Exception as e:
            self.dialog_status.setText(f"❌ Export failed: {str(e)}")
            self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")

    def _do_export_mfgpn_xml(self, dialog):
        """Export MFG PN XML (part numbers)"""
        try:
            output_folder = self._get_output_folder()
            if not output_folder:
                self.dialog_status.setText("❌ Output folder not configured")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            if self.new_df is None or 'MFG' not in self.new_df.columns or 'MFG_PN' not in self.new_df.columns:
                self.dialog_status.setText("❌ No MFG/MFG_PN data to export")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            project_name = self.project_name_input.text().strip() or DEFAULT_PROJECT_NAME
            catalog = self.catalog_input.text().strip() or DEFAULT_CATALOG

            # Build MFG PN data
            mfgpn_data = []
            for _, row in self.new_df.iterrows():
                mfg = str(row.get('MFG', '')).strip()
                mfg_pn = str(row.get('MFG_PN', '')).strip()
                description = str(row.get('Description', '')).strip() if 'Description' in self.new_df.columns else ''

                if mfg and mfg_pn:
                    mfgpn_data.append({
                        'MFG': mfg,
                        'MFG_PN': mfg_pn,
                        'Description': description
                    })

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            xml_path = Path(output_folder) / f"MFGPN_{timestamp}.xml"

            count = create_mfgpn_xml(mfgpn_data, str(xml_path), project_name, catalog)

            self.dialog_status.setText(f"✓ Exported {count} part numbers: {xml_path.name}")
            self.dialog_status.setStyleSheet("color: green; padding: 10px; background-color: #f0f0f0;")

        except Exception as e:
            self.dialog_status.setText(f"❌ Export failed: {str(e)}")
            self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")

    def _do_export_both_xml(self, dialog):
        """Export both MFG and MFG PN XMLs"""
        try:
            output_folder = self._get_output_folder()
            if not output_folder:
                self.dialog_status.setText("❌ Output folder not configured")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            if self.new_df is None or 'MFG' not in self.new_df.columns or 'MFG_PN' not in self.new_df.columns:
                self.dialog_status.setText("❌ No MFG/MFG_PN data to export")
                self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
                return

            project_name = self.project_name_input.text().strip() or DEFAULT_PROJECT_NAME
            catalog = self.catalog_input.text().strip() or DEFAULT_CATALOG
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Export MFG XML
            manufacturers = self.new_df['MFG'].dropna().unique().tolist()
            manufacturers = [m for m in manufacturers if str(m).strip()]
            mfg_xml_path = Path(output_folder) / f"MFG_{timestamp}.xml"
            mfg_count = create_mfg_xml(manufacturers, str(mfg_xml_path), project_name, catalog)

            # Export MFG PN XML
            mfgpn_data = []
            for _, row in self.new_df.iterrows():
                mfg = str(row.get('MFG', '')).strip()
                mfg_pn = str(row.get('MFG_PN', '')).strip()
                description = str(row.get('Description', '')).strip() if 'Description' in self.new_df.columns else ''

                if mfg and mfg_pn:
                    mfgpn_data.append({
                        'MFG': mfg,
                        'MFG_PN': mfg_pn,
                        'Description': description
                    })

            mfgpn_xml_path = Path(output_folder) / f"MFGPN_{timestamp}.xml"
            mfgpn_count = create_mfgpn_xml(mfgpn_data, str(mfgpn_xml_path), project_name, catalog)

            self.dialog_status.setText(
                f"✓ Exported:\n"
                f"  • {mfg_count} manufacturers: {mfg_xml_path.name}\n"
                f"  • {mfgpn_count} part numbers: {mfgpn_xml_path.name}"
            )
            self.dialog_status.setStyleSheet("color: green; padding: 10px; background-color: #f0f0f0;")

        except Exception as e:
            self.dialog_status.setText(f"❌ Export failed: {str(e)}")
            self.dialog_status.setStyleSheet("color: red; padding: 10px; background-color: #f0f0f0;")
