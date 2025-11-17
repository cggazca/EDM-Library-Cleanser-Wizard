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
        QMessageBox
    )
    from PyQt5.QtCore import Qt
    from PyQt5.QtGui import QColor
except ImportError:
    print("Error: PyQt5 is required. Install it with: pip install PyQt5")


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

        # Export options
        export_group = QGroupBox("Export Options")
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
        """Initialize by loading Combined and Combined_New sheets"""
        # Get the output Excel file path from ColumnMappingPage
        column_mapping_page = self.wizard().page(2)

        if not hasattr(column_mapping_page, 'output_excel_path') or not column_mapping_page.output_excel_path:
            self.summary_label.setText("Output Excel file not found. Please go back and complete Step 2.")
            return

        excel_path = column_mapping_page.output_excel_path

        if not os.path.exists(excel_path):
            self.summary_label.setText(f"Excel file not found: {excel_path}")
            return

        try:
            # Load Combined (original) sheet
            if 'Combined' in pd.ExcelFile(excel_path).sheet_names:
                self.original_df = pd.read_excel(excel_path, sheet_name='Combined')
            else:
                self.summary_label.setText("'Combined' sheet not found in Excel file")
                return

            # Load Combined_New (after changes) sheet
            if 'Combined_New' in pd.ExcelFile(excel_path).sheet_names:
                self.new_df = pd.read_excel(excel_path, sheet_name='Combined_New')
            else:
                # If Combined_New doesn't exist yet, use Combined as placeholder
                self.new_df = self.original_df.copy()
                self.summary_label.setText("'Combined_New' sheet not found - showing original data only")

            # Build comparison
            self.build_comparison()

        except Exception as e:
            self.summary_label.setText(f"Error loading data: {str(e)}")

    def get_mapped_columns(self):
        """Get only the mapped columns from Column Mapping step"""
        # These are the standard column names after combination
        mapped_columns = ['MFG', 'MFG_PN', 'Part_Number', 'Description', 'Source_Sheet']

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
            'Source_Sheet': 'Source Sheet'
        }
        return display_names.get(col, col)

    def build_comparison(self):
        """Build side-by-side comparison with Beyond Compare styling"""
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
        self.summary_label.setText(
            f"<b>Total Rows:</b> {total} | "
            f"<b>Changed:</b> {changed_count} ({changed_count/total*100:.1f}%) | "
            f"<b>Unchanged:</b> {unchanged} ({unchanged/total*100:.1f}%)"
        )

        # Populate tables
        self.populate_tables()

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
