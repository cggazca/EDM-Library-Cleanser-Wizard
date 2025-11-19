#!/usr/bin/env python3
"""
EDM Library Wizard
A comprehensive wizard for converting Access databases to Excel and generating XML files for EDM Library Creator
"""

import sys
import os
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import sqlalchemy as sa
import urllib
from sqlalchemy import inspect
import xml.etree.ElementTree as ET
from xml.dom import minidom
import time
import requests
import threading

try:
    from PyQt5.QtWidgets import (
        QApplication, QWizard, QWizardPage, QVBoxLayout, QHBoxLayout,
        QRadioButton, QPushButton, QLabel, QLineEdit, QFileDialog,
        QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QComboBox,
        QGroupBox, QMessageBox, QTextEdit, QProgressBar, QSpacerItem,
        QSizePolicy, QGridLayout, QWidget, QSplitter, QScrollArea, QMenu,
        QTabWidget, QButtonGroup, QSpinBox, QInputDialog, QDialog,
        QListWidget, QListWidgetItem, QDialogButtonBox
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings
    from PyQt5.QtGui import QFont, QIcon, QColor
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

try:
    from fuzzywuzzy import fuzz, process
    FUZZYWUZZY_AVAILABLE = True
except ImportError:
    FUZZYWUZZY_AVAILABLE = False


# Import wizard pages from separate modules
# Use relative imports to avoid shadowing with edm_wizard.py script
import sys
import os as _os
_script_dir = _os.path.dirname(_os.path.abspath(__file__))
if _script_dir not in sys.path:
    sys.path.insert(0, _script_dir)

from edm_wizard.ui.pages import (
    StartPage,
    DataSourcePage,
    ColumnMappingPage,
    PASSearchPage,
    XMLGenerationPage,
    SupplyFrameReviewPage
)

class CollapsibleGroupBox(QGroupBox):
    """A QGroupBox that can be collapsed/expanded by clicking the title"""

    def __init__(self, title="", parent=None):
        super().__init__(title, parent)
        self.setCheckable(True)
        self.setChecked(True)  # Expanded by default
        self.toggled.connect(self.on_toggled)

        # Store the content widget
        self._content_widget = None

    def setContentLayout(self, layout):
        """Set the content layout that will be shown/hidden"""
        if self._content_widget:
            self._content_widget.deleteLater()

        self._content_widget = QWidget()
        self._content_widget.setLayout(layout)

        # Create main layout for the group box
        main_layout = QVBoxLayout()
        main_layout.addWidget(self._content_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        super().setLayout(main_layout)

    def on_toggled(self, checked):
        """Show/hide content when toggled"""
        if self._content_widget:
            self._content_widget.setVisible(checked)


class ComparisonPage(QWizardPage):
    """Step 5: Side-by-Side Comparison - Beyond Compare Style"""

    def __init__(self):
        super().__init__()
        self.setTitle("Step 5: Review Changes - Side-by-Side Comparison")
        self.setSubTitle("Compare Combined (original) vs Combined_New (with normalization)")

        layout = QVBoxLayout()

        # Summary section
        summary_group = QGroupBox("📊 Changes Summary")
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
        left_group = QGroupBox("📄 Combined (Original)")
        left_layout = QVBoxLayout()

        self.left_table = QTableWidget()
        self.left_table.setSortingEnabled(True)  # Enable sorting
        self.left_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.left_table.setSelectionMode(QTableWidget.SingleSelection)
        self.left_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.left_table.verticalScrollBar().valueChanged.connect(self.sync_scroll_right)

        left_layout.addWidget(self.left_table)
        left_group.setLayout(left_layout)
        comparison_layout.addWidget(left_group)

        # Right table: Combined_New (After Changes)
        right_group = QGroupBox("📄 Combined_New (After Changes)")
        right_layout = QVBoxLayout()

        self.right_table = QTableWidget()
        self.right_table.setSortingEnabled(True)  # Enable sorting
        self.right_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.right_table.setSelectionMode(QTableWidget.SingleSelection)
        self.right_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.right_table.verticalScrollBar().valueChanged.connect(self.sync_scroll_left)

        right_layout.addWidget(self.right_table)
        right_group.setLayout(right_layout)
        comparison_layout.addWidget(right_group)

        layout.addLayout(comparison_layout, stretch=1)

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
        """Initialize by loading Combined and Combined_New sheets"""
        # Get the output Excel file path from ColumnMappingPage
        column_mapping_page = self.wizard().page(2)

        if not hasattr(column_mapping_page, 'output_excel_path') or not column_mapping_page.output_excel_path:
            self.summary_label.setText("❌ Output Excel file not found. Please go back and complete Step 2.")
            return

        excel_path = column_mapping_page.output_excel_path

        if not os.path.exists(excel_path):
            self.summary_label.setText(f"❌ Excel file not found: {excel_path}")
            return

        try:
            # Load Combined (original) sheet
            if 'Combined' in pd.ExcelFile(excel_path).sheet_names:
                self.original_df = pd.read_excel(excel_path, sheet_name='Combined')
            else:
                self.summary_label.setText("❌ 'Combined' sheet not found in Excel file")
                return

            # Load Combined_New (after changes) sheet
            if 'Combined_New' in pd.ExcelFile(excel_path).sheet_names:
                self.new_df = pd.read_excel(excel_path, sheet_name='Combined_New')
            else:
                # If Combined_New doesn't exist yet, use Combined as placeholder
                self.new_df = self.original_df.copy()
                self.summary_label.setText("⚠ 'Combined_New' sheet not found - showing original data only")

            # Build comparison
            self.build_comparison()

        except Exception as e:
            self.summary_label.setText(f"❌ Error loading data: {str(e)}")

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

            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_path = Path(output_folder) / f"Comparison_{timestamp}.csv"

            import csv
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Part Number', 'Original MFG', 'New MFG', 'Change Type', 'Notes'])

                for change in self.changes:
                    writer.writerow([
                        change['part_number'],
                        change['original_mfg'],
                        change['new_mfg'],
                        change['change_type'],
                        change['notes']
                    ])

            self.export_status.setText(f"✓ Exported to: {csv_path.name}")
            self.export_status.setStyleSheet("color: green;")

        except Exception as e:
            self.export_status.setText(f"✗ Export failed: {str(e)}")
            self.export_status.setStyleSheet("color: red;")

    def export_to_excel(self):
        """Export comparison to Excel"""
        try:
            start_page = self.wizard().page(0)
            output_folder = start_page.get_output_folder() if hasattr(start_page, 'get_output_folder') else None

            if not output_folder:
                QMessageBox.warning(self, "Error", "Output folder not configured")
                return

            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_path = Path(output_folder) / f"Comparison_{timestamp}.xlsx"

            # Create DataFrame
            df = pd.DataFrame(self.changes)
            df.columns = ['Part Number', 'Original MFG', 'New MFG', 'Change Type', 'Notes']

            # Write to Excel
            df.to_excel(excel_path, index=False, engine='xlsxwriter')

            self.export_status.setText(f"✓ Exported to: {excel_path.name}")
            self.export_status.setStyleSheet("color: green;")

        except Exception as e:
            self.export_status.setText(f"✗ Export failed: {str(e)}")
            self.export_status.setStyleSheet("color: red;")


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

        # Add pages - New 6-page flow
        self.start_page = StartPage()                          # Welcome: API credentials & output folder
        self.data_source_page = DataSourcePage()               # Step 1: Access DB export or Excel selection
        self.column_mapping_page = ColumnMappingPage()         # Step 2: Column mapping & combine
        self.pas_search_page = PASSearchPage()                 # Step 3: PAS API search
        self.review_page = SupplyFrameReviewPage()             # Step 4: Review matches & normalization
        self.comparison_page = ComparisonPage()                # Step 5: Old vs New comparison

        self.addPage(self.start_page)              # Page 0
        self.addPage(self.data_source_page)        # Page 1
        self.addPage(self.column_mapping_page)     # Page 2
        self.addPage(self.pas_search_page)         # Page 3
        self.addPage(self.review_page)             # Page 4
        self.addPage(self.comparison_page)         # Page 5

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
