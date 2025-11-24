"""
PAS Search Page: Part Aggregation Service API Search
"""

import sys
import os
import time
from pathlib import Path
import pandas as pd
import json
from datetime import datetime

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel,
        QPushButton, QProgressBar, QMessageBox, QTextEdit, QWidget,
        QTableWidget, QTableWidgetItem, QHeaderView, QApplication
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
    from PyQt5.QtGui import QColor
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

from edm_wizard.api.pas_client import PASAPIClient
from edm_wizard.workers.threads import PASSearchThread



class PASSearchPage(QWizardPage):
    """Step 3: Search parts using PAS API"""

    def __init__(self):
        super().__init__()
        self.setTitle("Step 3: Part Search via PAS API")
        self.setSubTitle("Search for parts using the Part Aggregation Service")

        layout = QVBoxLayout()

        # Info section
        info_group = QGroupBox("🔍 Part Search Information")
        info_layout = QVBoxLayout()
        
        info_text = QLabel(
            "This step will search for each part in your data using the Part Aggregation Service (PAS).\n"
            "The search results will be saved as a CSV file for review in the next step."
        )
        info_text.setWordWrap(True)
        info_layout.addWidget(info_text)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # Search button
        self.search_button = QPushButton("🔍 Start Part Search")
        self.search_button.clicked.connect(self.start_search)
        self.search_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 10px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        layout.addWidget(self.search_button)

        # Progress
        self.progress_label = QLabel("")
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Results Preview (real-time grid)
        results_group = QGroupBox("📊 Search Results Preview")
        results_layout = QVBoxLayout()

        self.results_table = QTableWidget()
        self.results_table.setColumnCount(5)
        self.results_table.setHorizontalHeaderLabels([
            "Part Number",
            "Manufacturer",
            "Match Status",
            "Match Details",
            "Search Time"
        ])

        # Enable sorting
        self.results_table.setSortingEnabled(True)

        # Set column resize modes
        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)  # Part Number
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Manufacturer
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Match Status
        header.setSectionResizeMode(3, QHeaderView.Stretch)  # Match Details
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Search Time

        self.results_table.setAlternatingRowColors(True)
        self.results_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.results_table.setSelectionMode(QTableWidget.SingleSelection)
        self.results_table.setEditTriggers(QTableWidget.NoEditTriggers)  # Make read-only

        results_layout.addWidget(self.results_table)
        results_group.setLayout(results_layout)
        layout.addWidget(results_group, stretch=1)

        # Summary
        summary_group = QGroupBox("Search Summary")
        summary_layout = QVBoxLayout()
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setMaximumHeight(100)  # Limit height
        summary_layout.addWidget(self.summary_text)
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)

        self.setLayout(layout)

        self.search_completed = False
        self.search_results = []
        self.combined_data = []
        self.csv_output_path = None

    def initializePage(self):
        """Initialize and automatically load combined data from Step 2"""
        # Get combined data from ColumnMappingPage (Step 2)
        column_mapping_page = self.wizard().page(2)  # ColumnMappingPage is page 2

        # Check if combined data is available
        if hasattr(column_mapping_page, 'combined_data') and column_mapping_page.combined_data is not None:
            if isinstance(column_mapping_page.combined_data, pd.DataFrame) and not column_mapping_page.combined_data.empty:
                # Use the combined data directly from ColumnMappingPage
                self.combined_data = column_mapping_page.combined_data.copy()

                # Update info label to show data is loaded
                parts_count = len(self.combined_data)
                cols = list(self.combined_data.columns)

                self.progress_label.setText(
                    f"✓ Auto-loaded {parts_count} parts from Step 2 (Combined sheet)\n"
                    f"Columns: {', '.join(cols[:5])}{'...' if len(cols) > 5 else ''}\n"
                    f"Click 'Start Part Search' to begin."
                )
                self.progress_label.setStyleSheet("color: green; font-weight: bold;")
                self.search_button.setEnabled(True)

                print(f"SUCCESS: Loaded {parts_count} parts from combined DataFrame")
            else:
                self.progress_label.setText("⚠ No data available after filtering. Please go back to Step 2 and adjust filter conditions.")
                self.progress_label.setStyleSheet("color: orange;")
                self.search_button.setEnabled(False)
                self.combined_data = pd.DataFrame()
        else:
            self.progress_label.setText(
                "⚠ No combined data available. Please go back to Step 2 and click 'Next' to combine sheets.\n"
                "Make sure at least one sheet is selected and has MFG/MFG PN columns mapped."
            )
            self.progress_label.setStyleSheet("color: orange;")
            self.search_button.setEnabled(False)
            self.combined_data = pd.DataFrame()

    def start_search(self):
        """Start the PAS search process using preloaded data"""
        try:
            # Validate that data is loaded
            if self.combined_data is None or self.combined_data.empty:
                QMessageBox.warning(
                    self,
                    "No Data",
                    "No parts data available to search.\n\n"
                    "Please go back to Step 3 and ensure data is combined."
                )
                return

            # Get PAS credentials from Start Page
            start_page = self.wizard().page(0)
            pas_creds = start_page.get_pas_credentials() if hasattr(start_page, 'get_pas_credentials') else None

            if not pas_creds:
                QMessageBox.warning(
                    self,
                    "Missing Credentials",
                    "PAS API credentials are required.\n\n"
                    "Please go back to Step 1 and enter your credentials."
                )
                return

            # Get output folder from StartPage
            output_folder = start_page.get_output_folder() if hasattr(start_page, 'get_output_folder') else None
            if not output_folder:
                QMessageBox.warning(
                    self,
                    "Missing Output Folder",
                    "Output folder is not configured.\n\n"
                    "Please go back to Step 1 and select an output folder."
                )
                return

            # Get max matches setting from StartPage
            max_matches = start_page.get_max_matches() if hasattr(start_page, 'get_max_matches') else 10
            
            # Create PAS client and store it as instance variable for re-search functionality
            self.pas_client = PASAPIClient(
                client_id=pas_creds['client_id'],
                client_secret=pas_creds['client_secret'],
                max_matches=max_matches
            )

            # Disable button and show progress
            self.search_button.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setMaximum(len(self.combined_data))
            self.progress_bar.setValue(0)

            # Store output folder for later use
            self.output_folder = Path(output_folder)

            # Clear results table
            self.results_table.setRowCount(0)

            # Convert DataFrame to list of dictionaries for the search thread
            parts_list = self.combined_data.to_dict('records')

            # Start search thread with parallel execution
            # max_workers=30 means 30 concurrent PAS API calls (adjustable for performance)
            self.search_thread = PASSearchThread(self.pas_client, parts_list, max_workers=30)
            self.search_thread.progress.connect(self.on_search_progress)
            self.search_thread.result_ready.connect(self.on_result_ready)  # Real-time display
            self.search_thread.finished.connect(self.on_search_finished)
            self.search_thread.error.connect(self.on_search_error)
            self.search_thread.start()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to start search:\n{str(e)}")
            self.search_button.setEnabled(True)

    def extract_from_sheets(self, dataframes, mappings):
        """Extract part data from individual sheets"""
        included_sheets = self.wizard().page(2).get_included_sheets()

        for sheet_name, df in dataframes.items():
            if sheet_name not in included_sheets:
                continue

            mapping = mappings.get(sheet_name, {})
            if not mapping.get('MFG') or not mapping.get('MFG_PN'):
                continue

            mfg_col = mapping['MFG']
            mfgpn_col = mapping['MFG_PN']
            desc_col = mapping.get('Description', '')

            for _, row in df.iterrows():
                if pd.notna(row.get(mfg_col)) and pd.notna(row.get(mfgpn_col)):
                    self.combined_data.append({
                        'MFG': str(row[mfg_col]).strip(),
                        'MFG_PN': str(row[mfgpn_col]).strip(),
                        'Description': str(row[desc_col]) if desc_col and pd.notna(row.get(desc_col)) else ''
                    })

    def on_search_progress(self, message, current, total):
        """Update progress during search"""
        self.progress_label.setText(message)
        self.progress_bar.setValue(current)

    def on_result_ready(self, result):
        """Add individual result to table in real-time"""
        from datetime import datetime

        # Temporarily disable sorting while adding row
        self.results_table.setSortingEnabled(False)

        row_position = self.results_table.rowCount()
        self.results_table.insertRow(row_position)

        # Part Number (convert to string to handle numeric part numbers)
        self.results_table.setItem(row_position, 0, QTableWidgetItem(str(result['PartNumber'])))

        # Manufacturer (convert to string to handle any numeric values)
        self.results_table.setItem(row_position, 1, QTableWidgetItem(str(result['ManufacturerName'])))

        # Match Status (with color coding)
        status_item = QTableWidgetItem(result['MatchStatus'])
        if result['MatchStatus'] == 'Found':
            status_item.setBackground(QColor(230, 255, 230))  # Light green
        elif result['MatchStatus'] == 'Multiple':
            status_item.setBackground(QColor(255, 240, 200))  # Light orange
        elif result['MatchStatus'] == 'Need user review':
            status_item.setBackground(QColor(230, 240, 255))  # Light blue
        elif result['MatchStatus'] == 'None':
            status_item.setBackground(QColor(240, 240, 240))  # Light gray
        elif result['MatchStatus'] == 'Error':
            status_item.setBackground(QColor(255, 230, 230))  # Light red
        self.results_table.setItem(row_position, 2, status_item)

        # Match Details
        matches = result.get('matches', [])
        if matches:
            # Format matches for display (handles both dict and string formats)
            match_strings = []
            for match in matches[:3]:
                if isinstance(match, dict):
                    match_strings.append(match.get('match_string', str(match)))
                else:
                    match_strings.append(str(match))
            match_details = ', '.join(match_strings)
            if len(matches) > 3:
                match_details += f' ... (+{len(matches) - 3} more)'
        else:
            match_details = 'No matches found'
        self.results_table.setItem(row_position, 3, QTableWidgetItem(match_details))

        # Search Time
        current_time = datetime.now().strftime("%H:%M:%S")
        self.results_table.setItem(row_position, 4, QTableWidgetItem(current_time))

        # Re-enable sorting
        self.results_table.setSortingEnabled(True)

        # Auto-scroll to latest result
        self.results_table.scrollToBottom()

    def on_search_finished(self, results):
        """Handle search completion"""
        self.search_results = results
        self.search_completed = True

        # Save results to CSV in output folder from StartPage
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_filename = f"PAS_Search_Results_{timestamp}.csv"
        self.csv_output_path = self.output_folder / csv_filename

        try:
            # Update UI to indicate saving is in progress
            self.progress_label.setText("Saving results to CSV...")
            self.progress_label.setStyleSheet("color: blue; font-weight: bold;")
            QApplication.processEvents()  # Force UI update

            self.save_results_csv()

            # Count results
            exact = sum(1 for r in results if r['MatchStatus'] == 'Found')
            multiple = sum(1 for r in results if r['MatchStatus'] == 'Multiple')
            none = sum(1 for r in results if r['MatchStatus'] == 'None')
            review = sum(1 for r in results if r['MatchStatus'] == 'Need user review')

            # Show summary
            summary = f"✓ Part Search Completed!\n\n"
            summary += f"Total parts searched: {len(results)}\n"
            summary += f"  - Exact matches (Found): {exact}\n"
            summary += f"  - Multiple matches: {multiple}\n"
            summary += f"  - No matches: {none}\n"
            summary += f"  - Need review: {review}\n\n"
            summary += f"Results saved to:\n{self.csv_output_path}\n\n"
            summary += f"Proceed to Step 4 to review and normalize matches."

            self.summary_text.setText(summary)
            self.progress_label.setText("✓ Search completed successfully!")
            self.progress_label.setStyleSheet("color: green; font-weight: bold;")

            self.completeChanged.emit()

            QMessageBox.information(
                self,
                "Search Complete",
                f"Successfully searched {len(results)} parts!\n\n"
                f"Exact matches: {exact}\n"
                f"Multiple matches: {multiple}\n"
                f"No matches: {none}\n"
                f"Need review: {review}\n\n"
                f"Results saved to:\n{csv_filename}"
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save results:\n{str(e)}")

        self.search_button.setEnabled(True)
        self.progress_bar.setVisible(False)

    def on_search_error(self, error_msg):
        """Handle search error"""
        self.progress_label.setText(f"✗ Search failed: {error_msg[:50]}...")
        self.progress_label.setStyleSheet("color: red;")
        self.search_button.setEnabled(True)
        self.progress_bar.setVisible(False)

        QMessageBox.critical(self, "Search Error", f"Search failed:\n{error_msg}")

    def save_results_csv(self):
        """Save search results to CSV in SearchAndAssign format"""
        import csv

        with open(self.csv_output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # Write header with new fields
            writer.writerow([
                'PartNumber', 
                'ManufacturerName', 
                'MatchStatus', 
                'MatchValue(PartNumber@ManufacturerName)',
                'Lifecycle_Status',
                'Lifecycle_Code',
                'External_ID'
            ])
            
            # Write data
            for result in self.search_results:
                part_number = result['PartNumber']
                manufacturer = result['ManufacturerName']
                status = result['MatchStatus']
                matches = result.get('matches', [])
                
                # Write one row per match (or one row if no matches)
                if matches:
                    for match in matches:
                        # Extract match information (handles both dict and string formats)
                        if isinstance(match, dict):
                            match_string = match.get('match_string', '')
                            lifecycle_status = match.get('lifecycle_status', '')
                            lifecycle_code = match.get('lifecycle_code', '')
                            external_id = match.get('external_id', '')
                        elif isinstance(match, str):
                            match_string = match
                            lifecycle_status = ''
                            lifecycle_code = ''
                            external_id = ''
                        else:
                            match_string = str(match)
                            lifecycle_status = ''
                            lifecycle_code = ''
                            external_id = ''
                        
                        writer.writerow([
                            part_number, 
                            manufacturer, 
                            status, 
                            match_string,
                            lifecycle_status,
                            lifecycle_code,
                            external_id
                        ])
                else:
                    writer.writerow([part_number, manufacturer, status, '', '', '', ''])

    def isComplete(self):
        """Check if search is complete"""
        return self.search_completed

