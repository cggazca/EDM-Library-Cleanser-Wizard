"""
Supply Frame Review Page: Results Review and Normalization
"""

import sys
import os
from pathlib import Path
import pandas as pd
import json
from datetime import datetime
import difflib

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QLineEdit,
        QPushButton, QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox,
        QComboBox, QMessageBox, QWidget, QTabWidget, QScrollArea, QSpinBox,
        QInputDialog, QMenu, QTextEdit, QDialog, QDialogButtonBox, QSplitter,
        QButtonGroup
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
    from PyQt5.QtGui import QColor, QFont
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

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

from edm_wizard.utils.xml_generation import escape_xml
from edm_wizard.workers.threads import PartialMatchAIThread, ManufacturerNormalizationAIThread



class SupplyFrameReviewPage(QWizardPage):
    """Step 4: Review PAS Matches and Normalize Manufacturers"""

    def __init__(self):
        super().__init__()
        self.setTitle("Step 4: Review Matches & Manufacturer Normalization")
        self.setSubTitle("Review match results (Found/Multiple/Need Review/None) and normalize manufacturer names")

        self.search_results = []
        self.parts_needing_review = []
        self.manufacturer_normalizations = {}
        self.normalization_reasoning = {}  # Store fuzzy/AI reasoning for each normalization
        self.normalization_scopes = {}  # Store selected sheets for each normalization row {row_idx: [sheet1, sheet2, ...]}
        self.original_data = []  # Store original data for comparison
        self.api_key = None
        self.ai_cache = {}  # Cache AI normalization results to ensure consistency
        
        # Initialize categorized parts lists
        self.found_parts = []
        self.multiple_parts = []
        self.need_review_parts = []
        self.none_parts = []
        self.errors_parts = []

        # Main layout with vertical splitter for resizable sections
        page_layout = QVBoxLayout()

        # Create vertical splitter to allow resizing sections
        main_splitter = QSplitter(Qt.Vertical)
        main_splitter.setChildrenCollapsible(True)  # Allow sections to be collapsed via splitter

        # Section 1: Match Results Summary (auto-loaded from PASSearchPage)
        self.summary_group = self.create_summary_section_widget()
        main_splitter.addWidget(self.summary_group)

        # Section 2: Review Matches by Category
        self.review_group = self.create_review_section_widget()
        main_splitter.addWidget(self.review_group)

        # Section 3: Manufacturer Normalization
        self.norm_group = self.create_normalization_section_widget()
        main_splitter.addWidget(self.norm_group)

        # Set initial sizes for splitter sections (in pixels)
        # Summary: 150, Review: 500, Normalization: 400
        main_splitter.setSizes([150, 500, 400])

        # Set stretch factors (higher = takes more space when expanding)
        main_splitter.setStretchFactor(0, 0)  # Summary - minimal
        main_splitter.setStretchFactor(1, 3)  # Review - largest
        main_splitter.setStretchFactor(2, 2)  # Normalization - medium

        page_layout.addWidget(main_splitter)
        self.setLayout(page_layout)

    @staticmethod
    def _get_match_info(match):
        """
        Helper function to extract match information from either old string format or new dict format.
        Returns: (mpn, mfg, lifecycle_status, lifecycle_code, external_id, findchips_url, match_string)
        """
        if isinstance(match, dict):
            # New dict format
            return (
                match.get('mpn', ''),
                match.get('mfg', ''),
                match.get('lifecycle_status', ''),
                match.get('lifecycle_code', ''),
                match.get('external_id', ''),
                match.get('findchips_url', ''),
                match.get('match_string', '')
            )
        elif isinstance(match, str):
            # Old string format: "PartNumber@Manufacturer"
            if '@' in match:
                pn, mfg = match.split('@', 1)
            else:
                pn = match
                mfg = ''
            return (pn, mfg, '', '', '', '', match)
        else:
            # Invalid format
            return ('', '', '', '', '', '', '')

    def initializePage(self):
        """Initialize by loading data from CSV file created by PASSearchPage"""
        pas_search_page = self.wizard().page(3)  # PASSearchPage is page 3

        # Check if page exists
        if pas_search_page is None:
            QMessageBox.critical(
                self,
                "Error",
                "Could not find PAS Search Page (Step 4).\n\n"
                "This is an internal error."
            )
            return

        # Get CSV path from PASSearchPage
        if not hasattr(pas_search_page, 'csv_output_path') or pas_search_page.csv_output_path is None:
            QMessageBox.warning(
                self,
                "No Data",
                "No search results CSV file found.\n\n"
                "Please go back to Step 4 and complete the PAS search."
            )
            return

        # Read results from CSV
        csv_path = pas_search_page.csv_output_path
        if not Path(csv_path).exists():
            QMessageBox.warning(
                self,
                "File Not Found",
                f"Search results CSV file not found:\n{csv_path}\n\n"
                "Please go back to Step 4 and complete the PAS search."
            )
            return

        try:
            # Load search results from CSV
            self.search_results = self.load_results_from_csv(csv_path)

            # Store original data for comparison later (convert DataFrame to list of dicts)
            # Get combined_data from ColumnMappingPage (page 2), not PAS Search Page
            column_mapping_page = self.wizard().page(2)  # ColumnMappingPage
            if hasattr(column_mapping_page, 'combined_data') and column_mapping_page.combined_data is not None:
                if not column_mapping_page.combined_data.empty:
                    # Convert DataFrame to list of dictionaries for easier processing
                    self.original_data = column_mapping_page.combined_data.to_dict('records')
                else:
                    self.original_data = []
            else:
                self.original_data = []

            # Load and display the results
            self.load_search_results()

        except Exception as e:
            import traceback
            error_details = ''.join(traceback.format_exception(type(e), e, e.__traceback__))
            print(f"ERROR: {error_details}")  # Print to console for debugging
            QMessageBox.critical(
                self,
                "Error Loading Results",
                f"Failed to load search results from CSV:\n{str(e)}\n\n"
                f"File: {csv_path}\n\n"
                f"Details:\n{error_details}"
            )

    def load_results_from_csv(self, csv_path):
        """Load search results from CSV file"""
        import csv
        from collections import defaultdict

        results = []

        # Group rows by PartNumber+ManufacturerName (since multiple matches create multiple rows)
        grouped = defaultdict(lambda: {'matches': [], 'lifecycle_data': {}, 'external_ids': {}})

        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                key = (row['PartNumber'], row['ManufacturerName'])

                # First occurrence - set basic info
                if not grouped[key]['matches']:
                    grouped[key]['PartNumber'] = row['PartNumber']
                    grouped[key]['ManufacturerName'] = row['ManufacturerName']
                    grouped[key]['MatchStatus'] = row['MatchStatus']

                # Add match value if present (as dict with lifecycle and external ID)
                match_value = row.get('MatchValue(PartNumber@ManufacturerName)', '').strip()
                if match_value:
                    # Create match dict with all available fields
                    match_dict = {
                        'match_string': match_value,
                        'lifecycle_status': row.get('Lifecycle_Status', ''),
                        'lifecycle_code': row.get('Lifecycle_Code', ''),
                        'external_id': row.get('External_ID', '')
                    }
                    # Parse mpn and mfg from match string
                    if '@' in match_value:
                        mpn, mfg = match_value.split('@', 1)
                        match_dict['mpn'] = mpn
                        match_dict['mfg'] = mfg
                    else:
                        match_dict['mpn'] = match_value
                        match_dict['mfg'] = ''
                    
                    grouped[key]['matches'].append(match_dict)

        # Convert to list format
        for data in grouped.values():
            results.append(data)

        return results

    def load_search_results(self):
        """Process and display search results"""
        # Categorize results by match status
        found = [r for r in self.search_results if r['MatchStatus'] == 'Found']
        multiple = [r for r in self.search_results if r['MatchStatus'] == 'Multiple']
        need_review = [r for r in self.search_results if r['MatchStatus'] == 'Need user review']
        none = [r for r in self.search_results if r['MatchStatus'] == 'None']
        errors = [r for r in self.search_results if r['MatchStatus'] == 'Error']

        # Update summary
        self.update_summary_display(found, multiple, need_review, none, errors)

        # Update tab counts
        self.review_tabs.setTabText(0, f"⚠ Multiple ({len(multiple)})")
        self.review_tabs.setTabText(1, f"👁 Need Review ({len(need_review)})")
        self.review_tabs.setTabText(2, f"✓ Found ({len(found)})")
        self.review_tabs.setTabText(3, f"✗ None ({len(none)})")
        self.review_tabs.setTabText(4, f"❌ Errors ({len(errors)})")

        # Store categorized data
        self.found_parts = found
        self.multiple_parts = multiple
        self.need_review_parts = need_review
        self.none_parts = none
        self.errors_parts = errors

        # Populate all tabs
        self.populate_category_table(self.found_table, found, show_actions=False)
        self.populate_category_table(self.multiple_table, multiple, show_actions=True)
        self.populate_category_table(self.need_review_table, need_review, show_actions=True)
        self.populate_category_table(self.none_table, none, show_actions="editable")  # Editable mode for None tab
        self.populate_category_table(self.errors_table, errors, show_actions=False)

        # For backward compatibility
        self.parts_needing_review = multiple + need_review

        # Enable buttons if there are parts to review
        start_page = self.wizard().page(0)
        api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None
        
        if len(multiple) > 0:
            self.multiple_auto_select_btn.setEnabled(True)
            self.multiple_ai_suggest_btn.setEnabled(bool(api_key and ANTHROPIC_AVAILABLE))
        
        if len(need_review) > 0:
            self.need_review_auto_select_btn.setEnabled(True)
            self.need_review_ai_suggest_btn.setEnabled(bool(api_key and ANTHROPIC_AVAILABLE))

        # Identify parts needing normalization
        self.identify_normalization_candidates()

    def create_summary_section_widget(self):
        """Section 1: Match Results Summary"""
        summary_group = QGroupBox("📊 Match Results Summary")
        summary_layout = QVBoxLayout()

        self.summary_label = QLabel("Loading results...")
        self.summary_label.setWordWrap(True)
        summary_layout.addWidget(self.summary_label)

        summary_group.setLayout(summary_layout)
        return summary_group

    def update_summary_display(self, found, multiple, need_review, none, errors):
        """Update the summary display with match counts"""
        total = len(self.search_results)
        summary_text = f"""
<b>Total Parts:</b> {total}<br>
<span style='color: green;'><b>✓ Found:</b> {len(found)}</span> ({len(found)/total*100:.1f}%)<br>
<span style='color: orange;'><b>⚠ Multiple:</b> {len(multiple)}</span> ({len(multiple)/total*100:.1f}%)<br>
<span style='color: blue;'><b>👁 Need Review:</b> {len(need_review)}</span> ({len(need_review)/total*100:.1f}%)<br>
<span style='color: gray;'><b>✗ None:</b> {len(none)}</span> ({len(none)/total*100:.1f}%)
"""
        if errors:
            summary_text += f"<br><span style='color: red;'><b>❌ Errors:</b> {len(errors)}</span>"

        self.summary_label.setText(summary_text)

    def create_csv_section_widget(self):
        """Section 1: Load SearchAndAssign CSV"""
        csv_group = QGroupBox("1. Load SearchAndAssign Results")
        csv_layout = QVBoxLayout()

        # File browser
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("CSV File:"))
        self.csv_path_input = QLineEdit()
        self.csv_path_input.setPlaceholderText("Select SearchAndAssign result CSV file...")
        file_layout.addWidget(self.csv_path_input)

        self.browse_csv_btn = QPushButton("Browse...")
        self.browse_csv_btn.clicked.connect(self.browse_csv)
        file_layout.addWidget(self.browse_csv_btn)

        self.load_csv_btn = QPushButton("Load CSV")
        self.load_csv_btn.clicked.connect(self.load_csv)
        self.load_csv_btn.setEnabled(False)
        file_layout.addWidget(self.load_csv_btn)

        csv_layout.addLayout(file_layout)

        # Summary
        self.csv_summary = QLabel("No file loaded")
        self.csv_summary.setStyleSheet("padding: 5px; background-color: #f0f0f0; border-radius: 3px;")
        csv_layout.addWidget(self.csv_summary)

        csv_group.setLayout(csv_layout)
        return csv_group


    def identify_normalization_candidates(self):
        """Identify manufacturers that need normalization using fuzzy matching"""
        if not FUZZYWUZZY_AVAILABLE:
            self.norm_status.setText("⚠ Fuzzy matching not available (install fuzzywuzzy)")
            return

        # Collect all manufacturer names from search results
        original_mfgs = set()
        canonical_mfgs = set()

        # From original data in all search results
        for result in self.search_results:
            mfg = result.get('ManufacturerName', '').strip()
            if mfg:
                original_mfgs.add(mfg)

        # Build MASTER LIST from PAS search matches
        # These are canonical manufacturer names validated by Siemens PAS database
        for result in self.search_results:
            for match in result.get('matches', []):
                # Extract manufacturer from match (handles both dict and string)
                _, mfg, _, _, _, _, _ = self._get_match_info(match)
                mfg = mfg.strip()
                if mfg:
                    canonical_mfgs.add(mfg)

        # Track manufacturers from USER-SELECTED matches (review phase work)
        # These are manufacturers the user specifically chose during review
        selected_mfgs = set()
        for result in self.search_results:
            if result.get('selected_match'):
                # Extract manufacturer from selected match (handles both dict and string)
                _, mfg, _, _, _, _, _ = self._get_match_info(result['selected_match'])
                mfg = mfg.strip()
                if mfg:
                    selected_mfgs.add(mfg)

        # Store both lists for future use
        self.canonical_manufacturers = sorted(list(canonical_mfgs))
        self.selected_manufacturers = sorted(list(selected_mfgs))  # User's review work

        if not original_mfgs or not canonical_mfgs:
            self.norm_status.setText("No manufacturer variations detected")
            # Still store empty list
            self.canonical_manufacturers = []
            return

        # Display master list info
        print(f"DEBUG: Built master manufacturer list with {len(canonical_mfgs)} canonical names from PAS")
        print(f"DEBUG: Comparing against {len(original_mfgs)} original manufacturer names")

        # Use fuzzy matching to find variations
        normalizations = {}
        reasoning_map = {}

        # Show ALL manufacturers, not just high-confidence matches
        for original in original_mfgs:
            # Skip if original is already in canonical list (exact match)
            if original in canonical_mfgs:
                # Still add to table but with itself as suggestion
                normalizations[original] = original
                reasoning_map[original] = {
                    'method': 'exact',
                    'score': 100,
                    'reasoning': f"Exact match - already in PAS canonical list"
                }
                continue

            # Find best match in canonical names using fuzzy matching
            best_match = None
            best_score = 0

            if canonical_mfgs:  # Only search if we have canonical names
                result = process.extractOne(original, canonical_mfgs, scorer=fuzz.ratio)
                if result:
                    best_match, best_score = result[0], result[1]

            # Add ALL manufacturers to the table with their best suggestion (if any)
            if best_match and best_score >= 70:  # Lower threshold for suggestions
                normalizations[original] = best_match
                reasoning_map[original] = {
                    'method': 'fuzzy',
                    'score': best_score,
                    'reasoning': f"Fuzzy match against PAS master list: {best_score}% similarity"
                }
                print(f"DEBUG: Normalization suggestion: '{original}' -> '{best_match}' ({best_score}%)")
            elif best_match and best_score >= 50:  # Medium confidence match
                # Show the best match but mark as needing review
                normalizations[original] = best_match
                reasoning_map[original] = {
                    'method': 'manual',
                    'score': best_score,
                    'reasoning': f"Low confidence match ({best_score}%) - please review"
                }
                print(f"DEBUG: Low confidence match for '{original}' -> '{best_match}' ({best_score}%) - needs review")
            else:
                # No good match - leave blank for manual entry
                # Use empty string so dropdown shows first canonical manufacturer
                normalizations[original] = ""
                reasoning_map[original] = {
                    'method': 'manual',
                    'score': 0,
                    'reasoning': f"No automatic match found - requires manual review"
                }
                print(f"DEBUG: No match for '{original}' - left blank for manual review")

        # Always populate the table if we have manufacturers
        if normalizations:
            self.manufacturer_normalizations = normalizations
            self.normalization_reasoning = reasoning_map

            # Populate normalization table
            self.norm_table.setRowCount(len(normalizations))

            row_idx = 0
            for original, canonical in normalizations.items():
                method = reasoning_map.get(original, {}).get('method', 'manual')
                score = reasoning_map.get(original, {}).get('score', 0)
                
                # Column 0: Include checkbox - centered
                include_cb = QCheckBox()
                # Uncheck exact matches (original == canonical, no change needed)
                # Check fuzzy matches and manual review items
                include_cb.setChecked(method != 'exact' and original != canonical)

                # Create a widget to center the checkbox
                checkbox_widget = QWidget()
                checkbox_layout = QHBoxLayout(checkbox_widget)
                checkbox_layout.addWidget(include_cb)
                checkbox_layout.setAlignment(Qt.AlignCenter)
                checkbox_layout.setContentsMargins(0, 0, 0, 0)
                self.norm_table.setCellWidget(row_idx, 0, checkbox_widget)

                # Column 1: Status - show the method
                status_map = {
                    'exact': 'Exact',
                    'fuzzy': 'Fuzzy',
                    'ai': 'AI',
                    'manual': 'Manual'
                }
                status_text = status_map.get(method, 'Manual')
                status_item = QTableWidgetItem(status_text)
                status_item.setTextAlignment(Qt.AlignCenter)
                
                # Color code the status
                if method == 'exact':
                    status_item.setBackground(QColor(230, 255, 230))  # Light green
                elif method == 'fuzzy':
                    status_item.setBackground(QColor(255, 250, 205))  # Light yellow
                elif method == 'ai':
                    status_item.setBackground(QColor(230, 240, 255))  # Light blue
                else:  # manual
                    status_item.setBackground(QColor(255, 240, 200))  # Light orange
                
                self.norm_table.setItem(row_idx, 1, status_item)

                # Column 2: Original MFG (read-only)
                self.norm_table.setItem(row_idx, 2, QTableWidgetItem(original))

                # Column 3: Normalize To (editable combo box)
                normalize_combo = QComboBox()
                normalize_combo.setEditable(True)
                
                # Disable mouse wheel to prevent accidental changes while scrolling
                normalize_combo.wheelEvent = lambda event: event.ignore()
                normalize_combo.setFocusPolicy(Qt.StrongFocus)

                # Build list using ONLY PAS canonical manufacturers (from match values)
                # Sort and add to dropdown
                for mfg in sorted(self.canonical_manufacturers):
                    normalize_combo.addItem(mfg)

                # Set current suggestion
                normalize_combo.setCurrentText(canonical)
                self.norm_table.setCellWidget(row_idx, 3, normalize_combo)

                # Column 4: AI Score - show score percentage
                ai_score_item = QTableWidgetItem("")
                ai_score_item.setTextAlignment(Qt.AlignCenter)
                if method in ['fuzzy', 'ai'] and score > 0:
                    ai_score_item.setText(f"{score}%")
                    ai_score_item.setToolTip(f"{method.capitalize()} match confidence: {score}%")
                self.norm_table.setItem(row_idx, 4, ai_score_item)

                # Column 5: AI Analyze button
                ai_btn = QPushButton("🤖 AI")
                ai_btn.setMaximumWidth(60)
                ai_btn.setToolTip("Run AI analysis for this manufacturer")
                ai_btn.clicked.connect(lambda checked, r=row_idx, orig=original: self.analyze_single_manufacturer_ai(r, orig))

                # Disable if no API key available
                start_page = self.wizard().page(0)
                api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None
                if not api_key or not ANTHROPIC_AVAILABLE:
                    ai_btn.setEnabled(False)
                    ai_btn.setToolTip("AI analysis not available (no API key)")

                # Center the button in the cell
                ai_btn_widget = QWidget()
                ai_btn_layout = QHBoxLayout(ai_btn_widget)
                ai_btn_layout.addWidget(ai_btn)
                ai_btn_layout.setAlignment(Qt.AlignCenter)
                ai_btn_layout.setContentsMargins(0, 0, 0, 0)
                self.norm_table.setCellWidget(row_idx, 5, ai_btn_widget)

                # Column 6: Scope dropdown
                scope_combo = QComboBox()
                scope_combo.addItems(["All Catalogs", "Per Catalog"])
                # Disable mouse wheel to prevent accidental changes while scrolling
                scope_combo.wheelEvent = lambda event: event.ignore()
                scope_combo.setFocusPolicy(Qt.StrongFocus)
                self.norm_table.setCellWidget(row_idx, 6, scope_combo)

                row_idx += 1

            # Update status and enable buttons
            # Count suggestions vs manual review
            fuzzy_count = sum(1 for v in reasoning_map.values() if v.get('method') == 'fuzzy')
            exact_count = sum(1 for v in reasoning_map.values() if v.get('method') == 'exact')
            manual_count = sum(1 for v in reasoning_map.values() if v.get('method') == 'manual')

            self.norm_status.setText(
                f"✓ Showing {len(normalizations)} manufacturers: "
                f"{exact_count} exact matches, {fuzzy_count} suggested normalizations, {manual_count} need manual review"
            )
            self.norm_status.setStyleSheet("color: green; font-weight: bold;")
            self.save_normalizations_btn.setEnabled(True)

            # Enable AI button if API key is available for additional validation
            start_page = self.wizard().page(0)
            api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None
            if api_key and ANTHROPIC_AVAILABLE:
                self.ai_normalize_btn.setEnabled(True)
        else:
            self.norm_status.setText(
                f"No variations detected (compared {len(original_mfgs)} names against {len(canonical_mfgs)} PAS manufacturers)"
            )

            # Still enable AI button if available
            start_page = self.wizard().page(0)
            api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None
            if api_key and ANTHROPIC_AVAILABLE:
                self.ai_normalize_btn.setEnabled(True)

    def create_review_section_widget(self):
        """Section 2: Review Partial Matches - Tabbed by Match Status"""
        review_group = QGroupBox("2. Review Match Results by Category")
        review_layout = QVBoxLayout()

        # Create tab widget
        self.review_tabs = QTabWidget()
        
        # Create tabs for each category
        self.found_tab = self.create_category_tab("Found", show_actions=False)
        self.multiple_tab = self.create_category_tab("Multiple", show_actions=True)
        self.need_review_tab = self.create_category_tab("Need Review", show_actions=True)
        self.none_tab = self.create_category_tab("None", show_actions="editable")  # Special editable mode for None tab
        self.errors_tab = self.create_category_tab("Errors", show_actions=False)
        
        # Add tabs to widget with emoji indicators
        self.review_tabs.addTab(self.multiple_tab, "⚠ Multiple (0)")
        self.review_tabs.addTab(self.need_review_tab, "👁 Need Review (0)")
        self.review_tabs.addTab(self.found_tab, "✓ Found (0)")
        self.review_tabs.addTab(self.none_tab, "✗ None (0)")
        self.review_tabs.addTab(self.errors_tab, "❌ Errors (0)")
        
        review_layout.addWidget(self.review_tabs)
        review_group.setLayout(review_layout)
        return review_group

    def create_category_tab(self, category, show_actions=True):
        """Create a tab for a specific match category with two-panel layout"""
        tab_widget = QWidget()
        tab_layout = QVBoxLayout(tab_widget)
        
        # Create horizontal splitter for two-panel layout
        splitter = QSplitter(Qt.Horizontal)
        
        # Left panel: Parts list
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # Parts list header
        parts_header_layout = QHBoxLayout()
        parts_header_layout.addWidget(QLabel(f"{category} Parts:"))
        
        # Add review count for interactive categories
        if show_actions:
            review_label = QLabel("(0 of 0 reviewed)")
            review_label.setStyleSheet("color: #1976d2; font-weight: bold;")
            parts_header_layout.addWidget(review_label)
            # Store reference based on category
            if category == "Multiple":
                self.multiple_review_label = review_label
            elif category == "Need Review":
                self.need_review_label = review_label
        
        # Add color legend for all interactive tabs
        if show_actions or show_actions == "editable":
            legend_label = QLabel("<span style='background-color: #C8FFFF; padding: 2px 5px;'>█</span> Re-searched from None")
            legend_label.setStyleSheet("color: #666; font-size: 9pt;")
            parts_header_layout.addWidget(legend_label)
        
        parts_header_layout.addStretch()
        left_layout.addLayout(parts_header_layout)
        
        # Parts list table
        parts_table = QTableWidget()
        if show_actions == "editable":
            # Special editable mode for None tab - editable MFG and Part Number with re-search action
            parts_table.setColumnCount(4)
            parts_table.setHorizontalHeaderLabels(["Part Number", "MFG", "Status", "Action"])
        elif show_actions:
            parts_table.setColumnCount(6)
            parts_table.setHorizontalHeaderLabels(["Part Number", "MFG", "Status", "Reviewed", "AI", "Action"])
        else:
            parts_table.setColumnCount(3)
            parts_table.setHorizontalHeaderLabels(["Part Number", "MFG", "Status"])

        # Set column resize modes
        parts_header = parts_table.horizontalHeader()
        parts_header.setSectionResizeMode(0, QHeaderView.Stretch)  # Part Number
        parts_header.setSectionResizeMode(1, QHeaderView.Stretch)  # MFG
        parts_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Status
        if show_actions == "editable":
            parts_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Action
        elif show_actions:
            parts_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Reviewed
            parts_header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # AI
            parts_header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Action

        parts_table.setSortingEnabled(True)  # Enable sorting
        parts_table.setSelectionBehavior(QTableWidget.SelectRows)
        parts_table.setSelectionMode(QTableWidget.SingleSelection)
        parts_table.itemSelectionChanged.connect(self.on_part_selected)
        left_layout.addWidget(parts_table)

        # Store table reference based on category
        if category == "Found":
            self.found_table = parts_table
        elif category == "Multiple":
            self.multiple_table = parts_table
        elif category == "Need Review":
            self.need_review_table = parts_table
        elif category == "None":
            self.none_table = parts_table
            # Enable context menu for None table to allow reverting changes
            parts_table.setContextMenuPolicy(Qt.CustomContextMenu)
            parts_table.customContextMenuRequested.connect(self.show_none_table_context_menu)
        elif category == "Errors":
            self.errors_table = parts_table
        
        # Bulk actions (only for interactive categories)
        if show_actions:
            bulk_layout = QHBoxLayout()
            auto_select_btn = QPushButton("Auto-Select Highest Similarity")
            auto_select_btn.clicked.connect(lambda: self.auto_select_highest_for_category(category))
            auto_select_btn.setEnabled(False)
            auto_select_btn.setToolTip(
                "Automatically selects the best match based on string similarity.\n\n"
                "How it works:\n"
                "• Uses difflib to compare part numbers character-by-character\n"
                "• Calculates similarity score (0-100%) for each match\n"
                "• Selects the match with highest similarity score\n"
                "• Fast and deterministic (no AI/API calls)\n"
                "• Best for exact or near-exact part number matches"
            )
            bulk_layout.addWidget(auto_select_btn)
            
            ai_suggest_btn = QPushButton("🤖 AI Suggest Best Matches")
            ai_suggest_btn.clicked.connect(lambda: self.ai_suggest_matches_for_category(category))
            ai_suggest_btn.setEnabled(False)
            ai_suggest_btn.setToolTip(
                "Uses Claude AI to intelligently suggest the best match.\n\n"
                "How it works:\n"
                "• Analyzes part number, manufacturer, and description\n"
                "• Considers manufacturer acquisitions (e.g., EPCOS → TDK)\n"
                "• Understands context and component semantics\n"
                "• Provides confidence score with reasoning\n"
                "• Skips parts with only 1 match\n"
                "• Skips already AI-processed parts\n"
                "• Best for complex matches requiring context understanding"
            )
            bulk_layout.addWidget(ai_suggest_btn)
            
            left_layout.addLayout(bulk_layout)
            
            # Store button references
            if category == "Multiple":
                self.multiple_auto_select_btn = auto_select_btn
                self.multiple_ai_suggest_btn = ai_suggest_btn
            elif category == "Need Review":
                self.need_review_auto_select_btn = auto_select_btn
                self.need_review_ai_suggest_btn = ai_suggest_btn
            
            # AI Status summary (only on Multiple tab for backward compatibility)
            if category == "Multiple":
                self.csv_summary = QLabel("")
                self.csv_summary.setStyleSheet("padding: 5px; background-color: #f0f0f0; border-radius: 3px;")
                self.csv_summary.setWordWrap(True)
                left_layout.addWidget(self.csv_summary)
        
        # Right panel: Match options (only for interactive categories)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        if show_actions == "editable":
            # For editable None tab, show instructions
            info_label = QLabel(
                "<h3>No Matches Found</h3>"
                "<p>These parts had no matches in the PAS database. You can:</p>"
                "<ul>"
                "<li><b>Edit</b> the Part Number or MFG fields directly in the table</li>"
                "<li>Click <b>🔍 Re-search</b> to search again with the modified values</li>"
                "<li>If a match is found, the part will move to the appropriate tab</li>"
                "</ul>"
                "<p style='color: #666; font-size: 10pt;'>"
                "💡 Tip: Try variations of the manufacturer name or part number format"
                "</p>"
            )
            info_label.setWordWrap(True)
            info_label.setStyleSheet("padding: 20px; background-color: #f9f9f9; border-radius: 5px;")
            right_layout.addWidget(info_label)
            right_layout.addStretch()
        elif show_actions:
            right_layout.addWidget(QLabel("Available Matches:"))

            matches_table = QTableWidget()
            matches_table.setColumnCount(7)  # Increased from 5 to 7
            matches_table.setHorizontalHeaderLabels(["Select", "Part Number", "Manufacturer", "Lifecycle Status", "External ID", "Similarity", "AI Score"])
            matches_table.setSortingEnabled(True)  # Enable sorting
            matches_table.setContextMenuPolicy(Qt.CustomContextMenu)
            matches_table.customContextMenuRequested.connect(self.show_match_context_menu)

            # Set column resize modes
            header = matches_table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Select
            header.setSectionResizeMode(1, QHeaderView.Stretch)  # Part Number
            header.setSectionResizeMode(2, QHeaderView.Stretch)  # Manufacturer
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Lifecycle Status
            header.setSectionResizeMode(4, QHeaderView.Stretch)  # External ID
            header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Similarity
            header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # AI Score
            
            right_layout.addWidget(matches_table)
            
            # Store matches table reference
            if category == "Multiple":
                self.multiple_matches_table = matches_table
            elif category == "Need Review":
                self.need_review_matches_table = matches_table
            
            none_correct_checkbox = QCheckBox("None of these are correct (keep original)")
            right_layout.addWidget(none_correct_checkbox)
            
            # Store checkbox reference
            if category == "Multiple":
                self.multiple_none_correct_checkbox = none_correct_checkbox
            elif category == "Need Review":
                self.need_review_none_correct_checkbox = none_correct_checkbox
            
            # Save button for selections
            save_layout = QHBoxLayout()
            save_selection_btn = QPushButton("💾 Save Selection")
            save_selection_btn.clicked.connect(self.save_current_selection)
            save_selection_btn.setEnabled(False)
            save_selection_btn.setToolTip("Save your current match selection for this part")
            save_layout.addWidget(save_selection_btn)
            save_layout.addStretch()
            right_layout.addLayout(save_layout)
            
            # Store save button reference
            if category == "Multiple":
                self.multiple_save_btn = save_selection_btn
            elif category == "Need Review":
                self.need_review_save_btn = save_selection_btn
        else:
            # For non-interactive tabs, just show a message
            info_label = QLabel(f"This tab shows parts with '{category}' status.\nNo action required.")
            info_label.setWordWrap(True)
            info_label.setStyleSheet("padding: 20px; color: #666;")
            right_layout.addWidget(info_label)
            right_layout.addStretch()
        
        # Add panels to splitter
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([400, 600])
        
        tab_layout.addWidget(splitter)
        return tab_widget

    # Compatibility properties - these point to the Multiple tab widgets by default
    # since that's where most of the interactive functionality is
    @property
    def parts_list(self):
        """For backward compatibility - points to multiple_table"""
        return self.multiple_table
    
    @property
    def matches_table(self):
        """For backward compatibility - points to multiple_matches_table"""
        return self.multiple_matches_table
    
    @property
    def auto_select_btn(self):
        """For backward compatibility - points to multiple_auto_select_btn"""
        return self.multiple_auto_select_btn
    
    @property
    def ai_suggest_btn(self):
        """For backward compatibility - points to multiple_ai_suggest_btn"""
        return self.multiple_ai_suggest_btn
    
    @property
    def save_selection_btn(self):
        """For backward compatibility - points to multiple_save_btn"""
        return self.multiple_save_btn
    
    @property
    def none_correct_checkbox(self):
        """For backward compatibility - points to multiple_none_correct_checkbox"""
        return self.multiple_none_correct_checkbox
    
    @property
    def review_count_label(self):
        """For backward compatibility - points to multiple_review_label"""
        return self.multiple_review_label

    def populate_category_table(self, table, parts_list, show_actions=True):
        """Populate a category table with parts"""
        from PyQt5.QtGui import QColor as _QColor  # Explicit local import to avoid scoping issues
        print(f"DEBUG populate_category_table: {len(parts_list)} parts, show_actions={show_actions}")
        table.setRowCount(len(parts_list))

        for row_idx, part in enumerate(parts_list):
            # Ensure part is a dict and has required keys
            if not isinstance(part, dict):
                print(f"ERROR: Part at row {row_idx} is not a dict: {type(part)} - {part}")
                continue

            # Ensure matches key exists
            if 'matches' not in part:
                part['matches'] = []

            if row_idx < 5:  # Log first 5
                print(f"DEBUG: Adding row {row_idx}: {part.get('PartNumber', 'N/A')} | {part.get('ManufacturerName', 'N/A')} | {part.get('MatchStatus', 'N/A')}")

            # Create items for Part Number and MFG
            pn_item = QTableWidgetItem(part.get('PartNumber', 'N/A'))
            mfg_item = QTableWidgetItem(part.get('ManufacturerName', 'N/A'))
            status_item = QTableWidgetItem(part.get('MatchStatus', 'N/A'))

            # Make editable for "editable" mode (None tab)
            if show_actions == "editable":
                # Store original values if not already stored (for revert functionality)
                if 'original_pn' not in part:
                    part['original_pn'] = part.get('PartNumber', 'N/A')
                if 'original_mfg' not in part:
                    part['original_mfg'] = part.get('ManufacturerName', 'N/A')

                # Set tooltips showing original values
                pn_item.setToolTip(f"Original: {part['original_pn']}\nRight-click to revert")
                mfg_item.setToolTip(f"Original: {part['original_mfg']}\nRight-click to revert")

                pn_item.setFlags(pn_item.flags() | Qt.ItemIsEditable)
                mfg_item.setFlags(mfg_item.flags() | Qt.ItemIsEditable)
                status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)  # Status not editable
            else:
                # Make all non-editable for other tabs
                pn_item.setFlags(pn_item.flags() & ~Qt.ItemIsEditable)
                mfg_item.setFlags(mfg_item.flags() & ~Qt.ItemIsEditable)
                status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)

            table.setItem(row_idx, 0, pn_item)
            table.setItem(row_idx, 1, mfg_item)
            table.setItem(row_idx, 2, status_item)
            
            # Color-code re-searched parts that moved from None to other categories
            if part.get('re_searched') and part.get('original_status') == 'None':
                # Light cyan background to indicate this part was re-searched from None
                highlight_color = _QColor(200, 255, 255)  # Light cyan
                pn_item.setBackground(highlight_color)
                mfg_item.setBackground(highlight_color)
                status_item.setBackground(highlight_color)

            if show_actions == "editable":
                # Add Re-search button for None tab
                research_btn = QPushButton("🔍 Re-search")
                research_btn.setToolTip("Re-search with modified values")
                research_btn.clicked.connect(lambda checked, idx=row_idx: self.research_single_part(idx))
                table.setCellWidget(row_idx, 3, research_btn)
            elif show_actions:
                # Reviewed indicator
                reviewed_item = QTableWidgetItem("✓" if part.get('selected_match') else "")
                reviewed_item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_idx, 3, reviewed_item)

                # AI indicator
                ai_status = ""
                if part.get('ai_processed'):
                    ai_status = "🤖"
                elif part.get('ai_processing'):
                    ai_status = "⏳"
                ai_item = QTableWidgetItem(ai_status)
                ai_item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_idx, 4, ai_item)

                # Action button - AI Suggest (only if >1 match and not already processed)
                matches = part.get('matches', [])
                if len(matches) > 1 and not part.get('ai_processed'):
                    ai_btn = QPushButton("🤖 AI")
                    ai_btn.setToolTip("Use AI to suggest best match for this part")
                    ai_btn.clicked.connect(lambda checked, idx=row_idx: self.ai_suggest_single(idx))
                    table.setCellWidget(row_idx, 5, ai_btn)

        print(f"DEBUG: Table populated with {table.rowCount()} rows")

    def research_single_part(self, row_idx):
        """Re-search a single part from the None tab with modified values"""
        if row_idx >= len(self.none_parts):
            return

        # Get the updated values from the table cells
        pn_item = self.none_table.item(row_idx, 0)
        mfg_item = self.none_table.item(row_idx, 1)

        if not pn_item or not mfg_item:
            return

        new_pn = pn_item.text().strip()
        new_mfg = mfg_item.text().strip()

        # Only require part number (MFG can be empty)
        if not new_pn:
            QMessageBox.warning(self, "Missing Data", "Part Number is required for search.")
            return

        # Get the part data
        part = self.none_parts[row_idx]

        # Store original values if not already stored (for tracking in search_results)
        if 'original_pn' not in part:
            part['original_pn'] = part.get('PartNumber', '')
        if 'original_mfg' not in part:
            part['original_mfg'] = part.get('ManufacturerName', '')

        # Update the part data with new values
        part['PartNumber'] = new_pn
        part['ManufacturerName'] = new_mfg

        # Disable the button while searching
        btn = self.none_table.cellWidget(row_idx, 3)
        if btn:
            btn.setEnabled(False)
            btn.setText("⏳ Searching...")

        # Get PAS search page to access the PAS client
        pas_page = self.wizard().page(3)  # PASSearchPage is page 3
        if not pas_page or not hasattr(pas_page, 'pas_client'):
            QMessageBox.warning(self, "Error", "PAS API client not available.")
            if btn:
                btn.setEnabled(True)
                btn.setText("🔍 Re-search")
            return

        try:
            # Perform PAS search using the PAS client's search_part method
            match_result, match_type = pas_page.pas_client.search_part(new_pn, new_mfg)

            # Map match_type to status
            if match_type in ['Found', 'Multiple', 'Need user review', 'None', 'Error']:
                status = match_type
            else:
                status = 'None'

            # Update the part data
            part['MatchStatus'] = status
            part['matches'] = match_result.get('matches', [])
            
            # Mark as re-searched so we can color-code it
            part['re_searched'] = True
            part['original_status'] = 'None'

            # Update the table to show new status
            status_item = self.none_table.item(row_idx, 2)
            if status_item:
                status_item.setText(status)

            # Move part to appropriate category if match found
            if status != 'None' and status != 'Error':
                # Remove from none_parts
                self.none_parts.pop(row_idx)

                # Add to appropriate category
                if status == 'Found':
                    self.found_parts.append(part)
                    # Auto-select the match for Found parts
                    if part['matches']:
                        part['selected_match'] = part['matches'][0]
                elif status == 'Multiple':
                    self.multiple_parts.append(part)
                elif status == 'Need user review':
                    self.need_review_parts.append(part)

                # Update search_results to reflect the change
                for result in self.search_results:
                    if (result.get('PartNumber') == part.get('original_pn', part['PartNumber']) and
                        result.get('ManufacturerName') == part.get('original_mfg', part['ManufacturerName'])):
                        result['MatchStatus'] = status
                        result['matches'] = part['matches']
                        break

                # Re-populate all tabs to reflect changes
                self.populate_category_table(self.found_table, self.found_parts, show_actions=False)
                self.populate_category_table(self.multiple_table, self.multiple_parts, show_actions=True)
                self.populate_category_table(self.need_review_table, self.need_review_parts, show_actions=True)
                self.populate_category_table(self.none_table, self.none_parts, show_actions="editable")

                # Update tab counts
                self.review_tabs.setTabText(0, f"⚠ Multiple ({len(self.multiple_parts)})")
                self.review_tabs.setTabText(1, f"👁 Need Review ({len(self.need_review_parts)})")
                self.review_tabs.setTabText(2, f"✓ Found ({len(self.found_parts)})")
                self.review_tabs.setTabText(3, f"✗ None ({len(self.none_parts)})")

                # Show result message
                if status == 'Found':
                    QMessageBox.information(self, "Match Found!",
                        f"Found exact match for {new_pn}!\n\nThe part has been moved to the 'Found' tab.")
                elif status == 'Multiple':
                    QMessageBox.information(self, "Multiple Matches Found",
                        f"Found {len(part['matches'])} matches for {new_pn}.\n\n"
                        f"The part has been moved to the 'Multiple' tab where you can select the correct match.")
                elif status == 'Need user review':
                    QMessageBox.information(self, "Match Needs Review",
                        f"Found match(es) for {new_pn} that need review.\n\n"
                        f"The part has been moved to the 'Need Review' tab.")
            else:
                QMessageBox.information(self, "No Match Found",
                    f"Still no matches found for {new_pn} with manufacturer {new_mfg}.\n\n"
                    f"Try editing the values and searching again.")

        except Exception as e:
            QMessageBox.critical(self, "Search Error", f"Error during search: {str(e)}")
            part['MatchStatus'] = 'Error'
            status_item = self.none_table.item(row_idx, 2)
            if status_item:
                status_item.setText('Error')

        finally:
            # Re-enable the button
            if btn:
                btn.setEnabled(True)
                btn.setText("🔍 Re-search")

    def show_none_table_context_menu(self, position):
        """Show context menu for None table cells to allow reverting changes"""
        # Get the item at the clicked position
        item = self.none_table.itemAt(position)
        if not item:
            return

        row = item.row()
        col = item.column()

        # Only show menu for MFG PN (col 0) and MFG (col 1) columns
        if col not in [0, 1]:
            return

        # Get the part data
        if row >= len(self.none_parts):
            return

        part = self.none_parts[row]

        # Check if original values exist
        has_original_pn = 'original_pn' in part
        has_original_mfg = 'original_mfg' in part

        if not has_original_pn and not has_original_mfg:
            return

        # Create context menu
        menu = QMenu(self)

        # Add revert actions based on which column was clicked
        if col == 0 and has_original_pn:  # MFG PN column
            current_pn = part.get('PartNumber', '')
            original_pn = part['original_pn']

            # Only show revert option if value has changed
            if current_pn != original_pn:
                revert_pn_action = menu.addAction(f"⟲ Revert MFG PN to Original")
                revert_pn_action.setToolTip(f"Change '{current_pn}' back to '{original_pn}'")
            else:
                # Show info that it's already at original value
                info_action = menu.addAction("✓ Already at original value")
                info_action.setEnabled(False)

        elif col == 1 and has_original_mfg:  # MFG column
            current_mfg = part.get('ManufacturerName', '')
            original_mfg = part['original_mfg']

            # Only show revert option if value has changed
            if current_mfg != original_mfg:
                revert_mfg_action = menu.addAction(f"⟲ Revert MFG to Original")
                revert_mfg_action.setToolTip(f"Change '{current_mfg}' back to '{original_mfg}'")
            else:
                # Show info that it's already at original value
                info_action = menu.addAction("✓ Already at original value")
                info_action.setEnabled(False)

        # Add separator and preview option
        if menu.actions():
            menu.addSeparator()
            preview_action = menu.addAction("👁 Show Original Values")

        # Execute menu and handle selection
        selected_action = menu.exec_(self.none_table.viewport().mapToGlobal(position))

        if not selected_action:
            return

        # Handle revert actions
        if selected_action.text().startswith("⟲ Revert MFG PN"):
            self.revert_none_field(row, 'pn')
        elif selected_action.text().startswith("⟲ Revert MFG"):
            self.revert_none_field(row, 'mfg')
        elif selected_action.text().startswith("👁 Show Original"):
            self.show_original_values_preview(row)

    def revert_none_field(self, row_idx, field_type):
        """Revert a specific field (MFG PN or MFG) back to its original value"""
        if row_idx >= len(self.none_parts):
            return

        part = self.none_parts[row_idx]

        if field_type == 'pn':
            # Revert Part Number
            if 'original_pn' not in part:
                return

            original_value = part['original_pn']
            part['PartNumber'] = original_value

            # Update table cell
            pn_item = self.none_table.item(row_idx, 0)
            if pn_item:
                pn_item.setText(original_value)

            QMessageBox.information(
                self,
                "Reverted",
                f"MFG PN has been reverted to original value:\n'{original_value}'"
            )

        elif field_type == 'mfg':
            # Revert Manufacturer
            if 'original_mfg' not in part:
                return

            original_value = part['original_mfg']
            part['ManufacturerName'] = original_value

            # Update table cell
            mfg_item = self.none_table.item(row_idx, 1)
            if mfg_item:
                mfg_item.setText(original_value)

            QMessageBox.information(
                self,
                "Reverted",
                f"MFG has been reverted to original value:\n'{original_value}'"
            )

    def show_original_values_preview(self, row_idx):
        """Show a preview dialog with original vs current values"""
        if row_idx >= len(self.none_parts):
            return

        part = self.none_parts[row_idx]

        original_pn = part.get('original_pn', 'N/A')
        original_mfg = part.get('original_mfg', 'N/A')
        current_pn = part.get('PartNumber', 'N/A')
        current_mfg = part.get('ManufacturerName', 'N/A')

        # Check if values have changed
        pn_changed = original_pn != current_pn
        mfg_changed = original_mfg != current_mfg

        # Build message
        message = "<b>Original vs Current Values:</b><br><br>"
        message += "<table border='1' cellpadding='5' cellspacing='0'>"
        message += "<tr><th>Field</th><th>Original</th><th>Current</th><th>Status</th></tr>"

        # MFG PN row
        pn_status = "<span style='color: orange;'>Changed</span>" if pn_changed else "<span style='color: green;'>Unchanged</span>"
        message += f"<tr><td><b>MFG PN</b></td><td>{original_pn}</td><td>{current_pn}</td><td>{pn_status}</td></tr>"

        # MFG row
        mfg_status = "<span style='color: orange;'>Changed</span>" if mfg_changed else "<span style='color: green;'>Unchanged</span>"
        message += f"<tr><td><b>MFG</b></td><td>{original_mfg}</td><td>{current_mfg}</td><td>{mfg_status}</td></tr>"

        message += "</table>"

        QMessageBox.information(self, "Original Values Preview", message)

    def auto_select_highest_for_category(self, category):
        """Auto-select highest similarity matches for a specific category"""
        if category == "Multiple":
            parts_list = self.multiple_parts
        elif category == "Need Review":
            parts_list = self.need_review_parts
        else:
            return
        
        from difflib import SequenceMatcher
        selected_count = 0
        
        for part in parts_list:
            if not part['matches']:
                continue

            original_pn = part['PartNumber'].upper().strip()
            original_mfg = part['ManufacturerName'].upper().strip()
            best_match = None
            best_similarity = 0.0

            # Calculate similarity for each match
            for match in part['matches']:
                # Use helper method to extract match info (handles both dict and string formats)
                match_pn, match_mfg, _, _, _, _, _ = self._get_match_info(match)

                # Ensure match_pn is a string before calling string methods
                match_pn = str(match_pn).upper().strip() if match_pn else ""
                match_mfg = str(match_mfg).upper().strip() if match_mfg else ""

                # Calculate combined similarity (60% part number, 40% manufacturer)
                pn_similarity = SequenceMatcher(None, original_pn, match_pn).ratio()
                mfg_similarity = SequenceMatcher(None, original_mfg, match_mfg).ratio() if match_mfg else 0

                # Weighted average
                combined_similarity = (pn_similarity * 0.6) + (mfg_similarity * 0.4)

                if combined_similarity > best_similarity:
                    best_similarity = combined_similarity
                    best_match = match

            if best_match:
                part['selected_match'] = best_match
                part['auto_selected'] = True
                part['similarity_score'] = best_similarity
                selected_count += 1

        QMessageBox.information(self, "Auto-Select Complete",
                              f"Selected best match for {selected_count} parts in {category} category using similarity analysis.")

        # Refresh the appropriate table
        if category == "Multiple":
            self.populate_category_table(self.multiple_table, self.multiple_parts, show_actions=True)
        elif category == "Need Review":
            self.populate_category_table(self.need_review_table, self.need_review_parts, show_actions=True)

    def ai_suggest_matches_for_category(self, category):
        """Use AI to suggest best matches for all unprocessed parts in a specific category"""
        if category == "Multiple":
            parts_list = self.multiple_parts
        elif category == "Need Review":
            parts_list = self.need_review_parts
        else:
            return
        
        # Get API key
        start_page = self.wizard().page(0)
        self.api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None

        if not self.api_key or not ANTHROPIC_AVAILABLE:
            QMessageBox.warning(self, "AI Not Available",
                              "Claude AI is not available. Please provide an API key.")
            return

        # Get combined data from previous step
        xml_gen_page = self.wizard().page(3)
        if hasattr(xml_gen_page, 'combined_data'):
            self.combined_data = xml_gen_page.combined_data

        # Filter out already processed parts
        unprocessed_parts = [part for part in parts_list if not part.get('ai_processed')]

        if not unprocessed_parts:
            QMessageBox.information(self, "All Processed",
                                  f"All parts in {category} have already been processed by AI.")
            return

        # Disable buttons
        if category == "Multiple":
            self.multiple_auto_select_btn.setEnabled(False)
            self.multiple_ai_suggest_btn.setEnabled(False)
        elif category == "Need Review":
            self.need_review_auto_select_btn.setEnabled(False)
            self.need_review_ai_suggest_btn.setEnabled(False)

        # Mark unprocessed parts as processing
        for part in unprocessed_parts:
            part['ai_processing'] = True

        # Refresh the appropriate table
        if category == "Multiple":
            self.populate_category_table(self.multiple_table, self.multiple_parts, show_actions=True)
        elif category == "Need Review":
            self.populate_category_table(self.need_review_table, self.need_review_parts, show_actions=True)

        # Start AI thread with only unprocessed parts
        self.ai_match_thread = PartialMatchAIThread(
            self.api_key,
            unprocessed_parts,
            self.combined_data
        )
        self.ai_match_thread.progress.connect(self.on_ai_match_progress)
        self.ai_match_thread.part_analyzed.connect(self.on_part_analyzed)
        self.ai_match_thread.finished.connect(lambda suggestions: self.on_ai_match_finished_for_category(suggestions, category))
        self.ai_match_thread.error.connect(self.on_ai_match_error)
        self.ai_match_thread.start()

    def on_ai_match_finished_for_category(self, suggestions, category):
        """Handle AI completion for category-specific processing"""
        self.csv_summary.setText(f"✓ AI analysis complete for {category}")
        self.csv_summary.setStyleSheet("padding: 5px; background-color: #d4edda; border-radius: 3px;")

        # Re-enable buttons
        if category == "Multiple":
            self.multiple_auto_select_btn.setEnabled(True)
            start_page = self.wizard().page(0)
            api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None
            self.multiple_ai_suggest_btn.setEnabled(bool(api_key and ANTHROPIC_AVAILABLE))
            # Refresh table
            self.populate_category_table(self.multiple_table, self.multiple_parts, show_actions=True)
        elif category == "Need Review":
            self.need_review_auto_select_btn.setEnabled(True)
            start_page = self.wizard().page(0)
            api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None
            self.need_review_ai_suggest_btn.setEnabled(bool(api_key and ANTHROPIC_AVAILABLE))
            # Refresh table
            self.populate_category_table(self.need_review_table, self.need_review_parts, show_actions=True)

        # Show summary
        processed_count = sum(1 for part in (self.multiple_parts + self.need_review_parts) if part.get('ai_processed'))
        QMessageBox.information(self, "AI Analysis Complete",
                              f"AI has analyzed {len(suggestions)} parts in {category}.\n"
                              f"Total AI-processed parts: {processed_count}")

    def create_review_section_widget_OLD_REMOVED(self):
        """Section 2: Review Partial Matches"""
        review_group = QGroupBox("2. Review Partial Matches")
        review_layout = QHBoxLayout()

        # Left panel: Parts list
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        # Parts list header with review count
        parts_header_layout = QHBoxLayout()
        parts_header_layout.addWidget(QLabel("Parts Needing Review:"))
        self.review_count_label = QLabel("(0 of 0 reviewed)")
        self.review_count_label.setStyleSheet("color: #1976d2; font-weight: bold;")
        parts_header_layout.addWidget(self.review_count_label)
        parts_header_layout.addStretch()
        left_layout.addLayout(parts_header_layout)

        self.parts_list = QTableWidget()
        self.parts_list.setColumnCount(6)
        self.parts_list.setHorizontalHeaderLabels(["Part Number", "MFG", "Status", "Reviewed", "AI", "Action"])

        # Set column resize modes
        parts_header = self.parts_list.horizontalHeader()
        parts_header.setSectionResizeMode(0, QHeaderView.Stretch)  # Part Number
        parts_header.setSectionResizeMode(1, QHeaderView.Stretch)  # MFG
        parts_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Status
        parts_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Reviewed
        parts_header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # AI
        parts_header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Action

        self.parts_list.setSortingEnabled(True)  # Enable sorting
        self.parts_list.setSelectionBehavior(QTableWidget.SelectRows)
        self.parts_list.setSelectionMode(QTableWidget.SingleSelection)
        self.parts_list.itemSelectionChanged.connect(self.on_part_selected)
        left_layout.addWidget(self.parts_list)

        # Bulk actions
        bulk_layout = QHBoxLayout()
        self.auto_select_btn = QPushButton("Auto-Select Highest Similarity")
        self.auto_select_btn.clicked.connect(self.auto_select_highest)
        self.auto_select_btn.setEnabled(False)
        self.auto_select_btn.setToolTip(
            "Automatically selects the best match based on string similarity.\n\n"
            "How it works:\n"
            "• Uses difflib to compare both MFG and MFG PN\n"
            "• Weighted scoring: 60% part number + 40% manufacturer\n"
            "• Calculates combined similarity (0-100%) for each match\n"
            "• Selects the match with highest combined score\n"
            "• Fast and deterministic (no AI/API calls)\n"
            "• Best for exact or near-exact matches"
        )
        bulk_layout.addWidget(self.auto_select_btn)

        self.ai_suggest_btn = QPushButton("🤖 AI Suggest Best Matches")
        self.ai_suggest_btn.clicked.connect(self.ai_suggest_matches)
        self.ai_suggest_btn.setEnabled(False)
        self.ai_suggest_btn.setToolTip(
            "Uses Claude AI to intelligently suggest the best match.\n\n"
            "How it works:\n"
            "• Analyzes part number, manufacturer, and description\n"
            "• Considers manufacturer acquisitions (e.g., EPCOS → TDK)\n"
            "• Understands context and component semantics\n"
            "• Provides confidence score with reasoning\n"
            "• Skips parts with only 1 match\n"
            "• Skips already AI-processed parts\n"
            "• Best for complex matches requiring context understanding"
        )
        bulk_layout.addWidget(self.ai_suggest_btn)

        left_layout.addLayout(bulk_layout)

        # Right panel: Match options
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        right_layout.addWidget(QLabel("Available Matches:"))

        self.matches_table = QTableWidget()
        self.matches_table.setColumnCount(5)
        self.matches_table.setHorizontalHeaderLabels(["Select", "Part Number", "Manufacturer", "Similarity", "AI Score"])
        self.matches_table.setSortingEnabled(True)  # Enable sorting
        self.matches_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.matches_table.customContextMenuRequested.connect(self.show_match_context_menu)

        # Set column resize modes
        header = self.matches_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Select column - fit to content
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Part Number - stretch
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Manufacturer - stretch
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Similarity - fit to content
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # AI Score - fit to content

        right_layout.addWidget(self.matches_table)

        self.none_correct_checkbox = QCheckBox("None of these are correct (keep original)")
        right_layout.addWidget(self.none_correct_checkbox)

        # Save button for selections
        save_layout = QHBoxLayout()
        self.save_selection_btn = QPushButton("💾 Save Selection")
        self.save_selection_btn.clicked.connect(self.save_current_selection)
        self.save_selection_btn.setEnabled(False)
        self.save_selection_btn.setToolTip("Save your current match selection for this part")
        save_layout.addWidget(self.save_selection_btn)
        save_layout.addStretch()
        right_layout.addLayout(save_layout)

        # Splitter
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([400, 600])

        review_layout.addWidget(splitter)
        review_group.setLayout(review_layout)
        return review_group

    def create_normalization_section_widget(self):
        """Section 3: Manufacturer Normalization"""
        norm_group = QGroupBox("3. Manufacturer Normalization")
        norm_layout = QVBoxLayout()

        # AI button
        ai_layout = QHBoxLayout()
        self.ai_normalize_btn = QPushButton("🤖 AI Detect Normalizations")
        self.ai_normalize_btn.clicked.connect(self.ai_detect_normalizations)
        self.ai_normalize_btn.setEnabled(False)
        self.ai_normalize_btn.setToolTip(
            "Use Claude AI to analyze ALL manufacturers and detect variations.\n"
            "AI will identify abbreviations, acquisitions, and alternate spellings.\n"
            "This is pure AI analysis - not fuzzy matching."
        )
        ai_layout.addWidget(self.ai_normalize_btn)

        self.norm_status = QLabel("")
        ai_layout.addWidget(self.norm_status)
        ai_layout.addStretch()
        norm_layout.addLayout(ai_layout)

        # Normalization table
        self.norm_table = QTableWidget()
        self.norm_table.setColumnCount(7)  # Increased from 6 to 7
        self.norm_table.setHorizontalHeaderLabels(["Include", "Status", "Original MFG", "Normalize To", "AI Score", "AI Analyze", "Scope"])
        self.norm_table.setSortingEnabled(True)  # Enable sorting
        self.norm_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.norm_table.customContextMenuRequested.connect(self.show_normalization_context_menu)

        # Set column resize modes
        header = self.norm_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Include - fit to checkbox
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Status
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Original MFG
        header.setSectionResizeMode(3, QHeaderView.Stretch)  # Normalize To
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # AI Score
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # AI Analyze button
        header.setSectionResizeMode(6, QHeaderView.Stretch)  # Scope
        
        norm_layout.addWidget(self.norm_table)

        # Help text with color legend
        help_label = QLabel(
            "<b>Instructions:</b> Review each manufacturer. "
            "Uncheck 'Include' to skip normalization. "
            "Edit 'Normalize To' dropdown to change suggestion. "
            "Use 'AI Analyze' button for individual AI suggestions. "
            "Right-click rows to see detection reasoning.<br><br>"
            "<b>Color Legend:</b> "
            "<span style='background-color: #E6FFE6; padding: 2px 5px;'>Green</span> = Exact match | "
            "<span style='background-color: #FFFACD; padding: 2px 5px;'>Yellow</span> = Fuzzy match | "
            "<span style='background-color: #E6F0FF; padding: 2px 5px;'>Blue</span> = AI suggestion | "
            "<span style='background-color: #FFF0C8; padding: 2px 5px;'>Orange</span> = Manual review | "
            "<span style='background-color: #F8F8F8; padding: 2px 5px;'>Gray</span> = No change needed"
        )
        help_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border-radius: 3px; font-size: 9pt;")
        help_label.setWordWrap(True)
        norm_layout.addWidget(help_label)

        # Save button
        save_norm_layout = QHBoxLayout()
        self.save_normalizations_btn = QPushButton("💾 Save Normalizations")
        self.save_normalizations_btn.clicked.connect(self.save_normalizations)
        self.save_normalizations_btn.setEnabled(False)
        self.save_normalizations_btn.setToolTip("Save manufacturer normalization settings")
        save_norm_layout.addWidget(self.save_normalizations_btn)
        save_norm_layout.addStretch()
        norm_layout.addLayout(save_norm_layout)

        norm_group.setLayout(norm_layout)
        return norm_group

    def create_comparison_section_widget(self):
        """Section 4: Comparison View"""
        comp_group = QGroupBox("4. Review Changes")
        comp_layout = QVBoxLayout()

        # Summary
        self.comparison_summary = QLabel("Apply changes to see comparison")
        self.comparison_summary.setStyleSheet("padding: 5px; background-color: #e3f2fd; border-radius: 3px; font-weight: bold;")
        comp_layout.addWidget(self.comparison_summary)

        # Side-by-side tables
        tables_layout = QHBoxLayout()

        # Old data
        old_widget = QWidget()
        old_layout = QVBoxLayout(old_widget)
        old_layout.addWidget(QLabel("Original Data:"))
        self.old_data_table = QTableWidget()
        self.old_data_table.setColumnCount(3)
        self.old_data_table.setHorizontalHeaderLabels(["MFG", "MFG PN", "Description"])
        self.old_data_table.setSortingEnabled(True)  # Enable sorting
        self.old_data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        old_layout.addWidget(self.old_data_table)

        # New data
        new_widget = QWidget()
        new_layout = QVBoxLayout(new_widget)
        new_layout.addWidget(QLabel("Updated Data:"))
        self.new_data_table = QTableWidget()
        self.new_data_table.setColumnCount(3)
        self.new_data_table.setHorizontalHeaderLabels(["MFG", "MFG PN", "Description"])
        self.new_data_table.setSortingEnabled(True)  # Enable sorting
        self.new_data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        new_layout.addWidget(self.new_data_table)

        # Sync scroll
        self.old_data_table.verticalScrollBar().valueChanged.connect(
            self.new_data_table.verticalScrollBar().setValue
        )
        self.new_data_table.verticalScrollBar().valueChanged.connect(
            self.old_data_table.verticalScrollBar().setValue
        )

        tables_layout.addWidget(old_widget)
        tables_layout.addWidget(new_widget)
        comp_layout.addLayout(tables_layout)

        comp_group.setLayout(comp_layout)
        return comp_group

    def create_actions_section_widget(self):
        """Section 5: Final Actions"""
        actions_widget = QWidget()
        actions_layout = QHBoxLayout(actions_widget)

        self.apply_changes_btn = QPushButton("Apply Changes & Generate Comparison")
        self.apply_changes_btn.clicked.connect(self.apply_changes)
        self.apply_changes_btn.setEnabled(False)
        actions_layout.addWidget(self.apply_changes_btn)

        actions_layout.addStretch()
        return actions_widget

    def browse_csv(self):
        """Browse for CSV file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select SearchAndAssign CSV",
            "", "CSV Files (*.csv);;All Files (*.*)"
        )
        if file_path:
            self.csv_path_input.setText(file_path)
            self.load_csv_btn.setEnabled(True)

    def load_csv(self):
        """Load and parse SearchAndAssign CSV"""
        csv_path = self.csv_path_input.text()
        if not csv_path or not Path(csv_path).exists():
            QMessageBox.warning(self, "File Not Found", "Please select a valid CSV file.")
            return

        # Create log file
        log_path = Path(csv_path).parent / "step4_load_debug.log"

        try:
            # Parse CSV with varying column counts - read as raw text lines
            import csv
            import datetime

            with open(log_path, 'w', encoding='utf-8') as log:
                log.write(f"=== Step 4 CSV Load Debug Log ===\n")
                log.write(f"Timestamp: {datetime.datetime.now()}\n")
                log.write(f"CSV Path: {csv_path}\n")
                log.write(f"CSV Exists: {Path(csv_path).exists()}\n")
                log.write(f"CSV Size: {Path(csv_path).stat().st_size} bytes\n\n")

                self.search_assign_data = []
                self.parts_needing_review = []

                exact_matches = 0
                partial_matches = 0
                needs_review = 0
                no_match = 0

                log.write("Starting CSV parsing...\n")

                with open(csv_path, 'r', encoding='utf-8') as f:
                    csv_reader = csv.reader(f)
                    header = next(csv_reader)  # Skip header
                    log.write(f"Header: {header}\n\n")

                    row_count = 0
                    for row in csv_reader:
                        row_count += 1

                        if len(row) < 3:  # Need at least PartNumber, ManufacturerName, MatchStatus
                            log.write(f"Row {row_count}: SKIPPED (too few columns: {len(row)})\n")
                            continue

                        part_num = row[0]
                        mfg = row[1]
                        status = row[2]

                        # Collect all match values (column 3 onwards)
                        matches = []
                        for col_idx in range(3, len(row)):
                            if row[col_idx].strip():  # Non-empty
                                matches.append(row[col_idx])

                        part_data = {
                            'PartNumber': part_num,
                            'ManufacturerName': mfg,
                            'MatchStatus': status,
                            'matches': matches,
                            'selected_match': None
                        }

                        self.search_assign_data.append(part_data)

                        # Log first 10 rows
                        if row_count <= 10:
                            log.write(f"Row {row_count}: PN={part_num}, MFG={mfg}, Status={status}, Matches={len(matches)}\n")

                        # Categorize
                        if status == "Found":
                            exact_matches += 1
                            if matches:
                                part_data['selected_match'] = matches[0]  # Auto-select Found match
                        elif status == "Multiple" or status == "Need user review":
                            if status == "Multiple":
                                partial_matches += 1
                            else:
                                needs_review += 1
                            self.parts_needing_review.append(part_data)
                        else:  # None
                            no_match += 1

                    log.write(f"\n=== Parsing Complete ===\n")
                    log.write(f"Total rows processed: {row_count}\n")
                    log.write(f"Total parts loaded: {len(self.search_assign_data)}\n")
                    log.write(f"Exact matches: {exact_matches}\n")
                    log.write(f"Partial matches: {partial_matches}\n")
                    log.write(f"Needs review: {needs_review}\n")
                    log.write(f"No match: {no_match}\n")
                    log.write(f"Parts needing review list: {len(self.parts_needing_review)}\n")

            # Update summary
            total = len(self.search_assign_data)
            print(f"DEBUG: Loaded {total} parts from {csv_path}")
            print(f"DEBUG: exact={exact_matches}, partial={partial_matches}, needs_review={needs_review}, no_match={no_match}")
            print(f"DEBUG: Log file written to: {log_path}")

            self.csv_summary.setText(
                f"✓ Loaded {total} parts: {exact_matches} exact, "
                f"{partial_matches} partial, {needs_review} need review, {no_match} no match"
            )
            self.csv_summary.setStyleSheet("padding: 5px; background-color: #c8e6c9; border-radius: 3px; font-weight: bold;")

            # Populate parts list
            self.populate_parts_list()

            # Populate manufacturer list for normalization
            self.populate_manufacturer_list()

            # Update review count
            self.update_review_count()

            # Enable buttons
            self.csv_loaded = True
            self.auto_select_btn.setEnabled(len(self.parts_needing_review) > 0)
            self.ai_suggest_btn.setEnabled(len(self.parts_needing_review) > 0)
            self.ai_normalize_btn.setEnabled(True)
            self.apply_changes_btn.setEnabled(True)

            QMessageBox.information(self, "CSV Loaded", f"Successfully loaded {total} parts from SearchAndAssign CSV.")

        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"Failed to load CSV:\n{str(e)}")

    def populate_manufacturer_list(self):
        """Populate manufacturer list from loaded data (without AI)"""
        # Collect all manufacturers from user data and SupplyFrame
        all_mfgs = set()
        supplyframe_mfgs = set()

        # From original data (Step 3)
        xml_gen_page = self.wizard().page(3)
        if hasattr(xml_gen_page, 'combined_data'):
            # Convert DataFrame to list of dictionaries if needed
            data = xml_gen_page.combined_data
            if hasattr(data, 'to_dict'):
                data = data.to_dict('records')
            for row in data:
                if isinstance(row, dict) and row.get('MFG'):
                    all_mfgs.add(row['MFG'])

        # From search results data
        if hasattr(self, 'search_results'):
            for part in self.search_results:
                # Original manufacturers
                if part.get('ManufacturerName'):
                    all_mfgs.add(part['ManufacturerName'])

                # SupplyFrame manufacturers from matches
                # Extract canonical manufacturer names from MatchValue column
                for match in part.get('matches', []):
                    # Match is a dict with 'mfg' key containing the canonical manufacturer
                    if isinstance(match, dict):
                        mfg = match.get('mfg', '').strip()
                        if mfg:
                            supplyframe_mfgs.add(mfg)
                    # Legacy support: if match is a string (old format)
                    elif isinstance(match, str) and '@' in match:
                        _, mfg = match.split('@', 1)
                        supplyframe_mfgs.add(mfg.strip())

        # Show unique manufacturer counts in status
        self.norm_status.setText(
            f"ℹ️ Found {len(all_mfgs)} unique manufacturers in your data, "
            f"{len(supplyframe_mfgs)} from SupplyFrame. Click AI Detect to find variations."
        )
        self.norm_status.setStyleSheet("color: #1976d2; font-weight: bold;")

    def populate_parts_list(self):
        """Populate the parts needing review list"""
        print(f"DEBUG populate_parts_list: parts_needing_review count = {len(self.parts_needing_review)}")
        self.parts_list.setRowCount(len(self.parts_needing_review))

        for row_idx, part in enumerate(self.parts_needing_review):
            if row_idx < 5:  # Log first 5
                print(f"DEBUG: Adding row {row_idx}: {part['PartNumber']} | {part['ManufacturerName']} | {part['MatchStatus']}")

            self.parts_list.setItem(row_idx, 0, QTableWidgetItem(part['PartNumber']))
            self.parts_list.setItem(row_idx, 1, QTableWidgetItem(part['ManufacturerName']))
            self.parts_list.setItem(row_idx, 2, QTableWidgetItem(part['MatchStatus']))

            # Reviewed indicator
            reviewed_item = QTableWidgetItem("✓" if part.get('selected_match') else "")
            reviewed_item.setTextAlignment(Qt.AlignCenter)
            self.parts_list.setItem(row_idx, 3, reviewed_item)

            # AI indicator
            ai_status = ""
            if part.get('ai_processed'):
                ai_status = "🤖"
            elif part.get('ai_processing'):
                ai_status = "⏳"
            ai_item = QTableWidgetItem(ai_status)
            ai_item.setTextAlignment(Qt.AlignCenter)
            self.parts_list.setItem(row_idx, 4, ai_item)

            # Action button - AI Suggest (only if >1 match and not already processed)
            if len(part['matches']) > 1 and not part.get('ai_processed'):
                ai_btn = QPushButton("🤖 AI")
                ai_btn.setToolTip("Use AI to suggest best match for this part")
                ai_btn.clicked.connect(lambda checked, idx=row_idx: self.ai_suggest_single(idx))
                self.parts_list.setCellWidget(row_idx, 5, ai_btn)

        print(f"DEBUG: Parts list populated with {self.parts_list.rowCount()} rows")

    def update_part_row(self, row_idx):
        """Update a single row in the parts list (for real-time AI updates)"""
        if row_idx >= len(self.parts_needing_review):
            return

        part = self.parts_needing_review[row_idx]

        # Update Reviewed indicator
        reviewed_item = QTableWidgetItem("✓" if part.get('selected_match') else "")
        reviewed_item.setTextAlignment(Qt.AlignCenter)
        self.parts_list.setItem(row_idx, 3, reviewed_item)

        # Update AI indicator
        ai_status = ""
        if part.get('ai_processed'):
            ai_status = "🤖"
        elif part.get('ai_processing'):
            ai_status = "⏳"
        ai_item = QTableWidgetItem(ai_status)
        ai_item.setTextAlignment(Qt.AlignCenter)
        self.parts_list.setItem(row_idx, 4, ai_item)

        # Update/Remove Action button - remove if already processed
        if part.get('ai_processed'):
            # Remove the button if AI has processed this part
            self.parts_list.setCellWidget(row_idx, 5, None)
        elif len(part['matches']) > 1 and not part.get('ai_processing'):
            # Re-add button if it's not processing and has multiple matches
            ai_btn = QPushButton("🤖 AI")
            ai_btn.setToolTip("Use AI to suggest best match for this part")
            ai_btn.clicked.connect(lambda checked, idx=row_idx: self.ai_suggest_single(idx))
            self.parts_list.setCellWidget(row_idx, 5, ai_btn)

    def on_part_selected(self):
        """Handle part selection - show matches"""
        # Determine which table triggered the selection
        sender = self.sender()

        # Default to checking all tables
        parts_list = None
        matches_table = None
        parts_data = None
        selected_rows = None

        # Check which table triggered the selection using sender()
        # This ensures we use the correct table even if other tables have selections
        if sender == self.multiple_table:
            if self.multiple_table.selectedIndexes():
                selected_rows = self.multiple_table.selectedIndexes()
                parts_list = self.multiple_table
                matches_table = self.multiple_matches_table
                parts_data = self.multiple_parts
        elif sender == self.need_review_table:
            if self.need_review_table.selectedIndexes():
                selected_rows = self.need_review_table.selectedIndexes()
                parts_list = self.need_review_table
                matches_table = self.need_review_matches_table
                parts_data = self.need_review_parts
        else:
            # Fallback to checking all tables if sender is not recognized
            # This handles cases where the method might be called manually
            if hasattr(self, 'multiple_table') and self.multiple_table.selectedIndexes():
                selected_rows = self.multiple_table.selectedIndexes()
                parts_list = self.multiple_table
                matches_table = self.multiple_matches_table
                parts_data = self.multiple_parts
            elif hasattr(self, 'need_review_table') and self.need_review_table.selectedIndexes():
                selected_rows = self.need_review_table.selectedIndexes()
                parts_list = self.need_review_table
                matches_table = self.need_review_matches_table
                parts_data = self.need_review_parts
            else:
                return

        if not selected_rows or not parts_data:
            return

        row_idx = selected_rows[0].row()
        if row_idx >= len(parts_data):
            return
            
        part = parts_data[row_idx]
        
        # Ensure part is a dict and has required keys
        if not isinstance(part, dict):
            print(f"ERROR on_part_selected: part is not a dict: {type(part)} - {part}")
            return
        
        if 'matches' not in part:
            part['matches'] = []
        
        # Populate matches table
        matches = part.get('matches', [])
        matches_table.setRowCount(len(matches))

        # Create a button group to ensure only one radio button can be selected at a time
        button_group = QButtonGroup()
        # Store the button group to prevent garbage collection
        matches_table.button_group = button_group

        # Calculate similarity scores for confidence
        from difflib import SequenceMatcher
        original_pn = part.get('PartNumber', '').upper().strip()

        for match_idx, match in enumerate(matches):
            # Extract match information using helper function
            mpn, mfg, lifecycle_status, lifecycle_code, external_id, findchips_url, match_string = self._get_match_info(match)

            # Radio button for selection - centered in cell
            radio = QRadioButton()
            # Check against match_string for compatibility
            selected = part.get('selected_match')
            if isinstance(selected, dict):
                is_selected = (selected.get('match_string') == match_string)
            else:
                is_selected = (selected == match or selected == match_string)
            radio.setChecked(is_selected)
            radio.toggled.connect(lambda checked, p=part, m=match: self.on_match_selected(p, m, checked))

            # Add radio button to the button group to ensure mutual exclusivity
            button_group.addButton(radio)

            # Create a widget to center the radio button
            radio_widget = QWidget()
            radio_layout = QHBoxLayout(radio_widget)
            radio_layout.addWidget(radio)
            radio_layout.setAlignment(Qt.AlignCenter)
            radio_layout.setContentsMargins(0, 0, 0, 0)

            matches_table.setCellWidget(match_idx, 0, radio_widget)
            matches_table.setItem(match_idx, 1, QTableWidgetItem(mpn))
            matches_table.setItem(match_idx, 2, QTableWidgetItem(mfg))
            
            # Lifecycle Status column
            lifecycle_item = QTableWidgetItem(lifecycle_status or '')
            lifecycle_item.setToolTip(f"Lifecycle Status Code: {lifecycle_code}" if lifecycle_code else "No lifecycle info")
            matches_table.setItem(match_idx, 3, lifecycle_item)
            
            # External ID column (link)
            external_item = QTableWidgetItem('')
            if external_id:
                # Truncate long URLs for display
                display_url = external_id if len(external_id) <= 40 else external_id[:37] + '...'
                external_item.setText(display_url)
                external_item.setToolTip(f"Click to open: {external_id}")
                external_item.setForeground(QColor(0, 0, 255))  # Blue for links
            matches_table.setItem(match_idx, 4, external_item)

            # Calculate similarity score
            match_pn = mpn.upper().strip()
            similarity = SequenceMatcher(None, original_pn, match_pn).ratio()
            similarity_pct = int(similarity * 100)
            similarity_item = QTableWidgetItem(f"{similarity_pct}%")
            similarity_item.setTextAlignment(Qt.AlignCenter)
            similarity_item.setToolTip("String similarity using difflib (part number matching)")
            matches_table.setItem(match_idx, 5, similarity_item)

            # AI Score - only show if AI has processed this part
            ai_score_item = QTableWidgetItem("")
            ai_score_item.setTextAlignment(Qt.AlignCenter)
            if part.get('ai_processed') and part.get('ai_match_scores'):
                # Get AI confidence for this specific match
                ai_scores = part.get('ai_match_scores', {})
                # Check both match and match_string
                score_key = match if not isinstance(match, dict) else match_string
                if score_key in ai_scores:
                    ai_conf = ai_scores[score_key]
                    ai_score_item.setText(f"{ai_conf}%")
                    ai_score_item.setToolTip("AI confidence score (considers context, manufacturer, description)")
            matches_table.setItem(match_idx, 6, ai_score_item)

    def on_match_selected(self, part, match, checked):
        """Handle match selection"""
        if checked:
            part['selected_match'] = match
            
            # Determine which tab's widgets to update
            if hasattr(self, 'multiple_parts') and part in self.multiple_parts:
                self.multiple_none_correct_checkbox.setChecked(False)
                self.multiple_save_btn.setEnabled(True)
            elif hasattr(self, 'need_review_parts') and part in self.need_review_parts:
                self.need_review_none_correct_checkbox.setChecked(False)
                self.need_review_save_btn.setEnabled(True)

    def refresh_matches_display(self):
        """Refresh the matches table for the currently selected part"""
        selected_rows = self.parts_list.selectedIndexes()
        if not selected_rows:
            return

        row_idx = selected_rows[0].row()
        if row_idx >= len(self.parts_needing_review):
            return

        part = self.parts_needing_review[row_idx]

        # Re-populate matches table
        self.matches_table.setRowCount(len(part['matches']))

        # Create a button group to ensure only one radio button can be selected at a time
        button_group = QButtonGroup()
        # Store the button group to prevent garbage collection
        self.matches_table.button_group = button_group

        from difflib import SequenceMatcher
        original_pn = part['PartNumber'].upper().strip()

        for match_idx, match in enumerate(part['matches']):
            # Extract match information using helper function
            mpn, mfg, lifecycle_status, lifecycle_code, external_id, findchips_url, match_string = self._get_match_info(match)

            # Radio button
            radio = QRadioButton()
            # Check against match_string for compatibility
            selected = part.get('selected_match')
            if isinstance(selected, dict):
                is_selected = (selected.get('match_string') == match_string)
            else:
                is_selected = (selected == match or selected == match_string)
            radio.setChecked(is_selected)
            radio.toggled.connect(lambda checked, p=part, m=match: self.on_match_selected(p, m, checked))

            # Add radio button to the button group to ensure mutual exclusivity
            button_group.addButton(radio)

            radio_widget = QWidget()
            radio_layout = QHBoxLayout(radio_widget)
            radio_layout.addWidget(radio)
            radio_layout.setAlignment(Qt.AlignCenter)
            radio_layout.setContentsMargins(0, 0, 0, 0)

            self.matches_table.setCellWidget(match_idx, 0, radio_widget)
            self.matches_table.setItem(match_idx, 1, QTableWidgetItem(mpn))
            self.matches_table.setItem(match_idx, 2, QTableWidgetItem(mfg))
            
            # Lifecycle Status column
            lifecycle_item = QTableWidgetItem(lifecycle_status or '')
            lifecycle_item.setToolTip(f"Lifecycle Status Code: {lifecycle_code}" if lifecycle_code else "No lifecycle info")
            self.matches_table.setItem(match_idx, 3, lifecycle_item)
            
            # External ID column (link)
            external_item = QTableWidgetItem('')
            if external_id:
                # Truncate long URLs for display
                display_url = external_id if len(external_id) <= 40 else external_id[:37] + '...'
                external_item.setText(display_url)
                external_item.setToolTip(f"Click to open: {external_id}")
                external_item.setForeground(QColor(0, 0, 255))  # Blue for links
            self.matches_table.setItem(match_idx, 4, external_item)

            # Similarity score
            match_pn = mpn.upper().strip()
            similarity = SequenceMatcher(None, original_pn, match_pn).ratio()
            similarity_pct = int(similarity * 100)
            similarity_item = QTableWidgetItem(f"{similarity_pct}%")
            similarity_item.setTextAlignment(Qt.AlignCenter)
            similarity_item.setToolTip("String similarity using difflib (part number matching)")
            self.matches_table.setItem(match_idx, 5, similarity_item)

            # AI Score - show if available
            ai_score_item = QTableWidgetItem("")
            ai_score_item.setTextAlignment(Qt.AlignCenter)
            if part.get('ai_processed') and part.get('ai_match_scores'):
                ai_scores = part.get('ai_match_scores', {})
                # Check both match and match_string
                score_key = match if not isinstance(match, dict) else match_string
                if score_key in ai_scores:
                    ai_conf = ai_scores[score_key]
                    ai_score_item.setText(f"{ai_conf}%")
                    ai_score_item.setToolTip("AI confidence score (considers context, manufacturer, description)")
            self.matches_table.setItem(match_idx, 6, ai_score_item)

    def show_match_context_menu(self, position):
        """Show context menu for matches table"""
        row = self.matches_table.rowAt(position.y())
        if row < 0:
            return

        # Get the currently selected part
        selected_rows = self.parts_list.selectedIndexes()
        if not selected_rows:
            return

        part_idx = selected_rows[0].row()
        if part_idx >= len(self.parts_needing_review):
            return

        part = self.parts_needing_review[part_idx]

        # Only show menu if AI has processed this part
        if not part.get('ai_processed') or not part.get('ai_reasoning'):
            return

        # Get the match at this row
        if row >= len(part['matches']):
            return

        match = part['matches'][row]

        # Only show reasoning if this is the AI-suggested match
        if match != part.get('selected_match'):
            return

        # Create context menu
        menu = QMenu(self)
        action = menu.addAction("🤖 Show AI Reasoning")

        selected_action = menu.exec_(self.matches_table.viewport().mapToGlobal(position))

        if selected_action == action:
            self.show_ai_reasoning(part)

    def show_ai_reasoning(self, part):
        """Show AI reasoning in a dialog"""
        reasoning = part.get('ai_reasoning', 'No reasoning available')
        confidence = part.get('ai_confidence', 0)
        selected_match = part.get('selected_match', 'N/A')

        # Parse match using helper function
        if isinstance(selected_match, dict):
            pn = selected_match.get('mpn', 'N/A')
            mfg = selected_match.get('mfg', 'N/A')
        elif isinstance(selected_match, str):
            if '@' in selected_match:
                pn, mfg = selected_match.split('@', 1)
            else:
                pn = selected_match
                mfg = "N/A"
        else:
            pn = str(selected_match)
            mfg = "N/A"

        dialog = QMessageBox(self)
        dialog.setWindowTitle("AI Reasoning")
        dialog.setIcon(QMessageBox.Information)
        dialog.setText(f"<b>AI Suggested Match:</b><br>"
                      f"Part Number: {pn}<br>"
                      f"Manufacturer: {mfg}<br>"
                      f"Confidence: {confidence}%<br><br>"
                      f"<b>Reasoning:</b>")
        dialog.setInformativeText(reasoning)
        dialog.setStandardButtons(QMessageBox.Ok)
        dialog.exec_()

    def update_review_count(self):
        """Update the reviewed parts count label"""
        total = len(self.parts_needing_review)
        reviewed = sum(1 for part in self.parts_needing_review if part.get('selected_match'))
        self.review_count_label.setText(f"({reviewed} of {total} reviewed)")

    def save_current_selection(self):
        """Save the current match selection"""
        selected_rows = self.parts_list.selectedIndexes()
        if not selected_rows:
            return

        row_idx = selected_rows[0].row()
        if row_idx >= len(self.parts_needing_review):
            return

        part = self.parts_needing_review[row_idx]

        # Mark as user-reviewed
        part['user_reviewed'] = True

        # Update the row
        self.update_part_row(row_idx)

        # Update count
        self.update_review_count()

        # Disable save button
        self.save_selection_btn.setEnabled(False)

        # Show confirmation
        QMessageBox.information(self, "Selection Saved",
                              f"Match selection saved for {part['PartNumber']}")

    def auto_select_highest(self):
        """Auto-select match with highest similarity using difflib (MFG + MFG PN)"""
        from difflib import SequenceMatcher

        selected_count = 0
        for part in self.parts_needing_review:
            if not part['matches']:
                continue

            original_pn = part['PartNumber'].upper().strip()
            original_mfg = part['ManufacturerName'].upper().strip()
            best_match = None
            best_similarity = 0.0

            # Calculate similarity for each match
            for match in part['matches']:
                # Extract match information using helper function
                match_pn, match_mfg, _, _, _, _, _ = self._get_match_info(match)

                match_pn = match_pn.upper().strip()
                match_mfg = match_mfg.upper().strip()

                # Calculate combined similarity (60% part number, 40% manufacturer)
                pn_similarity = SequenceMatcher(None, original_pn, match_pn).ratio()
                mfg_similarity = SequenceMatcher(None, original_mfg, match_mfg).ratio() if match_mfg else 0

                # Weighted average
                combined_similarity = (pn_similarity * 0.6) + (mfg_similarity * 0.4)

                if combined_similarity > best_similarity:
                    best_similarity = combined_similarity
                    best_match = match

            if best_match:
                part['selected_match'] = best_match
                part['auto_selected'] = True
                part['similarity_score'] = best_similarity
                selected_count += 1

        QMessageBox.information(self, "Auto-Select Complete",
                              f"Selected best match for {selected_count} parts using similarity analysis.")

        # Update review count
        self.update_review_count()

        # Refresh current selection if any
        self.on_part_selected()

    def ai_suggest_single(self, row_idx):
        """Use AI to suggest best match for a single part"""
        if row_idx >= len(self.parts_needing_review):
            return

        part = self.parts_needing_review[row_idx]

        # Skip if already processed or processing
        if part.get('ai_processed') or part.get('ai_processing'):
            return

        # Skip if only one match
        if len(part['matches']) <= 1:
            QMessageBox.information(self, "No AI Needed",
                                  "This part has only one match. No AI analysis needed.")
            return

        # Get API key
        start_page = self.wizard().page(0)
        self.api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None

        if not self.api_key or not ANTHROPIC_AVAILABLE:
            QMessageBox.warning(self, "AI Not Available",
                              "Claude AI is not available. Please provide an API key.")
            return

        # Get combined data from previous step
        xml_gen_page = self.wizard().page(3)
        if hasattr(xml_gen_page, 'combined_data'):
            self.combined_data = xml_gen_page.combined_data

        # Mark this part as processing
        part['ai_processing'] = True
        self.update_part_row(row_idx)

        # Create a list with just this one part
        parts_to_process = [part]

        # Start AI thread for single part
        self.ai_match_thread = PartialMatchAIThread(
            self.api_key,
            parts_to_process,
            self.combined_data
        )
        self.ai_match_thread.progress.connect(lambda msg, cur, tot: self.csv_summary.setText(f"🤖 Analyzing part..."))
        self.ai_match_thread.part_analyzed.connect(lambda idx, result: self.on_part_analyzed(row_idx, result))
        self.ai_match_thread.finished.connect(lambda suggestions: self.csv_summary.setText(f"✓ AI analysis complete"))
        self.ai_match_thread.error.connect(lambda err: QMessageBox.critical(self, "AI Error", f"AI analysis failed:\n{err}"))
        self.ai_match_thread.start()

    def ai_suggest_matches(self):
        """Use AI to suggest best matches for all unprocessed parts"""
        # Get API key
        start_page = self.wizard().page(0)
        self.api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None

        if not self.api_key or not ANTHROPIC_AVAILABLE:
            QMessageBox.warning(self, "AI Not Available",
                              "Claude AI is not available. Please provide an API key.")
            return

        # Get combined data from previous step
        xml_gen_page = self.wizard().page(3)
        if hasattr(xml_gen_page, 'combined_data'):
            self.combined_data = xml_gen_page.combined_data

        # Filter out already processed parts
        unprocessed_parts = [part for part in self.parts_needing_review
                            if not part.get('ai_processed')]

        if not unprocessed_parts:
            QMessageBox.information(self, "All Processed",
                                  "All parts have already been processed by AI.")
            return

        # Disable buttons
        self.ai_suggest_btn.setEnabled(False)
        self.auto_select_btn.setEnabled(False)

        # Mark unprocessed parts as processing
        for part in unprocessed_parts:
            part['ai_processing'] = True

        # Refresh parts list to show processing indicators
        self.populate_parts_list()

        # Start AI thread with only unprocessed parts
        self.ai_match_thread = PartialMatchAIThread(
            self.api_key,
            unprocessed_parts,
            self.combined_data
        )
        self.ai_match_thread.progress.connect(self.on_ai_match_progress)
        self.ai_match_thread.part_analyzed.connect(self.on_part_analyzed)  # NEW: real-time updates
        self.ai_match_thread.finished.connect(self.on_ai_match_finished)
        self.ai_match_thread.error.connect(self.on_ai_match_error)
        self.ai_match_thread.start()

    def on_ai_match_progress(self, message, current, total):
        """Update AI progress"""
        self.csv_summary.setText(message)
        self.csv_summary.setStyleSheet("padding: 5px; background-color: #e3f2fd; border-radius: 3px;")

    def on_part_analyzed(self, row_idx, result):
        """Handle real-time part analysis completion"""
        # Ensure result is a dict
        if not isinstance(result, dict):
            print(f"ERROR on_part_analyzed: result is not a dict: {type(result)} - {result}")
            return
        
        # The row_idx from AI thread is relative to the unprocessed parts list
        # We need to find the actual part in our categorized lists
        # For now, we'll search through both Multiple and Need Review lists
        part = None
        part_number = result.get('part_number')  # Assuming AI thread provides this
        
        # Try to find the part in multiple_parts and need_review_parts
        for p in self.multiple_parts + self.need_review_parts:
            if isinstance(p, dict) and p.get('PartNumber') == part_number:
                part = p
                break
        
        # Fallback: try using row_idx on parts_needing_review if we have it
        if part is None and row_idx < len(self.parts_needing_review):
            potential_part = self.parts_needing_review[row_idx]
            if isinstance(potential_part, dict):
                part = potential_part
        
        if part is None or not isinstance(part, dict):
            print(f"ERROR on_part_analyzed: Could not find part for row_idx {row_idx}")
            return

        # Clear processing flag
        part['ai_processing'] = False

        if result.get('skipped'):
            # Part was skipped (single match)
            part['ai_processed'] = False
        elif result.get('error'):
            # AI failed for this part
            part['ai_processed'] = False
        else:
            # AI successfully analyzed
            part['ai_processed'] = True

            # Store AI confidence score for the suggested match
            idx = result.get('suggested_index')
            confidence = result.get('confidence', 0)

            # Create a dictionary to store AI scores for each match
            # (For now, only the suggested match has an AI score)
            if not part.get('ai_match_scores'):
                part['ai_match_scores'] = {}

            matches = part.get('matches', [])
            if idx is not None and 0 <= idx < len(matches):
                suggested_match = matches[idx]
                part['selected_match'] = suggested_match
                part['ai_confidence'] = confidence
                part['ai_reasoning'] = result.get('reasoning', '')
                # Store the AI score for this specific match
                # Use match_string as key (hashable) instead of the dict itself
                match_key = suggested_match.get('match_string', '') if isinstance(suggested_match, dict) else str(suggested_match)
                part['ai_match_scores'][match_key] = confidence

        # Refresh the appropriate tables
        if part in self.multiple_parts:
            self.populate_category_table(self.multiple_table, self.multiple_parts, show_actions=True)
        elif part in self.need_review_parts:
            self.populate_category_table(self.need_review_table, self.need_review_parts, show_actions=True)

        # If this part is currently selected, refresh the matches display to show AI scores
        selected_rows = self.parts_list.selectedIndexes()
        if selected_rows:
            row_idx = selected_rows[0].row()
            if row_idx < len(self.parts_needing_review):
                selected_part = self.parts_needing_review[row_idx]
                if selected_part is part:
                    self.refresh_matches_display()

    def on_ai_match_finished(self, suggestions):
        """Apply AI suggestions"""
        applied = 0
        for part in self.parts_needing_review:
            # Ensure part is a dict
            if not isinstance(part, dict):
                continue
            
            pn = part.get('PartNumber')
            if pn and pn in suggestions:
                suggestion = suggestions[pn]
                idx = suggestion.get('suggested_index')
                matches = part.get('matches', [])
                if idx is not None and 0 <= idx < len(matches):
                    part['selected_match'] = matches[idx]
                    part['ai_confidence'] = suggestion.get('confidence', 0)
                    part['ai_reasoning'] = suggestion.get('reasoning', '')
                    applied += 1

        self.csv_summary.setText(f"✓ AI suggestions applied to {applied} parts")
        self.csv_summary.setStyleSheet("padding: 5px; background-color: #c8e6c9; border-radius: 3px; font-weight: bold;")

        # Re-enable buttons
        self.ai_suggest_btn.setEnabled(True)
        self.auto_select_btn.setEnabled(True)

        # Refresh view
        self.on_part_selected()

        QMessageBox.information(self, "AI Suggestions Complete",
                              f"AI suggested matches for {applied} parts.\n"
                              f"Review the suggestions and adjust as needed.")

    def on_ai_match_error(self, error_msg):
        """Handle AI error"""
        self.csv_summary.setText(f"✗ AI Error: {error_msg[:50]}")
        self.csv_summary.setStyleSheet("padding: 5px; background-color: #ffcdd2; border-radius: 3px;")

        self.ai_suggest_btn.setEnabled(True)
        self.auto_select_btn.setEnabled(True)

        QMessageBox.critical(self, "AI Error", f"AI suggestion failed:\n{error_msg}")

    def ai_detect_normalizations(self):
        """Use AI to detect manufacturer normalizations"""
        # Get API key
        start_page = self.wizard().page(0)
        self.api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None

        if not self.api_key or not ANTHROPIC_AVAILABLE:
            QMessageBox.warning(self, "AI Not Available",
                              "Claude AI is not available. Please provide an API key.")
            return

        # Collect manufacturers - CRITICAL: Keep source and target separate!
        source_mfgs = set()  # Only manufacturers from user's original data
        canonical_mfgs = set()  # Only canonical manufacturers from PAS

        # From original data (ONLY manufacturers from Step 3 - the SOURCE)
        xml_gen_page = self.wizard().page(3)
        if hasattr(xml_gen_page, 'combined_data'):
            # Convert DataFrame to list of dictionaries if needed
            data = xml_gen_page.combined_data
            if hasattr(data, 'to_dict'):
                data = data.to_dict('records')
            for row in data:
                if isinstance(row, dict) and row.get('MFG'):
                    source_mfgs.add(row['MFG'])

        # From search results - collect ONLY canonical manufacturers from PAS (the TARGET)
        # DO NOT add user's original manufacturers here - they're already in source_mfgs
        if hasattr(self, 'search_results'):
            for part in self.search_results:
                # Extract canonical manufacturer names from MatchValue column
                for match in part.get('matches', []):
                    # Match is a dict with 'mfg' key containing the canonical manufacturer
                    if isinstance(match, dict):
                        mfg = match.get('mfg', '').strip()
                        if mfg:
                            canonical_mfgs.add(mfg)
                    # Legacy support: if match is a string (old format)
                    elif isinstance(match, str) and '@' in match:
                        _, mfg = match.split('@', 1)
                        canonical_mfgs.add(mfg.strip())

        self.norm_status.setText("🤖 Analyzing manufacturers...")
        self.norm_status.setStyleSheet("color: blue;")
        self.ai_normalize_btn.setEnabled(False)

        # Start AI thread with SEPARATE source and target lists
        self.ai_norm_thread = ManufacturerNormalizationAIThread(
            self.api_key,
            list(source_mfgs),  # Only user's original manufacturers
            list(canonical_mfgs)  # Only PAS canonical manufacturers
        )
        self.ai_norm_thread.progress.connect(lambda msg: self.norm_status.setText(msg))
        self.ai_norm_thread.finished.connect(self.on_ai_norm_finished)
        self.ai_norm_thread.error.connect(self.on_ai_norm_error)
        self.ai_norm_thread.start()

    def on_ai_norm_finished(self, normalizations, reasoning_map):
        """Apply pure AI normalization suggestions"""
        self.manufacturer_normalizations = normalizations
        self.normalization_reasoning = reasoning_map  # Store reasoning for context menu

        # Populate AI cache with bulk results for consistency
        for original, canonical in normalizations.items():
            if original not in self.ai_cache:  # Don't overwrite existing cache
                self.ai_cache[original] = {
                    'canonical': canonical,
                    'reasoning': reasoning_map.get(original, {}).get('reasoning', 'AI suggested normalization')
                }

        # Collect manufacturers - keep lists separate to avoid contamination
        canonical_mfgs = set()  # PAS canonical manufacturers
        selected_mfgs = set()  # User-selected manufacturers from review phase

        # From search results - collect canonical manufacturers (for dropdown population)
        if hasattr(self, 'search_results'):
            for result in self.search_results:
                # Collect all canonical manufacturers from matches using helper function
                for match in result.get('matches', []):
                    _, mfg, _, _, _, _, _ = self._get_match_info(match)
                    mfg = mfg.strip()
                    if mfg:
                        canonical_mfgs.add(mfg)
                        # DO NOT add to all_mfgs - only use for dropdown

                # Track user-selected manufacturers (their review work)
                if result.get('selected_match'):
                    _, mfg, _, _, _, _, _ = self._get_match_info(result['selected_match'])
                    mfg = mfg.strip()
                    if mfg:
                        selected_mfgs.add(mfg)
                        # DO NOT add to all_mfgs - only use for tracking

        # Get all unique manufacturers from original data (ensure we show EVERYTHING)
        unique_original_mfgs = set()
        if hasattr(xml_gen_page, 'combined_data'):
            data = xml_gen_page.combined_data
            if hasattr(data, 'to_dict'):
                data = data.to_dict('records')
            for row in data:
                if isinstance(row, dict) and row.get('MFG'):
                    unique_original_mfgs.add(row['MFG'])

        # Create entries for ALL manufacturers (suggestions + no-change entries)
        all_entries = {}
        # Add all original manufacturers with identity mapping (no change by default)
        for mfg in unique_original_mfgs:
            all_entries[mfg] = mfg
        # Override with AI/fuzzy suggestions where available
        for original, canonical in normalizations.items():
            all_entries[original] = canonical

        # Populate normalization table with ALL manufacturers
        self.norm_table.setRowCount(len(all_entries))

        row_idx = 0
        for original, canonical in sorted(all_entries.items()):
            # Include checkbox - center it in the cell
            include_cb = QCheckBox()
            # Check if this is from AI/fuzzy suggestions (in normalizations dict)
            has_suggestion = original in normalizations
            if has_suggestion:
                method = reasoning_map.get(original, {}).get('method', 'manual')
                # Check fuzzy/AI matches, uncheck exact matches and identity mappings
                include_cb.setChecked(method != 'exact' and original != canonical)
            else:
                # Identity mapping (no change) - uncheck by default
                include_cb.setChecked(False)
            
            # Create a widget to center the checkbox
            checkbox_widget = QWidget()
            checkbox_layout = QHBoxLayout(checkbox_widget)
            checkbox_layout.addWidget(include_cb)
            checkbox_layout.setAlignment(Qt.AlignCenter)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            self.norm_table.setCellWidget(row_idx, 0, checkbox_widget)

            # Column 1: Status - show the method (MOVED TO COLUMN 1)
            if has_suggestion:
                method = reasoning_map.get(original, {}).get('method', 'manual')
                score = reasoning_map.get(original, {}).get('score', 0)
                status_map = {
                    'exact': 'Exact',
                    'fuzzy': 'Fuzzy',
                    'ai': 'AI',
                    'manual': 'Manual'
                }
                status_text = status_map.get(method, 'Manual')
            else:
                status_text = 'No Change'
                method = 'no_change'
                score = 0

            status_item = QTableWidgetItem(status_text)
            status_item.setTextAlignment(Qt.AlignCenter)
            
            # Color code the status
            if method == 'exact':
                status_item.setBackground(QColor(230, 255, 230))  # Light green
            elif method == 'fuzzy':
                status_item.setBackground(QColor(255, 250, 205))  # Light yellow
            elif method == 'ai':
                status_item.setBackground(QColor(230, 240, 255))  # Light blue
            elif method == 'manual':
                status_item.setBackground(QColor(255, 240, 200))  # Light orange
            else:  # no_change
                status_item.setBackground(QColor(248, 248, 248))  # Very light gray
            
            self.norm_table.setItem(row_idx, 1, status_item)

            # Column 2: Original MFG (read-only)
            self.norm_table.setItem(row_idx, 2, QTableWidgetItem(original))

            # Column 3: Normalize To (editable combo box)
            normalize_combo = QComboBox()
            normalize_combo.setEditable(True)
            # Disable mouse wheel
            normalize_combo.wheelEvent = lambda event: event.ignore()
            normalize_combo.setFocusPolicy(Qt.StrongFocus)

            # Add only PAS canonical manufacturers to dropdown
            if hasattr(self, 'canonical_manufacturers'):
                for mfg in sorted(self.canonical_manufacturers):
                    normalize_combo.addItem(mfg)
            else:
                # Fallback to canonical_mfgs from this function
                for mfg in sorted(canonical_mfgs):
                    normalize_combo.addItem(mfg)

            # Set current suggestion
            normalize_combo.setCurrentText(canonical)
            self.norm_table.setCellWidget(row_idx, 3, normalize_combo)

            # Column 4: AI Score - show score percentage
            ai_score_item = QTableWidgetItem("")
            ai_score_item.setTextAlignment(Qt.AlignCenter)
            if method in ['fuzzy', 'ai'] and score > 0:
                ai_score_item.setText(f"{score}%")
                ai_score_item.setToolTip(f"{method.capitalize()} match confidence: {score}%")
            self.norm_table.setItem(row_idx, 4, ai_score_item)

            # Column 5: AI Analyze button
            ai_btn = QPushButton("🤖 AI")
            ai_btn.setMaximumWidth(60)
            ai_btn.setToolTip("Run AI analysis for this manufacturer")
            ai_btn.clicked.connect(lambda checked, r=row_idx, orig=original: self.analyze_single_manufacturer_ai(r, orig))

            # Disable if no API key available
            start_page = self.wizard().page(0)
            api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None
            if not api_key or not ANTHROPIC_AVAILABLE:
                ai_btn.setEnabled(False)
                ai_btn.setToolTip("AI analysis not available (no API key)")

            # Center the button in the cell
            ai_btn_widget = QWidget()
            ai_btn_layout = QHBoxLayout(ai_btn_widget)
            ai_btn_layout.addWidget(ai_btn)
            ai_btn_layout.setAlignment(Qt.AlignCenter)
            ai_btn_layout.setContentsMargins(0, 0, 0, 0)
            self.norm_table.setCellWidget(row_idx, 5, ai_btn_widget)

            # Column 6: Scope dropdown
            scope_combo = QComboBox()
            scope_combo.addItems(["All Catalogs", "Per Catalog"])
            # Disable mouse wheel
            scope_combo.wheelEvent = lambda event: event.ignore()
            scope_combo.setFocusPolicy(Qt.StrongFocus)
            self.norm_table.setCellWidget(row_idx, 6, scope_combo)

            # Note: Color coding is already applied to Status column above
            # No need for additional row color coding

            row_idx += 1

        # Count method types
        fuzzy_count = sum(1 for v in reasoning_map.values() if v.get('method') == 'fuzzy')
        ai_count = sum(1 for v in reasoning_map.values() if v.get('method') == 'ai')
        exact_count = sum(1 for v in reasoning_map.values() if v.get('method') == 'exact')
        manual_count = sum(1 for v in reasoning_map.values() if v.get('method') == 'manual')
        no_change_count = len(all_entries) - len(normalizations)

        self.norm_status.setText(
            f"✓ Showing all {len(all_entries)} manufacturers: "
            f"{no_change_count} no change, {exact_count} exact, {fuzzy_count} fuzzy, {ai_count} AI-validated, {manual_count} manual"
        )
        self.norm_status.setStyleSheet("color: green; font-weight: bold;")
        self.ai_normalize_btn.setEnabled(True)
        self.save_normalizations_btn.setEnabled(True)

        QMessageBox.information(self, "Normalization Detection Complete",
                              f"Analysis complete! Showing all {len(all_entries)} manufacturers:\n\n"
                              f"• {no_change_count} already correct (no changes needed)\n"
                              f"• {exact_count} exact matches in PAS\n"
                              f"• {fuzzy_count} fuzzy match suggestions\n"
                              f"• {ai_count} AI-validated suggestions\n"
                              f"• {manual_count} need manual review\n\n"
                              f"Right-click any row to see detection reasoning.\n"
                              f"Review and adjust as needed.")

    def on_ai_norm_error(self, error_msg):
        """Handle AI normalization error"""
        self.norm_status.setText(f"✗ Error: {error_msg[:30]}")
        self.norm_status.setStyleSheet("color: red;")
        self.ai_normalize_btn.setEnabled(True)

        QMessageBox.critical(self, "AI Error", f"AI normalization failed:\n{error_msg}")

    def analyze_single_manufacturer_ai(self, row_idx, original_mfg):
        """Analyze a single manufacturer using AI with caching"""
        if not ANTHROPIC_AVAILABLE:
            QMessageBox.warning(self, "AI Not Available", "Claude AI package not installed.")
            return

        start_page = self.wizard().page(0)
        api_key = start_page.get_api_key() if hasattr(start_page, 'get_api_key') else None
        if not api_key:
            QMessageBox.warning(self, "No API Key", "Please configure Claude AI API key in Step 1.")
            return

        # Check cache first
        if original_mfg in self.ai_cache:
            cached_result = self.ai_cache[original_mfg]
            QMessageBox.information(
                self,
                "Cached Result",
                f"Using cached AI result for '{original_mfg}':\n\n"
                f"Suggested normalization: {cached_result['canonical']}\n\n"
                f"Reasoning: {cached_result['reasoning']}"
            )
            # Update the table with cached result
            self.update_table_row_with_ai_result(row_idx, original_mfg, cached_result['canonical'], cached_result['reasoning'])
            return

        # Collect canonical manufacturer names from PAS
        canonical_mfgs = set()
        if hasattr(self, 'search_results'):
            for result in self.search_results:
                for match in result.get('matches', []):
                    # Match is a dict with 'mfg' key containing the canonical manufacturer
                    if isinstance(match, dict):
                        mfg = match.get('mfg', '').strip()
                        if mfg:
                            canonical_mfgs.add(mfg)
                    # Legacy support: if match is a string (old format)
                    elif isinstance(match, str) and '@' in match:
                        _, mfg = match.split('@', 1)
                        mfg = mfg.strip()
                        if mfg:
                            canonical_mfgs.add(mfg)

        try:
            # Call AI for single manufacturer
            from anthropic import Anthropic
            client = Anthropic(api_key=api_key)

            prompt = f"""Analyze this manufacturer name and suggest a normalized form.

Manufacturer to analyze: "{original_mfg}"

PAS/SupplyFrame canonical manufacturer names (prefer these if applicable):
{json.dumps(sorted(canonical_mfgs), indent=2)}

Instructions:
1. If this manufacturer name matches or is a variation of a PAS canonical name, use that
2. If it's an abbreviation, expand it to the full company name
3. If it's an acquired company, map to the parent company
4. If it's already correct and complete, use the same name
5. Provide brief reasoning for your decision

Return ONLY valid JSON with this structure:
{{
    "canonical_name": "Suggested Manufacturer Name",
    "reasoning": "Brief explanation of why this normalization is suggested"
}}

IMPORTANT: Return ONLY valid JSON, no markdown, no other text."""

            response = client.messages.create(
                model="claude-sonnet-4-5-20250929",
                max_tokens=1024,
                temperature=0,  # Ensure consistent results
                messages=[{"role": "user", "content": prompt}]
            )

            response_text = response.content[0].text.strip()

            # Clean up code blocks
            if response_text.startswith('```'):
                parts = response_text.split('```')
                if len(parts) >= 2:
                    response_text = parts[1]
                    if response_text.startswith('json'):
                        response_text = response_text[4:]
                    response_text = response_text.strip()

            # Parse JSON
            import re
            try:
                ai_result = json.loads(response_text)
            except json.JSONDecodeError:
                # Try to extract JSON object
                json_match = re.search(r'\{[\s\S]*\}', response_text)
                if json_match:
                    ai_result = json.loads(json_match.group())
                else:
                    raise ValueError("Could not parse AI response")

            canonical_name = ai_result.get('canonical_name', original_mfg)
            reasoning = ai_result.get('reasoning', 'AI suggested normalization')

            # Cache the result
            self.ai_cache[original_mfg] = {
                'canonical': canonical_name,
                'reasoning': reasoning
            }

            # Update the table row
            self.update_table_row_with_ai_result(row_idx, original_mfg, canonical_name, reasoning)

            QMessageBox.information(
                self,
                "AI Analysis Complete",
                f"Manufacturer: {original_mfg}\n\n"
                f"Suggested normalization: {canonical_name}\n\n"
                f"Reasoning: {reasoning}"
            )

        except Exception as e:
            QMessageBox.critical(self, "AI Error", f"AI analysis failed:\n{str(e)}")

    def update_table_row_with_ai_result(self, row_idx, original_mfg, canonical_name, reasoning):
        """Update a table row with AI analysis result"""
        # Update the "Normalize To" combo box (Column 3)
        normalize_combo = self.norm_table.cellWidget(row_idx, 3)
        if normalize_combo:
            normalize_combo.setCurrentText(canonical_name)

        # Update the Status column (Column 1)
        status_item = QTableWidgetItem("AI")
        status_item.setTextAlignment(Qt.AlignCenter)
        status_item.setBackground(QColor(230, 240, 255))  # Light blue for AI
        self.norm_table.setItem(row_idx, 1, status_item)

        # Update the reasoning map
        self.normalization_reasoning[original_mfg] = {
            'method': 'ai',
            'reasoning': reasoning
        }

        # Update the normalizations dict
        self.manufacturer_normalizations[original_mfg] = canonical_name

        # Check the Include checkbox if normalization is different
        include_widget = self.norm_table.cellWidget(row_idx, 0)
        if include_widget:
            checkbox = include_widget.findChild(QCheckBox)
            if checkbox and original_mfg != canonical_name:
                checkbox.setChecked(True)

        # Update AI Score column
        ai_score_item = self.norm_table.item(row_idx, 4)
        if ai_score_item:
            ai_score_item.setText("")
            ai_score_item.setToolTip("AI analysis performed")

    def on_scope_changed(self, row_idx, combo_idx):
        """Handle scope dropdown changes"""
        scope_combo = self.norm_table.cellWidget(row_idx, 6)  # Column 6: Scope
        if not scope_combo:
            return

        scope_data = scope_combo.currentData()
        if scope_data == "specific":
            # Get available sheets from the combined data
            xml_gen_page = self.wizard().page(3)
            available_sheets = []

            if hasattr(xml_gen_page, 'combined_data'):
                data = xml_gen_page.combined_data
                if hasattr(data, 'to_dict'):
                    # DataFrame - get unique Source_Sheet values
                    if 'Source_Sheet' in data.columns:
                        available_sheets = sorted(data['Source_Sheet'].unique().tolist())
                else:
                    # List of dictionaries
                    sheet_set = set()
                    for row in data:
                        if isinstance(row, dict) and 'Source_Sheet' in row:
                            sheet_set.add(row['Source_Sheet'])
                    available_sheets = sorted(list(sheet_set))

            if not available_sheets:
                QMessageBox.warning(self, "No Sheets Found",
                                  "Could not find sheet information in the data.\n"
                                  "Make sure the data has a 'Source_Sheet' column.")
                scope_combo.setCurrentIndex(0)
                return

            # Show multi-select dialog
            selected_sheets = self.show_sheet_selection_dialog(available_sheets, row_idx)

            if selected_sheets:
                # Store the selected sheets for this row
                self.normalization_scopes[row_idx] = selected_sheets

                # Update the combo box display
                if len(selected_sheets) == len(available_sheets):
                    # All sheets selected - same as "All Catalogs"
                    scope_combo.setCurrentIndex(0)
                    if row_idx in self.normalization_scopes:
                        del self.normalization_scopes[row_idx]
                else:
                    # Show abbreviated list of selected sheets
                    if len(selected_sheets) <= 3:
                        sheets_text = ", ".join(selected_sheets)
                    else:
                        sheets_text = f"{', '.join(selected_sheets[:2])}, +{len(selected_sheets)-2} more"

                    # Update combo text to show selection
                    scope_combo.setItemText(combo_idx, f"Sheets: {sheets_text}")
            else:
                # User cancelled or selected none - revert to "All Catalogs"
                scope_combo.setCurrentIndex(0)
                if row_idx in self.normalization_scopes:
                    del self.normalization_scopes[row_idx]

    def show_sheet_selection_dialog(self, available_sheets, row_idx):
        """Show a dialog for multi-selecting sheets/catalogs"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Select Sheets/Catalogs")
        dialog.setMinimumWidth(400)
        dialog.setMinimumHeight(300)

        layout = QVBoxLayout(dialog)

        # Instructions
        label = QLabel("Select which sheets/catalogs to apply this normalization to:")
        layout.addWidget(label)

        # List widget with checkboxes
        list_widget = QListWidget()

        # Pre-select previously selected sheets if they exist
        previously_selected = self.normalization_scopes.get(row_idx, [])

        for sheet in available_sheets:
            item = QListWidgetItem(sheet)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            # Check if previously selected, otherwise check all by default
            if previously_selected:
                item.setCheckState(Qt.Checked if sheet in previously_selected else Qt.Unchecked)
            else:
                item.setCheckState(Qt.Checked)  # All selected by default
            list_widget.addItem(item)

        layout.addWidget(list_widget)

        # Select All / Deselect All buttons
        button_layout = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        deselect_all_btn = QPushButton("Deselect All")

        def select_all():
            for i in range(list_widget.count()):
                list_widget.item(i).setCheckState(Qt.Checked)

        def deselect_all():
            for i in range(list_widget.count()):
                list_widget.item(i).setCheckState(Qt.Unchecked)

        select_all_btn.clicked.connect(select_all)
        deselect_all_btn.clicked.connect(deselect_all)

        button_layout.addWidget(select_all_btn)
        button_layout.addWidget(deselect_all_btn)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        # OK / Cancel buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        # Show dialog
        if dialog.exec_() == QDialog.Accepted:
            # Get selected sheets
            selected = []
            for i in range(list_widget.count()):
                item = list_widget.item(i)
                if item.checkState() == Qt.Checked:
                    selected.append(item.text())
            return selected

        return None

    def apply_changes(self):
        """Apply all changes and generate comparison"""
        try:
            # Get the combined data from ColumnMappingPage (Step 2)
            column_mapping_page = self.wizard().page(2)  # ColumnMappingPage
            if not hasattr(column_mapping_page, 'combined_data') or column_mapping_page.combined_data is None or column_mapping_page.combined_data.empty:
                QMessageBox.warning(self, "No Data",
                                  "No combined data available from Step 2.\n"
                                  "Please complete Step 2 first.")
                return

            # Create a copy of the original data (convert DataFrame to list of dicts)
            import copy
            old_data = column_mapping_page.combined_data.to_dict('records')
            new_data = column_mapping_page.combined_data.to_dict('records')

            # Track changes for summary
            matches_applied = 0
            normalizations_applied = 0

            # Step 1: Apply selected partial matches
            if hasattr(self, 'search_results'):
                for part_data in self.search_results:
                    if 'selected_match' in part_data and part_data['selected_match']:
                        # Parse the selected match (handles both dict and string formats)
                        selected_match = part_data['selected_match']

                        # Extract PN and MFG from match
                        if isinstance(selected_match, dict):
                            new_pn = selected_match.get('mpn', '').strip()
                            new_mfg = selected_match.get('mfg', '').strip()
                        elif isinstance(selected_match, str) and '@' in selected_match:
                            new_pn, new_mfg = selected_match.split('@', 1)
                            new_pn = new_pn.strip()
                            new_mfg = new_mfg.strip()
                        else:
                            continue  # Skip invalid format

                        if new_pn and new_mfg:
                            # Find and update all matching records in new_data
                            original_pn = part_data['PartNumber']
                            original_mfg = part_data['ManufacturerName']

                            for record in new_data:
                                if (record['MFG_PN'] == original_pn and
                                    record['MFG'] == original_mfg):
                                    record['MFG_PN'] = new_pn
                                    record['MFG'] = new_mfg
                                    matches_applied += 1

            # Step 2: Apply manufacturer normalizations
            for row_idx in range(self.norm_table.rowCount()):
                # Check if this normalization is included
                include_widget = self.norm_table.cellWidget(row_idx, 0)
                if not include_widget:
                    continue
                include_checkbox = include_widget.findChild(QCheckBox)
                if not include_checkbox or not include_checkbox.isChecked():
                    continue

                variation_item = self.norm_table.item(row_idx, 2)  # Column 2: Original MFG
                canonical_combo = self.norm_table.cellWidget(row_idx, 3)  # Column 3: Normalize To
                scope_combo = self.norm_table.cellWidget(row_idx, 6)  # Column 6: Scope

                if not variation_item or not canonical_combo or not scope_combo:
                    continue

                variation = variation_item.text().strip()
                canonical = canonical_combo.currentText().strip()
                scope = scope_combo.currentText()

                # Get selected sheets for this row (if specific sheets were selected)
                selected_sheets = self.normalization_scopes.get(row_idx, None)

                # Apply normalization based on scope
                if scope == "All Catalogs" or selected_sheets is None:
                    # Apply to all records with this manufacturer variation
                    for record in new_data:
                        if record['MFG'] == variation:
                            record['MFG'] = canonical
                            normalizations_applied += 1
                else:
                    # Apply only to records from selected sheets
                    for record in new_data:
                        if record['MFG'] == variation and record.get('Source_Sheet') in selected_sheets:
                            record['MFG'] = canonical
                            normalizations_applied += 1

            # Step 3: Populate comparison tables
            self.populate_comparison_tables(old_data, new_data)

            # Step 4: Store the new data for later use
            self.updated_data = new_data

            # Show summary
            QMessageBox.information(self, "Changes Applied",
                                  f"Changes applied successfully!\n\n"
                                  f"• {matches_applied} parts updated from SupplyFrame matches\n"
                                  f"• {normalizations_applied} manufacturer names normalized\n\n"
                                  f"Review the comparison below.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to apply changes:\n{str(e)}")

    def populate_comparison_tables(self, old_data, new_data):
        """Populate side-by-side comparison tables with highlighting"""
        # Clear existing data
        self.old_data_table.setRowCount(0)
        self.new_data_table.setRowCount(0)

        # Set row count
        row_count = len(old_data)
        self.old_data_table.setRowCount(row_count)
        self.new_data_table.setRowCount(row_count)

        # Track changes
        changed_rows = 0

        for idx, (old_record, new_record) in enumerate(zip(old_data, new_data)):
            row_changed = False

            # Old data table
            old_mfg_item = QTableWidgetItem(old_record['MFG'])
            old_pn_item = QTableWidgetItem(old_record['MFG_PN'])
            old_desc_item = QTableWidgetItem(old_record['Description'][:50] + "..."
                                            if len(old_record['Description']) > 50
                                            else old_record['Description'])

            # New data table
            new_mfg_item = QTableWidgetItem(new_record['MFG'])
            new_pn_item = QTableWidgetItem(new_record['MFG_PN'])
            new_desc_item = QTableWidgetItem(new_record['Description'][:50] + "..."
                                            if len(new_record['Description']) > 50
                                            else new_record['Description'])

            # Highlight changes
            if old_record['MFG'] != new_record['MFG']:
                new_mfg_item.setBackground(QColor(255, 255, 200))  # Light yellow
                row_changed = True

            if old_record['MFG_PN'] != new_record['MFG_PN']:
                new_pn_item.setBackground(QColor(255, 255, 200))  # Light yellow
                row_changed = True

            if row_changed:
                changed_rows += 1

            # Set items
            self.old_data_table.setItem(idx, 0, old_mfg_item)
            self.old_data_table.setItem(idx, 1, old_pn_item)
            self.old_data_table.setItem(idx, 2, old_desc_item)

            self.new_data_table.setItem(idx, 0, new_mfg_item)
            self.new_data_table.setItem(idx, 1, new_pn_item)
            self.new_data_table.setItem(idx, 2, new_desc_item)

        # Update summary label
        self.comparison_summary.setText(
            f"Summary: {changed_rows} of {row_count} records modified "
            f"({changed_rows * 100 // row_count if row_count > 0 else 0}% changed)"
        )

    def regenerate_xml(self):
        """Regenerate XML files with updated data"""
        try:
            if not hasattr(self, 'updated_data') or not self.updated_data:
                QMessageBox.warning(self, "No Data",
                                  "No updated data available.\n"
                                  "Please apply changes first.")
                return

            # Get configuration from XMLGenerationPage (Step 3)
            prev_page_3 = self.wizard().page(3)  # XMLGenerationPage

            # Get project settings
            project_name = prev_page_3.project_name.text()
            catalog = prev_page_3.catalog.text()
            output_dir = Path(prev_page_3.output_path.text())

            # Get Excel file path to determine base name
            prev_page_1 = self.wizard().page(1)  # DataSourcePage
            excel_path = prev_page_1.get_excel_path()

            if not excel_path:
                QMessageBox.warning(self, "No File",
                                  "No Excel file path available.")
                return

            base_name = Path(excel_path).stem

            # Create output file paths with "_Updated" suffix
            mfg_xml_path = output_dir / f"{base_name}_MFG_Updated.xml"
            mfgpn_xml_path = output_dir / f"{base_name}_MFGPN_Updated.xml"

            # Extract unique manufacturers from updated data
            unique_mfgs = sorted(set(record['MFG'] for record in self.updated_data
                                    if record['MFG'] and record['MFG'].strip()))

            # Prepare MFGPN data
            mfgpn_data = []
            for record in self.updated_data:
                if record['MFG'] and record['MFG_PN']:
                    mfgpn_data.append({
                        'MFG': record['MFG'],
                        'MFG_PN': record['MFG_PN'],
                        'Description': record.get('Description', 'This is the PN description.')
                    })

            # Generate MFG XML
            mfg_count = self.create_mfg_xml(unique_mfgs, mfg_xml_path, project_name, catalog)

            # Generate MFGPN XML
            mfgpn_count = self.create_mfgpn_xml(mfgpn_data, mfgpn_xml_path, project_name, catalog)

            # Show success message
            QMessageBox.information(self, "XML Generated",
                                  f"Updated XML files generated successfully!\n\n"
                                  f"Files created:\n"
                                  f"• {mfg_xml_path.name} ({mfg_count} manufacturers)\n"
                                  f"• {mfgpn_xml_path.name} ({mfgpn_count} part numbers)\n\n"
                                  f"Output folder:\n{output_dir}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to regenerate XML:\n{str(e)}")

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

    def show_normalization_context_menu(self, position):
        """Show context menu for normalization table"""
        row = self.norm_table.rowAt(position.y())
        if row < 0:
            return

        # Get the original manufacturer name from this row (Column 2)
        original_item = self.norm_table.item(row, 2)
        if not original_item:
            return

        original_mfg = original_item.text()

        # Check if we have reasoning for this manufacturer
        if original_mfg not in self.normalization_reasoning:
            return

        # Create context menu
        menu = QMenu(self)
        action = menu.addAction("🔍 Show Detection Reasoning")

        selected_action = menu.exec_(self.norm_table.viewport().mapToGlobal(position))

        if selected_action == action:
            self.show_normalization_reasoning(original_mfg)

    def show_normalization_reasoning(self, original_mfg):
        """Show reasoning for manufacturer normalization"""
        reasoning_data = self.normalization_reasoning.get(original_mfg, {})
        method = reasoning_data.get('method', 'unknown')
        reasoning = reasoning_data.get('reasoning', 'No reasoning available')

        # Get canonical name from table
        for row_idx in range(self.norm_table.rowCount()):
            orig_item = self.norm_table.item(row_idx, 2)  # Column 2: Original MFG
            if orig_item and orig_item.text() == original_mfg:
                canonical_combo = self.norm_table.cellWidget(row_idx, 3)  # Column 3: Normalize To
                if canonical_combo:
                    canonical = canonical_combo.currentText()
                    break
        else:
            canonical = self.manufacturer_normalizations.get(original_mfg, 'N/A')

        # Create dialog
        dialog = QMessageBox(self)
        dialog.setWindowTitle("Normalization Detection Reasoning")
        dialog.setIcon(QMessageBox.Information)

        if method == 'fuzzy':
            score = reasoning_data.get('score', 0)
            dialog.setText(f"<b>Detection Method: Fuzzy Matching</b><br><br>"
                          f"Original: <b>{original_mfg}</b><br>"
                          f"Normalized to: <b>{canonical}</b><br>"
                          f"Similarity Score: <b>{score}%</b><br><br>"
                          f"<b>Reasoning:</b>")
            dialog.setInformativeText(reasoning)
        elif method == 'ai':
            dialog.setText(f"<b>Detection Method: AI Analysis</b><br><br>"
                          f"Original: <b>{original_mfg}</b><br>"
                          f"Normalized to: <b>{canonical}</b><br><br>"
                          f"<b>AI Reasoning:</b>")
            dialog.setInformativeText(reasoning)
        else:
            dialog.setText(f"<b>Original:</b> {original_mfg}<br>"
                          f"<b>Normalized to:</b> {canonical}")
            dialog.setInformativeText("No reasoning information available")

        dialog.setStandardButtons(QMessageBox.Ok)
        dialog.exec_()

    def save_normalizations(self):
        """Save manufacturer normalization settings"""
        if not self.manufacturer_normalizations:
            QMessageBox.warning(self, "No Normalizations",
                              "No normalizations to save.\n"
                              "Run AI detection first.")
            return

        # Count enabled normalizations
        enabled_count = 0
        for row_idx in range(self.norm_table.rowCount()):
            include_widget = self.norm_table.cellWidget(row_idx, 0)
            if include_widget:
                include_checkbox = include_widget.findChild(QCheckBox)
                if include_checkbox and include_checkbox.isChecked():
                    enabled_count += 1

        QMessageBox.information(self, "Normalizations Saved",
                              f"Manufacturer normalization settings saved!\n\n"
                              f"• {enabled_count} of {len(self.manufacturer_normalizations)} normalizations enabled\n"
                              f"• Settings will be applied when you click 'Apply Changes'\n\n"
                              f"You can continue editing the normalizations as needed.")

    def validatePage(self):
        """Apply normalizations and create Combined_New sheet before proceeding"""
        try:
            # Get the combined data from Step 2
            column_mapping_page = self.wizard().page(2)
            if not hasattr(column_mapping_page, 'combined_data') or column_mapping_page.combined_data is None:
                QMessageBox.warning(self, "No Data", "No combined data found from Step 2.")
                return False

            # Get the output Excel path
            if not hasattr(column_mapping_page, 'output_excel_path') or not column_mapping_page.output_excel_path:
                QMessageBox.warning(self, "No Output File", "Output Excel file not found.")
                return False

            output_excel = column_mapping_page.output_excel_path

            # Start with a copy of the combined data
            new_data = column_mapping_page.combined_data.copy()

            # Apply selected partial matches from search results
            if hasattr(self, 'search_results') and self.search_results:
                for part_data in self.search_results:
                    if 'selected_match' in part_data and part_data['selected_match']:
                        selected = part_data['selected_match']
                        # Find matching row in new_data
                        mask = (new_data['MFG_PN'] == part_data.get('part_number', ''))
                        if mask.any():
                            new_data.loc[mask, 'MFG'] = selected.get('manufacturer', '')

            # Apply manufacturer normalizations
            if hasattr(self, 'manufacturer_normalizations') and self.manufacturer_normalizations:
                normalizations_applied = 0
                for row_idx in range(self.norm_table.rowCount()):
                    include_widget = self.norm_table.cellWidget(row_idx, 0)
                    if not include_widget:
                        continue
                    include_checkbox = include_widget.findChild(QCheckBox)
                    if include_checkbox and include_checkbox.isChecked():
                        original_item = self.norm_table.item(row_idx, 2)  # Column 2: Original MFG
                        normalize_combo = self.norm_table.cellWidget(row_idx, 3)  # Column 3: Normalize To

                        if original_item and normalize_combo:
                            original_mfg = original_item.text()
                            canonical_mfg = normalize_combo.currentText()

                            # Get selected sheets for this row (if specific sheets were selected)
                            selected_sheets = self.normalization_scopes.get(row_idx, None)

                            # Apply normalization
                            if 'MFG' in new_data.columns:
                                if selected_sheets is None:
                                    # Apply to all records
                                    matches = (new_data['MFG'] == original_mfg).sum()
                                    if matches > 0:
                                        new_data.loc[new_data['MFG'] == original_mfg, 'MFG'] = canonical_mfg
                                        normalizations_applied += 1
                                        print(f"DEBUG: Normalized '{original_mfg}' → '{canonical_mfg}' ({matches} rows, all sheets)")
                                else:
                                    # Apply only to records from selected sheets
                                    if 'Source_Sheet' in new_data.columns:
                                        mask = (new_data['MFG'] == original_mfg) & (new_data['Source_Sheet'].isin(selected_sheets))
                                        matches = mask.sum()
                                        if matches > 0:
                                            new_data.loc[mask, 'MFG'] = canonical_mfg
                                            normalizations_applied += 1
                                            print(f"DEBUG: Normalized '{original_mfg}' → '{canonical_mfg}' ({matches} rows, sheets: {', '.join(selected_sheets)})")

                if normalizations_applied > 0:
                    print(f"DEBUG: Applied {normalizations_applied} manufacturer normalizations")

            # Read existing sheets from output Excel
            with pd.ExcelFile(output_excel) as xls:
                existing_sheets = {sheet: pd.read_excel(output_excel, sheet_name=sheet)
                                 for sheet in xls.sheet_names}

            # CRITICAL: Preserve the original Combined sheet - ensure it stays untouched
            # Use the original data from Step 3 instead of re-reading from Excel
            if hasattr(column_mapping_page, 'combined_data') and column_mapping_page.combined_data is not None:
                existing_sheets['Combined'] = column_mapping_page.combined_data.copy()

            # Add Combined_New sheet with normalized data
            existing_sheets['Combined_New'] = new_data

            # Write back to Excel
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                for sheet_name, df in existing_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Store for comparison page
            self.updated_data = new_data

            return True

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to apply changes: {str(e)}")
            return False


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

