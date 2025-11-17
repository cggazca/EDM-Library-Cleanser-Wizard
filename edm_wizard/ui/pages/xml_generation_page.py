"""
XML Generation Page: Legacy XML Output
"""

import sys
import os
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import json
import threading
import time
import xml.etree.ElementTree as ET
from xml.dom import minidom
import requests

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QLineEdit,
        QTableWidget, QTableWidgetItem, QHeaderView,
        QPushButton, QMessageBox, QWidget, QScrollArea, QCheckBox,
        QTextEdit, QFileDialog, QDialog, QDialogButtonBox, QSplitter,
        QGridLayout, QApplication
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
    from PyQt5.QtGui import QColor, QFont
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

try:
    import sqlalchemy as sa
    from sqlalchemy import inspect
    SQLALCHEMY_AVAILABLE = True
except ImportError:
    SQLALCHEMY_AVAILABLE = False

try:
    from anthropic import Anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

from edm_wizard.utils.xml_generation import escape_xml



class XMLGenerationPage(QWizardPage):
    """Step 3: Generate XML files (DEPRECATED - replaced by PASSearchPage)"""

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
        prev_page = self.wizard().page(1)  # DataSourcePage is page 1
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

            prev_page_0 = self.wizard().page(1)  # DataSourcePage is page 1
            prev_page_1 = self.wizard().page(2)  # ColumnMappingPage is page 2

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
        prev_page_1 = self.wizard().page(2)  # ColumnMappingPage is now page 2
        included_sheets = prev_page_1.get_included_sheets()

        all_mfg = set()
        all_mfgpn = []
        self.combined_data = []

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
                data_row = {
                    'MFG': str(row['MFG']).strip(),
                    'MFG_PN': str(row['MFG_PN']).strip(),
                    'Description': str(row['Description']) if pd.notna(row['Description']) else "This is the PN description."
                }
                all_mfgpn.append(data_row)
                self.combined_data.append(data_row)

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

        # Collect MFG/MFGPN pairs and store combined data
        self.combined_data = []
        for _, row in df_copy.iterrows():
            if pd.notna(row[mfg_col]) and pd.notna(row[mfgpn_col]):
                desc = row[desc_col] if desc_col and pd.notna(row[desc_col]) else "This is the PN description."
                data_row = {
                    'MFG': str(row[mfg_col]).strip(),
                    'MFG_PN': str(row[mfgpn_col]).strip(),
                    'Description': str(desc)
                }
                all_mfgpn.append(data_row)
                self.combined_data.append(data_row)

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

