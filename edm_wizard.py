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
        QTabWidget
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


class StartPage(QWizardPage):
    """Start Page: Claude AI API Key and PAS API Configuration"""

    def __init__(self):
        super().__init__()
        self.setTitle("Welcome to EDM Library Wizard")
        self.setSubTitle("Configure API credentials for intelligent column mapping and part search")

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

        # Claude API Key input section
        api_group = QGroupBox("Claude API Configuration")
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

        # PAS API Configuration section
        pas_group = QGroupBox("🔍 Part Aggregation Service (PAS) API Configuration")
        pas_layout = QVBoxLayout()

        pas_info = QLabel(
            "The Part Aggregation Service (PAS) is used to search for parts and get distributor information.\n"
            "Enter your PAS API credentials below."
        )
        pas_info.setWordWrap(True)
        pas_layout.addWidget(pas_info)

        # Client ID
        client_id_layout = QHBoxLayout()
        client_id_label = QLabel("Client ID:")
        client_id_label.setMinimumWidth(100)
        client_id_layout.addWidget(client_id_label)
        self.client_id_input = QLineEdit()
        self.client_id_input.setPlaceholderText("Enter PAS Client ID...")
        self.client_id_input.setMinimumWidth(400)  # Make wider to show full text
        self.client_id_input.textChanged.connect(self.on_pas_credentials_changed)
        client_id_layout.addWidget(self.client_id_input)
        client_id_layout.addStretch()
        pas_layout.addLayout(client_id_layout)

        # Client Secret
        secret_layout = QHBoxLayout()
        secret_label = QLabel("Client Secret:")
        secret_label.setMinimumWidth(100)
        secret_layout.addWidget(secret_label)
        self.client_secret_input = QLineEdit()
        self.client_secret_input.setPlaceholderText("Enter PAS Client Secret...")
        self.client_secret_input.setMinimumWidth(400)  # Make wider to show full text
        self.client_secret_input.setEchoMode(QLineEdit.Password)
        self.client_secret_input.textChanged.connect(self.on_pas_credentials_changed)
        secret_layout.addWidget(self.client_secret_input)

        # Show/Hide button for secret
        self.show_secret_btn = QPushButton("Show")
        self.show_secret_btn.setMaximumWidth(60)
        self.show_secret_btn.clicked.connect(self.toggle_secret_visibility)
        secret_layout.addWidget(self.show_secret_btn)
        secret_layout.addStretch()
        pas_layout.addLayout(secret_layout)

        # Save PAS credentials checkbox
        self.save_pas_checkbox = QCheckBox("Remember PAS credentials for future sessions")
        self.save_pas_checkbox.setChecked(True)
        pas_layout.addWidget(self.save_pas_checkbox)

        # Test PAS connection button
        test_pas_layout = QHBoxLayout()
        self.test_pas_btn = QPushButton("Test PAS Connection")
        self.test_pas_btn.clicked.connect(self.test_pas_credentials)
        self.test_pas_btn.setEnabled(False)
        test_pas_layout.addWidget(self.test_pas_btn)

        self.test_pas_status = QLabel("")
        test_pas_layout.addWidget(self.test_pas_status)
        test_pas_layout.addStretch()
        pas_layout.addLayout(test_pas_layout)

        pas_group.setLayout(pas_layout)
        layout.addWidget(pas_group)

        # SDD_HOME Directory Configuration
        tool_group = QGroupBox("🔧 SDD_HOME Directory")
        tool_layout = QVBoxLayout()

        tool_info = QLabel(
            "Specify the SDD_HOME directory for Siemens EDA tools (optional)."
        )
        tool_info.setWordWrap(True)
        tool_layout.addWidget(tool_info)

        # SDD_HOME path
        sdd_layout = QHBoxLayout()
        sdd_label = QLabel("SDD_HOME:")
        sdd_label.setMinimumWidth(100)
        sdd_layout.addWidget(sdd_label)
        self.mglaunch_input = QLineEdit()
        self.mglaunch_input.setPlaceholderText("C:\\SiemensEDA\\XPED2510\\SDD_HOME")
        self.mglaunch_input.setMinimumWidth(400)
        sdd_layout.addWidget(self.mglaunch_input)

        mglaunch_browse = QPushButton("Browse...")
        mglaunch_browse.clicked.connect(self.browse_mglaunch)
        sdd_layout.addWidget(mglaunch_browse)
        sdd_layout.addStretch()
        tool_layout.addLayout(sdd_layout)

        # Auto-detect button
        detect_layout = QHBoxLayout()
        self.detect_btn = QPushButton("Auto-Detect SDD_HOME")
        self.detect_btn.clicked.connect(self.auto_detect_mglaunch)
        detect_layout.addWidget(self.detect_btn)

        self.detect_status = QLabel("")
        detect_layout.addWidget(self.detect_status)
        detect_layout.addStretch()
        tool_layout.addLayout(detect_layout)

        tool_group.setLayout(tool_layout)
        layout.addWidget(tool_group)

        # Output Settings section
        output_group = QGroupBox("📁 Output Settings")
        output_layout = QVBoxLayout()

        output_info = QLabel(
            "Specify the folder where output files (CSV, Excel) will be saved."
        )
        output_info.setWordWrap(True)
        output_layout.addWidget(output_info)

        # Output folder selection
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("Output Folder:"))
        self.output_folder_input = QLineEdit()
        self.output_folder_input.setPlaceholderText("Select output folder...")
        self.output_folder_input.setReadOnly(True)
        folder_layout.addWidget(self.output_folder_input)

        browse_output_btn = QPushButton("Browse...")
        browse_output_btn.clicked.connect(self.browse_output_folder)
        folder_layout.addWidget(browse_output_btn)

        auto_folder_btn = QPushButton("Auto-Generate")
        auto_folder_btn.setToolTip("Create timestamped output folder in current directory")
        auto_folder_btn.clicked.connect(self.auto_generate_output_folder)
        folder_layout.addWidget(auto_folder_btn)

        output_layout.addLayout(folder_layout)
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        # Skip AI section
        skip_layout = QHBoxLayout()
        skip_layout.addStretch()
        self.skip_ai_btn = QPushButton("Continue without AI")
        self.skip_ai_btn.clicked.connect(self.skip_ai)
        skip_layout.addWidget(self.skip_ai_btn)
        layout.addLayout(skip_layout)

        layout.addStretch()
        self.setLayout(layout)

        # Load saved credentials if available
        self.load_saved_credentials()

        # Store whether APIs are validated
        self.api_validated = False
        self.pas_validated = False
        self.skip_ai_mode = False

    def load_saved_credentials(self):
        """Load API credentials from config file if it exists"""
        config_file = Path.home() / ".edm_wizard_config.json"
        if config_file.exists():
            try:
                with open(config_file, 'r') as f:
                    config = json.load(f)
                    if 'api_key' in config:
                        self.api_key_input.setText(config['api_key'])
                        self.test_status.setText("✓ Loaded saved Claude API key")
                        self.test_status.setStyleSheet("color: green;")
                    if 'client_id' in config:
                        self.client_id_input.setText(config['client_id'])
                    if 'client_secret' in config:
                        self.client_secret_input.setText(config['client_secret'])
                        if config.get('client_id') and config.get('client_secret'):
                            self.test_pas_status.setText("✓ Loaded saved PAS credentials")
                            self.test_pas_status.setStyleSheet("color: green;")
            except Exception as e:
                pass

    def save_credentials(self):
        """Save all credentials to config file"""
        config_file = Path.home() / ".edm_wizard_config.json"
        try:
            config = {}
            if self.save_key_checkbox.isChecked() and self.api_key_input.text().strip():
                config['api_key'] = self.api_key_input.text()
            if self.save_pas_checkbox.isChecked():
                if self.client_id_input.text().strip():
                    config['client_id'] = self.client_id_input.text()
                if self.client_secret_input.text().strip():
                    config['client_secret'] = self.client_secret_input.text()
            
            if config:
                with open(config_file, 'w') as f:
                    json.dump(config, f)
        except Exception as e:
            QMessageBox.warning(self, "Save Error", f"Could not save credentials: {str(e)}")

    def clear_saved_credentials(self):
        """Clear saved credentials from config file"""
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

    def on_pas_credentials_changed(self):
        """Enable test button when PAS credentials are entered"""
        has_both = (len(self.client_id_input.text().strip()) > 0 and 
                   len(self.client_secret_input.text().strip()) > 0)
        self.test_pas_btn.setEnabled(has_both)
        self.pas_validated = False
        self.test_pas_status.setText("")

    def toggle_key_visibility(self):
        """Toggle Claude API key visibility"""
        if self.api_key_input.echoMode() == QLineEdit.Password:
            self.api_key_input.setEchoMode(QLineEdit.Normal)
            self.show_key_btn.setText("Hide")
        else:
            self.api_key_input.setEchoMode(QLineEdit.Password)
            self.show_key_btn.setText("Show")

    def toggle_secret_visibility(self):
        """Toggle PAS client secret visibility"""
        if self.client_secret_input.echoMode() == QLineEdit.Password:
            self.client_secret_input.setEchoMode(QLineEdit.Normal)
            self.show_secret_btn.setText("Hide")
        else:
            self.client_secret_input.setEchoMode(QLineEdit.Password)
            self.show_secret_btn.setText("Show")

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

            # Save credentials if checkbox is checked
            self.save_credentials()

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

    def test_pas_credentials(self):
        """Test the PAS API connection"""
        client_id = self.client_id_input.text().strip()
        client_secret = self.client_secret_input.text().strip()

        if not client_id or not client_secret:
            self.test_pas_status.setText("⚠ Please enter both credentials")
            self.test_pas_status.setStyleSheet("color: orange;")
            return

        self.test_pas_status.setText("Testing connection...")
        self.test_pas_status.setStyleSheet("color: blue;")
        self.test_pas_btn.setEnabled(False)
        QApplication.processEvents()

        try:
            import requests
            import urllib.parse

            # PAS authentication endpoint
            auth_url = "https://samauth.us-east-1.sws.siemens.com/token"
            
            # Use basic auth with client credentials
            auth = (client_id, client_secret)
            
            auth_data = {
                'grant_type': 'client_credentials',
                'scope': 'sws.icarus.api.read'
            }
            
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            response = requests.post(
                auth_url,
                auth=auth,
                data=auth_data,
                headers=headers,
                timeout=10
            )
            response.raise_for_status()
            
            token_data = response.json()
            if 'access_token' in token_data:
                self.pas_validated = True
                self.test_pas_status.setText("✓ Connection successful!")
                self.test_pas_status.setStyleSheet("color: green;")
                
                # Save credentials if checkbox is checked
                self.save_credentials()
            else:
                raise Exception("No access token in response")

        except Exception as e:
            self.pas_validated = False
            error_msg = str(e)
            self.test_pas_status.setText(f"✗ Failed: {error_msg[:50]}...")
            self.test_pas_status.setStyleSheet("color: red;")

            QMessageBox.critical(
                self,
                "PAS Connection Test Failed",
                f"Failed to connect to PAS API:\n\n{error_msg}\n\n"
                "Please check:\n"
                "1. Your Client ID is correct\n"
                "2. Your Client Secret is correct\n"
                "3. You have internet connectivity\n"
                "4. Your credentials have proper permissions"
            )

        self.test_pas_btn.setEnabled(True)

    def browse_mglaunch(self):
        """Browse for SDD_HOME directory"""
        directory = QFileDialog.getExistingDirectory(
            self, "Select SDD_HOME Directory",
            "C:\\SiemensEDA"
        )
        if directory:
            self.mglaunch_input.setText(directory)
            self.detect_status.setText("✓ Directory set manually")
            self.detect_status.setStyleSheet("color: green;")

    def auto_detect_mglaunch(self):
        """Attempt to auto-detect SDD_HOME directory by searching for XPED installations"""
        self.detect_status.setText("Searching...")
        self.detect_status.setStyleSheet("color: blue;")
        self.detect_btn.setEnabled(False)
        QApplication.processEvents()

        # Search for any XPED installation in common root directories
        found_paths = []
        try:
            for root_path in [r"C:\SiemensEDA", r"C:\MentorGraphics", r"C:\Program Files\SiemensEDA", r"C:\Program Files\MentorGraphics"]:
                if os.path.exists(root_path):
                    # Search for directories matching *XPED* pattern
                    for item in os.listdir(root_path):
                        item_path = os.path.join(root_path, item)
                        # Check if it's a directory and contains "XPED" (case-insensitive)
                        if os.path.isdir(item_path) and "XPED" in item.upper():
                            # Check if SDD_HOME subdirectory exists
                            sdd_home_path = os.path.join(item_path, "SDD_HOME")
                            if os.path.exists(sdd_home_path) and os.path.isdir(sdd_home_path):
                                found_paths.append((sdd_home_path, item))
        except Exception as e:
            pass

        # If we found any XPED installations with SDD_HOME, use the first one (or latest version)
        if found_paths:
            # Sort by version number (extract from name) - prefer higher versions
            def extract_version(name):
                # Extract numeric part from names like "XPED2510"
                import re
                match = re.search(r'XPED(\d+)', name.upper())
                return int(match.group(1)) if match else 0

            found_paths.sort(key=lambda x: extract_version(x[1]), reverse=True)
            sdd_path, version_name = found_paths[0]

            self.mglaunch_input.setText(sdd_path)
            self.detect_status.setText(f"✓ Found: {version_name}")
            self.detect_status.setStyleSheet("color: green;")
            self.detect_btn.setEnabled(True)
            return

        # Not found
        self.detect_status.setText("✗ Not found - please browse manually")
        self.detect_status.setStyleSheet("color: orange;")
        self.detect_btn.setEnabled(True)

    def browse_output_folder(self):
        """Browse for output directory"""
        directory = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if directory:
            self.output_folder_input.setText(directory)

    def auto_generate_output_folder(self):
        """Auto-generate timestamped output folder"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_dir = Path.cwd()
        output_folder = base_dir / f"EDM_Output_{timestamp}"
        output_folder.mkdir(exist_ok=True)
        self.output_folder_input.setText(str(output_folder))

    def skip_ai(self):
        """Skip AI features and continue without API key"""
        self.skip_ai_mode = True
        self.wizard().next()

    def validatePage(self):
        """Validate before proceeding to next page"""
        # If skipping AI, always allow
        if self.skip_ai_mode:
            # Save or clear credentials based on checkbox
            if self.save_key_checkbox.isChecked() or self.save_pas_checkbox.isChecked():
                self.save_credentials()
            else:
                self.clear_saved_credentials()
            return True

        # Check if PAS credentials are provided
        has_pas_creds = (self.client_id_input.text().strip() and 
                        self.client_secret_input.text().strip())

        if not has_pas_creds:
            reply = QMessageBox.warning(
                self,
                "PAS Credentials Required",
                "PAS API credentials are required to search for parts.\n\n"
                "Please enter your Client ID and Client Secret, or contact your administrator.",
                QMessageBox.Ok
            )
            return False

        # If PAS credentials entered but not tested
        if has_pas_creds and not self.pas_validated:
            reply = QMessageBox.question(
                self,
                "PAS Credentials Not Tested",
                "You entered PAS credentials but haven't tested them.\n\n"
                "Do you want to continue without testing?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return False

        # If API key is entered but not tested (optional)
        if self.api_key_input.text().strip() and not self.api_validated:
            reply = QMessageBox.question(
                self,
                "Claude API Key Not Tested",
                "You entered a Claude API key but haven't tested it.\n\n"
                "Do you want to continue without testing?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return False

        # Check if output folder is selected
        if not self.output_folder_input.text().strip():
            reply = QMessageBox.warning(
                self,
                "Output Folder Required",
                "Please select an output folder for the results.\n\n"
                "Click 'Browse...' to select a folder or 'Auto-Generate' to create one automatically.",
                QMessageBox.Ok
            )
            return False

        # Save credentials based on checkboxes
        self.save_credentials()

        return True

    def get_api_key(self):
        """Get the entered Claude API key"""
        if self.skip_ai_mode:
            return None
        return self.api_key_input.text().strip() if self.api_key_input.text().strip() else None

    def get_pas_credentials(self):
        """Get the entered PAS credentials"""
        client_id = self.client_id_input.text().strip()
        client_secret = self.client_secret_input.text().strip()
        if client_id and client_secret:
            return {
                'client_id': client_id,
                'client_secret': client_secret
            }
        return None

    def get_output_folder(self):
        """Get the selected output folder"""
        return self.output_folder_input.text().strip()


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
        self.mapping_table.setSortingEnabled(True)  # Enable sorting
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
        self.filter_mfg.setChecked(True)  # Require MFG by default
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

        return True

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
        else:
            # No data after filtering - set empty dataframe
            self.combined_data = pd.DataFrame()


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
        """Initialize and automatically load data from Step 3"""
        # Get data from ColumnMappingPage (Step 3)
        column_mapping_page = self.wizard().page(2)  # ColumnMappingPage is page 2

        # Check if data is available
        if hasattr(column_mapping_page, 'combined_data') and column_mapping_page.combined_data is not None:
            # Check if DataFrame is not empty
            if not column_mapping_page.combined_data.empty:
                # Use the combined data directly from ColumnMappingPage
                self.combined_data = column_mapping_page.combined_data

                # Update info label to show data is loaded
                parts_count = len(self.combined_data)
                self.progress_label.setText(f"✓ Loaded {parts_count} parts from Step 3. Click 'Start Part Search' to begin.")
                self.progress_label.setStyleSheet("color: green; font-weight: bold;")
                self.search_button.setEnabled(True)
            else:
                self.progress_label.setText("⚠ No data available after filtering. Please go back to Step 3 and adjust filter conditions.")
                self.progress_label.setStyleSheet("color: orange;")
                self.search_button.setEnabled(False)
        else:
            self.progress_label.setText("⚠ No data available. Please go back to Step 3 and ensure data is combined.")
            self.progress_label.setStyleSheet("color: orange;")
            self.search_button.setEnabled(False)

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

            # Create PAS client
            pas_client = PASAPIClient(
                client_id=pas_creds['client_id'],
                client_secret=pas_creds['client_secret']
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
            # max_workers=15 means 15 concurrent PAS API calls (adjustable for performance)
            self.search_thread = PASSearchThread(pas_client, parts_list, max_workers=15)
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
            match_details = ', '.join(matches[:3])  # Show first 3 matches
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
        csv_filename = f"SearchAndAssign_Result_{timestamp}.csv"
        self.csv_output_path = self.output_folder / csv_filename

        try:
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
            
            # Write header
            writer.writerow(['PartNumber', 'ManufacturerName', 'MatchStatus', 'MatchValue(PartNumber@ManufacturerName)'])
            
            # Write data
            for result in self.search_results:
                part_number = result['PartNumber']
                manufacturer = result['ManufacturerName']
                status = result['MatchStatus']
                matches = result.get('matches', [])
                
                # Write one row per match (or one row if no matches)
                if matches:
                    for match in matches:
                        writer.writerow([part_number, manufacturer, status, match])
                else:
                    writer.writerow([part_number, manufacturer, status, ''])

    def isComplete(self):
        """Check if search is complete"""
        return self.search_completed


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


class PartialMatchAIThread(QThread):
    """Background thread for AI-powered partial match suggestions"""
    progress = pyqtSignal(str, int, int)  # message, current, total
    part_analyzed = pyqtSignal(int, dict)  # row_idx, analysis_result
    finished = pyqtSignal(dict)  # part_number -> suggested_match_index
    error = pyqtSignal(str)

    def __init__(self, api_key, parts_needing_review, combined_data):
        super().__init__()
        self.api_key = api_key
        self.parts_needing_review = parts_needing_review
        self.combined_data = combined_data

    def run(self):
        try:
            client = Anthropic(api_key=self.api_key)
            suggestions = {}

            total = len(self.parts_needing_review)
            for idx, part in enumerate(self.parts_needing_review):
                # Skip parts with only one match - no AI needed
                if len(part['matches']) <= 1:
                    self.progress.emit(f"⏭️ Skipping part {idx + 1} of {total} (only one match)...", idx + 1, total)
                    # Still mark as processed
                    self.part_analyzed.emit(idx, {'skipped': True, 'reason': 'single_match'})
                    continue

                self.progress.emit(f"🤖 Analyzing part {idx + 1} of {total}...", idx, total)

                # Get original description from combined data
                description = self.get_description_for_part(part['PartNumber'], part['ManufacturerName'])

                # Create prompt for AI
                matches_text = "\n".join([f"{i+1}. {m}" for i, m in enumerate(part['matches'])])

                prompt = f"""Analyze this electronic component and suggest the best matching part number from SupplyFrame.

Original Part:
- Part Number: {part['PartNumber']}
- Manufacturer: {part['ManufacturerName']}
- Description: {description if description else 'Not available'}

Available Matches from SupplyFrame:
{matches_text}

Instructions:
1. Compare the original part number with each match
2. Consider manufacturer variations (e.g., "EPCOS" vs "TDK Electronics")
3. Look for exact or closest part number matches
4. If the manufacturer has been acquired, prefer the current company name

Return a JSON object with:
{{
    "suggested_index": <0-based index of best match, or null if none are suitable>,
    "confidence": <0-100>,
    "reasoning": "<brief explanation>"
}}

Only return the JSON, no other text."""

                try:
                    response = client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=500,
                        messages=[{"role": "user", "content": prompt}]
                    )

                    response_text = response.content[0].text.strip()
                    if response_text.startswith('```'):
                        response_text = response_text.split('```')[1]
                        if response_text.startswith('json'):
                            response_text = response_text[4:]
                        response_text = response_text.strip()

                    result = json.loads(response_text)
                    suggestions[part['PartNumber']] = result

                    # Emit per-part update for real-time UI refresh
                    self.part_analyzed.emit(idx, result)

                except Exception as e:
                    # If AI fails for this part, emit error result
                    self.part_analyzed.emit(idx, {'error': str(e)})
                    continue

            self.finished.emit(suggestions)

        except Exception as e:
            self.error.emit(str(e))

    def get_description_for_part(self, part_number, mfg):
        """Find description from combined data"""
        for row in self.combined_data:
            if row.get('MFG_PN') == part_number and row.get('MFG') == mfg:
                return row.get('Description', '')
        return ''


class ManufacturerNormalizationAIThread(QThread):
    """Background thread for hybrid fuzzy+AI manufacturer normalization"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(dict, dict)  # (normalizations, reasoning_map)
    error = pyqtSignal(str)

    def __init__(self, api_key, all_manufacturers, supplyframe_manufacturers):
        super().__init__()
        self.api_key = api_key
        self.all_manufacturers = all_manufacturers
        self.supplyframe_manufacturers = supplyframe_manufacturers

    def run(self):
        try:
            # Phase 1: Fuzzy matching pre-filter
            self.progress.emit("🔍 Phase 1: Fuzzy matching analysis...")
            fuzzy_matches = {}
            reasoning_map = {}

            if FUZZYWUZZY_AVAILABLE and self.supplyframe_manufacturers:
                for user_mfg in self.all_manufacturers:
                    # Find best match in SupplyFrame manufacturers
                    best_match = process.extractOne(
                        user_mfg,
                        self.supplyframe_manufacturers,
                        scorer=fuzz.token_sort_ratio
                    )

                    if best_match:
                        canonical, score = best_match[0], best_match[1]

                        # High confidence matches (>90%): auto-accept
                        if score >= 90 and user_mfg != canonical:
                            fuzzy_matches[user_mfg] = canonical
                            reasoning_map[user_mfg] = {
                                'method': 'fuzzy',
                                'score': score,
                                'reasoning': f"High confidence fuzzy match ({score}% similarity)"
                            }
                        # Medium confidence matches (70-89%): send to AI for validation
                        elif 70 <= score < 90 and user_mfg != canonical:
                            # Will be validated by AI in phase 2
                            pass

            self.progress.emit(f"✓ Phase 1 complete: {len(fuzzy_matches)} high-confidence matches found")

            # Phase 2: AI analysis for ambiguous cases
            if ANTHROPIC_AVAILABLE and self.api_key:
                self.progress.emit("🤖 Phase 2: AI validation of ambiguous cases...")

                # Collect manufacturers that weren't auto-matched by fuzzy
                unmatched_mfgs = [m for m in self.all_manufacturers if m not in fuzzy_matches]

                if unmatched_mfgs:
                    client = Anthropic(api_key=self.api_key)

                    # Create prompt for AI to analyze remaining manufacturers
                    prompt = f"""Analyze these manufacturer names and detect variations that should be normalized.

Manufacturers needing analysis:
{json.dumps(sorted(unmatched_mfgs), indent=2)}

SupplyFrame canonical manufacturer names (prefer these):
{json.dumps(sorted(self.supplyframe_manufacturers), indent=2)}

Instructions:
1. Identify manufacturer name variations (e.g., "TI" vs "Texas Instruments")
2. Detect abbreviations, acquired companies, and alternate spellings
3. Map each variation to the canonical SupplyFrame name when available
4. For companies not in SupplyFrame, suggest the most complete/official name
5. IMPORTANT: For each mapping, provide a brief reasoning (acquisitions, abbreviations, etc.)
6. CRITICAL: Ensure all string values are properly escaped for JSON (escape quotes, newlines, backslashes)

Return ONLY valid JSON with this structure:
{{
    "normalizations": {{
        "<variation>": "<canonical_name>",
        "TI": "Texas Instruments",
        "EPCOS": "TDK Electronics"
    }},
    "reasoning": {{
        "TI": "Common abbreviation for Texas Instruments",
        "EPCOS": "EPCOS was acquired by TDK Electronics in 2009"
    }}
}}

IMPORTANT:
- Only include entries that need normalization
- Return ONLY valid JSON, no markdown, no other text
- Ensure all quotes inside strings are escaped with backslash"""

                    response = client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=2048,
                        messages=[{"role": "user", "content": prompt}]
                    )

                    response_text = response.content[0].text.strip()

                    # Clean up code blocks
                    if response_text.startswith('```'):
                        # Extract content between code blocks
                        parts = response_text.split('```')
                        if len(parts) >= 2:
                            response_text = parts[1]
                            # Remove 'json' language identifier if present
                            if response_text.startswith('json'):
                                response_text = response_text[4:]
                            response_text = response_text.strip()

                    # Try to parse JSON with better error handling
                    try:
                        ai_result = json.loads(response_text)
                    except json.JSONDecodeError as je:
                        # Log the error and try to extract what we can
                        self.progress.emit(f"⚠️ JSON parse error at char {je.pos}: {je.msg}")

                        # Fallback: Try to find JSON object in the response
                        import re
                        json_match = re.search(r'\{[\s\S]*\}', response_text)
                        if json_match:
                            try:
                                ai_result = json.loads(json_match.group())
                            except:
                                # If all parsing fails, return empty results for AI phase
                                self.progress.emit("⚠️ Could not parse AI response, using fuzzy matches only")
                                ai_result = {"normalizations": {}, "reasoning": {}}
                        else:
                            ai_result = {"normalizations": {}, "reasoning": {}}

                    ai_normalizations = ai_result.get('normalizations', {})
                    ai_reasoning = ai_result.get('reasoning', {})

                    # Merge AI results with fuzzy matches
                    for variation, canonical in ai_normalizations.items():
                        fuzzy_matches[variation] = canonical
                        reasoning_map[variation] = {
                            'method': 'ai',
                            'reasoning': ai_reasoning.get(variation, 'AI suggested normalization')
                        }

                    if ai_normalizations:
                        self.progress.emit(f"✓ Phase 2 complete: {len(ai_normalizations)} AI-validated matches")
                    else:
                        self.progress.emit("✓ Phase 2 complete: No additional AI matches")

            # Emit combined results
            self.finished.emit(fuzzy_matches, reasoning_map)

        except Exception as e:
            self.error.emit(str(e))


class PASSearchThread(QThread):
    """Background thread for searching parts via PAS API with parallel execution"""
    progress = pyqtSignal(str, int, int)  # message, current, total
    result_ready = pyqtSignal(dict)  # individual result for real-time display
    finished = pyqtSignal(list)  # all search results
    error = pyqtSignal(str)

    def __init__(self, pas_client, parts_data, max_workers=10):
        super().__init__()
        self.pas_client = pas_client
        self.parts_data = parts_data  # List of {'MFG': ..., 'MFG_PN': ..., 'Description': ...}
        self.max_workers = max_workers  # Number of parallel threads
        self.completed_count = 0
        self.lock = threading.Lock()

    def search_single_part(self, idx, part, total):
        """Search a single part with retry logic"""
        manufacturer = part.get('MFG', '')
        part_number = part.get('MFG_PN', '')

        # Handle NaN values from pandas (convert to empty string)
        import math
        if isinstance(manufacturer, float) and math.isnan(manufacturer):
            manufacturer = ''
        if isinstance(part_number, float) and math.isnan(part_number):
            part_number = ''

        # Convert to string and strip whitespace
        manufacturer = str(manufacturer).strip() if manufacturer else ''
        part_number = str(part_number).strip() if part_number else ''

        if not manufacturer or not part_number:
            with self.lock:
                self.completed_count += 1
                self.progress.emit(f"Skipping part {self.completed_count}/{total} (missing MFG or Manufacturer PN)...", self.completed_count, total)
            return {
                'PartNumber': part_number if part_number else '(empty)',
                'ManufacturerName': manufacturer if manufacturer else '(empty)',
                'MatchStatus': 'None',
                'matches': []
            }

        with self.lock:
            self.completed_count += 1
            current = self.completed_count

        self.progress.emit(
            f"Searching Manufacturer PN {current}/{total}: {manufacturer} - {part_number}...",
            current,
            total
        )

        # Search with retry logic (like SearchAndAssignApp - 3 retries)
        match_result = None
        match_type = None
        retry_count = 0
        max_retries = 3

        while retry_count < max_retries:
            try:
                match_result, match_type = self.pas_client.search_part(part_number, manufacturer)
                break  # Success
            except Exception as e:
                retry_count += 1
                if retry_count < max_retries:
                    self.progress.emit(
                        f"Retry {retry_count}/{max_retries} for {manufacturer} {part_number}...",
                        current,
                        total
                    )
                    time.sleep(3)  # Wait 3 seconds before retry
                else:
                    match_result = {'error': str(e)}
                    match_type = 'Error'

        # Map match_type to status (using SearchAndAssign terminology)
        if match_type in ['Found', 'Multiple', 'Need user review', 'None', 'Error']:
            status = match_type
        else:
            # Legacy mapping for backwards compatibility
            if match_type == 'exact':
                status = 'Found'
            elif match_type == 'partial':
                matches = match_result.get('matches', [])
                if len(matches) > 1:
                    status = 'Multiple'
                elif len(matches) == 1:
                    status = 'Found'
                else:
                    status = 'None'
            elif match_type == 'no_match':
                status = 'None'
            else:  # error
                status = 'Error'

        result_dict = {
            'PartNumber': part_number,
            'ManufacturerName': manufacturer,
            'MatchStatus': status,
            'matches': match_result.get('matches', []) if match_type != 'Error' else []
        }

        # Emit individual result for real-time display
        self.result_ready.emit(result_dict)

        return result_dict

    def run(self):
        try:
            from concurrent.futures import ThreadPoolExecutor, as_completed

            results = [None] * len(self.parts_data)  # Pre-allocate to maintain order
            total = len(self.parts_data)
            self.completed_count = 0

            # Use ThreadPoolExecutor for parallel execution
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                # Submit all tasks
                future_to_idx = {
                    executor.submit(self.search_single_part, idx, part, total): idx
                    for idx, part in enumerate(self.parts_data)
                }

                # Collect results as they complete
                for future in as_completed(future_to_idx):
                    idx = future_to_idx[future]
                    try:
                        result = future.result()
                        results[idx] = result
                    except Exception as e:
                        # Handle unexpected errors
                        self.progress.emit(f"Error processing part {idx + 1}: {str(e)}", idx + 1, total)
                        results[idx] = {
                            'PartNumber': self.parts_data[idx].get('MFG_PN', ''),
                            'ManufacturerName': self.parts_data[idx].get('MFG', ''),
                            'MatchStatus': 'Error',
                            'matches': []
                        }

            self.finished.emit(results)

        except Exception as e:
            self.error.emit(str(e))


class PASAPIClient:
    """Part Aggregation Service API Client"""
    
    def __init__(self, client_id, client_secret):
        """Initialize PAS API client with credentials"""
        self.client_id = client_id
        self.client_secret = client_secret
        self.pas_url = "https://api.pas.partquest.com"
        self.auth_url = "https://samauth.us-east-1.sws.siemens.com/token"
        self.access_token = None
        self.token_expires_at = None
        
    def _get_access_token(self):
        """Get or refresh the access token"""
        if self.access_token and self.token_expires_at:
            if datetime.now() < self.token_expires_at:
                return self.access_token
        
        # Request new token
        auth = (self.client_id, self.client_secret)
        auth_data = {
            'grant_type': 'client_credentials',
            'scope': 'sws.icarus.api.read'
        }
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        
        response = requests.post(
            self.auth_url,
            auth=auth,
            data=auth_data,
            headers=headers,
            timeout=10
        )
        response.raise_for_status()
        
        token_data = response.json()
        self.access_token = token_data['access_token']
        expires_in = token_data.get('expires_in', 7200)
        self.token_expires_at = datetime.now() + timedelta(seconds=expires_in - 60)
        
        return self.access_token
    
    def search_part(self, manufacturer_pn, manufacturer):
        """
        Search for a part using PAS API with SearchAndAssign matching algorithm

        Implements the exact matching logic from SearchAndAssignApp.java:
        1. Exact match: PartNumber AND ManufacturerName both exact
        2. Partial manufacturer match: PartNumber exact, ManufacturerName contains
        3. Alphanumeric-only match: Strip special chars, compare alphanumeric only
        4. Leading zero suppression: Remove leading zeros and compare
        5. PartNumber-only search: If manufacturer empty/unknown

        Returns: (result_dict, match_type)
        match_type: 'Found', 'Multiple', 'Need user review', 'None', or 'Error'
        """
        import re

        try:
            # Perform PAS search
            search_results = self._perform_pas_search(manufacturer_pn, manufacturer)

            if 'error' in search_results:
                return {'error': search_results['error']}, 'Error'

            parts = search_results.get('results', [])

            if not parts:
                return {'matches': []}, 'None'

            # Apply SearchAndAssign matching algorithm
            match_result = self._apply_searchandassign_matching(
                manufacturer_pn, manufacturer, parts
            )

            return match_result

        except Exception as e:
            return {'error': str(e)}, 'Error'

    def _perform_pas_search(self, manufacturer_pn, manufacturer):
        """Perform the actual PAS API search"""
        try:
            token = self._get_access_token()

            # Search endpoint
            endpoint = '/api/v2/search-providers/44/2/free-text/search'

            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json',
                'X-Siemens-Correlation-Id': f'corr-{int(time.time() * 1000)}',
                'X-Siemens-Session-Id': f'session-{int(time.time())}',
                'X-Siemens-Ebs-User-Country-Code': 'US',
                'X-Siemens-Ebs-User-Currency': 'USD'
            }

            # Combine manufacturer and part number for more accurate search
            # This matches how Java SearchAndAssign searches with both parameters
            if manufacturer and manufacturer.strip():
                search_term = f"{manufacturer_pn} {manufacturer}"
            else:
                search_term = manufacturer_pn

            request_body = {
                "ftsParameters": {
                    "match": {
                        "term": search_term
                    },
                    "paging": {
                        "requestedPageSize": 20
                    }
                }
            }

            url = f"{self.pas_url}{endpoint}"
            response = requests.post(
                url,
                headers=headers,
                json=request_body,
                timeout=60
            )

            if response.status_code == 401:
                # Token expired, retry once
                self.access_token = None
                self.token_expires_at = None
                token = self._get_access_token()
                headers['Authorization'] = f'Bearer {token}'
                response = requests.post(
                    url,
                    headers=headers,
                    json=request_body,
                    timeout=60
                )

            response.raise_for_status()
            result = response.json()

            if not result.get('success', False):
                error = result.get('error', {})
                error_msg = error.get('message', 'Unknown error')
                return {'error': error_msg}

            # Return results list
            if result.get('result') and result['result'].get('results'):
                return {
                    'results': result['result']['results'],
                    'totalCount': result['result'].get('totalCount', 0)
                }
            else:
                return {'results': []}

        except Exception as e:
            return {'error': str(e)}

    def _apply_searchandassign_matching(self, edm_pn, edm_mfg, parts):
        """
        Apply the exact SearchAndAssign matching algorithm from Java code

        Algorithm steps (matching SearchAndAssignApp.java):
        1. Search by PartNumber + ManufacturerName
           a. Exact match on both
           b. Partial match on ManufacturerName (contains)
           c. Alphanumeric-only match
           d. Leading zero suppression match
        2. If no match and manufacturer is empty/Unknown, search by PartNumber only
        """
        import re

        matches = []

        # Step 1: Search with both PartNumber and ManufacturerName
        if edm_mfg and edm_mfg not in ['', 'Unknown']:
            # Step 1a: Exact match on both PartNumber and ManufacturerName
            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')
                pas_mfg = part.get('manufacturerName', '')

                if pas_pn == edm_pn and pas_mfg == edm_mfg:
                    matches.append(part_data)

            if len(matches) > 1:
                return self._format_match_result(matches, 'Multiple')
            elif len(matches) == 1:
                return self._format_match_result(matches, 'Found')

            # Step 1b: Partial match on ManufacturerName (PN exact, MFG contains)
            matches.clear()
            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')
                pas_mfg = part.get('manufacturerName', '')

                if pas_pn == edm_pn and edm_mfg in pas_mfg:
                    matches.append(part_data)

            if len(matches) > 1:
                return self._format_match_result(matches, 'Multiple')
            elif len(matches) == 1:
                return self._format_match_result(matches, 'Found')

            # Step 1c: Alphanumeric-only match (strip special characters)
            matches.clear()
            pattern = re.compile(r'[^A-Za-z0-9]')
            edm_pn_alpha = pattern.sub('', edm_pn)

            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')
                pas_pn_alpha = pattern.sub('', pas_pn)

                if pas_pn_alpha == edm_pn_alpha:
                    matches.append(part_data)

            if len(matches) == 0:
                # Step 1d: Leading zero suppression
                edm_pn_no_zeros = edm_pn_alpha.lstrip('0')

                for part_data in parts:
                    part = part_data.get('searchProviderPart', {})
                    pas_pn = part.get('manufacturerPartNumber', '')
                    pas_pn_alpha = pattern.sub('', pas_pn)
                    pas_pn_no_zeros = pas_pn_alpha.lstrip('0')

                    if pas_pn_no_zeros == edm_pn_no_zeros:
                        matches.append(part_data)

                if len(matches) == 1:
                    return self._format_match_result(matches, 'Found')
                elif len(matches) > 1:
                    # Multiple matches - take first one
                    return self._format_match_result([matches[0]], 'Found')
            else:
                if len(matches) == 1:
                    return self._format_match_result(matches, 'Found')
                else:
                    # Multiple matches - take first one
                    return self._format_match_result([matches[0]], 'Found')

        # Step 2: Search by PartNumber only (if manufacturer empty/Unknown or no matches found)
        if not edm_mfg or edm_mfg in ['', 'Unknown'] or len(matches) == 0:
            matches.clear()

            # Exact PartNumber match
            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')

                if pas_pn == edm_pn:
                    matches.append(part_data)

            if len(matches) == 0:
                # Try alphanumeric-only
                pattern = re.compile(r'[^A-Za-z0-9]')
                edm_pn_alpha = pattern.sub('', edm_pn)

                for part_data in parts:
                    part = part_data.get('searchProviderPart', {})
                    pas_pn = part.get('manufacturerPartNumber', '')
                    pas_pn_alpha = pattern.sub('', pas_pn)

                    if pas_pn_alpha == edm_pn_alpha:
                        matches.append(part_data)

                if len(matches) == 0:
                    # Try leading zero suppression
                    edm_pn_no_zeros = edm_pn_alpha.lstrip('0')

                    for part_data in parts:
                        part = part_data.get('searchProviderPart', {})
                        pas_pn = part.get('manufacturerPartNumber', '')
                        pas_pn_alpha = pattern.sub('', pas_pn)
                        pas_pn_no_zeros = pas_pn_alpha.lstrip('0')

                        if pas_pn_no_zeros == edm_pn_no_zeros:
                            matches.append(part_data)

                    if len(matches) == 0:
                        # Partial matches - return all as Multiple
                        return self._format_match_result(parts, 'Multiple')
                    elif len(matches) == 1:
                        return self._format_match_result(matches, 'Need user review')
                    else:
                        # Multiple matches - take first one
                        return self._format_match_result([matches[0]], 'Found')
                else:
                    if len(matches) == 1:
                        return self._format_match_result(matches, 'Need user review')
                    else:
                        # Multiple matches - take first one
                        return self._format_match_result([matches[0]], 'Found')
            else:
                if len(matches) == 1:
                    return self._format_match_result(matches, 'Need user review')
                else:
                    return self._format_match_result(matches, 'Multiple')

        # No matches found
        return {'matches': []}, 'None'

    def _format_match_result(self, part_data_list, match_type):
        """Format the match result in a consistent way"""
        matches = []
        for part_data in part_data_list:
            part = part_data.get('searchProviderPart', {})
            mpn = part.get('manufacturerPartNumber', '')
            mfg = part.get('manufacturerName', '')
            matches.append(f"{mpn}@{mfg}")

        return {'matches': matches[:10]}, match_type


class SupplyFrameReviewPage(QWizardPage):
    """Step 5: Review PAS Matches and Normalize Manufacturers"""

    def __init__(self):
        super().__init__()
        self.setTitle("Step 5: Review Matches & Manufacturer Normalization")
        self.setSubTitle("Review match results (Found/Multiple/Need Review/None) and normalize manufacturer names")

        self.search_results = []
        self.parts_needing_review = []
        self.manufacturer_normalizations = {}
        self.normalization_reasoning = {}  # Store fuzzy/AI reasoning for each normalization
        self.original_data = []  # Store original data for comparison
        self.api_key = None

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

    def initializePage(self):
        """Initialize by loading data from PASSearchPage"""
        pas_search_page = self.wizard().page(3)  # PASSearchPage is page 3

        # Get search results from PASSearchPage
        if hasattr(pas_search_page, 'search_results') and pas_search_page.search_results:
            self.search_results = pas_search_page.search_results

            # Store original data for comparison later (convert DataFrame to list of dicts)
            if hasattr(pas_search_page, 'combined_data') and pas_search_page.combined_data is not None:
                if not pas_search_page.combined_data.empty:
                    # Convert DataFrame to list of dictionaries for easier processing
                    self.original_data = pas_search_page.combined_data.to_dict('records')
                else:
                    self.original_data = []
            else:
                self.original_data = []

            # Load and display the results
            self.load_search_results()
        else:
            QMessageBox.warning(
                self,
                "No Data",
                "No search results available.\n\n"
                "Please go back to Step 4 and complete the PAS search."
            )

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

        # Populate review tables
        self.populate_review_tables(found, multiple, need_review, none, errors)

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

    def populate_review_tables(self, found, multiple, need_review, none, errors):
        """Populate the review tables with categorized results"""
        # This will be called from load_search_results()
        # Placeholder for now - we'll implement the table population
        pass

    def identify_normalization_candidates(self):
        """Identify manufacturers that need normalization"""
        # Placeholder - will detect variations in manufacturer names
        pass

    def create_review_section_widget(self):
        """Section 2: Review Matches by Category"""
        review_group = QGroupBox("🔍 Review Match Results")
        review_layout = QVBoxLayout()

        # Create tab widget for different match categories
        self.review_tabs = QTabWidget()

        # Tab 1: Found (exact matches)
        self.found_table = self.create_results_table()
        self.review_tabs.addTab(self.found_table, "✓ Found")

        # Tab 2: Multiple matches
        self.multiple_table = self.create_results_table()
        self.review_tabs.addTab(self.multiple_table, "⚠ Multiple")

        # Tab 3: Need user review
        self.need_review_table = self.create_results_table()
        self.review_tabs.addTab(self.need_review_table, "👁 Need Review")

        # Tab 4: None (no matches)
        self.none_table = self.create_results_table()
        self.review_tabs.addTab(self.none_table, "✗ None")

        # Tab 5: Errors (optional - only if there are errors)
        self.errors_table = self.create_results_table()
        self.review_tabs.addTab(self.errors_table, "❌ Errors")

        review_layout.addWidget(self.review_tabs)
        review_group.setLayout(review_layout)
        return review_group

    def create_results_table(self):
        """Create a table widget for displaying match results"""
        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Part Number", "Manufacturer", "Match Status", "Match Details"])

        # Set column resize modes
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)  # Part Number
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Manufacturer
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Match Status
        header.setSectionResizeMode(3, QHeaderView.Stretch)  # Match Details

        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(True)  # Enable sorting

        return table

    def create_review_section_widget_OLD(self):
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
        self.ai_normalize_btn = QPushButton("🤖 AI Detect Manufacturer Variations")
        self.ai_normalize_btn.clicked.connect(self.ai_detect_normalizations)
        self.ai_normalize_btn.setEnabled(False)
        ai_layout.addWidget(self.ai_normalize_btn)

        self.norm_status = QLabel("")
        ai_layout.addWidget(self.norm_status)
        ai_layout.addStretch()
        norm_layout.addLayout(ai_layout)

        # Normalization table
        self.norm_table = QTableWidget()
        self.norm_table.setColumnCount(4)
        self.norm_table.setHorizontalHeaderLabels(["Include", "Original MFG", "Normalize To", "Scope"])
        self.norm_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.norm_table.customContextMenuRequested.connect(self.show_normalization_context_menu)
        self.norm_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        norm_layout.addWidget(self.norm_table)

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
        self.old_data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        old_layout.addWidget(self.old_data_table)

        # New data
        new_widget = QWidget()
        new_layout = QVBoxLayout(new_widget)
        new_layout.addWidget(QLabel("Updated Data:"))
        self.new_data_table = QTableWidget()
        self.new_data_table.setColumnCount(3)
        self.new_data_table.setHorizontalHeaderLabels(["MFG", "MFG PN", "Description"])
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

        self.regenerate_xml_btn = QPushButton("Regenerate XML Files")
        self.regenerate_xml_btn.clicked.connect(self.regenerate_xml)
        self.regenerate_xml_btn.setEnabled(False)
        actions_layout.addWidget(self.regenerate_xml_btn)

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
            for row in xml_gen_page.combined_data:
                if row.get('MFG'):
                    all_mfgs.add(row['MFG'])

        # From SearchAndAssign data
        for part in self.search_assign_data:
            # Original manufacturers
            if part.get('ManufacturerName'):
                all_mfgs.add(part['ManufacturerName'])

            # SupplyFrame manufacturers from matches
            for match in part.get('matches', []):
                if '@' in match:
                    _, mfg = match.split('@', 1)
                    supplyframe_mfgs.add(mfg)

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
        selected_rows = self.parts_list.selectedIndexes()
        if not selected_rows:
            return

        row_idx = selected_rows[0].row()
        part = self.parts_needing_review[row_idx]

        # Populate matches table
        self.matches_table.setRowCount(len(part['matches']))

        # Calculate similarity scores for confidence
        from difflib import SequenceMatcher
        original_pn = part['PartNumber'].upper().strip()

        for match_idx, match in enumerate(part['matches']):
            # Parse match: "PartNumber@Manufacturer"
            if '@' in match:
                pn, mfg = match.split('@', 1)
            else:
                pn = match
                mfg = ""

            # Radio button for selection - centered in cell
            radio = QRadioButton()
            if part.get('selected_match') == match:
                radio.setChecked(True)
            radio.toggled.connect(lambda checked, p=part, m=match: self.on_match_selected(p, m, checked))

            # Create a widget to center the radio button
            radio_widget = QWidget()
            radio_layout = QHBoxLayout(radio_widget)
            radio_layout.addWidget(radio)
            radio_layout.setAlignment(Qt.AlignCenter)
            radio_layout.setContentsMargins(0, 0, 0, 0)

            self.matches_table.setCellWidget(match_idx, 0, radio_widget)
            self.matches_table.setItem(match_idx, 1, QTableWidgetItem(pn))
            self.matches_table.setItem(match_idx, 2, QTableWidgetItem(mfg))

            # Calculate similarity score
            match_pn = pn.upper().strip()
            similarity = SequenceMatcher(None, original_pn, match_pn).ratio()
            similarity_pct = int(similarity * 100)
            similarity_item = QTableWidgetItem(f"{similarity_pct}%")
            similarity_item.setTextAlignment(Qt.AlignCenter)
            similarity_item.setToolTip("String similarity using difflib (part number matching)")
            self.matches_table.setItem(match_idx, 3, similarity_item)

            # AI Score - only show if AI has processed this part
            ai_score_item = QTableWidgetItem("")
            ai_score_item.setTextAlignment(Qt.AlignCenter)
            if part.get('ai_processed') and part.get('ai_match_scores'):
                # Get AI confidence for this specific match
                ai_scores = part.get('ai_match_scores', {})
                if match in ai_scores:
                    ai_conf = ai_scores[match]
                    ai_score_item.setText(f"{ai_conf}%")
                    ai_score_item.setToolTip("AI confidence score (considers context, manufacturer, description)")
            self.matches_table.setItem(match_idx, 4, ai_score_item)

    def on_match_selected(self, part, match, checked):
        """Handle match selection"""
        if checked:
            part['selected_match'] = match
            self.none_correct_checkbox.setChecked(False)
            self.save_selection_btn.setEnabled(True)

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

        from difflib import SequenceMatcher
        original_pn = part['PartNumber'].upper().strip()

        for match_idx, match in enumerate(part['matches']):
            # Parse match
            if '@' in match:
                pn, mfg = match.split('@', 1)
            else:
                pn = match
                mfg = ""

            # Radio button
            radio = QRadioButton()
            if part.get('selected_match') == match:
                radio.setChecked(True)
            radio.toggled.connect(lambda checked, p=part, m=match: self.on_match_selected(p, m, checked))

            radio_widget = QWidget()
            radio_layout = QHBoxLayout(radio_widget)
            radio_layout.addWidget(radio)
            radio_layout.setAlignment(Qt.AlignCenter)
            radio_layout.setContentsMargins(0, 0, 0, 0)

            self.matches_table.setCellWidget(match_idx, 0, radio_widget)
            self.matches_table.setItem(match_idx, 1, QTableWidgetItem(pn))
            self.matches_table.setItem(match_idx, 2, QTableWidgetItem(mfg))

            # Similarity score
            match_pn = pn.upper().strip()
            similarity = SequenceMatcher(None, original_pn, match_pn).ratio()
            similarity_pct = int(similarity * 100)
            similarity_item = QTableWidgetItem(f"{similarity_pct}%")
            similarity_item.setTextAlignment(Qt.AlignCenter)
            similarity_item.setToolTip("String similarity using difflib (part number matching)")
            self.matches_table.setItem(match_idx, 3, similarity_item)

            # AI Score - show if available
            ai_score_item = QTableWidgetItem("")
            ai_score_item.setTextAlignment(Qt.AlignCenter)
            if part.get('ai_processed') and part.get('ai_match_scores'):
                ai_scores = part.get('ai_match_scores', {})
                if match in ai_scores:
                    ai_conf = ai_scores[match]
                    ai_score_item.setText(f"{ai_conf}%")
                    ai_score_item.setToolTip("AI confidence score (considers context, manufacturer, description)")
            self.matches_table.setItem(match_idx, 4, ai_score_item)

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

        # Parse match
        if '@' in selected_match:
            pn, mfg = selected_match.split('@', 1)
        else:
            pn = selected_match
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
                # Parse match: "PartNumber@Manufacturer"
                if '@' in match:
                    match_pn, match_mfg = match.split('@', 1)
                else:
                    match_pn = match
                    match_mfg = ""

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
        if row_idx < len(self.parts_needing_review):
            part = self.parts_needing_review[row_idx]

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

                if idx is not None and 0 <= idx < len(part['matches']):
                    suggested_match = part['matches'][idx]
                    part['selected_match'] = suggested_match
                    part['ai_confidence'] = confidence
                    part['ai_reasoning'] = result.get('reasoning', '')
                    # Store the AI score for this specific match
                    part['ai_match_scores'][suggested_match] = confidence

            # Update UI for this specific row
            self.update_part_row(row_idx)

            # Auto-refresh matches table if this is the currently selected part
            selected_rows = self.parts_list.selectedIndexes()
            if selected_rows and selected_rows[0].row() == row_idx:
                self.refresh_matches_display()

            # Update review count
            self.update_review_count()

    def on_ai_match_finished(self, suggestions):
        """Apply AI suggestions"""
        applied = 0
        for part in self.parts_needing_review:
            pn = part['PartNumber']
            if pn in suggestions:
                suggestion = suggestions[pn]
                idx = suggestion.get('suggested_index')
                if idx is not None and 0 <= idx < len(part['matches']):
                    part['selected_match'] = part['matches'][idx]
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

        # Collect all manufacturers
        all_mfgs = set()
        supplyframe_mfgs = set()

        # From original data (ALL manufacturers from Step 3)
        xml_gen_page = self.wizard().page(3)
        if hasattr(xml_gen_page, 'combined_data'):
            for row in xml_gen_page.combined_data:
                if row.get('MFG'):
                    all_mfgs.add(row['MFG'])

        # From SearchAndAssign CSV - collect ALL manufacturers from ALL matches
        # (Not just selected matches, but all possible matches from SupplyFrame)
        for part in self.search_assign_data:
            # Add original manufacturer from the CSV
            if part.get('ManufacturerName'):
                all_mfgs.add(part['ManufacturerName'])

            # Add ALL manufacturers from matches (not just selected)
            for match in part.get('matches', []):
                if '@' in match:
                    _, mfg = match.split('@', 1)
                    supplyframe_mfgs.add(mfg)

        self.norm_status.setText("🤖 Analyzing manufacturers...")
        self.norm_status.setStyleSheet("color: blue;")
        self.ai_normalize_btn.setEnabled(False)

        # Start AI thread
        self.ai_norm_thread = ManufacturerNormalizationAIThread(
            self.api_key,
            list(all_mfgs),
            list(supplyframe_mfgs)
        )
        self.ai_norm_thread.progress.connect(lambda msg: self.norm_status.setText(msg))
        self.ai_norm_thread.finished.connect(self.on_ai_norm_finished)
        self.ai_norm_thread.error.connect(self.on_ai_norm_error)
        self.ai_norm_thread.start()

    def on_ai_norm_finished(self, normalizations, reasoning_map):
        """Apply hybrid fuzzy+AI normalization suggestions"""
        self.manufacturer_normalizations = normalizations
        self.normalization_reasoning = reasoning_map  # Store reasoning for context menu

        # Collect all unique manufacturers from both sources
        all_mfgs = set()

        # From original data
        xml_gen_page = self.wizard().page(3)
        if hasattr(xml_gen_page, 'combined_data'):
            for row in xml_gen_page.combined_data:
                if row.get('MFG'):
                    all_mfgs.add(row['MFG'])

        # From SearchAndAssign (SupplyFrame canonical names)
        for part in self.search_assign_data:
            if part.get('selected_match') and '@' in part['selected_match']:
                _, mfg = part['selected_match'].split('@', 1)
                all_mfgs.add(mfg)

        # From normalization suggestions
        for original, canonical in normalizations.items():
            all_mfgs.add(original)
            all_mfgs.add(canonical)

        # Sort manufacturers for easier selection
        sorted_mfgs = sorted(list(all_mfgs))

        # Populate normalization table
        self.norm_table.setRowCount(len(normalizations))

        row_idx = 0
        for original, canonical in normalizations.items():
            # Include checkbox
            include_cb = QCheckBox()
            include_cb.setChecked(True)
            self.norm_table.setCellWidget(row_idx, 0, include_cb)

            # Original MFG (read-only)
            self.norm_table.setItem(row_idx, 1, QTableWidgetItem(original))

            # Normalize To (editable combo box)
            normalize_combo = QComboBox()
            normalize_combo.setEditable(True)
            normalize_combo.addItems(sorted_mfgs)
            normalize_combo.setCurrentText(canonical)
            self.norm_table.setCellWidget(row_idx, 2, normalize_combo)

            # Scope dropdown
            scope_combo = QComboBox()
            scope_combo.addItems(["All Catalogs", "Per Catalog"])
            self.norm_table.setCellWidget(row_idx, 3, scope_combo)

            row_idx += 1

        # Count fuzzy vs AI matches
        fuzzy_count = sum(1 for v in reasoning_map.values() if v.get('method') == 'fuzzy')
        ai_count = sum(1 for v in reasoning_map.values() if v.get('method') == 'ai')

        self.norm_status.setText(
            f"✓ Found {len(normalizations)} variations "
            f"({fuzzy_count} fuzzy, {ai_count} AI-validated)"
        )
        self.norm_status.setStyleSheet("color: green; font-weight: bold;")
        self.ai_normalize_btn.setEnabled(True)
        self.save_normalizations_btn.setEnabled(True)

        QMessageBox.information(self, "Normalization Detected",
                              f"Hybrid analysis complete!\n\n"
                              f"• {fuzzy_count} high-confidence fuzzy matches\n"
                              f"• {ai_count} AI-validated matches\n"
                              f"• Total: {len(normalizations)} normalizations\n\n"
                              f"Right-click any row to see detection reasoning.\n"
                              f"Review and adjust as needed.")

    def on_ai_norm_error(self, error_msg):
        """Handle AI normalization error"""
        self.norm_status.setText(f"✗ Error: {error_msg[:30]}")
        self.norm_status.setStyleSheet("color: red;")
        self.ai_normalize_btn.setEnabled(True)

        QMessageBox.critical(self, "AI Error", f"AI normalization failed:\n{error_msg}")

    def apply_changes(self):
        """Apply all changes and generate comparison"""
        try:
            # Get the combined data from XMLGenerationPage (Step 3)
            prev_page_3 = self.wizard().page(3)  # XMLGenerationPage
            if not hasattr(prev_page_3, 'combined_data') or prev_page_3.combined_data is None or prev_page_3.combined_data.empty:
                QMessageBox.warning(self, "No Data",
                                  "No combined data available from Step 3.\n"
                                  "Please complete Step 3 first.")
                return

            # Create a copy of the original data
            import copy
            old_data = copy.deepcopy(prev_page_3.combined_data)
            new_data = copy.deepcopy(prev_page_3.combined_data)

            # Track changes for summary
            matches_applied = 0
            normalizations_applied = 0

            # Step 1: Apply selected partial matches
            for part_data in self.search_assign_data:
                if 'selected_match' in part_data and part_data['selected_match']:
                    # Parse the selected match: "PartNumber@ManufacturerName"
                    match_str = part_data['selected_match']
                    if '@' in match_str:
                        new_pn, new_mfg = match_str.split('@', 1)

                        # Find and update all matching records in new_data
                        original_pn = part_data['PartNumber']
                        original_mfg = part_data['ManufacturerName']

                        for record in new_data:
                            if (record['MFG_PN'] == original_pn and
                                record['MFG'] == original_mfg):
                                record['MFG_PN'] = new_pn.strip()
                                record['MFG'] = new_mfg.strip()
                                matches_applied += 1

            # Step 2: Apply manufacturer normalizations
            for row_idx in range(self.norm_table.rowCount()):
                # Check if this normalization is included
                include_checkbox = self.norm_table.cellWidget(row_idx, 0)
                if not include_checkbox or not include_checkbox.isChecked():
                    continue

                variation_item = self.norm_table.item(row_idx, 1)
                canonical_combo = self.norm_table.cellWidget(row_idx, 2)
                scope_combo = self.norm_table.cellWidget(row_idx, 3)

                if not variation_item or not canonical_combo or not scope_combo:
                    continue

                variation = variation_item.text().strip()
                canonical = canonical_combo.currentText().strip()
                scope = scope_combo.currentText()

                # Apply normalization based on scope
                if scope == "All Catalogs":
                    # Apply to all records with this manufacturer variation
                    for record in new_data:
                        if record['MFG'] == variation:
                            record['MFG'] = canonical
                            normalizations_applied += 1
                else:
                    # "Per Catalog" - only apply within this catalog
                    # Since we're working with a single dataset, treat same as "All Catalogs"
                    # In a multi-catalog scenario, you'd filter by catalog ID here
                    for record in new_data:
                        if record['MFG'] == variation:
                            record['MFG'] = canonical
                            normalizations_applied += 1

            # Step 3: Populate comparison tables
            self.populate_comparison_tables(old_data, new_data)

            # Step 4: Enable XML regeneration button
            self.regenerate_xml_btn.setEnabled(True)

            # Step 5: Store the new data for XML generation
            self.updated_data = new_data

            # Show summary
            QMessageBox.information(self, "Changes Applied",
                                  f"Changes applied successfully!\n\n"
                                  f"• {matches_applied} parts updated from SupplyFrame matches\n"
                                  f"• {normalizations_applied} manufacturer names normalized\n\n"
                                  f"Review the comparison below and regenerate XML when ready.")

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

        # Get the original manufacturer name from this row
        original_item = self.norm_table.item(row, 1)
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
            orig_item = self.norm_table.item(row_idx, 1)
            if orig_item and orig_item.text() == original_mfg:
                canonical_combo = self.norm_table.cellWidget(row_idx, 2)
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
            include_checkbox = self.norm_table.cellWidget(row_idx, 0)
            if include_checkbox and include_checkbox.isChecked():
                enabled_count += 1

        QMessageBox.information(self, "Normalizations Saved",
                              f"Manufacturer normalization settings saved!\n\n"
                              f"• {enabled_count} of {len(self.manufacturer_normalizations)} normalizations enabled\n"
                              f"• Settings will be applied when you click 'Apply Changes'\n\n"
                              f"You can continue editing the normalizations as needed.")


class ComparisonPage(QWizardPage):
    """Step 6: Old vs New Comparison - Show changes made"""

    def __init__(self):
        super().__init__()
        self.setTitle("Step 6: Review Changes")
        self.setSubTitle("Review all changes made to manufacturer names and match selections")

        layout = QVBoxLayout()

        # Summary section
        summary_group = QGroupBox("📊 Changes Summary")
        summary_layout = QVBoxLayout()

        self.summary_label = QLabel("No changes to display")
        self.summary_label.setWordWrap(True)
        summary_layout.addWidget(self.summary_label)

        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)

        # Comparison table
        comparison_group = QGroupBox("🔄 Old vs New Comparison")
        comparison_layout = QVBoxLayout()

        self.comparison_table = QTableWidget()
        self.comparison_table.setColumnCount(5)
        self.comparison_table.setHorizontalHeaderLabels([
            "Part Number",
            "Original MFG",
            "New MFG",
            "Change Type",
            "Notes"
        ])

        # Set column resize modes
        header = self.comparison_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)  # Part Number
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Original MFG
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # New MFG
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Change Type
        header.setSectionResizeMode(4, QHeaderView.Stretch)  # Notes

        self.comparison_table.setAlternatingRowColors(True)
        self.comparison_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.comparison_table.setSortingEnabled(True)  # Enable sorting

        comparison_layout.addWidget(self.comparison_table)
        comparison_group.setLayout(comparison_layout)
        layout.addWidget(comparison_group, stretch=1)

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
        self.original_data = []
        self.normalized_data = []
        self.changes = []

    def initializePage(self):
        """Initialize by loading data from SupplyFrameReviewPage"""
        review_page = self.wizard().page(4)  # SupplyFrameReviewPage is page 4

        # Get original and normalized data
        if hasattr(review_page, 'original_data'):
            self.original_data = review_page.original_data

        if hasattr(review_page, 'manufacturer_normalizations'):
            self.normalizations = review_page.manufacturer_normalizations
        else:
            self.normalizations = {}

        # Build comparison
        self.build_comparison()

    def build_comparison(self):
        """Build the comparison between original and normalized data"""
        self.changes = []

        # Ensure original_data is a list of dictionaries
        if not hasattr(self, 'original_data') or self.original_data is None:
            self.original_data = []

        # If original_data is a DataFrame, convert to list of dicts
        if hasattr(self.original_data, 'to_dict'):
            self.original_data = self.original_data.to_dict('records')

        # Compare original data with normalizations
        for orig_item in self.original_data:
            if not isinstance(orig_item, dict):
                continue  # Skip non-dict items

            orig_mfg = orig_item.get('MFG', '')
            part_num = orig_item.get('MFG_PN', '')

            # Check if this manufacturer was normalized
            if orig_mfg in self.normalizations:
                new_mfg = self.normalizations[orig_mfg]
                if new_mfg != orig_mfg:
                    self.changes.append({
                        'part_number': part_num,
                        'original_mfg': orig_mfg,
                        'new_mfg': new_mfg,
                        'change_type': 'Normalized',
                        'notes': f'Manufacturer name standardized'
                    })

        # Update display
        self.update_comparison_display()

    def update_comparison_display(self):
        """Update the comparison table with changes"""
        # Update summary
        total_parts = len(self.original_data)
        changed_parts = len(self.changes)

        if changed_parts > 0:
            summary_text = f"""
<b>Total Parts:</b> {total_parts}<br>
<b>Parts Changed:</b> {changed_parts} ({changed_parts/total_parts*100:.1f}%)<br>
<b>Parts Unchanged:</b> {total_parts - changed_parts} ({(total_parts-changed_parts)/total_parts*100:.1f}%)
"""
            self.summary_label.setText(summary_text)
        else:
            self.summary_label.setText("✓ No changes were made to the data")

        # Populate comparison table
        self.comparison_table.setRowCount(len(self.changes))

        for row, change in enumerate(self.changes):
            # Part Number
            self.comparison_table.setItem(row, 0, QTableWidgetItem(change['part_number']))

            # Original MFG
            orig_item = QTableWidgetItem(change['original_mfg'])
            orig_item.setBackground(QColor(255, 230, 230))  # Light red
            self.comparison_table.setItem(row, 1, orig_item)

            # New MFG
            new_item = QTableWidgetItem(change['new_mfg'])
            new_item.setBackground(QColor(230, 255, 230))  # Light green
            self.comparison_table.setItem(row, 2, new_item)

            # Change Type
            self.comparison_table.setItem(row, 3, QTableWidgetItem(change['change_type']))

            # Notes
            self.comparison_table.setItem(row, 4, QTableWidgetItem(change['notes']))

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
        self.start_page = StartPage()                          # Step 1: API credentials & output folder
        self.data_source_page = DataSourcePage()               # Step 2: Access DB export or Excel selection
        self.column_mapping_page = ColumnMappingPage()         # Step 3: Column mapping & combine
        self.pas_search_page = PASSearchPage()                 # Step 4: PAS API search
        self.review_page = SupplyFrameReviewPage()             # Step 5: Review matches & normalization
        self.comparison_page = ComparisonPage()                # Step 6: Old vs New comparison

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
