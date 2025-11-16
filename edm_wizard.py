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
        QTabWidget, QButtonGroup, QSpinBox
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

        # Model selection
        model_layout = QHBoxLayout()
        model_layout.addWidget(QLabel("Claude Model:"))
        self.model_selector = QComboBox()
        self.model_selector.addItem("Claude Sonnet 4.5 (Latest, Recommended)", "claude-sonnet-4-5-20250929")
        self.model_selector.addItem("Claude Haiku 4.5 (Fastest)", "claude-haiku-4-5-20251001")
        self.model_selector.addItem("Claude Opus 4.1 (Most Capable)", "claude-opus-4-1-20250805")
        self.model_selector.setCurrentIndex(0)  # Default to Sonnet 4.5
        model_layout.addWidget(self.model_selector)
        model_layout.addStretch()
        api_layout.addLayout(model_layout)

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

        # Advanced Settings section
        advanced_group = QGroupBox("⚙️ Advanced Settings")
        advanced_layout = QVBoxLayout()

        # Max matches per part setting
        max_matches_layout = QHBoxLayout()
        max_matches_label = QLabel("Max Matches per Part:")
        max_matches_label.setMinimumWidth(150)
        max_matches_label.setToolTip(
            "Maximum number of matches to display for parts with multiple matches.\n"
            "Lower values make the UI faster but may hide some options.\n"
            "Higher values show more options but may slow down the UI."
        )
        max_matches_layout.addWidget(max_matches_label)
        
        self.max_matches_spinner = QSpinBox()
        self.max_matches_spinner.setMinimum(5)
        self.max_matches_spinner.setMaximum(100)
        self.max_matches_spinner.setValue(10)  # Default to 10
        self.max_matches_spinner.setSuffix(" matches")
        self.max_matches_spinner.setToolTip(
            "Recommended: 10-25 for best performance\n"
            "Default: 10 (original behavior)"
        )
        max_matches_layout.addWidget(self.max_matches_spinner)
        max_matches_layout.addStretch()
        advanced_layout.addLayout(max_matches_layout)

        advanced_group.setLayout(advanced_layout)
        layout.addWidget(advanced_group)

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

    def get_selected_model(self):
        """Get the selected Claude model"""
        return self.model_selector.currentData()  # Returns the model ID

    def get_max_matches(self):
        """Get the maximum number of matches to display per part"""
        return self.max_matches_spinner.value()


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


class SQLiteExportThread(QThread):
    """Background thread for exporting SQLite database"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(str, object)  # excel_path, dataframes_dict
    error = pyqtSignal(str)

    def __init__(self, sqlite_file, output_file):
        super().__init__()
        self.sqlite_file = sqlite_file
        self.output_file = output_file

    def run(self):
        try:
            import sqlite3

            self.progress.emit("Connecting to SQLite database...")

            # Connect to SQLite database
            conn = sqlite3.connect(self.sqlite_file)
            cursor = conn.cursor()

            # Get table names
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';")
            tables = [row[0] for row in cursor.fetchall()]

            if not tables:
                self.error.emit("No tables found in SQLite database.")
                conn.close()
                return

            self.progress.emit(f"Found {len(tables)} tables. Exporting...")

            # Export all tables
            dataframes = {}
            with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
                for idx, table in enumerate(tables, 1):
                    self.progress.emit(f"Exporting table {idx}/{len(tables)}: {table}")

                    # Read table data (SQLite uses double quotes for identifiers)
                    df = pd.read_sql_query(f'SELECT * FROM "{table}"', conn)

                    # Clean sheet name
                    sheet_name = self.clean_sheet_name(table)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    dataframes[sheet_name] = df

            conn.close()
            self.progress.emit("Export completed successfully!")
            self.finished.emit(self.output_file, dataframes)

        except Exception as e:
            self.error.emit(f"Error exporting SQLite database: {str(e)}")

    @staticmethod
    def clean_sheet_name(name):
        """Clean Excel sheet names"""
        invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
        for char in invalid_chars:
            name = name.replace(char, '')
        return name[:31]


class SheetDetectionWorker(QThread):
    """Worker thread for detecting columns in a single sheet"""
    finished = pyqtSignal(str, dict)  # sheet_name, mapping
    error = pyqtSignal(str, str)  # sheet_name, error_msg

    def __init__(self, api_key, sheet_name, dataframe, model="claude-sonnet-4-5-20250929", max_retries=5):
        super().__init__()
        self.api_key = api_key
        self.sheet_name = sheet_name
        self.dataframe = dataframe
        self.model = model
        self.max_retries = max_retries

    def run(self):
        import time
        retry_count = 0
        base_delay = 10  # Start with 10 second delay
        
        while retry_count <= self.max_retries:
            try:
                client = Anthropic(api_key=self.api_key)

                # Prepare column information
                columns = self.dataframe.columns.tolist()

                # Filter out rows that are mostly empty (less than 30% of columns have data)
                min_fields_threshold = max(2, len(columns) * 0.3)
                non_empty_counts = self.dataframe.notna().sum(axis=1)
                df_filtered = self.dataframe[non_empty_counts >= min_fields_threshold].copy()

                if len(df_filtered) == 0:
                    df_filtered = self.dataframe.copy()

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
                    'total_rows': len(self.dataframe),
                    'rows_with_data': len(df_filtered),
                    'non_empty_counts': {}
                }

                for col in columns:
                    non_empty = df_filtered[col].notna().sum()
                    stats['non_empty_counts'][col] = non_empty

                sheet_info = {
                    'sheet_name': self.sheet_name,
                    'columns': columns,
                    'sample_data': sample_rows,
                    'statistics': stats
                }

                # Create prompt for Claude
                prompt = f"""Analyze the following Excel sheet and its columns. Identify which columns correspond to:
1. MFG (Manufacturer name) - Look for manufacturer names like "Siemens", "ABB", "Schneider", etc.
2. MFG_PN (Manufacturer Part Number) - The primary part number from the manufacturer
3. MFG_PN_2 (Secondary/alternative Manufacturer Part Number) - An alternative or backup part number
4. Part_Number (Internal part number) - Internal reference numbers
5. Description (Part description) - Text description of the part

Here is the sheet with its columns, sample data (up to 50 rows), and statistics:

{json.dumps(sheet_info, indent=2, default=str)}

Note: Rows with little to no information (less than 30% of columns filled) have been filtered out.

Analyze the sample data carefully. Look at:
- Column names (they might have hints like "Mfg", "Manufacturer", "PN", "Part", "Description", etc.)
- Data patterns (manufacturer names vs part numbers vs descriptions)
- Data completeness (statistics show total_rows, rows_with_data after filtering, and non_empty_counts per column)
- Data consistency across the sample rows

Return a JSON object with the mapping and confidence scores (0-100). Base confidence on:
- How well the column name matches the expected field
- How consistent the data pattern is with the expected field type
- How complete the data is (columns with mostly empty values should have lower confidence)

Format:
{{
  "{self.sheet_name}": {{
    "MFG": {{"column": "column_name or null", "confidence": 0-100}},
    "MFG_PN": {{"column": "column_name or null", "confidence": 0-100}},
    "MFG_PN_2": {{"column": "column_name or null", "confidence": 0-100}},
    "Part_Number": {{"column": "column_name or null", "confidence": 0-100}},
    "Description": {{"column": "column_name or null", "confidence": 0-100}}
  }}
}}

Only return the JSON, no other text."""

                # Call Claude API with selected model
                response = client.messages.create(
                    model=self.model,
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

                mapping = json.loads(response_text)

                # Emit the mapping for this sheet
                if self.sheet_name in mapping:
                    self.finished.emit(self.sheet_name, mapping[self.sheet_name])
                else:
                    self.error.emit(self.sheet_name, "Sheet mapping not found in response")
                
                # Success - exit retry loop
                break

            except Exception as e:
                error_str = str(e)
                
                # Check if it's a rate limit error (429)
                is_rate_limit = '429' in error_str or 'rate_limit' in error_str.lower() or 'overloaded' in error_str.lower()
                
                if is_rate_limit and retry_count < self.max_retries:
                    # Calculate exponential backoff delay
                    delay = base_delay * (2 ** retry_count)  # 10s, 20s, 40s, 80s, 160s
                    retry_count += 1
                    
                    # Sleep and retry
                    time.sleep(delay)
                    continue  # Retry the request
                else:
                    # Not a rate limit error, or max retries reached
                    if retry_count >= self.max_retries:
                        self.error.emit(self.sheet_name, f"Max retries ({self.max_retries}) exceeded. Last error: {error_str}")
                    else:
                        self.error.emit(self.sheet_name, error_str)
                    break


class AIDetectionThread(QThread):
    """Coordinator thread for parallel AI column detection across all sheets"""
    progress = pyqtSignal(str, int, int)  # message, current, total
    finished = pyqtSignal(dict)  # mappings
    error = pyqtSignal(str)

    def __init__(self, api_key, dataframes, model="claude-sonnet-4-5-20250929"):
        super().__init__()
        self.api_key = api_key
        self.dataframes = dataframes
        self.model = model
        self.all_mappings = {}
        self.completed_count = 0
        self.error_count = 0
        self.workers = []

    def run(self):
        try:
            sheet_names = list(self.dataframes.keys())
            total_sheets = len(sheet_names)

            self.progress.emit(f"🔄 Starting parallel analysis of {total_sheets} sheets...", 0, total_sheets)

            # Create a worker for each sheet
            for sheet_name in sheet_names:
                worker = SheetDetectionWorker(
                    self.api_key,
                    sheet_name,
                    self.dataframes[sheet_name],
                    self.model
                )
                worker.finished.connect(self.on_sheet_completed)
                worker.error.connect(self.on_sheet_error)
                self.workers.append(worker)

            # Start workers with staggered delays to avoid rate limiting
            # Conservative approach: process one at a time with longer delays
            import time
            batch_size = 1  # Process one sheet at a time to avoid rate limits
            delay_between_requests = 12.0  # 12 second delay between requests (safe for most API tiers)

            for i in range(0, len(self.workers), batch_size):
                batch = self.workers[i:i + batch_size]

                # Start workers in this batch
                for worker in batch:
                    worker.start()

                # Wait for this batch to complete before starting next
                for worker in batch:
                    worker.wait()

                # If not the last batch, wait before starting next request
                if i + batch_size < len(self.workers):
                    self.progress.emit(
                        f"⏳ Rate limit protection: waiting {delay_between_requests}s before next request...",
                        self.completed_count,
                        total_sheets
                    )
                    time.sleep(delay_between_requests)

            # All workers have already completed (waited in the loop above)
            # No need to wait again

            # Check if we got at least some results
            if len(self.all_mappings) > 0:
                # Report summary including failures
                success_count = len(self.all_mappings)
                failed_count = self.error_count

                if failed_count > 0:
                    # Build error report
                    failed_list = getattr(self, 'failed_sheets', [])
                    error_details = "\n".join([f"  - {item['sheet']}: {item['error'][:80]}" for item in failed_list[:10]])
                    if len(failed_list) > 10:
                        error_details += f"\n  ... and {len(failed_list) - 10} more"

                    self.progress.emit(
                        f"⚠️ Completed with {failed_count} errors. Successfully mapped {success_count}/{total_sheets} sheets.",
                        total_sheets,
                        total_sheets
                    )
                else:
                    self.progress.emit("✅ All sheets mapped successfully!", total_sheets, total_sheets)

                self.finished.emit(self.all_mappings)
            else:
                self.error.emit("No sheets were successfully analyzed. Please check your API key and try again.")

        except Exception as e:
            self.error.emit(str(e))

    def on_sheet_completed(self, sheet_name, mapping):
        """Handle completion of a single sheet detection"""
        self.all_mappings[sheet_name] = mapping
        self.completed_count += 1
        total = len(self.dataframes)
        self.progress.emit(
            f"🤖 Completed {self.completed_count}/{total} sheets ('{sheet_name}')",
            self.completed_count,
            total
        )

    def on_sheet_error(self, sheet_name, error_msg):
        """Handle error from a single sheet detection"""
        self.error_count += 1
        self.completed_count += 1

        # Track failed sheet
        if not hasattr(self, 'failed_sheets'):
            self.failed_sheets = []
        self.failed_sheets.append({'sheet': sheet_name, 'error': error_msg})

        total = len(self.dataframes)
        self.progress.emit(
            f"⚠️ Error on sheet '{sheet_name}': {error_msg[:50]}... ({self.completed_count}/{total})",
            self.completed_count,
            total
        )


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
        self.file_path.setPlaceholderText("Select Access DB (.mdb/.accdb), SQLite DB (.db/.sqlite/.sqlite3), or Excel (.xlsx/.xls)...")
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
            "All Supported Files (*.mdb *.accdb *.db *.sqlite *.sqlite3 *.xlsx *.xls);;"
            "Access Database (*.mdb *.accdb);;"
            "SQLite Database (*.db *.sqlite *.sqlite3);;"
            "Excel Files (*.xlsx *.xls);;"
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
        elif self.detected_file_type == 'excel':
            # Excel files just need to be loaded
            return len(self.dataframes) > 0
        return False

    def get_excel_path(self):
        """Get the Excel file path"""
        if self.detected_file_type in ['access', 'sqlite']:
            return self.exported_excel_path
        elif self.detected_file_type == 'excel':
            return self.file_path.text()
        return None

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
            # max_workers=15 means 15 concurrent PAS API calls (adjustable for performance)
            self.search_thread = PASSearchThread(self.pas_client, parts_list, max_workers=15)
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
        csv_filename = f"PAS_Search_Results_{timestamp}.csv"
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

        # Only require part_number (MFG can be empty)
        if not part_number:
            with self.lock:
                self.completed_count += 1
                self.progress.emit(f"Skipping part {self.completed_count}/{total} (missing Manufacturer PN)...", self.completed_count, total)
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
    
    def __init__(self, client_id, client_secret, max_matches=10):
        """Initialize PAS API client with credentials"""
        self.client_id = client_id
        self.client_secret = client_secret
        self.pas_url = "https://api.pas.partquest.com"
        self.auth_url = "https://samauth.us-east-1.sws.siemens.com/token"
        self.access_token = None
        self.token_expires_at = None
        self.max_matches = max_matches  # Maximum matches to display per part
        
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
        """
        Perform PAS API parametric search (matches Java AggregationServiceWebCall.searchExactMatchSF)

        Uses parametric/search endpoint with property-based filters instead of free-text search.
        This provides more accurate results by filtering on specific properties:
        - 6230417e = Manufacturer Name
        - d8ac8dcc = Manufacturer Part Number
        """
        try:
            token = self._get_access_token()

            # Parametric search endpoint (not free-text!)
            endpoint = '/api/v2/search-providers/44/2/parametric/search'

            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json',
                'X-Siemens-Correlation-Id': f'corr-{int(time.time() * 1000)}',
                'X-Siemens-Session-Id': f'session-{int(time.time())}',
                'X-Siemens-Ebs-User-Country-Code': 'US',
                'X-Siemens-Ebs-User-Currency': 'USD'
            }

            # Build filter based on whether manufacturer is provided
            if manufacturer and manufacturer.strip():
                # Two-parameter search: AND filter for both Part Number and Manufacturer
                # Matches Java line 510 in AggregationServiceWebCall.java
                search_filter = {
                    "__logicalOperator__": "And",
                    "__expression__": "LogicalExpression",
                    "left": {
                        "__valueOperator__": "SmartMatch",
                        "__expression__": "ValueExpression",
                        "propertyId": "6230417e",  # Manufacturer Name
                        "term": manufacturer
                    },
                    "right": {
                        "__valueOperator__": "SmartMatch",
                        "__expression__": "ValueExpression",
                        "propertyId": "d8ac8dcc",  # Manufacturer Part Number
                        "term": manufacturer_pn
                    }
                }
                page_size = 10  # Java uses 10 for two-parameter search
            else:
                # One-parameter search: filter by Part Number only
                # Matches Java line 550 in AggregationServiceWebCall.java
                search_filter = {
                    "__valueOperator__": "SmartMatch",
                    "__expression__": "ValueExpression",
                    "propertyId": "d8ac8dcc",  # Manufacturer Part Number
                    "term": manufacturer_pn
                }
                page_size = 50  # Java uses 50 for one-parameter search

            request_body = {
                "searchParameters": {
                    "partClassId": "76f2225d",  # Root part class
                    "customParameters": {},
                    "outputs": ["6230417e", "d8ac8dcc", "750a45c8", "2a2b1476", "e1aa6f26"],
                    "sort": [],
                    "paging": {
                        "requestedPageSize": page_size
                    },
                    "filter": search_filter
                }
            }

            # Collect all results (handle pagination like Java does)
            all_results = []
            url = f"{self.pas_url}{endpoint}"

            while True:
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

                # Add results from this page
                if result.get('result') and result['result'].get('results'):
                    all_results.extend(result['result']['results'])

                # Check for next page (Java fetches ALL pages)
                next_page_token = result.get('result', {}).get('nextPageToken')
                if not next_page_token:
                    break

                # Prepare next page request
                endpoint = '/api/v2/search-providers/44/2/parametric/get-next-page'
                url = f"{self.pas_url}{endpoint}"
                request_body = {
                    "pageToken": next_page_token
                }

            return {
                'results': all_results,
                'totalCount': len(all_results)
            }

        except Exception as e:
            return {'error': str(e)}

    def _apply_searchandassign_matching(self, edm_pn, edm_mfg, parts):
        """
        Apply the exact SearchAndAssign matching algorithm from Java code (SearchAndAssignApp.java lines 291-540)

        Mirrors Java implementation precisely:
        - If manufacturer is NOT empty/Unknown: Try exact + partial + alphanumeric + zero suppression WITH manufacturer validation
        - If no match OR manufacturer IS empty/Unknown: Search by PartNumber only (returns "Need user review" for single matches)
        """
        import re

        pattern = re.compile(r'[^A-Za-z0-9]')
        matches = []
        result_record = None

        # ========== STEP 1: Search with Manufacturer (if provided and not empty/Unknown) ==========
        if edm_mfg and edm_mfg not in ['', 'Unknown']:
            # 1a. Exact match on BOTH PartNumber AND ManufacturerName (Java lines 317-338)
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

            # 1b. Partial match on ManufacturerName (Java lines 341-364)
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

            # 1c. Alphanumeric-only match (Java lines 366-415)
            # MUST also check manufacturer - we're still in Step 1 (with manufacturer)
            matches.clear()
            edm_pn_alpha = pattern.sub('', edm_pn)
            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')
                pas_mfg = part.get('manufacturerName', '')
                pas_pn_alpha = pattern.sub('', pas_pn)
                # Check both part number AND manufacturer (partial match OK)
                if pas_pn_alpha == edm_pn_alpha and (pas_mfg == edm_mfg or edm_mfg in pas_mfg):
                    matches.append(part_data)

            if len(matches) == 0:
                # 1d. Leading zero suppression (Java lines 378-415)
                # MUST also check manufacturer - we're still in Step 1 (with manufacturer)
                edm_pn_no_zeros = edm_pn_alpha.lstrip('0')
                for part_data in parts:
                    part = part_data.get('searchProviderPart', {})
                    pas_pn = part.get('manufacturerPartNumber', '')
                    pas_mfg = part.get('manufacturerName', '')
                    pas_pn_alpha = pattern.sub('', pas_pn)
                    pas_pn_no_zeros = pas_pn_alpha.lstrip('0')
                    # Check both part number AND manufacturer (partial match OK)
                    if pas_pn_no_zeros == edm_pn_no_zeros and (pas_mfg == edm_mfg or edm_mfg in pas_mfg):
                        matches.append(part_data)

                if len(matches) == 1:
                    return self._format_match_result(matches, 'Found')
                elif len(matches) > 1:
                    return self._format_match_result([matches[0]], 'Found')
            else:
                if len(matches) >= 1:
                    return self._format_match_result([matches[0]], 'Found')

        # ========== STEP 2: Search by PartNumber only (Java lines 419-540) ==========
        # Triggered if: manufacturer is empty/Unknown OR no matches found in Step 1
        if result_record is None and (not edm_mfg or edm_mfg in ['', 'Unknown'] or len(matches) == 0):
            matches.clear()
            all_results = list(parts)  # Keep reference to all results for fallback

            if len(parts) == 0:
                return {'matches': []}, 'None'

            # Special case: If exactly 1 result from PAS search (Java lines 433-442)
            if len(parts) == 1:
                return self._format_match_result(parts, 'Need user review')

            # Multiple results from PAS - try to narrow down by PartNumber (Java lines 444-540)
            # 2a. Exact PartNumber match (Java lines 447-448)
            for part_data in parts:
                part = part_data.get('searchProviderPart', {})
                pas_pn = part.get('manufacturerPartNumber', '')
                if pas_pn == edm_pn:
                    matches.append(part_data)

            if len(matches) == 0:
                # 2b. Alphanumeric-only match (Java lines 451-521)
                edm_pn_alpha = pattern.sub('', edm_pn)
                for part_data in parts:
                    part = part_data.get('searchProviderPart', {})
                    pas_pn = part.get('manufacturerPartNumber', '')
                    pas_pn_alpha = pattern.sub('', pas_pn)
                    if pas_pn_alpha == edm_pn_alpha:
                        matches.append(part_data)

                if len(matches) == 0:
                    # 2c. Leading zero suppression (Java lines 463-521)
                    edm_pn_no_zeros = edm_pn_alpha.lstrip('0')
                    for part_data in parts:
                        part = part_data.get('searchProviderPart', {})
                        pas_pn = part.get('manufacturerPartNumber', '')
                        pas_pn_alpha = pattern.sub('', pas_pn)
                        pas_pn_no_zeros = pas_pn_alpha.lstrip('0')
                        if pas_pn_no_zeros == edm_pn_no_zeros:
                            matches.append(part_data)

                    if len(matches) == 0:
                        # No matches - return all as Multiple (Java lines 474-483)
                        return self._format_match_result(all_results, 'Multiple')
                    elif len(matches) == 1:
                        # Java line 487
                        return self._format_match_result(matches, 'Need user review')
                    else:
                        # Multiple matches - take first (Java lines 494-499)
                        return self._format_match_result([matches[0]], 'Found')
                else:
                    if len(matches) == 1:
                        # Java line 503
                        return self._format_match_result(matches, 'Need user review')
                    elif len(matches) > 1:
                        # Java lines 511-519
                        return self._format_match_result(matches, 'Multiple')
            else:
                if len(matches) == 1:
                    # Java line 525
                    return self._format_match_result(matches, 'Need user review')
                elif len(matches) > 1:
                    # Java lines 532-540 - Multiple exact matches
                    return self._format_match_result(matches, 'Multiple')

        # No matches found (fallback)
        return {'matches': []}, 'None'

    def _format_match_result(self, part_data_list, match_type):
        """Format the match result in a consistent way"""
        matches = []
        for part_data in part_data_list:
            part = part_data.get('searchProviderPart', {})
            mpn = part.get('manufacturerPartNumber', '')
            mfg = part.get('manufacturerName', '')
            matches.append(f"{mpn}@{mfg}")

        # Limit matches to user-configured maximum (default 10)
        return {'matches': matches[:self.max_matches]}, match_type


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
        self.original_data = []  # Store original data for comparison
        self.api_key = None
        
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

        except Exception as e:
            QMessageBox.critical(
                self,
                "Error Loading Results",
                f"Failed to load search results from CSV:\n{str(e)}\n\n"
                f"File: {csv_path}"
            )

    def load_results_from_csv(self, csv_path):
        """Load search results from CSV file"""
        import csv
        from collections import defaultdict

        results = []

        # Group rows by PartNumber+ManufacturerName (since multiple matches create multiple rows)
        grouped = defaultdict(lambda: {'matches': []})

        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                key = (row['PartNumber'], row['ManufacturerName'])

                # First occurrence - set basic info
                if not grouped[key]['matches'] and row['MatchValue(PartNumber@ManufacturerName)']:
                    grouped[key]['PartNumber'] = row['PartNumber']
                    grouped[key]['ManufacturerName'] = row['ManufacturerName']
                    grouped[key]['MatchStatus'] = row['MatchStatus']

                # Add match value if present
                match_value = row['MatchValue(PartNumber@ManufacturerName)'].strip()
                if match_value:
                    grouped[key]['matches'].append(match_value)
                elif not grouped[key]['matches']:
                    # Empty match (for None/Error status)
                    grouped[key]['PartNumber'] = row['PartNumber']
                    grouped[key]['ManufacturerName'] = row['ManufacturerName']
                    grouped[key]['MatchStatus'] = row['MatchStatus']

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
                if '@' in match:
                    _, mfg = match.split('@', 1)
                    mfg = mfg.strip()
                    if mfg:
                        canonical_mfgs.add(mfg)

        # Track manufacturers from USER-SELECTED matches (review phase work)
        # These are manufacturers the user specifically chose during review
        selected_mfgs = set()
        for result in self.search_results:
            if result.get('selected_match'):
                if '@' in result['selected_match']:
                    _, mfg = result['selected_match'].split('@', 1)
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

        for original in original_mfgs:
            # Skip if original is already in canonical list (exact match)
            if original in canonical_mfgs:
                continue

            # Find best match in canonical names using fuzzy matching
            result = process.extractOne(original, canonical_mfgs, scorer=fuzz.ratio)
            if result:
                match_name, score = result[0], result[1]

                # Only suggest normalization if:
                # 1. Names are different (not exact match)
                # 2. Score is high enough (>= 85) to suggest they're the same manufacturer
                if original != match_name and score >= 85:
                    normalizations[original] = match_name
                    reasoning_map[original] = {
                        'method': 'fuzzy',
                        'score': score,
                        'reasoning': f"Fuzzy match against PAS master list: {score}% similarity"
                    }
                    print(f"DEBUG: Normalization suggestion: '{original}' -> '{match_name}' ({score}%)")

        # If we found variations, populate the table
        if normalizations:
            self.manufacturer_normalizations = normalizations
            self.normalization_reasoning = reasoning_map

            # Populate normalization table
            self.norm_table.setRowCount(len(normalizations))

            row_idx = 0
            for original, canonical in normalizations.items():
                # Include checkbox (checked by default)
                include_cb = QCheckBox()
                include_cb.setChecked(True)
                self.norm_table.setCellWidget(row_idx, 0, include_cb)

                # Original MFG (read-only)
                self.norm_table.setItem(row_idx, 1, QTableWidgetItem(original))

                # Normalize To (editable combo box with color-coded manufacturers)
                normalize_combo = QComboBox()
                normalize_combo.setEditable(True)

                # Use QStandardItemModel for color coding
                from PyQt5.QtGui import QStandardItemModel, QStandardItem, QBrush
                model = QStandardItemModel()

                # Add canonical manufacturers with color coding
                for mfg in self.canonical_manufacturers:
                    item = QStandardItem(mfg)
                    # Color-code manufacturers from user's review selections
                    if mfg in self.selected_manufacturers:
                        # GREEN for user-selected manufacturers (their review work)
                        item.setForeground(QBrush(QColor(0, 128, 0)))  # Dark green
                        item.setToolTip("✓ Selected from review phase - preserves your work")
                        # Make it bold
                        font = item.font()
                        font.setBold(True)
                        item.setFont(font)
                    else:
                        # Normal black for other canonical manufacturers
                        item.setToolTip("Canonical manufacturer from PAS database")
                    model.appendRow(item)

                # Add original names not in canonical list (in gray)
                for mfg in sorted(original_mfgs - canonical_mfgs):
                    item = QStandardItem(mfg)
                    item.setForeground(QBrush(QColor(128, 128, 128)))  # Gray
                    item.setToolTip("Original manufacturer name (not in PAS canonical list)")
                    model.appendRow(item)

                normalize_combo.setModel(model)
                normalize_combo.setCurrentText(canonical)
                self.norm_table.setCellWidget(row_idx, 2, normalize_combo)

                # Scope dropdown
                scope_combo = QComboBox()
                scope_combo.addItems(["All Catalogs", "Per Catalog"])
                self.norm_table.setCellWidget(row_idx, 3, scope_combo)

                row_idx += 1

            # Update status and enable buttons
            self.norm_status.setText(
                f"✓ Found {len(normalizations)} variations using PAS master list ({len(canonical_mfgs)} canonical names)"
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
            matches_table.setColumnCount(5)
            matches_table.setHorizontalHeaderLabels(["Select", "Part Number", "Manufacturer", "Similarity", "AI Score"])
            matches_table.setContextMenuPolicy(Qt.CustomContextMenu)
            matches_table.customContextMenuRequested.connect(self.show_match_context_menu)
            
            # Set column resize modes
            header = matches_table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Select
            header.setSectionResizeMode(1, QHeaderView.Stretch)  # Part Number
            header.setSectionResizeMode(2, QHeaderView.Stretch)  # Manufacturer
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Similarity
            header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # AI Score
            
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
                from PyQt5.QtGui import QColor
                highlight_color = QColor(200, 255, 255)  # Light cyan
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

        # Color legend for dropdown
        legend_label = QLabel(
            "<b>Normalize To dropdown colors:</b> "
            "<span style='color: green; font-weight: bold;'>● Green/Bold</span> = Your review selections (preserves your work) | "
            "<span style='color: black;'>● Black</span> = PAS canonical manufacturers | "
            "<span style='color: gray;'>● Gray</span> = Original names (not in PAS)"
        )
        legend_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border-radius: 3px; font-size: 9pt;")
        legend_label.setWordWrap(True)
        norm_layout.addWidget(legend_label)

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
        # Determine which table triggered the selection
        sender = self.sender()
        
        # Default to checking all tables
        parts_list = None
        matches_table = None
        parts_data = None
        
        # Check which table has a selection
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

            # Add radio button to the button group to ensure mutual exclusivity
            button_group.addButton(radio)

            # Create a widget to center the radio button
            radio_widget = QWidget()
            radio_layout = QHBoxLayout(radio_widget)
            radio_layout.addWidget(radio)
            radio_layout.setAlignment(Qt.AlignCenter)
            radio_layout.setContentsMargins(0, 0, 0, 0)

            matches_table.setCellWidget(match_idx, 0, radio_widget)
            matches_table.setItem(match_idx, 1, QTableWidgetItem(pn))
            matches_table.setItem(match_idx, 2, QTableWidgetItem(mfg))

            # Calculate similarity score
            match_pn = pn.upper().strip()
            similarity = SequenceMatcher(None, original_pn, match_pn).ratio()
            similarity_pct = int(similarity * 100)
            similarity_item = QTableWidgetItem(f"{similarity_pct}%")
            similarity_item.setTextAlignment(Qt.AlignCenter)
            similarity_item.setToolTip("String similarity using difflib (part number matching)")
            matches_table.setItem(match_idx, 3, similarity_item)

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
            matches_table.setItem(match_idx, 4, ai_score_item)

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

            # Add radio button to the button group to ensure mutual exclusivity
            button_group.addButton(radio)

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
                part['ai_match_scores'][suggested_match] = confidence

        # Refresh the appropriate tables
        if part in self.multiple_parts:
            self.populate_category_table(self.multiple_table, self.multiple_parts, show_actions=True)
        elif part in self.need_review_parts:
            self.populate_category_table(self.need_review_table, self.need_review_parts, show_actions=True)

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

        # Collect all manufacturers
        all_mfgs = set()
        supplyframe_mfgs = set()

        # From original data (ALL manufacturers from Step 3)
        xml_gen_page = self.wizard().page(3)
        if hasattr(xml_gen_page, 'combined_data'):
            # Convert DataFrame to list of dictionaries if needed
            data = xml_gen_page.combined_data
            if hasattr(data, 'to_dict'):
                data = data.to_dict('records')
            for row in data:
                if isinstance(row, dict) and row.get('MFG'):
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
        canonical_mfgs = set()
        selected_mfgs = set()  # User-selected manufacturers from review phase

        # From original data
        xml_gen_page = self.wizard().page(3)
        if hasattr(xml_gen_page, 'combined_data'):
            # Convert DataFrame to list of dictionaries if needed
            data = xml_gen_page.combined_data
            if hasattr(data, 'to_dict'):
                data = data.to_dict('records')
            for row in data:
                if isinstance(row, dict) and row.get('MFG'):
                    all_mfgs.add(row['MFG'])

        # From search results - collect canonical and selected manufacturers
        if hasattr(self, 'search_results'):
            for result in self.search_results:
                # Collect all canonical manufacturers from matches
                for match in result.get('matches', []):
                    if '@' in match:
                        _, mfg = match.split('@', 1)
                        mfg = mfg.strip()
                        if mfg:
                            canonical_mfgs.add(mfg)
                            all_mfgs.add(mfg)

                # Track user-selected manufacturers (their review work)
                if result.get('selected_match') and '@' in result['selected_match']:
                    _, mfg = result['selected_match'].split('@', 1)
                    mfg = mfg.strip()
                    if mfg:
                        selected_mfgs.add(mfg)
                        all_mfgs.add(mfg)

        # From normalization suggestions
        for original, canonical in normalizations.items():
            all_mfgs.add(original)
            all_mfgs.add(canonical)

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

            # Normalize To (editable combo box with color-coded manufacturers)
            normalize_combo = QComboBox()
            normalize_combo.setEditable(True)

            # Use QStandardItemModel for color coding
            from PyQt5.QtGui import QStandardItemModel, QStandardItem, QBrush
            model = QStandardItemModel()

            # Add all manufacturers with color coding
            for mfg in sorted(all_mfgs):
                item = QStandardItem(mfg)

                # Color-code based on source
                if mfg in selected_mfgs:
                    # GREEN for user-selected manufacturers (their review work)
                    item.setForeground(QBrush(QColor(0, 128, 0)))  # Dark green
                    item.setToolTip("✓ Selected from review phase - preserves your work")
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                elif mfg in canonical_mfgs:
                    # Normal black for other canonical manufacturers
                    item.setToolTip("Canonical manufacturer from PAS database")
                else:
                    # Gray for original names
                    item.setForeground(QBrush(QColor(128, 128, 128)))
                    item.setToolTip("Original manufacturer name (not in PAS canonical list)")

                model.appendRow(item)

            normalize_combo.setModel(model)
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
            if hasattr(self, 'search_assign_data') and self.search_assign_data:
                for part_data in self.search_assign_data:
                    if 'selected_match' in part_data and part_data['selected_match']:
                        selected = part_data['selected_match']
                        # Find matching row in new_data
                        mask = (new_data['MFG_PN'] == part_data.get('part_number', ''))
                        if mask.any():
                            new_data.loc[mask, 'MFG'] = selected.get('manufacturer', '')

            # Apply manufacturer normalizations
            if hasattr(self, 'manufacturer_normalizations') and self.manufacturer_normalizations:
                for row_idx in range(self.norm_table.rowCount()):
                    include_checkbox = self.norm_table.cellWidget(row_idx, 0)
                    if include_checkbox and include_checkbox.isChecked():
                        original_item = self.norm_table.item(row_idx, 1)
                        canonical_item = self.norm_table.item(row_idx, 2)

                        if original_item and canonical_item:
                            original_mfg = original_item.text()
                            canonical_mfg = canonical_item.text()

                            # Apply normalization
                            if 'MFG' in new_data.columns:
                                new_data.loc[new_data['MFG'] == original_mfg, 'MFG'] = canonical_mfg

            # Read existing sheets from output Excel
            with pd.ExcelFile(output_excel) as xls:
                existing_sheets = {sheet: pd.read_excel(output_excel, sheet_name=sheet)
                                 for sheet in xls.sheet_names}

            # Add Combined_New sheet
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

    def build_comparison(self):
        """Build side-by-side comparison with Beyond Compare styling"""
        if self.original_df is None or self.new_df is None:
            return

        # Ensure both DataFrames have the same columns
        all_columns = list(set(self.original_df.columns) | set(self.new_df.columns))

        # Add missing columns
        for col in all_columns:
            if col not in self.original_df.columns:
                self.original_df[col] = ""
            if col not in self.new_df.columns:
                self.new_df[col] = ""

        # Reorder columns to match
        self.original_df = self.original_df[all_columns]
        self.new_df = self.new_df[all_columns]

        # Build row comparison data
        self.all_rows = []
        max_rows = max(len(self.original_df), len(self.new_df))

        changed_count = 0
        for i in range(max_rows):
            row_changed = False
            if i < len(self.original_df) and i < len(self.new_df):
                # Compare each cell
                for col in all_columns:
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

        # Set up columns
        columns = list(self.original_df.columns)
        self.left_table.setColumnCount(len(columns))
        self.left_table.setHorizontalHeaderLabels(columns)
        self.right_table.setColumnCount(len(columns))
        self.right_table.setHorizontalHeaderLabels(columns)

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
