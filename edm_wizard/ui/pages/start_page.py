"""
Start Page: Claude AI API Key and PAS API Configuration
"""

import sys
import os
from pathlib import Path
from datetime import datetime
import json

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QScrollArea, QWidget,
        QGroupBox, QLabel, QLineEdit, QPushButton, QCheckBox, QComboBox,
        QSpinBox, QFileDialog, QMessageBox, QApplication
    )
    from PyQt5.QtCore import Qt, QSettings, QThread, pyqtSignal
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

try:
    from anthropic import Anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False



class StartPage(QWizardPage):
    """Start Page: Claude AI API Key and PAS API Configuration"""

    def __init__(self):
        super().__init__()
        self.setTitle("Welcome to EDM Library Wizard")
        self.setSubTitle("Configure API credentials for intelligent column mapping and part search")

        # Create scroll area to handle overflow
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        
        # Content widget that will go inside the scroll area
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)

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
        
        # Set the scroll area content and add to page
        scroll.setWidget(content_widget)
        page_layout = QVBoxLayout()
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.addWidget(scroll)
        self.setLayout(page_layout)

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

