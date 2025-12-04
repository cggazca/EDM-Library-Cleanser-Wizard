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
    SupplyFrameReviewPage,
    ComparisonPage
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


# Note: SaveOptionsDialog and ComparisonPage are imported from edm_wizard.ui.pages.comparison_page


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
