"""
EDM Library Wizard - Main Application Window

The EDMWizard class manages the complete wizard workflow with 6 pages:
1. StartPage - API configuration and output folder selection
2. DataSourcePage - Select and export data source (Access/SQLite/Excel)
3. ColumnMappingPage - Map columns and combine sheets
4. PASSearchPage - Search parts via PAS API
5. SupplyFrameReviewPage - Review and normalize results
6. ComparisonPage - Compare original vs modified data

This module provides the main wizard window and styling.
"""

try:
    from PyQt5.QtWidgets import QWizard, QSizePolicy
    from PyQt5.QtCore import Qt
except ImportError:
    print("Error: PyQt5 is required. Install it with: pip install PyQt5")
    raise

from edm_wizard.ui.pages import (
    StartPage,
    DataSourcePage,
    ColumnMappingPage,
    PASSearchPage,
    XMLGenerationPage,
    SupplyFrameReviewPage,
    ComparisonPage
)


class EDMWizard(QWizard):
    """
    Main wizard window for EDM Library processing

    A 6-step wizard that guides users through:
    - Configuring API credentials and output location
    - Selecting and exporting data sources
    - Mapping and combining columns
    - Searching parts via PAS API
    - Reviewing and normalizing results
    - Comparing original vs modified data

    Attributes:
        start_page: StartPage instance (page 0)
        data_source_page: DataSourcePage instance (page 1)
        column_mapping_page: ColumnMappingPage instance (page 2)
        pas_search_page: PASSearchPage instance (page 3)
        review_page: SupplyFrameReviewPage instance (page 4)
        comparison_page: ComparisonPage instance (page 5)
    """

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

        # Create wizard pages
        self.start_page = StartPage()
        self.data_source_page = DataSourcePage()
        self.column_mapping_page = ColumnMappingPage()
        self.pas_search_page = PASSearchPage()
        self.review_page = SupplyFrameReviewPage()
        self.comparison_page = ComparisonPage()

        # Add pages in sequence
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
        self._apply_styling()

    def _apply_styling(self):
        """Apply custom stylesheet to wizard"""
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
