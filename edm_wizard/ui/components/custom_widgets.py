"""
Custom UI components for EDM Library Wizard
"""

from PyQt5.QtWidgets import QGroupBox, QComboBox, QVBoxLayout, QWidget


class CollapsibleGroupBox(QGroupBox):
    """
    A QGroupBox that can be collapsed/expanded by clicking the title checkbox

    Used in SupplyFrameReviewPage for dynamic section expansion.
    """

    def __init__(self, title="", parent=None):
        """
        Initialize collapsible group box

        Args:
            title: Title text for the group box
            parent: Parent widget
        """
        super().__init__(title, parent)
        self.setCheckable(True)
        self.setChecked(True)  # Expanded by default
        self.toggled.connect(self.on_toggled)

        # Store the content widget
        self._content_widget = None

    def setContentLayout(self, layout):
        """
        Set the content layout that will be shown/hidden

        Args:
            layout: QLayout to display inside the group box
        """
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
        """
        Show/hide content when toggled

        Args:
            checked: True if expanded, False if collapsed
        """
        if self._content_widget:
            self._content_widget.setVisible(checked)


class NoScrollComboBox(QComboBox):
    """
    ComboBox that ignores mouse wheel events

    Prevents accidental value changes when scrolling through forms.
    Used in ColumnMappingPage for column selection dropdowns.
    """

    def wheelEvent(self, event):
        """
        Override wheel event to prevent scrolling

        Args:
            event: QWheelEvent
        """
        event.ignore()
