#!/usr/bin/env python3
"""
EDM Library Wizard - Entry Point

A comprehensive wizard for converting Access databases to Excel and generating
XML files for EDM Library Creator v1.7.000.0130.

This is the main entry point that launches the modular EDM Wizard application.
"""

import sys

try:
    from PyQt5.QtWidgets import QApplication
except ImportError:
    print("Error: PyQt5 is required. Install it with: pip install PyQt5")
    sys.exit(1)

from edm_wizard.ui.wizard import EDMWizard


def main():
    """Main entry point for EDM Library Wizard"""
    app = QApplication(sys.argv)

    # Set application style
    app.setStyle('Fusion')

    # Create and show wizard
    wizard = EDMWizard()
    wizard.show()

    # Run application event loop
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
