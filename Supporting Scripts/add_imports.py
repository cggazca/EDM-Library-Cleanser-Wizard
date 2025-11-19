#!/usr/bin/env python3
"""
Add proper imports to extracted page modules
"""
from pathlib import Path

pages_config = {
    'start_page.py': {
        'title': 'Start Page: Claude AI API Key and PAS API Configuration',
        'imports': """import sys
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
    from PyQt5.QtCore import Qt, QSettings
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

try:
    from anthropic import Anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False
"""
    },
    'data_source_page.py': {
        'title': 'Data Source Page: File Selection and Database Export',
        'imports': """import sys
import os
from pathlib import Path
import pandas as pd
import shutil

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QLineEdit,
        QPushButton, QFileDialog, QComboBox, QTableWidget, QTableWidgetItem,
        QProgressBar, QMessageBox
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

from ..workers.threads import AccessExportThread, SQLiteExportThread
"""
    },
    'column_mapping_page.py': {
        'title': 'Column Mapping Page: AI-Assisted Column Detection',
        'imports': """import sys
import os
from pathlib import Path
import json
import pandas as pd

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QLineEdit,
        QPushButton, QFileDialog, QComboBox, QCheckBox, QTableWidget,
        QTableWidgetItem, QHeaderView, QProgressBar, QMessageBox, QWidget,
        QSplitter, QScrollArea, QSpinBox
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

try:
    from anthropic import Anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

from ..workers.threads import AIDetectionThread
"""
    },
    'pas_search_page.py': {
        'title': 'PAS Search Page: Part Aggregation Service API Search',
        'imports': """import sys
import os
import time
from pathlib import Path
import pandas as pd
import json

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel,
        QPushButton, QProgressBar, QMessageBox, QTextEdit, QWidget
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

from ..api.pas_client import PASAPIClient
"""
    },
    'xml_generation_page.py': {
        'title': 'XML Generation Page: Legacy XML Output',
        'imports': """import sys
import os
from pathlib import Path
import pandas as pd

try:
    from PyQt5.QtWidgets import (
        QWizardPage, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QLineEdit,
        QTableWidget, QTableWidgetItem, QHeaderView, QHeaderView,
        QPushButton, QMessageBox, QWidget, QScrollArea
    )
    from PyQt5.QtCore import Qt
except ImportError:
    print("Error: PyQt5 is required.")
    sys.exit(1)

from ...utils.xml_generation import escape_xml
"""
    },
    'review_page.py': {
        'title': 'Supply Frame Review Page: Results Review and Normalization',
        'imports': """import sys
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
        QInputDialog, QMenu, QHeaderView, QTextEdit, QDialog, QDialogButtonBox
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

from ...utils.xml_generation import escape_xml
"""
    }
}

pages_dir = Path("edm_wizard/ui/pages")

for filename, config in pages_config.items():
    file_path = pages_dir / filename
    
    if not file_path.exists():
        print(f"Warning: {file_path} does not exist")
        continue
    
    # Read current content
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Create header
    header = f'''"""
{config['title']}
"""

{config['imports']}


'''
    
    # Combine header with content
    new_content = header + content
    
    # Write back
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print(f"Added imports to {filename}")

print("\nImports added successfully!")
