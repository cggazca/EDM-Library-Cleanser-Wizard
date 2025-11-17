#!/usr/bin/env python3
"""
Create new edm_wizard.py that imports pages from modules
This script reads the original edm_wizard.py and extracts everything except the 6 main pages
"""

# Read original file
with open('edm_wizard.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Line ranges for content to keep
# Keep: imports (1-84), CollapsibleGroupBox (53-84), threads (752-1023, 3261-3667, 3522-3667), 
# helper classes, ComparisonPage (7610-7983), EDMWizard (7984+), main function

# Classes to exclude from the output (they go to separate modules)
excluded_classes = {
    'StartPage': (85, 1144),
    'DataSourcePage': (1145, 1473),
    'ColumnMappingPage': (1474, 2450),
    'PASSearchPage': (2451, 2879),
    'XMLGenerationPage': (2880, 4069),
    'SupplyFrameReviewPage': (4070, 7609),
    'NoScrollComboBox': (1468, 1473),  # Also exclude this helper
}

# Build list of line ranges to keep (0-indexed)
keep_ranges = [
    (0, 84),      # Imports and imports block
    (751, 1023),  # AccessExportThread, SQLiteExportThread, SheetDetectionWorker
    (1023, 1143), # AIDetectionThread
    (1467, 1473), # NoScrollComboBox (actually we'll exclude this)
    (3260, 3667), # PartialMatchAIThread, ManufacturerNormalizationAIThread, PASSearchThread, PASAPIClient
    (7609, 7983), # ComparisonPage
    (7983, 8090), # EDMWizard and main function
]

# Actually, it's easier to just exclude the page classes
# Let's build the output line by line

output_lines = []
i = 0

while i < len(lines):
    # Check if we're at the start of an excluded class
    excluded = False
    for class_name, (start, end) in excluded_classes.items():
        if i == start - 1:  # -1 because lines are 1-indexed but list is 0-indexed
            # Skip to end of this class
            i = end - 1  # -1 because we'll increment at loop end
            excluded = True
            break
    
    if not excluded:
        output_lines.append(lines[i])
    
    i += 1

# Now add imports for the pages at the appropriate location
# Find where to insert the imports (after standard imports)
import_insert_pos = 0
for idx, line in enumerate(output_lines):
    if line.strip() == 'FUZZYWUZZY_AVAILABLE = False':
        import_insert_pos = idx + 2
        break

# Create import statement
import_statement = '''
# Import wizard pages from separate modules
from edm_wizard.ui.pages import (
    StartPage,
    DataSourcePage,
    ColumnMappingPage,
    PASSearchPage,
    XMLGenerationPage,
    SupplyFrameReviewPage
)
'''

# Insert the imports
output_lines.insert(import_insert_pos, import_statement)

# Write output
with open('edm_wizard_refactored.py', 'w', encoding='utf-8') as f:
    f.writelines(output_lines)

print("Created edm_wizard_refactored.py")
print(f"Total lines: {len(output_lines)}")
