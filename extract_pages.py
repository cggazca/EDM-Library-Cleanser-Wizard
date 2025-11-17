#!/usr/bin/env python3
"""
Script to extract wizard pages from edm_wizard.py into separate modules
"""
import re
from pathlib import Path

# Read the main file
main_file = "edm_wizard.py"
with open(main_file, 'r', encoding='utf-8') as f:
    content = f.read()
    lines = content.split('\n')

# Define class ranges (line numbers are 1-indexed, but list is 0-indexed)
# From grep output: StartPage (85), DataSourcePage (1145), ColumnMappingPage (1474), 
# PASSearchPage (2451), XMLGenerationPage (2880), SupplyFrameReviewPage (4070), EDMWizard (7984)
class_ranges = {
    'StartPage': (84, 1144),  # 85-1145 (exclusive end)
    'DataSourcePage': (1144, 1473),  # 1145-1474
    'ColumnMappingPage': (1473, 2450),  # 1474-2451
    'PASSearchPage': (2450, 2879),  # 2451-2880
    'XMLGenerationPage': (2879, 4069),  # 2880-4070
    'SupplyFrameReviewPage': (4069, 7983),  # 4070-7984
    'EDMWizard': (7983, len(lines))  # 7984-end
}

# Extract each class
pages_dir = Path("edm_wizard/ui/pages")
pages_dir.mkdir(parents=True, exist_ok=True)

page_classes = {
    'StartPage': 'start_page',
    'DataSourcePage': 'data_source_page',
    'ColumnMappingPage': 'column_mapping_page',
    'PASSearchPage': 'pas_search_page',
    'XMLGenerationPage': 'xml_generation_page',
    'SupplyFrameReviewPage': 'review_page'
}

for class_name, module_name in page_classes.items():
    start, end = class_ranges[class_name]
    class_lines = lines[start:end]
    
    # Write to file
    output_file = pages_dir / f"{module_name}.py"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(class_lines))
    
    print(f"Extracted {class_name} to {output_file} (lines {start+1}-{end})")

print("\nExtraction complete!")
