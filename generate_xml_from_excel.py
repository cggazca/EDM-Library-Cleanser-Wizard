#!/usr/bin/env python3
"""
Generate XML files for xml-console from Excel data
Creates separate XML files for Manufacturers (MFG) and Manufacturer Part Numbers (MFGPN)
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET
from xml.dom import minidom
import sys


def escape_xml(text):
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


def create_mfg_xml(df, output_file, project_name="VarTrainingLab", catalog="VV"):
    """
    Create XML file for Manufacturers (class 090)

    Args:
        df: DataFrame containing manufacturer data
        output_file: Path to output XML file
        project_name: DDP Project name
        catalog: Catalog identifier
    """
    # Get unique manufacturers from the MFG column
    if 'MFG' not in df.columns:
        print("ERROR: 'MFG' column not found in Excel file")
        return False

    manufacturers = df['MFG'].dropna().unique()
    manufacturers = sorted([str(m).strip() for m in manufacturers if str(m).strip()])

    print(f"Found {len(manufacturers)} unique manufacturers")

    # Create XML structure
    root = ET.Element('data')

    # Add comment elements manually
    comment_lines = [
        f'Created By: EDM Library Creator v1.7.000.0130',
        f'DDP Project: {project_name}',
        f'Date: {datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")}'
    ]

    for mfg in manufacturers:
        obj = ET.SubElement(root, 'object')
        obj.set('objectid', escape_xml(mfg))
        obj.set('catalog', catalog)
        obj.set('class', '090')

        # Add fields
        field1 = ET.SubElement(obj, 'field')
        field1.set('id', '090obj_skn')
        field1.text = catalog

        field2 = ET.SubElement(obj, 'field')
        field2.set('id', '090obj_id')
        field2.text = escape_xml(mfg)

        field3 = ET.SubElement(obj, 'field')
        field3.set('id', '090her_name')
        field3.text = escape_xml(mfg)

    # Convert to string with proper formatting
    xml_str = ET.tostring(root, encoding='utf-8', method='xml')
    dom = minidom.parseString(xml_str)

    # Create custom XML with comments
    xml_lines = ['<?xml version="1.0" encoding="utf-8" standalone="yes"?>']
    for comment in comment_lines:
        xml_lines.append(f'<!--{comment}-->')

    # Get formatted XML (skip first line which is the XML declaration)
    formatted = dom.toprettyxml(indent='  ', encoding='utf-8').decode('utf-8')
    xml_content = '\n'.join(formatted.split('\n')[1:])

    # Write to file
    final_xml = '\n'.join(xml_lines) + '\n' + xml_content

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(final_xml)

    print(f"Created MFG XML: {output_file}")
    print(f"  - {len(manufacturers)} manufacturers")
    return True


def create_mfgpn_xml(df, output_file, project_name="VarTrainingLab", catalog="VV"):
    """
    Create XML file for Manufacturer Part Numbers (class 060)

    Args:
        df: DataFrame containing part number data
        output_file: Path to output XML file
        project_name: DDP Project name
        catalog: Catalog identifier
    """
    # Check required columns
    if 'MFG PN' not in df.columns or 'MFG' not in df.columns:
        print("ERROR: Required columns 'MFG PN' and/or 'MFG' not found in Excel file")
        return False

    # Filter rows with both MFG and MFG PN
    df_filtered = df[['MFG', 'MFG PN']].dropna()

    # Remove duplicates based on MFG:MFGPN combination
    df_filtered = df_filtered.drop_duplicates()

    print(f"Found {len(df_filtered)} unique MFG/MFG PN combinations")

    # Create XML structure
    root = ET.Element('data')

    # Add comment elements manually
    comment_lines = [
        f'Created By: EDM Library Creator v1.7.000.0130',
        f'DDP Project: {project_name}',
        f'Date: {datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")}'
    ]

    for idx, row in df_filtered.iterrows():
        mfg = str(row['MFG']).strip()
        mfg_pn = str(row['MFG PN']).strip()

        # objectid is "MFG:MFGPN"
        objectid = f"{mfg}:{mfg_pn}"

        obj = ET.SubElement(root, 'object')
        obj.set('objectid', escape_xml(objectid))
        obj.set('class', '060')

        # Add fields
        field1 = ET.SubElement(obj, 'field')
        field1.set('id', '060partnumber')
        field1.text = escape_xml(mfg_pn)

        field2 = ET.SubElement(obj, 'field')
        field2.set('id', '060mfgref')
        field2.text = escape_xml(mfg)

        field3 = ET.SubElement(obj, 'field')
        field3.set('id', '060komp_name')
        field3.text = "This is the PN description."

    # Convert to string with proper formatting
    xml_str = ET.tostring(root, encoding='utf-8', method='xml')
    dom = minidom.parseString(xml_str)

    # Create custom XML with comments
    xml_lines = ['<?xml version="1.0" encoding="utf-8" standalone="yes"?>']
    for comment in comment_lines:
        xml_lines.append(f'<!--{comment}-->')

    # Get formatted XML (skip first line which is the XML declaration)
    formatted = dom.toprettyxml(indent='  ', encoding='utf-8').decode('utf-8')
    xml_content = '\n'.join(formatted.split('\n')[1:])

    # Write to file
    final_xml = '\n'.join(xml_lines) + '\n' + xml_content

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(final_xml)

    print(f"Created MFGPN XML: {output_file}")
    print(f"  - {len(df_filtered)} part numbers")
    return True


def main():
    """Main function to generate XML files from Excel data"""

    # Get Excel file path
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        print("Usage: python generate_xml_from_excel.py <excel_file>")
        print("\nLooking for combined Excel file in current directory...")
        possible_files = list(Path('.').glob('*combined*.xlsx'))
        if possible_files:
            excel_file = str(possible_files[0])
            print(f"Found: {excel_file}")
        else:
            print("No combined Excel file found.")
            sys.exit(1)

    # Check if file exists
    if not Path(excel_file).exists():
        print(f"ERROR: File '{excel_file}' not found")
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f"Processing Excel file: {excel_file}")
    print(f"{'='*60}\n")

    # Read Excel file
    try:
        df = pd.read_excel(excel_file)
        print(f"Loaded Excel file with {len(df)} rows and {len(df.columns)} columns")
        print(f"Columns: {', '.join(df.columns)}\n")
    except Exception as e:
        print(f"ERROR reading Excel file: {e}")
        sys.exit(1)

    # Get project name from user or use default
    project_name = "VarTrainingLab"
    catalog = "VV"

    # Generate output file names
    base_name = Path(excel_file).stem.replace('_combined', '')
    output_dir = Path(excel_file).parent

    mfg_xml = output_dir / f"{base_name}_MFG.xml"
    mfgpn_xml = output_dir / f"{base_name}_MFGPN.xml"

    print(f"Output files:")
    print(f"  - {mfg_xml}")
    print(f"  - {mfgpn_xml}\n")

    # Create XML files
    print("Generating XML files...\n")

    success1 = create_mfg_xml(df, mfg_xml, project_name, catalog)
    print()
    success2 = create_mfgpn_xml(df, mfgpn_xml, project_name, catalog)

    if success1 and success2:
        print(f"\n{'='*60}")
        print("SUCCESS: XML files generated successfully!")
        print(f"{'='*60}")
    else:
        print("\nERROR: Failed to generate one or more XML files")
        sys.exit(1)


if __name__ == "__main__":
    main()
