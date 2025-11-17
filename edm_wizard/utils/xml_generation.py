"""
XML generation utilities for EDM Library Creator
"""

import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import pandas as pd

from .constants import XML_CLASS_MFG, XML_CLASS_MFGPN


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


def create_mfg_xml(manufacturers, output_file, project_name, catalog):
    """
    Create MFG XML file (Manufacturer class 090)

    Args:
        manufacturers: List of manufacturer names
        output_file: Output file path
        project_name: DDP project name
        catalog: Catalog code (e.g., "VV")

    Returns:
        Number of manufacturers written
    """
    manufacturers = sorted([m for m in manufacturers if m])

    root = ET.Element('data')

    for mfg in manufacturers:
        obj = ET.SubElement(root, 'object')
        obj.set('objectid', escape_xml(mfg))
        obj.set('catalog', catalog)
        obj.set('class', XML_CLASS_MFG)

        field1 = ET.SubElement(obj, 'field')
        field1.set('id', '090obj_skn')
        field1.text = catalog

        field2 = ET.SubElement(obj, 'field')
        field2.set('id', '090obj_id')
        field2.text = escape_xml(mfg)

        field3 = ET.SubElement(obj, 'field')
        field3.set('id', '090her_name')
        field3.text = escape_xml(mfg)

    save_xml(root, output_file, project_name)
    return len(manufacturers)


def create_mfgpn_xml(mfgpn_data, output_file, project_name, catalog):
    """
    Create MFGPN XML file (Manufacturer Part Number class 060)

    Args:
        mfgpn_data: List of dicts with 'MFG', 'MFG_PN', 'Description' keys
        output_file: Output file path
        project_name: DDP project name
        catalog: Catalog code (e.g., "VV")

    Returns:
        Number of unique part numbers written
    """
    # Remove duplicates
    unique_pairs = {}
    for item in mfgpn_data:
        key = (item['MFG'], item['MFG_PN'])
        if key not in unique_pairs:
            unique_pairs[key] = item.get('Description', '')

    root = ET.Element('data')

    for (mfg, mfg_pn), description in unique_pairs.items():
        objectid = f"{mfg}:{mfg_pn}"

        obj = ET.SubElement(root, 'object')
        obj.set('objectid', escape_xml(objectid))
        obj.set('class', XML_CLASS_MFGPN)

        field1 = ET.SubElement(obj, 'field')
        field1.set('id', '060partnumber')
        field1.text = escape_xml(mfg_pn)

        field2 = ET.SubElement(obj, 'field')
        field2.set('id', '060mfgref')
        field2.text = escape_xml(mfg)

        field3 = ET.SubElement(obj, 'field')
        field3.set('id', '060komp_name')
        field3.text = escape_xml(description)

    save_xml(root, output_file, project_name)
    return len(unique_pairs)


def save_xml(root, output_file, project_name):
    """
    Format and save XML file with EDM Library Creator headers

    Args:
        root: ET.Element root node
        output_file: Output file path
        project_name: DDP project name
    """
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
