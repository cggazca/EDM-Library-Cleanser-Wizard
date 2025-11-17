"""
Data processing utilities for EDM Library Wizard
"""

import pandas as pd
from .constants import EXCEL_MAX_SHEET_NAME_LENGTH, EXCEL_INVALID_SHEET_CHARS


def clean_sheet_name(name):
    """
    Clean Excel sheet names to meet Excel requirements

    Excel sheet name restrictions:
    - Maximum 31 characters
    - Cannot contain: \ / * ? : [ ]

    Args:
        name: Sheet name to clean

    Returns:
        Cleaned sheet name
    """
    for char in EXCEL_INVALID_SHEET_CHARS:
        name = name.replace(char, '')
    return name[:EXCEL_MAX_SHEET_NAME_LENGTH]


def combine_dataframes(dataframes, mappings, include_sheets=None, filter_conditions=None):
    """
    Combine multiple DataFrames with column mapping

    Args:
        dataframes: Dict of {sheet_name: DataFrame}
        mappings: Dict of {sheet_name: {'mfg_col': str, 'mfgpn_col': str}}
        include_sheets: List of sheet names to include (None = all)
        filter_conditions: Optional filter expressions (not implemented yet)

    Returns:
        Combined DataFrame with 'Source_Sheet', 'MFG', and 'MFG PN' columns
    """
    combined_rows = []

    sheets_to_process = include_sheets if include_sheets else list(dataframes.keys())

    for sheet_name in sheets_to_process:
        if sheet_name not in dataframes or sheet_name not in mappings:
            continue

        df = dataframes[sheet_name].copy()
        mapping = mappings[sheet_name]

        mfg_col = mapping.get('mfg_col')
        mfgpn_col = mapping.get('mfgpn_col')

        if not mfg_col or not mfgpn_col:
            continue

        # Create standardized columns
        df['Source_Sheet'] = sheet_name
        df['MFG'] = df[mfg_col]
        df['MFG PN'] = df[mfgpn_col]

        # Keep original columns plus standardized ones
        combined_rows.append(df)

    if not combined_rows:
        return pd.DataFrame()

    result = pd.concat(combined_rows, ignore_index=True)

    # Apply filters if provided (placeholder for future implementation)
    if filter_conditions:
        # TODO: Implement filter logic
        pass

    return result


def extract_unique_manufacturers(dataframe, mfg_column='MFG'):
    """
    Extract unique manufacturers from DataFrame

    Args:
        dataframe: DataFrame with manufacturer data
        mfg_column: Column name containing manufacturers

    Returns:
        Sorted list of unique manufacturer names (empty strings removed)
    """
    if mfg_column not in dataframe.columns:
        return []

    manufacturers = dataframe[mfg_column].dropna().unique().tolist()
    manufacturers = [str(m).strip() for m in manufacturers if str(m).strip()]
    return sorted(set(manufacturers))


def extract_mfgpn_data(dataframe, mfg_column='MFG', mfgpn_column='MFG PN', desc_column=None):
    """
    Extract manufacturer part number data for XML generation

    Args:
        dataframe: DataFrame with part data
        mfg_column: Column name containing manufacturers
        mfgpn_column: Column name containing part numbers
        desc_column: Optional column name containing descriptions

    Returns:
        List of dicts with 'MFG', 'MFG_PN', 'Description' keys
    """
    result = []

    for _, row in dataframe.iterrows():
        mfg = str(row.get(mfg_column, '')).strip()
        mfg_pn = str(row.get(mfgpn_column, '')).strip()

        if not mfg or not mfg_pn:
            continue

        description = ''
        if desc_column and desc_column in dataframe.columns:
            description = str(row.get(desc_column, '')).strip()

        result.append({
            'MFG': mfg,
            'MFG_PN': mfg_pn,
            'Description': description
        })

    return result


def normalize_manufacturer_names(dataframe, normalizations, mfg_column='MFG'):
    """
    Apply manufacturer name normalizations to DataFrame

    Args:
        dataframe: DataFrame to modify
        normalizations: Dict mapping {original_name: normalized_name}
        mfg_column: Column name containing manufacturers

    Returns:
        Modified DataFrame (in-place modification)
    """
    if mfg_column not in dataframe.columns:
        return dataframe

    dataframe[mfg_column] = dataframe[mfg_column].apply(
        lambda x: normalizations.get(str(x), str(x)) if pd.notna(x) else x
    )

    return dataframe
