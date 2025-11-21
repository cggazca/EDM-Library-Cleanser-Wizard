"""
Headless automation utilities for running the EDM wizard flows without a GUI.

The headless runner loads an Excel workbook, applies column mappings and filters,
builds the Combined sheet, and generates the MFG/MFGPN XML outputs. It is built
to mirror the Combine + XML steps the PyQt wizard performs so automated tests can
exercise the core data path without manual clicking.
"""

from __future__ import annotations

import argparse
import json
import urllib.parse
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, MutableMapping, Optional

import pandas as pd
import sqlalchemy as sa
from sqlalchemy import inspect

from . import constants
from .data_processing import clean_sheet_name, extract_mfgpn_data, extract_unique_manufacturers
from .xml_generation import create_mfg_xml, create_mfgpn_xml


DEFAULT_FILTERS = {
    "require_mfg": True,
    "require_mfg_pn": True,
    "require_part_number": False,
    "require_description": False,
}


@dataclass
class HeadlessConfig:
    """Configuration for a headless run."""

    input_path: Path
    output_dir: Path
    mappings: Mapping[str, Mapping[str, str]]
    include_sheets: Optional[Iterable[str]] = None
    filters: Mapping[str, bool] = None
    fill_tbd: bool = False
    project_name: str = constants.DEFAULT_PROJECT_NAME
    catalog: str = constants.DEFAULT_CATALOG


@dataclass
class HeadlessResult:
    """Result details from a headless run."""

    output_dir: Path
    output_excel: Path
    mfg_xml: Path
    mfgpn_xml: Path
    combined_rows: int


def _load_config(config_like) -> HeadlessConfig:
    """
    Accept a dict/path/str and return a typed HeadlessConfig.

    Paths are treated as JSON files to avoid adding a YAML dependency.
    """
    if isinstance(config_like, (str, Path)):
        config_path = Path(config_like)
        with config_path.open("r", encoding="utf-8") as fh:
            loaded = json.load(fh)
    elif isinstance(config_like, Mapping):
        loaded = dict(config_like)
    else:
        raise TypeError("config must be a dict, str path, or Path")

    filters = DEFAULT_FILTERS.copy()
    filters.update(loaded.get("filters", {}))

    return HeadlessConfig(
        input_path=Path(loaded["input_path"]),
        output_dir=Path(loaded.get("output_dir", "headless_output")),
        mappings=loaded["mappings"],
        include_sheets=loaded.get("include_sheets"),
        filters=filters,
        fill_tbd=loaded.get("fill_tbd", False),
        project_name=loaded.get("project_name", constants.DEFAULT_PROJECT_NAME),
        catalog=loaded.get("catalog", constants.DEFAULT_CATALOG),
    )


def _load_workbook(input_path: Path) -> Dict[str, pd.DataFrame]:
    """Load all sheets from an Excel workbook."""
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    excel_file = pd.ExcelFile(input_path)
    return {sheet: pd.read_excel(input_path, sheet_name=sheet) for sheet in excel_file.sheet_names}


def _export_access_to_excel(mdb_path: Path, output_dir: Path):
    """
    Export an Access database to Excel (Step 1 of the wizard, headless).

    Returns the DataFrames for immediate reuse (avoids re-reading from disk).
    """
    if not mdb_path.exists():
        raise FileNotFoundError(f"Access database not found: {mdb_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_excel = output_dir / f"{mdb_path.stem}.xlsx"

    try:
        conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(mdb_path)
        quoted_conn_str = urllib.parse.quote_plus(conn_str)
        engine = sa.create_engine(f"access+pyodbc:///?odbc_connect={quoted_conn_str}")

        inspector = inspect(engine)
        tables = inspector.get_table_names()
        if not tables:
            raise RuntimeError("No tables found in Access database.")

        dataframes: Dict[str, pd.DataFrame] = {}
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            for table in tables:
                df = pd.read_sql(f"SELECT * FROM [{table}]", engine)
                sheet_name = clean_sheet_name(table)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                dataframes[sheet_name] = df

        return output_excel, dataframes
    except Exception as exc:  # pragma: no cover - driver/config specific
        raise RuntimeError(
            f"Failed to export Access DB '{mdb_path}'. Ensure the Microsoft Access Database "
            f"Engine is installed. Original error: {exc}"
        ) from exc


def _apply_filters(df: pd.DataFrame, filters: Mapping[str, bool]) -> pd.DataFrame:
    """Apply data quality filters mirroring the UI selections."""
    mask = pd.Series(True, index=df.index)

    if filters.get("require_mfg") and "MFG" in df.columns:
        mask &= df["MFG"].notna() & (df["MFG"].astype(str).str.strip() != "")

    if filters.get("require_mfg_pn") and "MFG_PN" in df.columns:
        mask &= df["MFG_PN"].notna() & (df["MFG_PN"].astype(str).str.strip() != "")

    if filters.get("require_part_number") and "Part_Number" in df.columns:
        mask &= df["Part_Number"].notna() & (df["Part_Number"].astype(str).str.strip() != "")

    if filters.get("require_description") and "Description" in df.columns:
        mask &= df["Description"].notna() & (df["Description"].astype(str).str.strip() != "")

    return df.loc[mask].copy()


def _combine_sheets(
    dataframes: Mapping[str, pd.DataFrame],
    mappings: Mapping[str, Mapping[str, str]],
    include_sheets: Optional[Iterable[str]],
    filters: Mapping[str, bool],
    fill_tbd: bool,
) -> pd.DataFrame:
    """
    Combine sheets into a single DataFrame using the provided mappings.

    This closely mirrors ColumnMappingPage.combine_sheets but runs without the UI.
    """
    combined = []
    target_sheets = list(include_sheets) if include_sheets else list(dataframes.keys())

    for sheet_name in target_sheets:
        if sheet_name not in dataframes or sheet_name not in mappings:
            continue

        sheet_mapping: MutableMapping[str, str] = dict(mappings[sheet_name])
        df = dataframes[sheet_name].copy()
        df["Source_Sheet"] = sheet_name

        rename_dict = {
            col_name: target
            for target, col_name in sheet_mapping.items()
            if col_name and target != "MFG_PN_2"
        }

        if rename_dict:
            df = df.rename(columns=rename_dict)

        # Handle MFG_PN fallback to MFG_PN_2
        mfg_pn_primary = sheet_mapping.get("MFG_PN")
        mfg_pn_secondary = sheet_mapping.get("MFG_PN_2")
        if mfg_pn_primary and mfg_pn_secondary:
            if "MFG_PN" in df.columns and mfg_pn_secondary in dataframes[sheet_name].columns:
                secondary_values = dataframes[sheet_name][mfg_pn_secondary]
                empty_mask = df["MFG_PN"].isna() | (df["MFG_PN"].astype(str).str.strip() == "")
                df.loc[empty_mask, "MFG_PN"] = secondary_values[empty_mask].values

        if fill_tbd and {"MFG", "MFG_PN"} <= set(df.columns):
            mfg_pn_present = df["MFG_PN"].notna() & (df["MFG_PN"].astype(str).str.strip() != "")
            mfg_missing = df["MFG"].isna() | (df["MFG"].astype(str).str.strip() == "")
            df.loc[mfg_pn_present & mfg_missing, "MFG"] = "TBD"

        filtered = _apply_filters(df, filters)
        if not filtered.empty:
            combined.append(filtered)

    if not combined:
        return pd.DataFrame()

    return pd.concat(combined, ignore_index=True)


def run_headless_flow(config_like) -> HeadlessResult:
    """
    Execute the wizard flow headlessly (Access/Excel -> combine -> XML).

    Args:
        config_like: Dict/Path/str pointing to JSON with keys:
            - input_path: Access .mdb/.accdb OR Excel file to use
            - output_dir: folder where outputs will be written
            - mappings: {sheet: {"MFG": ..., "MFG_PN": ..., "MFG_PN_2": ..., "Part_Number": ..., "Description": ...}}
            - include_sheets: optional list of sheets to process (defaults to all)
            - filters: optional dict matching DEFAULT_FILTERS keys
            - fill_tbd: bool to fill missing MFG with 'TBD' when MFG_PN exists
            - project_name: optional project name for XML headers
            - catalog: optional catalog code for XML headers
    """
    config = _load_config(config_like)
    config.output_dir.mkdir(parents=True, exist_ok=True)

    # Step 1: load or export the data source
    if config.input_path.suffix.lower() in constants.SUPPORTED_ACCESS_EXTENSIONS:
        effective_input_path, dataframes = _export_access_to_excel(config.input_path, config.output_dir)
    elif config.input_path.suffix.lower() in constants.SUPPORTED_EXCEL_EXTENSIONS:
        dataframes = _load_workbook(config.input_path)
        effective_input_path = config.input_path
    else:
        raise ValueError(
            f"Unsupported input type '{config.input_path.suffix}'. "
            f"Use one of {constants.SUPPORTED_ACCESS_EXTENSIONS + constants.SUPPORTED_EXCEL_EXTENSIONS}."
        )

    combined_df = _combine_sheets(
        dataframes=dataframes,
        mappings=config.mappings,
        include_sheets=config.include_sheets,
        filters=config.filters,
        fill_tbd=config.fill_tbd,
    )

    if combined_df.empty:
        raise ValueError("No data remained after applying mappings/filters; nothing to process.")

    output_excel = config.output_dir / f"{effective_input_path.stem}_Combined.xlsx"
    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        # Persist originals for debugging and the Combined sheet for downstream steps
        for sheet_name, df in dataframes.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        combined_df.to_excel(writer, sheet_name="Combined", index=False)

    manufacturers = extract_unique_manufacturers(combined_df, mfg_column="MFG")
    mfgpn_records = extract_mfgpn_data(
        combined_df,
        mfg_column="MFG",
        mfgpn_column="MFG_PN",
        desc_column="Description",
    )

    mfg_xml_path = config.output_dir / f"{effective_input_path.stem}_MFG.xml"
    mfgpn_xml_path = config.output_dir / f"{effective_input_path.stem}_MFGPN.xml"

    create_mfg_xml(manufacturers, mfg_xml_path, config.project_name, config.catalog)
    create_mfgpn_xml(mfgpn_records, mfgpn_xml_path, config.project_name, config.catalog)

    return HeadlessResult(
        output_dir=config.output_dir,
        output_excel=output_excel,
        mfg_xml=mfg_xml_path,
        mfgpn_xml=mfgpn_xml_path,
        combined_rows=len(combined_df),
    )


def main():
    """CLI entry point for running the headless flow."""
    parser = argparse.ArgumentParser(description="Run the EDM wizard flow without the GUI.")
    parser.add_argument(
        "--config",
        required=True,
        help="Path to JSON configuration file describing mappings and filters.",
    )
    args = parser.parse_args()

    result = run_headless_flow(Path(args.config))
    print(f"Combined rows: {result.combined_rows}")
    print(f"Excel written to: {result.output_excel}")
    print(f"MFG XML: {result.mfg_xml}")
    print(f"MFGPN XML: {result.mfgpn_xml}")


if __name__ == "__main__":
    main()
