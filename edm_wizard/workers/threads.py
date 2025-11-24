"""
QThread worker classes for EDM Library Wizard

This module consolidates all background thread workers used for long-running
operations that need to keep the UI responsive:

- AccessExportThread: Export Access databases to Excel
- SQLiteExportThread: Export SQLite databases to Excel
- SheetDetectionWorker: AI-powered single sheet column detection
- AIDetectionThread: Coordinator for parallel AI sheet detection
- PartialMatchAIThread: AI suggestions for partial matches
- ManufacturerNormalizationAIThread: AI manufacturer name normalization
- PASSearchThread: Parallel PAS API part searching
"""

import json
import time
import threading
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import sqlalchemy as sa
import urllib.parse
from sqlalchemy import inspect

from PyQt5.QtCore import QThread, pyqtSignal

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

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

from ..utils.data_processing import clean_sheet_name


class AccessExportThread(QThread):
    """Background thread for exporting Access database to Excel"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(str, object)  # excel_path, dataframes_dict
    error = pyqtSignal(str)

    def __init__(self, mdb_file, output_file):
        super().__init__()
        self.mdb_file = mdb_file
        self.output_file = output_file

    def run(self):
        try:
            self.progress.emit("Connecting to Access database...")

            # Create connection string
            conn_str = (
                r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                r"DBQ=" + self.mdb_file
            )
            quoted_conn_str = urllib.parse.quote_plus(conn_str)
            engine = sa.create_engine(f"access+pyodbc:///?odbc_connect={quoted_conn_str}")

            # Get table names
            inspector = inspect(engine)
            tables = inspector.get_table_names()

            self.progress.emit(f"Found {len(tables)} tables. Exporting...")

            # Export all tables
            dataframes = {}
            with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
                for idx, table in enumerate(tables, 1):
                    self.progress.emit(f"Exporting table {idx}/{len(tables)}: {table}")
                    df = pd.read_sql(f"SELECT * FROM [{table}]", engine)

                    # Clean sheet name
                    sheet_name = clean_sheet_name(table)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    dataframes[sheet_name] = df

            self.progress.emit("Export completed successfully!")
            self.finished.emit(self.output_file, dataframes)

        except Exception as e:
            self.error.emit(f"Error exporting Access database: {str(e)}")


class SQLiteExportThread(QThread):
    """Background thread for exporting SQLite database to Excel"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(str, object)  # excel_path, dataframes_dict
    error = pyqtSignal(str)

    def __init__(self, sqlite_file, output_file):
        super().__init__()
        self.sqlite_file = sqlite_file
        self.output_file = output_file

    def run(self):
        try:
            import sqlite3

            self.progress.emit("Connecting to SQLite database...")

            # Connect to SQLite database
            conn = sqlite3.connect(self.sqlite_file)
            cursor = conn.cursor()

            # Get table names
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';")
            tables = [row[0] for row in cursor.fetchall()]

            if not tables:
                self.error.emit("No tables found in SQLite database.")
                conn.close()
                return

            self.progress.emit(f"Found {len(tables)} tables. Exporting...")

            # Export all tables
            dataframes = {}
            with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
                for idx, table in enumerate(tables, 1):
                    self.progress.emit(f"Exporting table {idx}/{len(tables)}: {table}")

                    # Read table data (SQLite uses double quotes for identifiers)
                    df = pd.read_sql_query(f'SELECT * FROM "{table}"', conn)

                    # Clean sheet name
                    sheet_name = clean_sheet_name(table)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    dataframes[sheet_name] = df

            conn.close()
            self.progress.emit("Export completed successfully!")
            self.finished.emit(self.output_file, dataframes)

        except Exception as e:
            self.error.emit(f"Error exporting SQLite database: {str(e)}")


class SheetDetectionWorker(QThread):
    """Worker thread for detecting columns in a single sheet using AI"""
    finished = pyqtSignal(str, dict)  # sheet_name, mapping
    error = pyqtSignal(str, str)  # sheet_name, error_msg

    def __init__(self, api_key, sheet_name, dataframe, model="claude-sonnet-4-5-20250929", max_retries=5):
        super().__init__()
        self.api_key = api_key
        self.sheet_name = sheet_name
        self.dataframe = dataframe
        self.model = model
        self.max_retries = max_retries

    def run(self):
        retry_count = 0
        base_delay = 10  # Start with 10 second delay

        while retry_count <= self.max_retries:
            try:
                client = Anthropic(api_key=self.api_key)

                # Prepare column information
                columns = self.dataframe.columns.tolist()

                # Filter out rows that are mostly empty (less than 30% of columns have data)
                min_fields_threshold = max(2, len(columns) * 0.3)
                non_empty_counts = self.dataframe.notna().sum(axis=1)
                df_filtered = self.dataframe[non_empty_counts >= min_fields_threshold].copy()

                if len(df_filtered) == 0:
                    df_filtered = self.dataframe.copy()

                # Increase sample size to 50 rows for better detection
                sample_rows = []

                # First 20 rows
                if len(df_filtered) > 0:
                    sample_rows.extend(df_filtered.head(20).to_dict('records'))

                # Random sample from middle (if we have more than 40 rows)
                if len(df_filtered) > 40:
                    middle_sample = df_filtered.iloc[20:-10].sample(n=min(20, len(df_filtered) - 30), random_state=42)
                    sample_rows.extend(middle_sample.to_dict('records'))

                # Last 10 rows (if we have more than 30 rows total)
                if len(df_filtered) > 30:
                    sample_rows.extend(df_filtered.tail(10).to_dict('records'))

                # Get basic statistics
                stats = {
                    'total_rows': len(self.dataframe),
                    'rows_with_data': len(df_filtered),
                    'non_empty_counts': {}
                }

                for col in columns:
                    non_empty = df_filtered[col].notna().sum()
                    stats['non_empty_counts'][col] = non_empty

                sheet_info = {
                    'sheet_name': self.sheet_name,
                    'columns': columns,
                    'sample_data': sample_rows,
                    'statistics': stats
                }

                # Create prompt for Claude
                prompt = f"""Analyze the following Excel sheet and its columns. Identify which columns correspond to:
1. MFG (Manufacturer name) - Look for manufacturer names like "Siemens", "ABB", "Schneider", etc.
2. MFG_PN (Manufacturer Part Number) - The primary part number from the manufacturer
3. MFG_PN_2 (Secondary/alternative Manufacturer Part Number) - An alternative or backup part number
4. Part_Number (Internal part number) - Internal reference numbers
5. Description (Part description) - Text description of the part

Here is the sheet with its columns, sample data (up to 50 rows), and statistics:

{json.dumps(sheet_info, indent=2, default=str)}

Note: Rows with little to no information (less than 30% of columns filled) have been filtered out.

Analyze the sample data carefully. Look at:
- Column names (they might have hints like "Mfg", "Manufacturer", "PN", "Part", "Description", etc.)
- Data patterns (manufacturer names vs part numbers vs descriptions)
- Data completeness (statistics show total_rows, rows_with_data after filtering, and non_empty_counts per column)
- Data consistency across the sample rows

Return a JSON object with the mapping and confidence scores (0-100). Base confidence on:
- How well the column name matches the expected field
- How consistent the data pattern is with the expected field type
- How complete the data is (columns with mostly empty values should have lower confidence)

Format:
{{
  "{self.sheet_name}": {{
    "MFG": {{"column": "column_name or null", "confidence": 0-100}},
    "MFG_PN": {{"column": "column_name or null", "confidence": 0-100}},
    "MFG_PN_2": {{"column": "column_name or null", "confidence": 0-100}},
    "Part_Number": {{"column": "column_name or null", "confidence": 0-100}},
    "Description": {{"column": "column_name or null", "confidence": 0-100}}
  }}
}}

Only return the JSON, no other text."""

                # Call Claude API with selected model
                response = client.messages.create(
                    model=self.model,
                    max_tokens=4096,
                    messages=[{"role": "user", "content": prompt}]
                )

                # Parse response
                response_text = response.content[0].text.strip()
                if response_text.startswith('```'):
                    response_text = response_text.split('```')[1]
                    if response_text.startswith('json'):
                        response_text = response_text[4:]
                    response_text = response_text.strip()

                mapping = json.loads(response_text)

                # Emit the mapping for this sheet
                if self.sheet_name in mapping:
                    self.finished.emit(self.sheet_name, mapping[self.sheet_name])
                else:
                    self.error.emit(self.sheet_name, "Sheet mapping not found in response")

                # Success - exit retry loop
                break

            except Exception as e:
                error_str = str(e)

                # Check if it's a rate limit error (429)
                is_rate_limit = '429' in error_str or 'rate_limit' in error_str.lower() or 'overloaded' in error_str.lower()

                if is_rate_limit and retry_count < self.max_retries:
                    # Calculate exponential backoff delay
                    delay = base_delay * (2 ** retry_count)  # 10s, 20s, 40s, 80s, 160s
                    retry_count += 1

                    # Sleep and retry
                    time.sleep(delay)
                    continue  # Retry the request
                else:
                    # Not a rate limit error, or max retries reached
                    if retry_count >= self.max_retries:
                        self.error.emit(self.sheet_name, f"Max retries ({self.max_retries}) exceeded. Last error: {error_str}")
                    else:
                        self.error.emit(self.sheet_name, error_str)
                    break


class AIDetectionThread(QThread):
    """Coordinator thread for parallel AI column detection across all sheets"""
    progress = pyqtSignal(str, int, int)  # message, current, total
    finished = pyqtSignal(dict)  # mappings
    error = pyqtSignal(str)

    def __init__(self, api_key, dataframes, model="claude-sonnet-4-5-20250929"):
        super().__init__()
        self.api_key = api_key
        self.dataframes = dataframes
        self.model = model
        self.all_mappings = {}
        self.completed_count = 0
        self.error_count = 0
        self.workers = []

    def run(self):
        try:
            sheet_names = list(self.dataframes.keys())
            total_sheets = len(sheet_names)

            self.progress.emit(f"Starting parallel analysis of {total_sheets} sheets...", 0, total_sheets)

            # Create a worker for each sheet
            for sheet_name in sheet_names:
                worker = SheetDetectionWorker(
                    self.api_key,
                    sheet_name,
                    self.dataframes[sheet_name],
                    self.model
                )
                worker.finished.connect(self.on_sheet_completed)
                worker.error.connect(self.on_sheet_error)
                self.workers.append(worker)

            # Start workers with staggered delays to avoid rate limiting
            # Conservative approach: process one at a time with longer delays
            batch_size = 1  # Process one sheet at a time to avoid rate limits
            delay_between_requests = 12.0  # 12 second delay between requests (safe for most API tiers)

            for i in range(0, len(self.workers), batch_size):
                batch = self.workers[i:i + batch_size]

                # Start workers in this batch
                for worker in batch:
                    worker.start()

                # Wait for this batch to complete before starting next
                for worker in batch:
                    worker.wait()

                # If not the last batch, wait before starting next request
                if i + batch_size < len(self.workers):
                    self.progress.emit(
                        f"Rate limit protection: waiting {delay_between_requests}s before next request...",
                        self.completed_count,
                        total_sheets
                    )
                    time.sleep(delay_between_requests)

            # All workers have already completed (waited in the loop above)
            # No need to wait again

            # Check if we got at least some results
            if len(self.all_mappings) > 0:
                # Report summary including failures
                success_count = len(self.all_mappings)
                failed_count = self.error_count

                if failed_count > 0:
                    # Build error report
                    failed_list = getattr(self, 'failed_sheets', [])
                    error_details = "\n".join([f"  - {item['sheet']}: {item['error'][:80]}" for item in failed_list[:10]])
                    if len(failed_list) > 10:
                        error_details += f"\n  ... and {len(failed_list) - 10} more"

                    self.progress.emit(
                        f"Completed with {failed_count} errors. Successfully mapped {success_count}/{total_sheets} sheets.",
                        total_sheets,
                        total_sheets
                    )
                else:
                    self.progress.emit("All sheets mapped successfully!", total_sheets, total_sheets)

                self.finished.emit(self.all_mappings)
            else:
                self.error.emit("No sheets were successfully analyzed. Please check your API key and try again.")

        except Exception as e:
            self.error.emit(str(e))

    def on_sheet_completed(self, sheet_name, mapping):
        """Handle completion of a single sheet detection"""
        self.all_mappings[sheet_name] = mapping
        self.completed_count += 1
        total = len(self.dataframes)
        self.progress.emit(
            f"Completed {self.completed_count}/{total} sheets ('{sheet_name}')",
            self.completed_count,
            total
        )

    def on_sheet_error(self, sheet_name, error_msg):
        """Handle error from a single sheet detection"""
        self.error_count += 1
        self.completed_count += 1

        # Track failed sheet
        if not hasattr(self, 'failed_sheets'):
            self.failed_sheets = []
        self.failed_sheets.append({'sheet': sheet_name, 'error': error_msg})

        total = len(self.dataframes)
        self.progress.emit(
            f"Error on sheet '{sheet_name}': {error_msg[:50]}... ({self.completed_count}/{total})",
            self.completed_count,
            total
        )


class PartialMatchAIThread(QThread):
    """Background thread for AI-powered partial match suggestions"""
    progress = pyqtSignal(str, int, int)  # message, current, total
    part_analyzed = pyqtSignal(int, dict)  # row_idx, analysis_result
    finished = pyqtSignal(dict)  # part_number -> suggested_match_index
    error = pyqtSignal(str)

    def __init__(self, api_key, parts_needing_review, combined_data):
        super().__init__()
        self.api_key = api_key
        self.parts_needing_review = parts_needing_review
        self.combined_data = combined_data

    def run(self):
        try:
            client = Anthropic(api_key=self.api_key)
            suggestions = {}

            total = len(self.parts_needing_review)
            for idx, part in enumerate(self.parts_needing_review):
                # Skip parts with only one match - no AI needed
                if len(part['matches']) <= 1:
                    self.progress.emit(f"Skipping part {idx + 1} of {total} (only one match)...", idx + 1, total)
                    # Still mark as processed
                    self.part_analyzed.emit(idx, {'skipped': True, 'reason': 'single_match'})
                    continue

                self.progress.emit(f"Analyzing part {idx + 1} of {total}...", idx, total)

                # Get original description from combined data
                description = self.get_description_for_part(part['PartNumber'], part['ManufacturerName'])

                # Create prompt for AI
                matches_text = "\n".join([f"{i+1}. {m}" for i, m in enumerate(part['matches'])])

                prompt = f"""Analyze this electronic component and suggest the best matching part number from SupplyFrame.

Original Part:
- Part Number: {part['PartNumber']}
- Manufacturer: {part['ManufacturerName']}
- Description: {description if description else 'Not available'}

Available Matches from SupplyFrame:
{matches_text}

Instructions:
1. Compare the original part number with each match
2. Consider manufacturer variations (e.g., "EPCOS" vs "TDK Electronics")
3. Look for exact or closest part number matches
4. If the manufacturer has been acquired, prefer the current company name

Return a JSON object with:
{{
    "suggested_index": <0-based index of best match, or null if none are suitable>,
    "confidence": <0-100>,
    "reasoning": "<brief explanation>"
}}

Only return the JSON, no other text."""

                try:
                    response = client.messages.create(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=500,
                        messages=[{"role": "user", "content": prompt}]
                    )

                    response_text = response.content[0].text.strip()
                    if response_text.startswith('```'):
                        response_text = response_text.split('```')[1]
                        if response_text.startswith('json'):
                            response_text = response_text[4:]
                        response_text = response_text.strip()

                    result = json.loads(response_text)
                    suggestions[part['PartNumber']] = result

                    # Emit per-part update for real-time UI refresh
                    self.part_analyzed.emit(idx, result)

                except Exception as e:
                    # If AI fails for this part, emit error result
                    self.part_analyzed.emit(idx, {'error': str(e)})
                    continue

            self.finished.emit(suggestions)

        except Exception as e:
            self.error.emit(str(e))

    def get_description_for_part(self, part_number, mfg):
        """Find description from combined data"""
        # Convert DataFrame to list of dictionaries if needed
        data = self.combined_data
        if hasattr(data, 'to_dict'):
            data = data.to_dict('records')

        for row in data:
            if isinstance(row, dict) and row.get('MFG_PN') == part_number and row.get('MFG') == mfg:
                return row.get('Description', '')
        return ''


class ManufacturerNormalizationAIThread(QThread):
    """Background thread for pure AI manufacturer normalization"""
    progress = pyqtSignal(str)
    finished = pyqtSignal(dict, dict)  # (normalizations, reasoning_map)
    error = pyqtSignal(str)

    def __init__(self, api_key, all_manufacturers, supplyframe_manufacturers):
        super().__init__()
        self.api_key = api_key
        self.all_manufacturers = all_manufacturers
        self.supplyframe_manufacturers = supplyframe_manufacturers

    def run(self):
        try:
            normalizations = {}
            reasoning_map = {}

            # Pure AI analysis - no fuzzy matching pre-filter
            if ANTHROPIC_AVAILABLE and self.api_key:
                self.progress.emit("AI analyzing all manufacturers...")

                if self.all_manufacturers:
                    client = Anthropic(api_key=self.api_key)

                    # Create prompt for AI to analyze ALL manufacturers
                    prompt = f"""Analyze these manufacturer names and detect variations that need normalization.

SOURCE manufacturers (from user's data - these are what need normalizing):
{json.dumps(sorted(self.all_manufacturers), indent=2)}

TARGET manufacturers (PAS/SupplyFrame canonical names - normalize TO these when applicable):
{json.dumps(sorted(self.supplyframe_manufacturers), indent=2)}

Instructions:
1. ONLY create mappings for manufacturers in the SOURCE list
2. Identify variations in SOURCE that should map to canonical TARGET names
3. Examples of variations to detect:
   - Abbreviations: "TI" → "Texas Instruments"
   - Acquisitions: "EPCOS" → "TDK Electronics"
   - Alternate spellings: "STMicro" → "STMicroelectronics"
4. CRITICAL RULES:
   - If a SOURCE name already matches a TARGET name exactly, DO NOT include it
   - NEVER create reverse mappings (e.g., DO NOT map "Texas Instruments" → "TI")
   - ONLY map FROM variations TO canonical names, never the reverse
   - Direction matters: abbreviation → full name, NOT full name → abbreviation
   - For companies not in TARGET list, suggest the most complete/official name
5. Provide brief reasoning for each mapping (acquisitions, abbreviations, etc.)

Return ONLY valid JSON with this structure:
{{
    "normalizations": {{
        "TI": "Texas Instruments",
        "EPCOS": "TDK Electronics"
    }},
    "reasoning": {{
        "TI": "Common abbreviation for Texas Instruments",
        "EPCOS": "EPCOS was acquired by TDK Electronics in 2009"
    }}
}}

IMPORTANT:
- Only include entries that need normalization (skip exact matches)
- Return ONLY valid JSON, no markdown, no other text
- Ensure all quotes inside strings are escaped with backslash"""

                    response = client.messages.create(
                        model="claude-sonnet-4-5-20250929",
                        max_tokens=4096,
                        temperature=0,  # Ensure consistent results
                        messages=[{"role": "user", "content": prompt}]
                    )

                    response_text = response.content[0].text.strip()

                    # Clean up code blocks
                    if response_text.startswith('```'):
                        # Extract content between code blocks
                        parts = response_text.split('```')
                        if len(parts) >= 2:
                            response_text = parts[1]
                            # Remove 'json' language identifier if present
                            if response_text.startswith('json'):
                                response_text = response_text[4:]
                            response_text = response_text.strip()

                    # Try to parse JSON with better error handling
                    try:
                        ai_result = json.loads(response_text)
                    except json.JSONDecodeError as je:
                        # Log the error and try to extract what we can
                        self.progress.emit(f"JSON parse error at char {je.pos}: {je.msg}")

                        # Fallback: Try to find JSON object in the response
                        import re
                        json_match = re.search(r'\{[\s\S]*\}', response_text)
                        if json_match:
                            try:
                                ai_result = json.loads(json_match.group())
                            except:
                                # If all parsing fails, return empty results
                                self.progress.emit("Could not parse AI response")
                                ai_result = {"normalizations": {}, "reasoning": {}}
                        else:
                            ai_result = {"normalizations": {}, "reasoning": {}}

                    ai_normalizations = ai_result.get('normalizations', {})
                    ai_reasoning = ai_result.get('reasoning', {})

                    # Validate and filter AI results to prevent incorrect normalizations
                    validated_count = 0
                    skipped_count = 0
                    for variation, canonical in ai_normalizations.items():
                        # Only include if variation is in the source data
                        if variation not in self.all_manufacturers:
                            self.progress.emit(f"Skipping '{variation}' → '{canonical}' (not in source data)")
                            skipped_count += 1
                            continue

                        # Skip if variation already equals canonical (no change needed)
                        if variation == canonical:
                            self.progress.emit(f"Skipping '{variation}' → '{canonical}' (already identical)")
                            skipped_count += 1
                            continue

                        # CRITICAL: Ensure canonical is ONLY from PAS list
                        # This is the primary purpose of normalization - map to PAS canonical names
                        if canonical not in self.supplyframe_manufacturers:
                            self.progress.emit(f"Skipping '{variation}' → '{canonical}' (target not in PAS canonical list)")
                            skipped_count += 1
                            continue

                        # Skip reverse mappings (canonical → variation)
                        # This catches cases where AI incorrectly maps full names to abbreviations
                        if canonical in self.all_manufacturers and variation in self.supplyframe_manufacturers:
                            self.progress.emit(f"Skipping '{variation}' → '{canonical}' (appears to be reverse mapping)")
                            skipped_count += 1
                            continue

                        # Validation passed - store the normalization
                        normalizations[variation] = canonical
                        reasoning_map[variation] = {
                            'method': 'ai',
                            'reasoning': ai_reasoning.get(variation, 'AI suggested normalization')
                        }
                        validated_count += 1

                    if validated_count > 0:
                        self.progress.emit(f"AI analysis complete: {validated_count} normalizations detected, {skipped_count} skipped")
                    else:
                        self.progress.emit(f"AI analysis complete: No valid normalizations suggested ({skipped_count} skipped)")

            # Emit results
            self.finished.emit(normalizations, reasoning_map)

        except Exception as e:
            self.error.emit(str(e))


class PASSearchThread(QThread):
    """Background thread for searching parts via PAS API with parallel execution"""
    progress = pyqtSignal(str, int, int)  # message, current, total
    result_ready = pyqtSignal(dict)  # individual result for real-time display
    finished = pyqtSignal(list)  # all search results
    error = pyqtSignal(str)

    def __init__(self, pas_client, parts_data, max_workers=10):
        super().__init__()
        self.pas_client = pas_client
        self.parts_data = parts_data  # List of {'MFG': ..., 'MFG_PN': ..., 'Description': ...}
        self.max_workers = max_workers  # Number of parallel threads
        self.completed_count = 0
        self.lock = threading.Lock()

    def search_single_part(self, idx, part, total):
        """Search a single part with retry logic"""
        manufacturer = part.get('MFG', '')
        part_number = part.get('MFG_PN', '')

        # Handle NaN values from pandas (convert to empty string)
        import math
        if isinstance(manufacturer, float) and math.isnan(manufacturer):
            manufacturer = ''
        if isinstance(part_number, float) and math.isnan(part_number):
            part_number = ''

        # Convert to string and strip whitespace
        manufacturer = str(manufacturer).strip() if manufacturer else ''
        part_number = str(part_number).strip() if part_number else ''

        # Only require part_number (MFG can be empty)
        if not part_number:
            with self.lock:
                self.completed_count += 1
                self.progress.emit(f"Skipping part {self.completed_count}/{total} (missing Manufacturer PN)...", self.completed_count, total)
            return {
                'PartNumber': part_number if part_number else '(empty)',
                'ManufacturerName': manufacturer if manufacturer else '(empty)',
                'MatchStatus': 'None',
                'matches': []
            }

        with self.lock:
            self.completed_count += 1
            current = self.completed_count

        self.progress.emit(
            f"Searching Manufacturer PN {current}/{total}: {manufacturer} - {part_number}...",
            current,
            total
        )

        # Search with retry logic (like SearchAndAssignApp - 3 retries)
        match_result = None
        match_type = None
        retry_count = 0
        max_retries = 3

        while retry_count < max_retries:
            try:
                match_result, match_type = self.pas_client.search_part(part_number, manufacturer)
                break  # Success
            except Exception as e:
                retry_count += 1
                if retry_count < max_retries:
                    self.progress.emit(
                        f"Retry {retry_count}/{max_retries} for {manufacturer} {part_number}...",
                        current,
                        total
                    )
                    time.sleep(3)  # Wait 3 seconds before retry
                else:
                    match_result = {'error': str(e)}
                    match_type = 'Error'

        # Map match_type to status (using SearchAndAssign terminology)
        if match_type in ['Found', 'Multiple', 'Need user review', 'None', 'Error']:
            status = match_type
        else:
            # Legacy mapping for backwards compatibility
            if match_type == 'exact':
                status = 'Found'
            elif match_type == 'partial':
                matches = match_result.get('matches', [])
                if len(matches) > 1:
                    status = 'Multiple'
                elif len(matches) == 1:
                    status = 'Found'
                else:
                    status = 'None'
            elif match_type == 'no_match':
                status = 'None'
            else:  # error
                status = 'Error'

        result_dict = {
            'PartNumber': part_number,
            'ManufacturerName': manufacturer,
            'MatchStatus': status,
            'matches': match_result.get('matches', []) if match_type != 'Error' else []
        }

        # Emit individual result for real-time display
        self.result_ready.emit(result_dict)

        return result_dict

    def run(self):
        try:
            results = [None] * len(self.parts_data)  # Pre-allocate to maintain order
            total = len(self.parts_data)
            self.completed_count = 0

            # Use ThreadPoolExecutor for parallel execution
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                # Submit all tasks
                future_to_idx = {
                    executor.submit(self.search_single_part, idx, part, total): idx
                    for idx, part in enumerate(self.parts_data)
                }

                # Collect results as they complete
                for future in as_completed(future_to_idx):
                    idx = future_to_idx[future]
                    try:
                        result = future.result()
                        results[idx] = result
                    except Exception as e:
                        # Handle unexpected errors
                        self.progress.emit(f"Error processing part {idx + 1}: {str(e)}", idx + 1, total)
                        results[idx] = {
                            'PartNumber': self.parts_data[idx].get('MFG_PN', ''),
                            'ManufacturerName': self.parts_data[idx].get('MFG', ''),
                            'MatchStatus': 'Error',
                            'matches': []
                        }

            self.finished.emit(results)

        except Exception as e:
            self.error.emit(str(e))
