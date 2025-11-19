# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

EDM (Engineering Data Management) library processing toolkit for Xpedition/PADS layout data. Converts between Microsoft Access databases (.mdb/.accdb), SQLite databases (.db/.sqlite/.sqlite3), Excel spreadsheets, and XML formats compatible with xml-console for EDM Library Creator (v1.7.000.0130).

**Primary tool**: `edm_wizard_refactored.py` - PyQt5-based wizard with AI-powered column mapping and PAS API integration for part search and normalization.

## Core Architecture

The codebase provides both individual command-line tools and an integrated wizard application:

### **RECOMMENDED: EDM Wizard (`edm_wizard_refactored.py`)**
**All-in-one GUI wizard** that combines all processing steps with an easy-to-use PyQt5 interface using a modular architecture.

**Architecture**: The wizard uses PyQt5's `QWizard` framework with 6 distinct pages:
1. **StartPage** - Claude AI and PAS API configuration + output folder selection
2. **DataSourcePage** - Access DB, SQLite DB, or Excel file selection with preview (auto-detects by file extension)
3. **ColumnMappingPage** - Column mapping with AI-assisted parallel detection and sheet combining
4. **PASSearchPage** - Part Aggregation Service (PAS) search (auto-loads data from Step 3)
5. **SupplyFrameReviewPage** - Review match results by category (Found/Multiple/Need Review/None) + manufacturer normalization
6. **ComparisonPage** - Old vs New comparison showing all changes made with export options

**Key UI Components**:
- `CollapsibleGroupBox` - Custom checkable QGroupBox that expands/collapses content when toggled
- Thread-based workers (`AccessExportThread`, `SQLiteExportThread`, `SheetDetectionWorker`, `PASSearchThread`) for long-running operations
- Parallel AI detection - Each sheet analyzed concurrently for faster processing
- Tabbed interface in Step 5 for categorized match review
- Scroll areas and dynamic section expansion for better UX
- Color-coded comparison table in Step 6 (red=old, green=new)

**AI Integration** (Optional, requires Claude API key):
- Column mapping auto-detection (Step 3) - **Parallelized for all sheets simultaneously**
- Part number match suggestions using similarity scoring (Step 5)
- Manufacturer normalization detection (Step 5)
- Uses `anthropic` package with Claude Sonnet 4.5, Haiku 4.5, or Opus 4.1 models
- Model selection available: Sonnet 4.5 (recommended), Haiku 4.5 (fastest), Opus 4.1 (most capable)

**PAS API Integration** (Optional, uses Client ID/Secret):
- Direct part search via Siemens Part Aggregation Service (Step 4)
- OAuth 2.0 client credentials flow with automatic token refresh
- Implements SearchAndAssign matching algorithm from legacy Java tool:
  - Step 1: Search with PN + MFG (exact → partial → alphanumeric → zero suppression)
  - Step 2: Search by PN only (if MFG empty/Unknown or no matches)
- Returns match types: "Found", "Multiple", "Need user review", "None", "Error"
- Retrieves distributor availability, pricing, and lifecycle data
- Supports enriching providers (Supply Chain enricher ID: 33, version 1)
- If PAS API is not configured, wizard can skip to legacy XML generation
- See `Part Aggeration Service/example.py` for standalone PAS client implementation

### Individual Command-Line Tools

The following standalone tools can be used separately if needed:

### 1. Access Database Export (`AccessToExcel/AccessToExcel.py`)
Standalone command-line tool to extract Microsoft Access databases to Excel.
- **Input**: `.mdb` or `.accdb` files
- **Output**: Single Excel workbook with multiple sheets (one per table) OR separate files per table
- **Technology**: SQLAlchemy with pyodbc and Microsoft Access ODBC driver
- **Configuration**: Set `output_as_single_spreadsheet = True/False` in script to control output format
- **Usage**: `python AccessToExcel/AccessToExcel.py "path/to/database.mdb"`

### 2. Excel Sheet Combiner (`excel_sheet_combiner.py`)
Standalone Tkinter GUI tool to merge multiple Excel sheets with column selection.
- **Input**: Excel workbook with multiple sheets
- **Output**: Combined Excel file + optional XML files (MFG and MFGPN)
- **Workflow**:
  1. Select columns to combine (shows common columns, sheet presence indicators)
  2. Combines data from all sheets into single sheet with `Source_Sheet` column
  3. Optional: Generate XML files via integrated mapping dialog
- **Usage**: `python excel_sheet_combiner.py "path/to/excel_file.xlsx"`

### 3. XML Generator (`generate_xml_from_excel.py`)
Standalone command-line tool to generate XML files for xml-console.
- **Input**: Excel file with `MFG` and `MFG PN` columns
- **Output**: Two XML files (MFG and MFGPN)
- **Usage**: `python generate_xml_from_excel.py [excel_file]`
- **Note**: Steps 2 & 3 can be combined via `excel_sheet_combiner.py`'s integrated XML generation

### 4. Utility Scripts
- **`AccessToExcel/Summary.py`**: Generates manufacturer statistics and catalog summaries from exported Excel files
- **`AccessToExcel/Unit_Test.py`**: Test suite for AccessToExcel functionality
- **`Part Aggeration Service/example.py`**: Standalone PAS API client with CSV batch processing and HTML report generation

## XML Format Specification

### MFG XML (Manufacturers - Class 090)
```xml
<object objectid="MANUFACTURER_NAME" catalog="VV" class="090">
  <field id="090obj_skn">VV</field>
  <field id="090obj_id">MANUFACTURER_NAME</field>
  <field id="090her_name">MANUFACTURER_NAME</field>
</object>
```

### MFGPN XML (Manufacturer Part Numbers - Class 060)
```xml
<object objectid="MFG:PARTNUMBER" class="060">
  <field id="060partnumber">PARTNUMBER</field>
  <field id="060mfgref">MFG</field>
  <field id="060komp_name">This is the PN description.</field>
</object>
```

## Development Commands

### Setup
```bash
# Install all dependencies
pip install -r requirements_wizard.txt

# Minimal install (for individual tools only)
pip install pandas sqlalchemy xlsxwriter pyodbc
```

### Run EDM Wizard (Recommended)
```bash
python edm_wizard_refactored.py
```

**Note**: The refactored version uses a modular architecture with separate page modules. The old monolithic `edm_wizard.py` file (351KB) is deprecated and should not be used.

### Build Standalone Executable
```bash
# Using batch file (Windows)
build_exe.bat

# Manual build
pip install pyinstaller
pyinstaller --onefile --windowed --name "EDM_Library_Wizard" edm_wizard_refactored.py

# Output: dist\EDM_Library_Wizard.exe
```

### Run Individual Tools
```bash
# Export Access database to Excel
python AccessToExcel/AccessToExcel.py "path/to/database.mdb"

# Combine Excel sheets (Tkinter GUI)
python excel_sheet_combiner.py "path/to/excel_file.xlsx"

# Generate XML files from Excel
python generate_xml_from_excel.py "path/to/combined_excel.xlsx"

# Generate statistics from Excel export
python AccessToExcel/Summary.py "path/to/excel_file.xlsx"

# Standalone PAS API client (CSV batch processing)
python "Part Aggeration Service/example.py"
```

### Testing
```bash
# Run Access export tests
python AccessToExcel/Unit_Test.py
```

## Data Flow

### EDM Wizard (Primary Workflow)
**Modern PAS API-based workflow with optional XML generation fallback**:
```
Access DB/SQLite/Excel → Column Mapping (AI) → PAS Search (optional) → Review/Normalize → Compare → Export
```

**Wizard Pages** (6 steps):
1. **StartPage**: Configure Claude AI key (optional), PAS API credentials (optional), output folder
2. **DataSourcePage**: Select Access DB (.mdb/.accdb), SQLite DB (.db/.sqlite/.sqlite3), or Excel file
   - Access/SQLite: Export to Excel with preview (threaded operation)
   - Excel: Direct import with preview
3. **ColumnMappingPage**: Map MFG/MFG PN columns across all sheets
   - AI-assisted parallel detection (analyzes all sheets concurrently)
   - Manual mapping via dropdowns
   - Optional: Combine selected sheets into single sheet
4. **PASSearchPage**: Batch part search via PAS API (auto-loads from Step 3)
   - SearchAndAssign algorithm (PN+MFG → PN only fallback)
   - Skippable if PAS not configured
5. **SupplyFrameReviewPage**: Review matches in tabbed interface
   - Categories: Found / Multiple / Need Review / None
   - Manufacturer normalization (fuzzy + AI hybrid)
   - CollapsibleGroupBox sections expand when data ready
6. **ComparisonPage**: Old vs New comparison table
   - Color-coded changes (red=old, green=new)
   - Export to CSV/Excel

**Data Flow**:
- Data flows automatically via wizard page attributes: `self.wizard().page(index).attribute_name`
- API credentials stored in `QSettings` for persistence
- All long operations (export, AI detection, PAS search) use QThread workers

**Fallback**: If PAS API skipped, wizard can fall back to legacy XML generation workflow

### Individual Tools Workflow (Legacy)
```
MS Access DB → AccessToExcel.py → Excel (Multi-Sheet)
Excel (Multi-Sheet) → excel_sheet_combiner.py → Excel (Combined) + [Optional: MFG.xml + MFGPN.xml]
Excel (Combined) → generate_xml_from_excel.py → MFG.xml + MFGPN.xml
```

## Key Data Structures

### Expected Excel Columns
- `MFG` or similar (Manufacturer Name/FIRM)
- `MFG PN` or similar (Manufacturer Part Number/DEVICE)
- `Source_Sheet` (added automatically by combiner)

### Project Configuration
- **Project Name**: Default "VarTrainingLab" (configurable)
- **Catalog**: Default "VV" (configurable)
- **EDM Library Creator Version**: v1.7.000.0130

## Dependencies

### EDM Wizard (Full Installation)
```bash
pip install -r requirements_wizard.txt
```

**Core dependencies** (required):
- `pandas>=2.0.0` - Data manipulation
- `sqlalchemy>=2.0.0` - Database connectivity
- `xlsxwriter>=3.0.0` - Excel file generation
- `pyodbc>=4.0.0` - Access database ODBC driver
- `PyQt5>=5.15.0` - GUI framework

**Optional dependencies**:
- `anthropic>=0.39.0` - AI column detection and normalization
- `fuzzywuzzy>=0.18.0` - Fuzzy manufacturer name matching
- `python-Levenshtein>=0.27.0` - Performance boost for fuzzywuzzy
- `requests>=2.31.0` - PAS API communication

**Build tools** (optional):
- `pyinstaller>=5.0.0` - Create standalone executable

### Individual Tools (Minimal)
```bash
pip install pandas sqlalchemy xlsxwriter pyodbc
```

### Platform-Specific Requirements

**Windows** (for Access database support):
- Microsoft Access Database Engine (ODBC Driver)
- Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920
- Required for `.mdb` and `.accdb` files

**SQLite support**: Built into Python (no additional driver needed)

## Database Configuration Files

The `DB_config/` directory contains Xpedition database configuration files (`.dbc`):
- `Xpedition_ELESdxdb.dbc` - Base configuration
- `Xpedition_ELESdxdbINT.dbc` - International variant
- `Xpedition_ELESdxdbINT_PART_SPEC_MAG.dbc` - With part specifications
- Backup configurations in `00_BK/` subdirectory

## File Naming Conventions

Generated files follow these patterns:
- Combined Excel: `{original_name}_combined.xlsx`
- MFG XML: `{base_name}_MFG.xml`
- MFGPN XML: `{base_name}_MFGPN.xml`

## Important Implementation Details

### PyQt5 Wizard State Management
- Data flows between wizard pages via attributes set on page objects
- Access previous page data: `self.wizard().page(page_index).attribute_name`
- Example: Step 4 accesses Step 3's `combined_data` via `self.wizard().page(3).combined_data`
- API key is stored in `QSettings` for persistence and passed through wizard pages

### Threading for Long Operations
All long-running operations use QThread workers to prevent UI freezing:
- **AccessExportThread**: Access DB → Excel export with progress updates
- **SQLiteExportThread**: SQLite DB → Excel export with progress updates
- **SheetDetectionWorker**: AI-powered column detection (parallelized per sheet)
- **PASSearchThread**: Batch PAS API searches with progress tracking

**Pattern**: Workers emit signals (`progress`, `finished`, `error`) → UI updates in main thread

### UI Patterns
- **CollapsibleGroupBox**: Custom widget for Step 5's sections, auto-expands when data is ready
- **Dynamic validation**: Wizard pages use `validatePage()` override to control Next/Finish button state
- **Scroll areas**: Large content areas (Steps 4 & 5) wrapped in `QScrollArea` for responsiveness
- **Context menus**: Right-click menus on tables for data manipulation (copy, export, etc.)

### XML Escaping
All XML generators properly escape special characters (`&`, `<`, `>`, `"`, `'`) using dedicated `escape_xml()` functions.

### Excel Sheet Names
Sheet names are cleaned to meet Excel's requirements:
- Maximum 31 characters
- No special characters: `\ / * ? : [ ]`

### Duplicate Handling
MFGPN XML generation automatically removes duplicate MFG:PN combinations before export.

### AI Features (Optional - requires Claude API key)
Gracefully degrades if `anthropic` package not installed or API key not provided.

**Column Detection** (Step 3 - ColumnMappingPage):
- Analyzes first 10 rows of each sheet in parallel
- Suggests MFG and MFG PN column mappings
- Model selection: Sonnet 4.5 (default), Haiku 4.5 (fastest), Opus 4.1 (most capable)

**Manufacturer Normalization** (Step 5 - SupplyFrameReviewPage):
- Hybrid approach: Fuzzy matching (fuzzywuzzy) → Claude AI fallback
- First tries fuzzy match against PAS canonical manufacturer names
- Falls back to Claude AI for ambiguous cases
- Returns normalized name + reasoning

**Part Number Matching**:
- Uses difflib similarity + Claude AI for intelligent suggestions
- Helps resolve "Multiple" and "Need Review" matches

### PAS API Features (Optional - requires Client ID/Secret)

**Authentication**:
- OAuth 2.0 client credentials flow
- Automatic token refresh (2-hour expiration)
- Token cached in `QSettings` for session persistence

**Part Search** (Step 4 - PASSearchPage):
- Batch processing with progress tracking
- SearchAndAssign algorithm (exact → partial → alphanumeric → zero suppression)
- Match types: "Found", "Multiple", "Need user review", "None", "Error"

**Enriching Data**:
- Distributor information (pricing tiers, stock, lead times)
- Supply Chain data (lifecycle status, risk rank, authorized distributors)
- Configurable enriching providers (default: Supply Chain enricher ID 33, version 1)

**Configuration**:
- API endpoint: `https://api.pas.partquest.com`
- Auth service: `https://samauth.us-east-1.sws.siemens.com/`
- Search provider ID: 44 (default)

**Fallback**: If PAS API not configured or fails, wizard can skip to legacy XML generation

## PAS API Architecture

The standalone `PASAPIClient` class (`Part Aggeration Service/example.py`) provides a reference implementation for PAS integration. The EDM Wizard implements similar functionality directly.

### PASAPIClient Class Features
- **Authentication**: OAuth 2.0 client credentials flow with automatic token refresh
- **Search Endpoint**: `/api/v2/search-providers/{providerId}/{version}/free-text/search`
- **Enriching Providers**: Configurable providers for extended data (e.g., Supply Chain enricher ID: 33)
- **Batch Processing**: CSV input → Excel/HTML output with color-coded match types
- **Output Formats**:
  - Excel with color-coded match types
  - HTML report with interactive filtering
  - Raw JSON responses for debugging

### EDM Wizard Integration
The wizard replicates PASAPIClient functionality inline:
- **StartPage**: Captures Client ID/Secret, stores in `QSettings`
- **PASSearchThread**: Performs OAuth authentication + batch searches in background
- **PASSearchPage**: Displays progress bar and status
- **SupplyFrameReviewPage**: Displays categorized results (Found/Multiple/Need Review/None)

### Configuration
- Client credentials stored in `QSettings` for persistence
- API endpoint: `https://api.pas.partquest.com`
- Auth service: `https://samauth.us-east-1.sws.siemens.com/`
- Search provider ID: 44 (default)
- Supply Chain enricher: ID 33, version 1

## Code Architecture Patterns

### Wizard Page Communication
Pages share data via wizard page attributes accessed using page indices:
```python
# In Step 4, access Step 3's data
step3_data = self.wizard().page(3).combined_data
```

### QSettings Persistence
API credentials and preferences stored using `QSettings`:
```python
settings = QSettings("VarIndustries", "EDMWizard")
settings.setValue("claude_api_key", api_key)
api_key = settings.value("claude_api_key", "")
```

### Worker Thread Pattern
All long operations follow this pattern:
1. Create QThread worker class with `pyqtSignal` for communication
2. Connect worker signals to UI update slots
3. Start worker in background thread
4. Worker emits `progress`, `finished`, `error` signals
5. UI updates in main thread via connected slots

### Database Export Pattern
Both Access and SQLite exports follow similar flow:
1. Connect to database (SQLAlchemy for Access, sqlite3 for SQLite)
2. Inspect schema to get table names
3. Query each table into pandas DataFrame
4. Write DataFrames to Excel using `ExcelWriter` with `xlsxwriter` engine
5. Return file path on success

### Column Mapping AI Pattern
Parallel detection for performance:
1. Create multiple `SheetDetectionWorker` threads (one per sheet)
2. Each worker analyzes first 10 rows independently
3. Results collected and displayed as suggestions
4. User can accept/reject/modify suggestions

## Common Modifications

### Adding New Claude AI Models
Update `StartPage.__init__()` model selector:
```python
self.model_selector.addItem("Model Name", "model-id")
```

### Changing Default XML Project Settings
Modify defaults in `XMLGenerationPage.__init__()`:
```python
self.project_input.setText("YourProjectName")
self.catalog_input.setText("YourCatalog")
```

### Adding New PAS Enriching Providers
Update enricher configuration in PAS search logic:
```python
enrichers = [{"id": "33", "version": "1"}]  # Add new enricher IDs
```

### Customizing Excel Sheet Name Cleaning
Modify `clean_sheet_name()` function to adjust character limits or replacements:
```python
def clean_sheet_name(name):
    # Excel limits: 31 chars, no special chars
    name = name[:31]
    for char in ['\\', '/', '*', '?', ':', '[', ']']:
        name = name.replace(char, '_')
    return name
```
