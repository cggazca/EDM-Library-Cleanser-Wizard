# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an EDM (Engineering Data Management) library processing toolkit for Xpedition/PADS layout data. The project handles conversion workflows between Microsoft Access databases (.mdb/.accdb), Excel spreadsheets, and XML formats compatible with xml-console for EDM Library Creator (v1.7.000.0130).

## Core Architecture

The codebase provides both individual command-line tools and an integrated wizard application:

### **RECOMMENDED: EDM Wizard (`edm_wizard.py`)**
**All-in-one GUI wizard** that combines all processing steps with an easy-to-use PyQt5 interface.

**Architecture**: The wizard uses PyQt5's `QWizard` framework with 6 distinct pages:
1. **StartPage** - Claude AI and PAS API configuration + output folder selection
2. **DataSourcePage** - Access DB export or Excel file selection with preview
3. **ColumnMappingPage** - Column mapping with AI-assisted detection and sheet combining
4. **PASSearchPage** - Part Aggregation Service (PAS) search (auto-loads data from Step 3)
5. **SupplyFrameReviewPage** - Review match results by category (Found/Multiple/Need Review/None) + manufacturer normalization
6. **ComparisonPage** - Old vs New comparison showing all changes made with export options

**Key UI Components**:
- `CollapsibleGroupBox` - Custom checkable QGroupBox that expands/collapses content when toggled
- Thread-based workers (`ExportThread`, `ColumnDetectionThread`, `PASSearchThread`) for long-running operations
- Tabbed interface in Step 5 for categorized match review
- Scroll areas and dynamic section expansion for better UX
- Color-coded comparison table in Step 6 (red=old, green=new)

**AI Integration** (Optional, requires Claude API key):
- Column mapping auto-detection (Step 3)
- Part number match suggestions using similarity scoring (Step 5)
- Manufacturer normalization detection (Step 5)
- Uses `anthropic` package with Claude Sonnet 4.5 model

**PAS API Integration** (Required, uses Client ID/Secret):
- Direct part search via Siemens Part Aggregation Service (Step 4)
- Implements exact SearchAndAssign matching algorithm from legacy Java tool:
  - Step 1: Search with PN + MFG (exact → partial → alphanumeric → zero suppression)
  - Step 2: Search by PN only (if MFG empty/Unknown or no matches)
- Returns match types: "Found", "Multiple", "Need user review", "None", "Error"
- Retrieves distributor availability, pricing, and lifecycle data
- Supports enriching providers for extended part information
- See `Part Aggeration Service/example.py` for standalone PAS client implementation

### Individual Command-Line Tools

The following standalone tools can be used separately if needed:

### 1. Access Database Export (`AccessToExcel/AccessToExcel.py`)
Extracts data from Microsoft Access databases into Excel format.
- **Input**: `.mdb` or `.accdb` files (e.g., `ELESdxdb.mdb`)
- **Output**: Single Excel workbook with multiple sheets (one per table) or separate Excel files per table
- **Technology**: Uses SQLAlchemy with pyodbc and the Microsoft Access ODBC driver
- **Configuration**: Set `output_as_single_spreadsheet = True/False` to control output format

### 2. Excel Sheet Combiner (`excel_sheet_combiner.py`)
Interactive GUI tool to merge multiple Excel sheets with column selection.
- **Input**: Excel workbook with multiple sheets
- **Output**: Combined Excel file with selected columns + optional XML files
- **Technology**: Tkinter GUI with pandas backend
- **Workflow**:
  1. User selects columns to combine (shows common columns, sheet presence info)
  2. Combines data from all sheets into single sheet with `Source_Sheet` column
  3. Optionally generates MFG and MFGPN XML files via mapping dialog

### 3. XML Generator (`generate_xml_from_excel.py`)
Command-line tool to generate XML files for xml-console from Excel data.
- **Input**: Excel file (preferably combined format with `MFG` and `MFG PN` columns)
- **Output**: Two XML files (MFG and MFGPN)
- **Usage**: `python generate_xml_from_excel.py [excel_file]`

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

### Run EDM Wizard (All-in-One - Recommended)
```bash
python edm_wizard.py
```

### Build Standalone Executable
```bash
build_exe.bat
# Output: dist\EDM_Library_Wizard.exe
```

Or manually:
```bash
pip install -r requirements_wizard.txt
pyinstaller --onefile --windowed --name "EDM_Library_Wizard" edm_wizard.py
```

### Alternative: Individual Command-Line Tools

#### Export Access Database to Excel
```bash
python AccessToExcel/AccessToExcel.py "path/to/database.mdb"
```

#### Combine Excel Sheets (GUI)
```bash
python excel_sheet_combiner.py "path/to/excel_file.xlsx"
```

#### Generate XML Files
```bash
python generate_xml_from_excel.py "path/to/combined_excel.xlsx"
```

### Run Tests
```bash
python AccessToExcel/Unit_Test.py
```

## Data Flow

### EDM Wizard (Primary Workflow)
**Modern PAS API-based workflow** (XML generation removed):
Access DB/Excel → Column Mapping → PAS API Search → Review Matches → Normalize → Compare Changes

1. **Step 1**: Configure Claude AI (optional) and PAS API credentials (required) + select output folder
2. **Step 2**: Import from Access DB or Excel file
3. **Step 3**: Map MFG/MFG PN columns (AI-assisted) and combine data
4. **Step 4**: Auto-loads data from Step 3, searches parts via PAS API using SearchAndAssign algorithm
5. **Step 5**: Review match results in tabs (Found/Multiple/Need Review/None), normalize manufacturer names
6. **Step 6**: Review old vs new comparison, export changes to CSV/Excel

**Key Changes from Legacy**:
- No XML generation (removed XMLGenerationPage)
- Data flows automatically between pages (no manual CSV loading)
- Output folder configured once in Step 1
- Match results categorized using exact SearchAndAssign algorithm
- Final comparison page shows all changes before finishing

### Individual Tools Workflow
1. **MS Access DB** → `AccessToExcel.py` → **Excel (Multiple Sheets)**
2. **Excel (Multiple Sheets)** → `excel_sheet_combiner.py` → **Excel (Combined Single Sheet)**
3. **Excel (Combined)** → `generate_xml_from_excel.py` → **MFG.xml + MFGPN.xml**

Note: Steps 2 & 3 can be combined via `excel_sheet_combiner.py`'s integrated XML generation feature.

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

### For EDM Wizard
Install all dependencies at once:
```bash
pip install -r requirements_wizard.txt
```

Or individually:
```
pandas>=2.0.0
sqlalchemy>=2.0.0
xlsxwriter>=3.0.0
pyodbc>=4.0.0
PyQt5>=5.15.0
pyinstaller>=5.0.0  # Only needed for building executable
anthropic>=0.39.0  # For AI features
fuzzywuzzy>=0.18.0  # For manufacturer normalization
python-Levenshtein>=0.27.0  # For fuzzy matching performance
requests>=2.31.0  # For PAS API calls
```

### For Individual Tools
```
pandas
sqlalchemy
xlsxwriter
pyodbc
```

### Windows-Specific Requirement
- Microsoft Access Database Engine (ODBC Driver) for `.mdb`/`.accdb` file access
- Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920

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
- **ExportThread**: Handles Access DB to Excel export without freezing UI
- **ColumnDetectionThread**: Runs AI column detection in background
- **PASSearchThread**: Performs batch part searches via PAS API
- All workers emit signals (`progress`, `finished`, `error`) for UI updates

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

### AI Features (Optional)
- **Column Detection**: Analyzes first 10 rows to suggest MFG/MFG PN columns
- **Part Matching**: Uses difflib similarity + Claude AI for intelligent part number matching
- **Manufacturer Normalization**: Hybrid approach using fuzzy matching + Claude AI
  - First tries fuzzy matching against PAS canonical manufacturer names
  - Falls back to Claude AI for ambiguous cases
  - Provides reasoning for normalization suggestions
- All AI features gracefully degrade if API key not provided or `anthropic` package not installed

### PAS API Features (Optional)
- **Part Search**: Search for parts using manufacturer + part number
- **Match Types**: Returns exact, partial, or no_match results
- **Enriching Data**: Retrieves distributor info, pricing tiers, stock, lead times
- **Supply Chain Data**: Lifecycle status, risk rank, authorized distributors
- **Batch Processing**: Searches multiple parts with progress tracking
- Configured via bearer token from Siemens OAuth service

## Utility Scripts

- `AccessToExcel/Summary.py`: Generates manufacturer statistics and catalog summaries from exported Excel files
- `AccessToExcel/Unit_Test.py`: Test suite for the AccessToExcel functionality
- `Part Aggeration Service/example.py`: Standalone PAS API client with batch CSV processing and HTML report generation

## PAS API Architecture

The PAS (Part Aggregation Service) integration provides a modern alternative to static XML generation:

### PASAPIClient Class (`Part Aggeration Service/example.py`)
- **Authentication**: OAuth 2.0 client credentials flow with automatic token refresh
- **Search Endpoint**: `/api/v2/search-providers/{providerId}/{version}/free-text/search`
- **Enriching**: Configurable enriching providers for extended data (Supply Chain enricher ID: 33)
- **Output Formats**: Excel with color-coded match types, HTML report with interactive filtering, raw JSON responses

### Integration in EDM Wizard
- **PASSearchPage**: Batch searches parts from combined data with progress bar
- **PASAPIClient instance**: Shared across wizard for authentication and searches
- **Match Review**: SupplyFrameReviewPage displays search results with distributor data
- **Fallback**: If PAS search is skipped, wizard falls back to legacy XML generation

### Configuration
- Bearer token stored in `QSettings` for persistence across sessions
- API endpoint: `https://api.pas.partquest.com`
- Auth service: `https://samauth.us-east-1.sws.siemens.com/`
- Search provider ID: 44 (default)
- Supply Chain enricher ID: 33 (version 1)
