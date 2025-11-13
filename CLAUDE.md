# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an EDM (Engineering Data Management) library processing toolkit for Xpedition/PADS layout data. The project handles conversion workflows between Microsoft Access databases (.mdb/.accdb), Excel spreadsheets, and XML formats compatible with xml-console for EDM Library Creator (v1.7.000.0130).

## Core Architecture

The codebase provides both individual command-line tools and an integrated wizard application:

### **RECOMMENDED: EDM Wizard (`edm_wizard.py`)**
**All-in-one GUI wizard** that combines all processing steps with an easy-to-use PyQt5 interface.

**Architecture**: The wizard uses PyQt5's `QWizard` framework with 5 distinct pages:
1. **StartPage** - Claude AI API key configuration (optional, enables AI features)
2. **DataSourcePage** - Access DB export or Excel file selection with preview
3. **ColumnMappingPage** - Column mapping with AI-assisted detection and sheet combining
4. **XMLGenerationPage** - XML file generation with project settings
5. **SupplyFrameReviewPage** - SupplyFrame match review and manufacturer normalization

**Key UI Components**:
- `CollapsibleGroupBox` - Custom checkable QGroupBox that expands/collapses content when toggled
- Thread-based workers (`ExportThread`, `ColumnDetectionThread`) for long-running operations
- Scroll areas and dynamic section expansion for better UX on Step 4

**AI Integration** (Optional, requires Claude API key):
- Column mapping auto-detection (Step 2)
- Part number match suggestions using similarity scoring (Step 4)
- Manufacturer normalization detection (Step 4)
- Uses `anthropic` package with Claude Sonnet 4.5 model

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

### EDM Wizard (Recommended)
**All-in-one workflow**: Access DB/Excel → Column Mapping → Combine (optional) → XML Generation

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
- Both emit signals (`progress`, `finished`, `error`) for UI updates

### UI Patterns
- **CollapsibleGroupBox**: Custom widget for Step 4's 5 sections, auto-expands when data is ready
- **Dynamic validation**: Wizard pages use `validatePage()` override to control Next/Finish button state
- **Scroll areas**: Large content areas (Step 4) wrapped in `QScrollArea` for responsiveness

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
- **Manufacturer Normalization**: Detects variations (e.g., "Texas Instruments" vs "TI")
- All AI features gracefully degrade if API key not provided or `anthropic` package not installed

## Utility Scripts

- `AccessToExcel/Summary.py`: Generates manufacturer statistics and catalog summaries from exported Excel files
- `AccessToExcel/Unit_Test.py`: Test suite for the AccessToExcel functionality
