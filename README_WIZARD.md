# EDM Library Wizard

A comprehensive wizard application for converting Access databases to Excel and generating XML files for EDM Library Creator v1.7.000.0130.

## Features

### Step 1: Data Source Selection
- **Convert Access Database to Excel**: Select an `.mdb` or `.accdb` file and export all tables to a single Excel workbook
- **Use Existing Excel File**: Browse to an existing Excel file with multiple sheets
- **Preview**: View the first 100 rows of your data before proceeding

### Step 2: Column Mapping & Combine Options
- **Bulk Assign**: Quickly assign the same column to all sheets at once
  - Select column type (MFG, MFG PN, Part Number, Description)
  - Select column name from dropdown
  - Click "Apply to All Sheets"

- **Column Mapping Table**: Map columns for each sheet with:
  - **Include Checkbox**: Select which sheets to include in processing
  - **Sheet Name**: Click on any row to preview that sheet's data
  - MFG (Manufacturer) - Required for XML
  - MFG PN (Manufacturer Part Number) - Required for XML
  - Part Number - Reserved for future use
  - Description - Used in MFGPN XML instead of hardcoded text

- **Sheet Preview**: View first 100 rows of selected sheet
  - Click on any sheet in the mapping table to preview
  - See column names and sample data
  - Helps verify correct column selection

- **Combine Sheets**: Optionally combine selected sheets into a single "Combined" sheet
  - Only checked sheets will be combined

- **Flexible Filters**: Choose which columns must not be empty (ALL conditions must be met):
  - MFG must not be empty
  - MFG PN must not be empty
  - Part Number must not be empty
  - Description must not be empty

### Step 3: XML Generation
- **Project Settings**: Customize Project Name and Catalog ID
- **Output Location**: Choose where to save XML files (defaults to Excel file location)
- **TBD Option**: If MFG PN exists but MFG is empty, automatically set MFG to "TBD"
- **Generates**:
  - `{filename}_MFG.xml` - Manufacturers (Class 090)
  - `{filename}_MFGPN.xml` - Manufacturer Part Numbers (Class 060)

## Installation

### Requirements
```bash
pip install pandas sqlalchemy xlsxwriter pyodbc PyQt5
```

### Windows-Specific Requirement
- Microsoft Access Database Engine (ODBC Driver) for `.mdb`/`.accdb` file access
- Download from: https://www.microsoft.com/en-us/download/details.aspx?id=54920

## Usage

### Run from Python
```bash
python edm_wizard.py
```

### Create Executable (Windows)

1. Install PyInstaller:
```bash
pip install pyinstaller
```

2. Create executable:
```bash
pyinstaller --onefile --windowed --name "EDM_Library_Wizard" edm_wizard.py
```

3. The executable will be created in the `dist/` folder

### Alternative: Create executable with icon
```bash
pyinstaller --onefile --windowed --name "EDM_Library_Wizard" --icon=app_icon.ico edm_wizard.py
```

## Workflow

1. **Launch the wizard**
2. **Step 1**: Choose your data source
   - Convert Access DB → Excel (with preview)
   - OR use existing Excel file
3. **Step 2**: Map columns for each sheet
   - Select MFG and MFG PN columns (required)
   - Optionally select Part Number and Description columns
   - Choose whether to combine sheets
   - Set filter conditions if combining
4. **Step 3**: Configure and generate XML
   - Set Project Name and Catalog
   - Choose output location
   - Enable TBD option if needed
   - Click "Generate XML Files"
5. **Finish**: Review summary and close wizard

## Navigation

- **Next**: Proceed to next step (validates current step)
- **Back**: Return to previous step to make changes
- **Cancel**: Exit wizard without saving
- **Finish**: Complete and close wizard (only enabled after XML generation)

## Validation & Warnings

- If no sheets have both MFG and MFG PN mapped, a warning will appear (can continue anyway)
- Export progress is shown with status messages
- XML generation success is confirmed with summary statistics

## Output Files

### Excel File
If combining sheets, a new "Combined" sheet is added to the original Excel file with:
- `Source_Sheet` column indicating origin
- Standardized column names (MFG, MFG_PN, Part_Number, Description)
- Filtered rows based on selected conditions

### XML Files
- **MFG XML**: Contains unique manufacturers (Class 090)
- **MFGPN XML**: Contains manufacturer part numbers with descriptions (Class 060)
- Both files include proper XML headers with metadata (creator, project, date)

## Notes

- The wizard preserves all original sheets when combining
- Duplicate MFG/MFGPN combinations are automatically removed
- XML files use proper character escaping for special characters
- Progress is saved at each step, allowing navigation back and forth

## Troubleshooting

### "ODBC Driver not found"
Install Microsoft Access Database Engine (see Requirements above)

### "PyQt5 not found"
```bash
pip install PyQt5
```

### Wizard doesn't start
Ensure all dependencies are installed:
```bash
pip install pandas sqlalchemy xlsxwriter pyodbc PyQt5
```

## Support

For issues or questions, refer to the CLAUDE.md file for technical details about the data processing pipeline.
