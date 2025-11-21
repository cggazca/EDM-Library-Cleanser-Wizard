# Quick Start Guide - EDM Library Wizard

Get up and running with the EDM Library Wizard in 3 simple steps!

## Prerequisites

1. **Install Python 3.8+** (if not already installed)
   - Download from: https://www.python.org/downloads/

2. **Install Microsoft Access Database Engine** (for .mdb/.accdb files)
   - Download from: https://www.microsoft.com/en-us/download/details.aspx?id=54920

## Installation

Open Command Prompt and navigate to the project folder, then run:

```bash
pip install -r requirements_wizard.txt
```

## Running the Wizard

### Option 1: Run from Python (Recommended for first-time use)
```bash
python edm_wizard.py
```

### Option 2: Create and Run Executable
```bash
# Build the executable (one-time setup)
build_exe.bat

# Run the executable
dist\EDM_Library_Wizard.exe
```

### Option 3: Run headless for automation/tests
```bash
# Add pytest for automation helpers
pip install -r requirements_test.txt
python -m edm_wizard.utils.headless_flow --config path\to\headless_config.json
pytest tests/test_headless_flow.py
```
Config JSON keys: `input_path` (Access `.mdb/.accdb` or Excel), `output_dir`, `mappings` per sheet (`MFG`, `MFG_PN`, `MFG_PN_2`, `Part_Number`, `Description`) plus optional `include_sheets`, `filters`, and `fill_tbd`. See `tests/test_headless_flow.py` for a working example.

## Using the Wizard

### Step 1: Choose Your Data Source
1. Select one option:
   - **Convert Access Database**: Browse to your `.mdb` or `.accdb` file, click "Export Access Database"
   - **Use Existing Excel**: Browse to your `.xlsx` file
2. Review the preview of your data
3. Click **Next**

### Step 2: Map Your Columns
1. **Quick Setup with Bulk Assign** (optional):
   - Select a column type (e.g., "MFG")
   - Select the column name (e.g., "Manufacturer Name")
   - Click "Apply to All Sheets"
   - Repeat for MFG PN, Description, etc.

2. **Review and Adjust Per Sheet**:
   - Check/uncheck "Include" checkbox for sheets you want to process
   - Click on any sheet row to preview its data
   - Adjust column mappings as needed using dropdowns:
     - **MFG Column**: Manufacturer name (required)
     - **MFG PN Column**: Manufacturer part number (required)
     - **Description Column**: Part description (optional, improves XML quality)
     - **Part Number Column**: For future use (optional)

3. **Optional**: Check "Combine selected sheets into single 'Combined' sheet"
   - If combining, select which columns must not be empty (filter conditions)
   - Only sheets with "Include" checked will be combined

4. Click **Next**

### Step 3: Generate XML Files
1. Review/edit Project Name (default: "VarTrainingLab")
2. Review/edit Catalog (default: "VV")
3. Review/edit Output Location (default: same as Excel file)
4. Check/uncheck "Set MFG to 'TBD' if empty" option
5. Click **"Generate XML Files"**
6. Review the summary
7. Click **Finish**

## Output Files

You'll find these files in your output location:

- **{filename}_MFG.xml** - Manufacturer database (Class 090)
- **{filename}_MFGPN.xml** - Part number database (Class 060)
- **{filename}.xlsx** - Original Excel file (with "Combined" sheet if you chose to combine)

## Example Workflow

1. Start wizard: `python edm_wizard.py`
2. Select "Convert Access Database to Excel"
3. Browse to `ELESdxdb.mdb`
4. Click "Export Access Database"
5. Review preview → Click "Next"
6. Map columns for each sheet (MFG, MFG PN, Description)
7. Check "Combine sheets" and select filter: "MFG PN must not be empty"
8. Click "Next"
9. Keep defaults (Project: VarTrainingLab, Catalog: VV)
10. Click "Generate XML Files"
11. Review summary → Click "Finish"
12. Done! Find your XML files next to the Excel file

## Troubleshooting

**Problem**: "PyQt5 not found"
**Solution**: Run `pip install PyQt5`

**Problem**: "ODBC Driver not found"
**Solution**: Install Microsoft Access Database Engine (see Prerequisites)

**Problem**: Wizard window is too small
**Solution**: Resize the window or maximize it

**Problem**: No data in preview
**Solution**: Ensure your Excel file has data and is not corrupted

## Need More Help?

- See `README_WIZARD.md` for detailed documentation
- See `CLAUDE.md` for technical details about the data processing
- For the individual command-line tools, see the relevant Python scripts

## Tips for Best Results

1. **Use descriptive column names** in your source data (helps with mapping)
2. **Fill in Description column** if available (improves XML quality)
3. **Combine sheets** if you have the same data structure across multiple sheets
4. **Use filter conditions** when combining to exclude incomplete records
5. **Enable TBD option** to handle parts without manufacturer information
6. **Review the preview** before proceeding to ensure data loaded correctly

Happy processing!
