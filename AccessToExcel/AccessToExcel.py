import pandas as pd
import os
import sys
import sqlalchemy as sa
import urllib
from sqlalchemy import inspect

def print_help():
    """Print help information about the script usage."""
    help_text = """
AccessToExcel - Export Microsoft Access database tables to Excel

Usage:
    python AccessToExcel.py [database_path]
    python AccessToExcel.py [-h | --help]

Arguments:
    database_path    Path to the Access database file (.mdb or .accdb)
                     If not provided, you will be prompted to enter it

Options:
    -h, --help       Show this help message and exit

Configuration:
    Edit the script to change these settings:
    - output_as_single_spreadsheet: True = one workbook with multiple sheets
                                    False = separate Excel file per table
    - output_folder: Directory where Excel files will be saved (default: 'Exported_Tables')

Examples:
    python AccessToExcel.py "C:\Path\To\Database.mdb"
    python AccessToExcel.py U:\nicest\PadsFlow\Master_comDB_v3.6.mdb
    python AccessToExcel.py

Requirements:
    - pandas
    - sqlalchemy
    - xlsxwriter
    - pyodbc
    - Microsoft Access Database Engine (ODBC Driver)
"""
    print(help_text)

# ----- Configuration -----
# Set this flag to True to output all tables into one Excel workbook (each table on a separate sheet).
# Set it to False to output each table as a separate Excel file.
output_as_single_spreadsheet = True

# Define the MDB file path and output folder
# Check if filepath is provided as command-line argument
if len(sys.argv) > 1:
    if sys.argv[1] in ['-h', '--help']:
        print_help()
        sys.exit(0)
    mdb_file = sys.argv[1]
else:
    # Prompt user for filepath if not provided
    mdb_file = input("Enter the path to the Access database (.mdb or .accdb): ").strip()
    # Remove quotes if user copied path with quotes
    mdb_file = mdb_file.strip('"').strip("'")

# Validate that the file exists
if not os.path.exists(mdb_file):
    print(f"Error: File not found at '{mdb_file}'")
    sys.exit(1)

if not (mdb_file.lower().endswith('.mdb') or mdb_file.lower().endswith('.accdb')):
    print(f"Error: File must be a .mdb or .accdb file")
    sys.exit(1)

print(f"Using database: {mdb_file}\n")
output_folder = 'Exported_Tables'

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Create a connection string for the Access database
conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=" + mdb_file
)
# URL-encode the connection string
quoted_conn_str = urllib.parse.quote_plus(conn_str)

# Create the SQLAlchemy engine; note that this requires an Access dialect.
engine = sa.create_engine(f"access+pyodbc:///?odbc_connect={quoted_conn_str}")

# Retrieve table names using SQLAlchemy's inspector
inspector = inspect(engine)
tables = inspector.get_table_names()

# Helper function to clean Excel sheet names (max 31 characters and no invalid characters)
def clean_sheet_name(name):
    # Excel sheet names cannot exceed 31 characters and cannot contain the characters: : \ / ? * [ ]
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '')
    return name[:31]

if output_as_single_spreadsheet:
    # Define the output file path for the single workbook.
    output_path = os.path.join(output_folder, "All_Tables.xlsx")
    print(f"Exporting all tables to a single workbook: {output_path}")
    
    # Create an Excel writer object
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for table in tables:
            print(f"Exporting table: {table}")
            df = pd.read_sql(f"SELECT * FROM [{table}]", engine)  # Use brackets for table names with spaces
            # Clean the sheet name to ensure compatibility with Excel
            sheet_name = clean_sheet_name(table)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Added {table} as sheet '{sheet_name}'")
        # The writer saves automatically on exit of the 'with' block.
    print(f"\nAll tables have been successfully exported to {output_path}!")
    
else:
    # Export each table as a separate Excel file.
    for table in tables:
        print(f"Exporting table: {table}")
        df = pd.read_sql(f"SELECT * FROM [{table}]", engine)  # Use brackets for table names with spaces
        output_file = os.path.join(output_folder, f"{table}.xlsx")
        df.to_excel(output_file, index=False)
        print(f"Saved {table} to {output_file}")
        
    print("\nAll tables have been successfully exported as separate files!")
