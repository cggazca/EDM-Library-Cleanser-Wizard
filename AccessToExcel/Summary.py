import pandas as pd

# Load the Excel file
file_path = r"C:\Users\z004ut2y\Downloads\Exported_Tables\PADSProfessionalVXLibrary_Access_To_Excel.xlsx"
xls = pd.ExcelFile(file_path)

# Initialize lists to store the results
manufacturer_counts = []
catalog_summary = []

# Process each sheet in the Excel file
for sheet in xls.sheet_names:
    # Load the sheet into a DataFrame
    df = xls.parse(sheet)
    
    # Check for relevant columns (Manufacturer Name and Manufacturer Part Number)
    manufacturer_col = next((col for col in df.columns if "Manufacturer Name" in col), None)
    part_number_col = next((col for col in df.columns if "Manufacturer Part Number" in col), None)

    # If the required columns exist, proceed with analysis
    if manufacturer_col and part_number_col:
        # Count the number of all part numbers per manufacturer
        counts = df.groupby(manufacturer_col)[part_number_col].count().reset_index()
        counts.columns = ["Manufacturer Name", "Total Part Numbers"]
        counts["Catalog"] = sheet  # Add the sheet name as the catalog name
        manufacturer_counts.append(counts)
        
        # Calculate summary statistics per catalog
        total_manufacturers = counts.shape[0]
        total_part_numbers = df[part_number_col].count()
        
        catalog_summary.append({
            "Catalog": sheet,
            "Total Manufacturers": total_manufacturers,
            "Total Part Numbers": total_part_numbers
        })

# Combine manufacturer counts into a single DataFrame
if manufacturer_counts:
    manufacturer_df = pd.concat(manufacturer_counts, ignore_index=True)
    manufacturer_df.to_excel("Manufacturer_Part_Numbers.xlsx", index=False)
    print("Manufacturer part numbers per manufacturer saved to Manufacturer_Part_Numbers.xlsx")

# Convert catalog summary into DataFrame
catalog_summary_df = pd.DataFrame(catalog_summary)
if not catalog_summary_df.empty:
    catalog_summary_df.to_excel("Catalog_Summary.xlsx", index=False)
    print("Catalog summary saved to Catalog_Summary.xlsx")
