#!/usr/bin/env python3
"""
Excel Sheet Combiner with GUI
Allows user to select columns from multiple sheets and combine them into a new sheet
Then generates XML files for xml-console
"""

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime


class ExcelSheetCombiner:
    def __init__(self, root, excel_file_path):
        self.root = root
        self.root.title("Excel Sheet Combiner")
        self.root.geometry("900x700")
        
        self.excel_file_path = excel_file_path
        self.xl_file = pd.ExcelFile(excel_file_path)
        self.sheet_names = self.xl_file.sheet_names
        
        # Gather all columns from all sheets
        self.all_columns = set()
        self.sheet_columns = {}
        
        for sheet in self.sheet_names:
            df = pd.read_excel(excel_file_path, sheet_name=sheet)
            cols = list(df.columns)
            self.sheet_columns[sheet] = cols
            self.all_columns.update(cols)
        
        self.all_columns = sorted(list(self.all_columns))
        
        # Find common columns (present in all sheets)
        self.common_columns = set(self.sheet_columns[self.sheet_names[0]])
        for sheet in self.sheet_names[1:]:
            self.common_columns &= set(self.sheet_columns[sheet])
        self.common_columns = sorted(list(self.common_columns))
        
        # Dictionary to store checkbutton variables
        self.column_vars = {}
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Select Columns to Combine", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, pady=10)
        
        # Info frame
        info_frame = ttk.LabelFrame(main_frame, text="File Information", padding="5")
        info_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        info_text = f"File: {Path(self.excel_file_path).name}\n"
        info_text += f"Sheets: {len(self.sheet_names)} ({', '.join(self.sheet_names[:5])}"
        if len(self.sheet_names) > 5:
            info_text += f", ... +{len(self.sheet_names)-5} more"
        info_text += ")\n"
        info_text += f"Total unique columns: {len(self.all_columns)}\n"
        info_text += f"Common columns (in all sheets): {len(self.common_columns)}"
        
        info_label = ttk.Label(info_frame, text=info_text, justify=tk.LEFT)
        info_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Column selection frame with scrollbar
        selection_frame = ttk.LabelFrame(main_frame, text="Select Columns", padding="5")
        selection_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        selection_frame.columnconfigure(0, weight=1)
        selection_frame.rowconfigure(0, weight=1)
        
        # Canvas and scrollbar for column list
        canvas = tk.Canvas(selection_frame, borderwidth=0)
        scrollbar = ttk.Scrollbar(selection_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Add filter section
        filter_frame = ttk.Frame(scrollable_frame)
        filter_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(filter_frame, text="Quick Select:").grid(row=0, column=0, padx=5)
        ttk.Button(filter_frame, text="Common Columns Only", 
                  command=self.select_common).grid(row=0, column=1, padx=5)
        ttk.Button(filter_frame, text="Select All", 
                  command=self.select_all).grid(row=0, column=2, padx=5)
        ttk.Button(filter_frame, text="Deselect All", 
                  command=self.deselect_all).grid(row=0, column=3, padx=5)
        
        # Column headers
        ttk.Label(scrollable_frame, text="Column Name", 
                 font=('Arial', 10, 'bold')).grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(scrollable_frame, text="Common", 
                 font=('Arial', 10, 'bold')).grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(scrollable_frame, text="Present in Sheets", 
                 font=('Arial', 10, 'bold')).grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)
        
        # Add checkboxes for each column
        for idx, column in enumerate(self.all_columns, start=2):
            var = tk.BooleanVar(value=False)
            self.column_vars[column] = var
            
            # Checkbox
            cb = ttk.Checkbutton(scrollable_frame, variable=var, text=column)
            cb.grid(row=idx, column=0, sticky=tk.W, padx=5, pady=2)
            
            # Common indicator
            is_common = column in self.common_columns
            common_label = ttk.Label(scrollable_frame, text="✓" if is_common else "")
            common_label.grid(row=idx, column=1, padx=5, pady=2)
            
            # Sheet count
            sheets_with_column = [s for s in self.sheet_names if column in self.sheet_columns[s]]
            count_text = f"{len(sheets_with_column)}/{len(self.sheet_names)}"
            if len(sheets_with_column) <= 3:
                count_text += f" ({', '.join(sheets_with_column)})"
            
            count_label = ttk.Label(scrollable_frame, text=count_text, foreground='gray')
            count_label.grid(row=idx, column=2, sticky=tk.W, padx=5, pady=2)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, pady=10)
        
        ttk.Button(button_frame, text="Combine Selected Columns", 
                  command=self.combine_sheets).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=self.root.quit).grid(row=0, column=1, padx=5)
    
    def select_common(self):
        """Select only common columns"""
        for col, var in self.column_vars.items():
            var.set(col in self.common_columns)
    
    def select_all(self):
        """Select all columns"""
        for var in self.column_vars.values():
            var.set(True)
    
    def deselect_all(self):
        """Deselect all columns"""
        for var in self.column_vars.values():
            var.set(False)
    
    def combine_sheets(self):
        """Combine selected columns from all sheets"""
        # Get selected columns
        selected_columns = [col for col, var in self.column_vars.items() if var.get()]
        
        if not selected_columns:
            messagebox.showwarning("No Selection", "Please select at least one column.")
            return
        
        # Ask for output file name
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"{Path(self.excel_file_path).stem}_combined.xlsx"
        )
        
        if not output_file:
            return
        
        try:
            # Create combined dataframe
            combined_data = []
            
            for sheet_name in self.sheet_names:
                df = pd.read_excel(self.excel_file_path, sheet_name=sheet_name)
                
                # Add sheet name column
                df['Source_Sheet'] = sheet_name
                
                # Keep only selected columns that exist in this sheet
                available_columns = [col for col in selected_columns if col in df.columns]
                columns_to_keep = ['Source_Sheet'] + available_columns
                
                df_subset = df[columns_to_keep].copy()
                
                # Add missing columns with NaN values
                for col in selected_columns:
                    if col not in df_subset.columns:
                        df_subset[col] = pd.NA
                
                combined_data.append(df_subset)
            
            # Concatenate all dataframes
            final_df = pd.concat(combined_data, ignore_index=True)
            
            # Reorder columns: Source_Sheet first, then selected columns in order
            column_order = ['Source_Sheet'] + selected_columns
            final_df = final_df[column_order]
            
            # Save to Excel
            final_df.to_excel(output_file, index=False, sheet_name='Combined')
            
            messagebox.showinfo("Success",
                              f"Combined data saved to:\n{output_file}\n\n"
                              f"Total rows: {len(final_df)}\n"
                              f"Columns: {len(selected_columns) + 1}")

            # Ask if user wants to generate XML files
            generate_xml = messagebox.askyesno("Generate XML Files",
                                              "Do you want to generate XML files for xml-console?")

            if generate_xml:
                self.show_xml_mapping_dialog(output_file, final_df)
            else:
                self.root.quit()
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

    def show_xml_mapping_dialog(self, excel_file, df):
        """Show dialog to map columns to MFG and MFG PN"""
        # Create new window
        dialog = tk.Toplevel(self.root)
        dialog.title("XML Generation - Column Mapping")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="Map Columns for XML Generation",
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Description
        desc_label = ttk.Label(main_frame,
                              text="Select which columns correspond to Manufacturer (MFG) and Manufacturer Part Number (MFG PN):",
                              wraplength=550, justify=tk.LEFT)
        desc_label.grid(row=1, column=0, columnspan=2, pady=(0, 20))

        # Get available columns (excluding Source_Sheet)
        columns = [col for col in df.columns if col != 'Source_Sheet']

        # MFG column selection
        ttk.Label(main_frame, text="Manufacturer (MFG) Column:",
                 font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5)

        mfg_var = tk.StringVar()
        mfg_combo = ttk.Combobox(main_frame, textvariable=mfg_var, values=columns, state='readonly', width=40)
        mfg_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))

        # Try to auto-select if column name contains "MFG" or "FIRM"
        for col in columns:
            if 'MFG' in col.upper() or 'FIRM' in col.upper():
                mfg_var.set(col)
                break

        # MFG PN column selection
        ttk.Label(main_frame, text="Manufacturer Part Number (MFG PN) Column:",
                 font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky=tk.W, pady=5)

        mfgpn_var = tk.StringVar()
        mfgpn_combo = ttk.Combobox(main_frame, textvariable=mfgpn_var, values=columns, state='readonly', width=40)
        mfgpn_combo.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))

        # Try to auto-select if column name contains "PN" or "DEVICE"
        for col in columns:
            if 'PN' in col.upper() or 'PART' in col.upper() or 'DEVICE' in col.upper():
                mfgpn_var.set(col)
                break

        # Project settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Project Settings", padding="10")
        settings_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        settings_frame.columnconfigure(1, weight=1)

        ttk.Label(settings_frame, text="Project Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        project_var = tk.StringVar(value="VarTrainingLab")
        project_entry = ttk.Entry(settings_frame, textvariable=project_var, width=30)
        project_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))

        ttk.Label(settings_frame, text="Catalog:").grid(row=1, column=0, sticky=tk.W, pady=5)
        catalog_var = tk.StringVar(value="VV")
        catalog_entry = ttk.Entry(settings_frame, textvariable=catalog_var, width=30)
        catalog_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 0))

        # Output file location
        ttk.Label(settings_frame, text="Output Location:").grid(row=2, column=0, sticky=tk.W, pady=5)
        output_label = ttk.Label(settings_frame, text=str(Path(excel_file).parent),
                                foreground='gray', wraplength=400)
        output_label.grid(row=2, column=1, sticky=tk.W, pady=5, padx=(10, 0))

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=20)

        def generate_xml_files():
            mfg_column = mfg_var.get()
            mfgpn_column = mfgpn_var.get()
            project_name = project_var.get()
            catalog = catalog_var.get()

            if not mfg_column or not mfgpn_column:
                messagebox.showwarning("Missing Selection",
                                      "Please select both MFG and MFG PN columns.")
                return

            if mfg_column == mfgpn_column:
                messagebox.showwarning("Invalid Selection",
                                      "MFG and MFG PN columns must be different.")
                return

            try:
                # Generate XML files
                base_name = Path(excel_file).stem.replace('_combined', '')
                output_dir = Path(excel_file).parent

                mfg_xml = output_dir / f"{base_name}_MFG.xml"
                mfgpn_xml = output_dir / f"{base_name}_MFGPN.xml"

                # Generate MFG XML
                self.create_mfg_xml(df, mfg_column, mfg_xml, project_name, catalog)

                # Generate MFGPN XML
                self.create_mfgpn_xml(df, mfg_column, mfgpn_column, mfgpn_xml, project_name, catalog)

                messagebox.showinfo("Success",
                                   f"XML files generated successfully!\n\n"
                                   f"MFG XML: {mfg_xml.name}\n"
                                   f"MFGPN XML: {mfgpn_xml.name}\n\n"
                                   f"Location: {output_dir}")

                dialog.destroy()
                self.root.quit()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate XML files:\n{str(e)}")

        ttk.Button(button_frame, text="Generate XML Files",
                  command=generate_xml_files).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel",
                  command=lambda: [dialog.destroy(), self.root.quit()]).grid(row=0, column=1, padx=5)

    def escape_xml(self, text):
        """Escape special XML characters"""
        if pd.isna(text):
            return ""
        text = str(text)
        text = text.replace("&", "&amp;")
        text = text.replace("<", "&lt;")
        text = text.replace(">", "&gt;")
        text = text.replace('"', "&quot;")
        text = text.replace("'", "&apos;")
        return text

    def create_mfg_xml(self, df, mfg_column, output_file, project_name, catalog):
        """Create XML file for Manufacturers (class 090)"""
        # Get unique manufacturers
        manufacturers = df[mfg_column].dropna().unique()
        manufacturers = sorted([str(m).strip() for m in manufacturers if str(m).strip()])

        # Create XML structure
        root = ET.Element('data')

        for mfg in manufacturers:
            obj = ET.SubElement(root, 'object')
            obj.set('objectid', self.escape_xml(mfg))
            obj.set('catalog', catalog)
            obj.set('class', '090')

            field1 = ET.SubElement(obj, 'field')
            field1.set('id', '090obj_skn')
            field1.text = catalog

            field2 = ET.SubElement(obj, 'field')
            field2.set('id', '090obj_id')
            field2.text = self.escape_xml(mfg)

            field3 = ET.SubElement(obj, 'field')
            field3.set('id', '090her_name')
            field3.text = self.escape_xml(mfg)

        # Format and save
        self.save_xml(root, output_file, project_name)

    def create_mfgpn_xml(self, df, mfg_column, mfgpn_column, output_file, project_name, catalog):
        """Create XML file for Manufacturer Part Numbers (class 060)"""
        # Filter rows with both MFG and MFG PN
        df_filtered = df[[mfg_column, mfgpn_column]].dropna()
        df_filtered = df_filtered.drop_duplicates()

        # Create XML structure
        root = ET.Element('data')

        for idx, row in df_filtered.iterrows():
            mfg = str(row[mfg_column]).strip()
            mfg_pn = str(row[mfgpn_column]).strip()

            objectid = f"{mfg}:{mfg_pn}"

            obj = ET.SubElement(root, 'object')
            obj.set('objectid', self.escape_xml(objectid))
            obj.set('class', '060')

            field1 = ET.SubElement(obj, 'field')
            field1.set('id', '060partnumber')
            field1.text = self.escape_xml(mfg_pn)

            field2 = ET.SubElement(obj, 'field')
            field2.set('id', '060mfgref')
            field2.text = self.escape_xml(mfg)

            field3 = ET.SubElement(obj, 'field')
            field3.set('id', '060komp_name')
            field3.text = "This is the PN description."

        # Format and save
        self.save_xml(root, output_file, project_name)

    def save_xml(self, root, output_file, project_name):
        """Format and save XML file with proper comments and formatting"""
        # Convert to string with proper formatting
        xml_str = ET.tostring(root, encoding='utf-8', method='xml')
        dom = minidom.parseString(xml_str)

        # Create custom XML with comments
        comment_lines = [
            f'Created By: EDM Library Creator v1.7.000.0130',
            f'DDP Project: {project_name}',
            f'Date: {datetime.now().strftime("%-m/%-d/%Y %-I:%M:%S %p") if sys.platform != "win32" else datetime.now().strftime("%#m/%#d/%Y %#I:%M:%S %p")}'
        ]

        xml_lines = ['<?xml version="1.0" encoding="utf-8" standalone="yes"?>']
        for comment in comment_lines:
            xml_lines.append(f'<!--{comment}-->')

        # Get formatted XML (skip first line which is the XML declaration)
        formatted = dom.toprettyxml(indent='  ', encoding='utf-8').decode('utf-8')
        xml_content = '\n'.join(formatted.split('\n')[1:])

        # Write to file
        final_xml = '\n'.join(xml_lines) + '\n' + xml_content

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(final_xml)


def main():
    # Create a temporary root window for file dialog
    temp_root = tk.Tk()
    temp_root.withdraw()  # Hide the main window
    
    # Check if file path provided as argument
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        # Ask user to browse for Excel file
        excel_file = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
    
    temp_root.destroy()
    
    # Check if user cancelled or file doesn't exist
    if not excel_file:
        print("No file selected. Exiting.")
        sys.exit(0)
    
    if not Path(excel_file).exists():
        messagebox.showerror("Error", f"File '{excel_file}' not found.")
        sys.exit(1)
    
    # Create main application window
    root = tk.Tk()
    app = ExcelSheetCombiner(root, excel_file)
    root.mainloop()


if __name__ == "__main__":
    main()
