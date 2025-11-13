import unittest
from unittest.mock import patch, Mock, call
import pandas as pd
import numpy as np
from Summary import *
import os

class TestSummaryAnalysis(unittest.TestCase):
    def setUp(self):
        self.sample_data = {
            'Sheet1': pd.DataFrame({
                'Manufacturer Name': ['Mfg1', 'Mfg1', 'Mfg2'],
                'Manufacturer Part Number': ['Part1', 'Part2', 'Part3']
            }),
            'Sheet2': pd.DataFrame({
                'Different Manufacturer Name': ['Mfg3', 'Mfg4'],
                'Different Part Number': ['Part4', 'Part5']
            })
        }
        
    @patch('pandas.ExcelFile')
    def test_excel_file_loading(self, mock_excel):
        mock_excel.return_value.sheet_names = ['Sheet1']
        mock_excel.return_value.parse.return_value = self.sample_data['Sheet1']
        
        xls = pd.ExcelFile(file_path)
        self.assertEqual(len(xls.sheet_names), 1)
        
    @patch('pandas.ExcelFile')
    def test_manufacturer_column_detection(self, mock_excel):
        mock_excel.return_value.sheet_names = ['Sheet1']
        mock_excel.return_value.parse.return_value = self.sample_data['Sheet1']
        
        df = mock_excel.return_value.parse('Sheet1')
        manufacturer_col = next((col for col in df.columns if "Manufacturer Name" in col), None)
        self.assertEqual(manufacturer_col, "Manufacturer Name")
        
    @patch('pandas.ExcelFile')
    def test_process_valid_sheet(self, mock_excel):
        mock_excel.return_value.sheet_names = ['Sheet1']
        mock_excel.return_value.parse.return_value = self.sample_data['Sheet1']
        
        with patch('pandas.DataFrame.to_excel') as mock_to_excel:
            xls = pd.ExcelFile(file_path)
            df = xls.parse('Sheet1')
            counts = df.groupby('Manufacturer Name')['Manufacturer Part Number'].count().reset_index()
            
            self.assertEqual(len(counts), 2)  # Two manufacturers
            self.assertEqual(counts['Manufacturer Name'].tolist(), ['Mfg1', 'Mfg2'])
            
    @patch('pandas.ExcelFile')
    def test_empty_data_handling(self, mock_excel):
        mock_excel.return_value.sheet_names = []
        
        with patch('pandas.DataFrame.to_excel') as mock_to_excel:
            xls = pd.ExcelFile(file_path)
            self.assertEqual(len(xls.sheet_names), 0)
            
    @patch('pandas.ExcelFile')
    def test_output_file_creation(self, mock_excel):
        mock_excel.return_value.sheet_names = ['Sheet1']
        mock_excel.return_value.parse.return_value = self.sample_data['Sheet1']
        
        with patch('pandas.DataFrame.to_excel') as mock_to_excel:
            manufacturer_df = pd.DataFrame(self.sample_data['Sheet1'])
            manufacturer_df.to_excel("Manufacturer_Part_Numbers.xlsx", index=False)
            
            mock_to_excel.assert_called_once_with("Manufacturer_Part_Numbers.xlsx", index=False)
            
    def test_catalog_summary_creation(self):
        test_summary = [
            {"Catalog": "Sheet1", "Total Manufacturers": 2, "Total Part Numbers": 3},
            {"Catalog": "Sheet2", "Total Manufacturers": 2, "Total Part Numbers": 2}
        ]
        
        catalog_summary_df = pd.DataFrame(test_summary)
        self.assertEqual(len(catalog_summary_df), 2)
        self.assertEqual(list(catalog_summary_df.columns), 
                        ['Catalog', 'Total Manufacturers', 'Total Part Numbers'])

if __name__ == '__main__':
    unittest.main()