import unittest
from unittest.mock import MagicMock, patch
import pandas as pd
import os
import tkinter as tk
from tkinter import ttk

# Import the class to test
# We need to mock selenium basics because strict import might fail if drivers aren't found or we don't want to start them
import sys
sys.modules['selenium'] = MagicMock()
sys.modules['selenium.webdriver'] = MagicMock()
sys.modules['selenium.webdriver.chrome.options'] = MagicMock()
sys.modules['selenium.webdriver.common.by'] = MagicMock()
sys.modules['selenium.webdriver.common.keys'] = MagicMock()
sys.modules['selenium.webdriver.support.ui'] = MagicMock()
sys.modules['selenium.webdriver.support'] = MagicMock()
sys.modules['selenium.common.exceptions'] = MagicMock()

# Now import the app
# We need to make sure we point to the right file. Since it is in the same dir as we run, we append path.
sys.path.append(r'f:\Python\dvc')
try:
    from ToolTraCuuVBDLIS_V2 import AppUI
except ImportError:
    # If run from a different CWD, this might fail, so we read file content and exec it to simulate import
    # But for now assume we run it such that import works or we paste code.
    # Actually, simpler to just start the test assuming the file is importable or just mock the logic being tested.
    # Let's try to simple-import.
    pass

class TestExcelExport(unittest.TestCase):
    
    @patch('ToolTraCuuVBDLIS_V2.filedialog.asksaveasfilename')
    @patch('ToolTraCuuVBDLIS_V2.messagebox.showinfo')
    @patch('ToolTraCuuVBDLIS_V2.os.startfile')
    def test_export_logic(self, mock_startfile, mock_showinfo, mock_asksaveas):
        # Setup
        root = MagicMock()
        
        # Patch show_login to do nothing so we don't hit tk errors
        with patch('ToolTraCuuVBDLIS_V2.AppUI.show_login') as mock_show_login, \
             patch('ToolTraCuuVBDLIS_V2.AppUI.load_config') as mock_load_config, \
             patch('ToolTraCuuVBDLIS_V2.MplisAutomation') as MockAutomation:
            
            app = AppUI(root)
            # Manually init search_results since we skipped the line in init if it was before show_login?
            # Creating AppUI calls __init__.
            # __init__ sets self.search_results = [] THEN calls show_login().
            # So search_results should be there.
            app.result_index = 0 # manually init because clear_root (called by show_login) is skipped
            
            # Manually init col_configs because build_result_grid is skipped
            app.col_configs = [
                {"width": 50, "weight": 0},  # #
                {"width": 300, "weight": 1}, # Cert
                {"width": 400, "weight": 1}, # Owner 
                {"width": 300, "weight": 1}  # Land
            ]
            
            # 1. Simulate adding results
            
            # 1. Simulate adding results
            # The logic in add_result_row:
            # col[1] = Cert, col[2] = Owner, col[3] = Land
            test_data_1 = {"raw_columns": ["", "Cert001", "Nguyen Van A", "To 1 Thua 2"]}
            test_data_2 = {"raw_columns": ["", "Cert002", "Tran Thi B", "To 3 Thua 4"]}
            
            # We must mock scrollable_frame to be Truthy so add_result_row works
            app.scrollable_frame = MagicMock()
            
            app.add_result_row(test_data_1)
            app.add_result_row(test_data_2)
            
            # Verify data stored
            self.assertEqual(len(app.search_results), 2)
            self.assertEqual(app.search_results[0]['Thông tin chủ sở hữu'], "Nguyen Van A")
            
            # 2. Simulate Export
            test_output_file = os.path.abspath("test_export_result.xlsx")
            if os.path.exists(test_output_file):
                os.remove(test_output_file)
                
            mock_asksaveas.return_value = test_output_file
            
            app.export_excel()
            
            # 3. Verify File Created
            self.assertTrue(os.path.exists(test_output_file))
            
            # 4. Verify Content
            df = pd.read_excel(test_output_file)
            self.assertEqual(len(df), 2)
            self.assertEqual(df.iloc[0]['Thông tin chủ sở hữu'], "Nguyen Van A")
            self.assertEqual(df.iloc[1]['Thửa đất'], "To 3 Thua 4")
            
            print("Verification Successful: Excel file created with correct data.")
            
            # Cleanup
            if os.path.exists(test_output_file):
                os.remove(test_output_file)

if __name__ == '__main__':
    unittest.main()
