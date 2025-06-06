import unittest
import pandas as pd
from pyxlsx import PyWorkbook # Assuming __init__.py exports this
import os
from openpyxl import load_workbook # For direct verification

# Define a directory for test outputs
TEST_OUTPUT_DIR = "test_outputs"
if not os.path.exists(TEST_OUTPUT_DIR):
    os.makedirs(TEST_OUTPUT_DIR)

class TestPyXLSX(unittest.TestCase):

    def test_create_workbook_and_add_sheet(self):
        wb_path = os.path.join(TEST_OUTPUT_DIR, "test_create_workbook.xlsx")

        wb = PyWorkbook()
        self.assertIsNotNone(wb.workbook, "Workbook object should be created.")

        # Test adding sheet by title
        ws1 = wb.add_worksheet(title="MySheet1")
        self.assertEqual(ws1.worksheet.title, "MySheet1", "Sheet title not set correctly.")
        self.assertIn("MySheet1", wb.workbook.sheetnames, "Sheet not found in workbook by title.")

        # Test adding sheet with default title
        ws2 = wb.add_worksheet()
        self.assertIsNotNone(ws2.worksheet.title, "Default sheet title should be assigned.")
        self.assertIn(ws2.worksheet.title, wb.workbook.sheetnames, "Default titled sheet not found.")

        wb.save(wb_path)
        self.assertTrue(os.path.exists(wb_path), "Workbook file not saved.")

        # Clean up
        if os.path.exists(wb_path):
            os.remove(wb_path)

    def test_write_simple_dataframe(self):
        wb_path = os.path.join(TEST_OUTPUT_DIR, "test_write_dataframe.xlsx")
        data = {'col1': [1, 2], 'col2': ['A', 'B']}
        df = pd.DataFrame(data)

        wb = PyWorkbook()
        ws = wb.add_worksheet(title="DataFrameSheet")
        ws.write_dataframe(df, start_row=1, start_col=1, header=True)
        wb.save(wb_path)
        self.assertTrue(os.path.exists(wb_path))

        # Verification using openpyxl directly
        verify_wb = load_workbook(wb_path)
        verify_ws = verify_wb["DataFrameSheet"]

        self.assertEqual(verify_ws.cell(row=1, column=1).value, 'col1')
        self.assertEqual(verify_ws.cell(row=1, column=2).value, 'col2')
        self.assertEqual(verify_ws.cell(row=2, column=1).value, 1)
        self.assertEqual(verify_ws.cell(row=2, column=2).value, 'A')
        self.assertEqual(verify_ws.cell(row=3, column=1).value, 2)
        self.assertEqual(verify_ws.cell(row=3, column=2).value, 'B')

        if os.path.exists(wb_path):
            os.remove(wb_path)

    def test_write_dataframe_with_formatting(self):
        wb_path = os.path.join(TEST_OUTPUT_DIR, "test_dataframe_formatting.xlsx")
        data = {'ID': [101, 102], 'Value': [123.45, 678.90], 'Status': ['Active', 'Inactive']}
        df = pd.DataFrame(data)

        wb = PyWorkbook()
        ws = wb.add_worksheet("FormattedDF")
        ws.write_dataframe(df, start_row=2, start_col=2, header=True, include_index=False,
                           header_font_bold=True, header_fill_color="C0C0C0", # Silver
                           font_name="Calibri", font_size=11,
                           number_formats={'Value': '0.00'},
                           column_widths={'B': 15, 'C': 20, 'D': 10}, # Test with column letters
                           data_alignment_horizontal='center',
                           header_alignment_horizontal='center'
                           )
        wb.save(wb_path)
        self.assertTrue(os.path.exists(wb_path))

        # Basic verification (more detailed style checking can be complex and added later)
        verify_wb = load_workbook(wb_path)
        verify_ws = verify_wb["FormattedDF"]

        # Header check
        self.assertEqual(verify_ws.cell(row=2, column=2).value, 'ID')
        self.assertTrue(verify_ws.cell(row=2, column=2).font.bold, "Header font should be bold.")
        self.assertIsNotNone(verify_ws.cell(row=2, column=2).fill.fgColor.rgb, "Header fill color not applied.")


        # Data check
        self.assertEqual(verify_ws.cell(row=3, column=3).value, 123.45)
        self.assertEqual(verify_ws.cell(row=3, column=3).number_format, '0.00', "Number format not applied.")
        self.assertEqual(verify_ws.cell(row=3, column=3).font.name, 'Calibri', "Data font name not applied.")
        self.assertEqual(verify_ws.cell(row=3, column=3).alignment.horizontal, 'center', "Data alignment not applied.")


        if os.path.exists(wb_path):
            os.remove(wb_path)

    def test_write_and_read_cell(self):
        wb_path = os.path.join(TEST_OUTPUT_DIR, "test_cell_ops.xlsx")
        wb = PyWorkbook()
        ws = wb.add_worksheet("CellOps")

        ws.write_cell(1, 1, "Hello", font_bold=True, font_color="FF0000") # Red text
        ws.write_cell(2, 1, 12345, number_format="#,##0")

        self.assertEqual(ws.read_cell(1, 1), "Hello")
        self.assertEqual(ws.read_cell(2, 1), 12345)

        wb.save(wb_path)
        self.assertTrue(os.path.exists(wb_path))

        # Verify with openpyxl
        verify_wb = load_workbook(wb_path)
        verify_ws = verify_wb["CellOps"]
        self.assertEqual(verify_ws.cell(1,1).value, "Hello")
        self.assertTrue(verify_ws.cell(1,1).font.bold)
        # self.assertEqual(verify_ws.cell(1,1).font.color.rgb, "FFFF0000") # Note: openpyxl might add FF for alpha

        self.assertEqual(verify_ws.cell(2,1).value, 12345)
        self.assertEqual(verify_ws.cell(2,1).number_format, "#,##0")

        if os.path.exists(wb_path):
            os.remove(wb_path)

if __name__ == '__main__':
    unittest.main()
