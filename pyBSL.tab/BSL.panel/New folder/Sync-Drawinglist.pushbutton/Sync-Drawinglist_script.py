#! python3
import os
import sys
import clr
import System
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog

PYREVIT_LIBPATH = os.path.join(os.path.join(os.getenv('APPDATA'), 'pyRevit-Master'), 'site-packages')
sys.path.append(PYREVIT_LIBPATH)

import xlrd

dialog = OpenFileDialog()
dialog.Filter = 'Excel Files|*.xlsx'

# Show the dialog and get the selected file
result = dialog.ShowDialog()

if result == System.Windows.Forms.DialogResult.OK:
    file_path = dialog.FileName

    excel_sheet = xlrd.open_workbook(file_path)
    sheets = excel_sheet.sheets()
    print(sheets[0].row(0)[0].value)