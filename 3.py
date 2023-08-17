import os, sys, glob
from win32com.client import Dispatch
import pandas as pd
from openpyxl import load_workbook

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

print(application_path)

inputPath = os.path.join(application_path, "INPUT\\")
outputPath = os.path.join(application_path, "OUTPUT\\")
inputFiles = glob.glob(inputPath+"*.xlsx")

excel = Dispatch("Excel.Application")
excel.Visible = 1
excel.DisplayAlerts = False

xlUp = -4162
xlLeft = -4159
xlShiftDown = -4121
xlAscending = 1
xlSortColumns = 1

invalid_domain_list = []
invalid_domain_tuple = ()

invalid_excel = excel.Workbooks.Open(rf'{application_path}\List Invalid Domain.xlsx', UpdateLinks = 0)
invalid_domain_sheet = invalid_excel.Worksheets(1)
invalid_domain_Row = invalid_domain_sheet.Cells(invalid_domain_sheet.Rows.Count, 1).End(xlUp).Row + 1
for i in range(2,invalid_domain_Row):
    invalid_domain_list.append(invalid_domain_sheet.Cells(i,1).Value)
invalid_domain_tuple = tuple(invalid_domain_list)

invalid_excel.Close(False)
excel.Quit()

for files in inputFiles:
    filename = files.split("\\")[-1]
    book = load_workbook(files)
    df = pd.read_excel(files, sheet_name = 0, dtype=object)
    writer = pd.ExcelWriter(rf"{outputPath}{filename}", engine="openpyxl")
    # writer = pd.ExcelWriter(r"E:\WORK\RPA\Email\OUTPUT\test.xlsx", engine="xlsxwriter", engine_kwargs={'options': {'strings_to_numbers': True}})
    writer.book = book
    valid_email = df[~df['Email'].str.endswith(invalid_domain_tuple,na=True) | df['Email'].str.contains("@") | ~df['Email'].str.contains("@.") | ~df['Email'].str.startswith(".") | ~df['Email'].str.contains("..") | df['Email'].str.contains(".com") | df['Email'].str.contains(".my")] 
    not_valid_email = df[df['Email'].str.endswith(invalid_domain_tuple, na=True) | pd.isna(df['Email']) | ~df['Email'].str.contains("@") | df['Email'].str.contains("@.") | df['Email'].str.startswith(".") | df['Email'].str.contains("..") | ~df['Email'].str.contains(".com") | ~df['Email'].str.contains(".my")]

    df['Phone'] = df['Phone'].astype("string")
    # df.T.reset_index().T.to_excel(writer, sheet_name="Data", header=None, index=None)
    valid_email.T.reset_index().T.to_excel(writer, sheet_name="Valid", header=None, index=None)
    not_valid_email.T.reset_index().T.to_excel(writer, sheet_name="Invalid", header=None, index=None)
    writer.close()