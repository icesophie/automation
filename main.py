# You may reuse the code from task 4 - win32.
# Find a way to import sourcecode.py and use the navigation function from the file.
# All import <package> statements must be at the top.


import sourcecode
import win32com.client as win32
from pathlib import Path
import sys


def main():

    #define the full path for the Excel file
    excel_file = Path.cwd() / "Master.xlsx"

    #Open up Excel and make it visible
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True

    #Open the Excel file
    workbook = excel.Workbooks.Open(str(excel_file))

    #Open the sheet by using the sheet name
    sheet = workbook.Sheets("Sheet1")

    #Determine the used range (rows and columns that have data)
    used_range = sheet.UsedRange
    num_rows = used_range.Rows.Count
    #num_columns = used_range.Columns.Count

    rows_to_check = [int(arg)  for arg in sys.argv[1:]]

    #If rows_to_check is provided, validate and check only those rows
    if rows_to_check:
        for row_to_check in rows_to_check:
            if row_to_check < 2 or row_to_check > num_rows:
                print(f"Invalid row number {row_to_check}. Please provide a row number between 2 and {num_rows}")
            else:

                # print(f"Checking row: {row_to_check}")
                url = sheet.Cells(row_to_check, 2).Value
                method = sheet.Cells(row_to_check, 3).Value
                ref = sheet.Cells(row_to_check, 4).Value
                #print(f"Row {row_to_check}: URL={url}, Method={method}, Ref={ref}")
                #Call sourcecode.navigation() and assume it returns a new url
                result = sourcecode.navigation(method, url, ref)
                sheet.Cells(row_to_check, 1).Value = result  # Write the text to the first column
                print(f'Web scraping for {row_to_check}: {result}')


    #if no row number is provided, check all rows from row 2 onwards
    else:
        for row in range(2, num_rows + 1):
            #print(f"Checking row: {row}")
            url = sheet.Cells(row, 2).Value
            method = sheet.Cells(row, 3).Value
            ref = sheet.Cells(row, 4).Value
            #print(f"Row {row}: URL={url}, Method={method}, Ref={ref}")
            result = sourcecode.navigation(method, url, ref)
            sheet.Cells(row, 1).Value = result  # Write the text to the first column
            print(f'Web scraping for {row}: {result}')


    # Save the changes to the workbook
    workbook.Save()

    #Close Excel
    workbook.Close()
    excel.Application.Quit()

if __name__ == "__main__":
    main()






