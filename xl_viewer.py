# https://github.com/love2spooge/python_xl_viewer



# /// import + other system stuff
import openpyxl
from prettytable import PrettyTable
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils import column_index_from_string


# /// variables
t = PrettyTable()
xl_column = []
xl_row = []

xl_search = ""
xl_search_column = ""

wb = openpyxl.load_workbook("sample.xlsx") # open file
sheet = wb.active

# DEF

# generate table
def generate_table():

    xl_column = []

    xl_column.insert(0, "")
    for i in range(1, sheet.max_column + 1):
        xl_column.insert(i, get_column_letter(i))
    t = PrettyTable(xl_column)

    xl_row = []
    for x in range(1, sheet.max_row + 1):
        xl_row = []
        for y in range(1, sheet.max_column + 1):
            if y == 1:
                xl_row.insert(0, x)

            if sheet.cell(row=x, column=y).value == None:
                xl_row.insert(y, "")
            else:
                xl_row.insert(y, sheet.cell(row=x, column=y).value)

            if y == (sheet.max_column):
                t.add_row(xl_row)
    print(wb.active)
    print(t)
    print(wb.sheetnames)

    t.clear_rows()

#CODE

# open spreadsheet file

generate_table()

while True:
    sheet = wb[str(input("Open Sheet: "))]
    generate_table()
    if input() == "exit":
       break
