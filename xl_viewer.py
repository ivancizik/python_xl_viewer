# https://github.com/love2spooge/python_xl_viewer



# /// import + other system stuff
import keyboard
from prettytable import PrettyTable
import openpyxl
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils import column_index_from_string


# /// variables
t = PrettyTable()
xl_column = []
xl_row = []

xl_input = ""
xl_search_column = ""

wb = openpyxl.load_workbook("sample.xlsx") # open file
sheet = wb.active

# DEF

# generate table and print
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
    print(sheet)
    print(t)
    print(wb.sheetnames)
    print("")
    print("F1 - open sheet. F3 - search. F10 - exit")



#CODE

generate_table()

while True:  # making a loop
    if keyboard.is_pressed('F1'):
        while True:
            print("")
            t.clear_rows()
            xl_input = input("Open Sheet: ")
            if xl_input in wb.sheetnames:
                sheet = wb[xl_input]
                generate_table()
                break
            else:
                print("Sheet does not exist")

    if keyboard.is_pressed('F3'):
        print("TO-DO")

    if keyboard.is_pressed('F10'):
        print("")
        print("Exiting...")
        break
