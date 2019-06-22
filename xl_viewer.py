# https://github.com/love2spooge/python_xl_viewer



# /// import + other system stuff
from prettytable import PrettyTable
import sys
import openpyxl
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils import column_index_from_string


# /// variables

xl_file = sys.argv[1] if len(sys.argv) > 1 else "" # argument from command line

if xl_file == "": # if argument is not specified
    print(
    '''
    You didn't specified the input file
    Make sure that you run script with argument:
    python xl_viewer.py sample.xlsx
    '''
    )
    quit()


t = PrettyTable()
xl_column = []  # table column
xl_row = []     # table row

xl_input = ""           # input for sheet option

xl_search_term = ""
xl_search_column = ""   # variable for search TO-DO
xl_compare = ""         # for search

try:
    wb = openpyxl.load_workbook(xl_file) # open file
except:
    # error if file is not found or not supported
    print("Error", xl_file,"not found or file not supported")
    quit()

sheet = wb.active # open active sheet in file

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
    print("1 - Open sheet. 2 - Search. 3 - Exit")



#CODE

generate_table()

while True:  # making a loop
    xl_input = input("Select action: ")
    if xl_input == "1":
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

    # Search
    if xl_input == "2":
        xl_search_term = input("Search for: ")
        xl_search_column = input("Select column: ")

        wb.create_sheet("Results")
        sheet_result = wb["Results"]

        for row in sheet.iter_rows():
            xl_compare = str(row[column_index_from_string(xl_search_column[0]) - 1].value)

            if xl_compare == "None":
                xl_compare = "something"
                
            if xl_search_term.lower() in xl_compare.lower():
                sheet_result.append((cell.value for cell in row))
        
        sheet = wb["Results"]
        generate_table()
        wb.remove(wb["Results"])
        sheet = wb[wb.sheetnames[0]]
       
    if xl_input == "3":
        print("")
        print("Exiting...")
        break
