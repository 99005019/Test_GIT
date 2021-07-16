import xlrd
from xlrd.sheet import ctype_text
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell


def read_excel_sheet(row_number):
    xl_workbook = xlrd.open_workbook("Book.xls", formatting_info=True)
    sheet_names = xl_workbook.sheet_names()
    # print('Sheet Names', sheet_names)
    xl_sheet = xl_workbook.sheet_by_name(sheet_names[1])
    # print(xl_sheet)
    row = xl_sheet.row(row_number)
    # print(row)
    # print('(Column #) type:value')
    for idx, cell_obj in enumerate(row):
        cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
        print('%s' % cell_obj.value)


def find_value_by_ps(sh, ps_number):
    for row in range(sh.nrows):
        my_cell = sh.cell(row, 0)
        if my_cell.value == ps_number:
            read_excel_sheet(row)
            return xl_rowcol_to_cell(row, 0)
    return -1


print("Enter PS NO. to get data")
ps_number = int(input())
for sh in xlrd.open_workbook("Book.xls").sheets():
    print(find_value_by_ps(sh, ps_number))
