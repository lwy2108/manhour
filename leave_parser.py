import xlrd
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import re


# from json import dump


def read_xls_to_xlsx(source, sheet):
    xls = xlrd.open_workbook(source)
    xlsx = Workbook()

    xls_sheet = xls.sheet_by_index(sheet)
    xls_sheet_name = xls.sheet_names()[0]
    xlsx_sheet = xlsx.active
    xlsx_sheet.title = xls_sheet_name

    for row in range(xls_sheet.nrows):
        for column in range(xls_sheet.ncols):
            xlsx_sheet.cell(row=row + 1, column=column + 1).value = xls_sheet.cell_value(row, column)

    return xlsx, xls_sheet_name


def load_file():
    file_path = input('Drag and drop file, then press enter: ')
    try:
        return load_workbook(file_path, read_only=True), None
    except openpyxl.utils.exceptions.InvalidFileException:
        return read_xls_to_xlsx(file_path, 0)


def parse_for_employees():
    count = 0
    data = {}
    for row in file_sheet['D10':'D3000']:
        for cell in row:
            if cell.value == 'Name : ':
                count += 1
                data[file_sheet['E' + str(cell.row)].value] = cell.coordinate
    return count, data


def parse_for_leave(emp):
    first_entry_start = file_sheet['E' + str(file_sheet[emp_summary[emp]].row + 4)]
    if re.match('^[0-9]{2}/[0-9]{2}/[0-9]{4}$', first_entry_start.value):
        pass
    else:
        first_entry_start = file_sheet['E' + str(first_entry_start.row + 1)]
    first_entry_end = file_sheet['F' + str(first_entry_start.row)]
    first_entry_days = file_sheet['G' + str(first_entry_start.row)]
    first_entry_approval = file_sheet['I' + str(first_entry_start.row)]
    emp_summary[emp] = {first_entry_start.value: [first_entry_end.value, first_entry_days.value,
                                                  first_entry_approval.value]}
    cell_in_focus = first_entry_start
    cell_attempt = cell_in_focus
    while True:
        cell_in_focus = cell_attempt
        for adjustment in range(1, 4):
            cell_attempt = file_sheet.cell(row=cell_in_focus.row+adjustment, column=cell_in_focus.column)
            if re.match('^[0-9]{2}/[0-9]{2}/[0-9]{4}$', str(cell_attempt.value)):
                cell_attempt_end = file_sheet['F'+str(cell_attempt.row)]
                cell_attempt_days = file_sheet['G'+str(cell_attempt.row)]
                cell_attempt_approval = file_sheet['I'+str(cell_attempt.row)]
                emp_summary[emp][cell_attempt.value] = [cell_attempt_end.value, cell_attempt_days.value,
                                                        cell_attempt_approval.value]
                break
            else:
                continue
        if cell_in_focus.value == 'Subtotal':
            break
        else:
            continue

    # reference search from next emp, or no. of emp?
    # filter dates - try/except
    # cancelled entries
    # use count


# def generate_leave_summary():
#     global leave_summary
#     leave_summary = {}
#     worksheet_leave = file['Sheet1']
#     # search cells for names (nested loop to search every cell)
#     for row in worksheet_leave.rows:
#         for cell in row:
#             if str(cell.value).strip() == 'Name :':
#                 # search for subtotals in the area of 24 rows below each name (header is in the next column)
#                 for search_row in range(cell.row + 1, cell.row + 25):
#                     search_cell = worksheet_leave.cell(row=search_row, column=(cell.column + 1))
#                     if str(search_cell.value) == 'Subtotal':
#                         # output matching names and subtotals to dictionary and returning to the second loop
#                         leave_summary[worksheet_leave.cell(row=cell.row, column=(cell.column + 1)).value] = []
#                         leave_summary[worksheet_leave.cell(row=cell.row, column=(cell.column + 1)).value].append(
#                             worksheet_leave.cell(row=search_row, column=(cell.column + 2)).value)
#                         break
#
#     return leave_summary


# def export_leave_summary():
#     with open('leave_summary.json', 'w') as f:
#         dump(leave_summary, f)


# def print_leave_summary():
#     for person in sorted(leave_summary.keys()):
#         print(person, '-->', leave_summary[person])


file, sheet_name = load_file()
if sheet_name is None:
    file_sheet = file.worksheets[0]
else:
    file_sheet = file[sheet_name]
emp_count, emp_summary = parse_for_employees()
for employee in emp_summary.keys():
    parse_for_leave(employee)
    print(employee, emp_summary[employee])
# split into individual days?
