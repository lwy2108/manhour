from openpyxl import load_workbook
from json import dump

def load_file():
    file_path = input('Drag and drop file, then press enter: ') # convert from xls to xlsx?, add or remove staff easily
    file = load_workbook(file_path, read_only=True)
    return file

def prep_leave_summary():
    file = load_file()
    wsheet_leave = file ['Sheet1'] # 'LeaveReport'
    return wsheet_leave

def generate_leave_summary():
    leave_summary = {}
    # search cells for names (nested loop to search every cell)
    for row in wsheet_leave.rows:
        for cell in row:
            if str(cell.value).strip() == 'Name :':
                # search for subtotals in the area of 24 rows below each name (header is in the next column)
                for search_row in range(cell.row + 1, cell.row + 25):
                    search_cell = wsheet_leave.cell(row=search_row, column=(cell.column + 1))
                    if str(search_cell.value) == 'Subtotal':
                        # output matching names and subtotals to dictionary and returning to the second loop
                        leave_summary[wsheet_leave.cell(row=cell.row, column=(cell.column + 1)).value] = [] # value? week?
                        # leave_summary[wsheet_leave.cell(row=cell.row, column=(cell.column + 1)).value][0] = \
                        # wsheet_leave.cell(row=search_row, column=(cell.column + 2)).value
                        leave_summary[wsheet_leave.cell(row=cell.row, column=(cell.column + 1)).value].append(
                            wsheet_leave.cell(row=search_row, column=(cell.column + 2)).value
                        )
                        break
                    # if str(search_cell.value):
                        # leave_summary[wsheet_leave.cell(row=cell.row, column=(cell.column + 1)).value][1] = 1
                        # date column instead?
                        break

    return leave_summary

def export_leave_summary():
    with open('leave_summary.json', 'w') as f:
        dump(leave_summary, f)

wsheet_leave = prep_leave_summary()
leave_summary = generate_leave_summary()
# export_leave_summary()
print(leave_summary)

# ! test on past few months before moving on

# ! get name 'Name :', search for 'Subtotal' until following 'CSE ITS PTE LTD' and get value, add to leave_summary[name]