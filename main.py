import weekly_dates as wd
import leave_parser as lp
import generate_report as gr
import datetime as dt


while True:
    try:
        first_day = wd.select_month_year()
    except ValueError:
        print('Input was invalid. Try again.')
        continue
    confirm = wd.confirm_selection()
    if confirm == 'y':
        break
    else:
        continue
weeks = wd.determine_number_of_weeks(first_day)
adjusted_first_day = wd.determine_adjusted_first_day(first_day)
wk1_start, wk2_start, wk3_start, wk4_start, wk5_start = wd.determine_weekly_start_dates(adjusted_first_day, weeks)
wk1_end, wk2_end, wk3_end, wk4_end, wk5_end = wd.determine_weekly_end_dates(wk1_start, wk2_start, wk3_start, wk4_start,
                                                                            wk5_start)
# print(wk1_start, wk2_start, wk3_start, wk4_start, wk5_start)
# print(wk1_end, wk2_end, wk3_end, wk4_end, wk5_end)
# add loop for confirmation
# add validation / strip

weeks_dates = {}
for week in range(1, weeks+1):
    weeks_dates[week] = []
    weeks_dates[week].append(globals()[f'wk{week}_start'])
    for day in range(1,5):
        day_date = globals()[f'wk{week}_start'] + dt.timedelta(days=day)
        if day_date == globals()[f'wk{week}_end']:
            weeks_dates[week].append(day_date)
            break
        else:
            weeks_dates[week].append(day_date)
    print('Week', week, ':', weeks_dates[week])

while True:
    try:
        file, sheet_name = lp.load_file()
        break
    except FileNotFoundError:
        print('File not found. Try again.')
        continue
if sheet_name is None:
    file_sheet = file.worksheets[0]
else:
    file_sheet = file[sheet_name]

emp_summary = lp.parse_for_employees(file_sheet)

# for bosses

while True:
    try:
        file_boss, sheet_name_boss = lp.load_file_boss()
        break
    except FileNotFoundError:
        print('File not found. Try again.')
        continue
if sheet_name_boss is None:
    file_sheet_boss = file_boss.worksheets[0]
else:
    file_sheet_boss = file_boss[sheet_name_boss]

boss_summary = lp.parse_for_employees(file_sheet_boss)

for employee in emp_summary.keys():
    lp.parse_for_leave(employee, emp_summary, file_sheet)
    print(employee, emp_summary[employee])

for boss in boss_summary:
    lp.parse_for_leave(boss, boss_summary,file_sheet_boss)
    print(boss, boss_summary[boss])

emp_summary = {**emp_summary, **boss_summary}

# add completion message

holidays = gr.load_holidays(first_day.year)
gr.remove_cancel(emp_summary)

report_wb = gr.load_template(weeks)
report = report_wb['ITS_D']
report_df = report_wb['DF']
report_vep_c = report_wb['VEP-C']
mth = dt.datetime.strftime(wk3_start, '%b').upper()
gr.write_title_date(report, adjusted_first_day, globals()[f'wk{weeks}_end'], mth)
# print(report['A1'].value)
emp_first_row = []
max_row = report.max_row - 25

for row in report[f'A8:A{max_row}']:
    for cell in row:
        if str(cell.value).strip():
            correct_count = gr.report_verify_line_count(weeks, report, cell.row, cell.column)
            if str(cell.value).strip() == 'None':
                continue
            elif correct_count == 1:
                emp_first_row.append(cell.row)
                gr.report_focus_next_line(report, cell.row + weeks - 1)
                print(cell.value)
            else:
                gr.report_focus_next_line(report, cell.row)

for row in emp_first_row:
    gr.fill_holidays(report, holidays, row, weeks_dates)
    report[f'D{row}'] = wk1_start
    report[f'E{row}'] = wk1_end
    for line in range(1, weeks):
        report[f'D{row+line}'] = globals()[f'wk{1+line}_start']
        report[f'E{row+line}'] = globals()[f'wk{1+line}_end']

print('Employee first rows --------------------------')
print(emp_first_row)
for row in emp_first_row:
    print(row, report[f'A{row}'].value)

matched_emp = {}

# matching, simple then complex (failsafe)
for emp in emp_summary.keys():
    print(emp, '--------------')
    match, name, row = gr.report_simple_match(report, emp_first_row, emp)
    if match == 1:
        print('1')
        matched_emp[emp] = row
        emp_first_row.remove(row)
    else:
        match, name, row = gr.report_match_failsafe(report, emp_first_row, emp)
        if match == 1:
            print('+1')
            matched_emp[emp] = row
            emp_first_row.remove(row)
        else:
            print('0')
for row in emp_first_row:
    print(report[f'A{row}'].value)
print(len(emp_first_row))

print(matched_emp)

for employee in emp_summary:
    try:
        first_row = matched_emp[employee]
        print('matched', employee, emp_summary[employee], first_row)
    except KeyError:
        continue
    for entry in emp_summary[employee]:
        if emp_summary[employee][entry][1] > 1.0:
            start_dt = dt.datetime.strptime(entry[0:10], '%d/%m/%Y')
            end_dt = dt.datetime.strptime(emp_summary[employee][entry][0], '%d/%m/%Y')
            entry_dates = gr.entry_dates(start_dt, end_dt)
        else:
            entry_dates = [dt.datetime.strptime(entry, '%d/%m/%Y')]
        entry_type = emp_summary[employee][entry][3]
        entry_duration = emp_summary[employee][entry][1]
        print(entry_type)
        gr.report_add_entry(report, first_row, weeks_dates, entry_dates, entry_type, entry_duration)

report_date = dt.datetime.strftime(wk3_start, '%m-%Y')
report_wb.save(f'manhour_{report_date}.xlsx')
