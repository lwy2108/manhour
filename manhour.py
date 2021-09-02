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
    for day in range(1,7):
        day_date = globals()[f'wk{week}_start'] + dt.timedelta(days=day)
        if day_date == globals()[f'wk{week}_end']:
            weeks_dates[week].append(day_date)
            break
        else:
            weeks_dates[week].append(day_date)
    # print('Week', week, ':', weeks_dates[week])

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
for employee in emp_summary.keys():
    lp.parse_for_leave(employee, emp_summary, file_sheet)
    print(employee, emp_summary[employee])

# add completion message

holidays = gr.load_holidays(first_day.year)
gr.remove_cancel(emp_summary)
for employee in list(emp_summary.keys()):
    for entry in emp_summary[employee].keys():
        print(entry, emp_summary[employee][entry])
        if emp_summary[employee][entry][1] > 1.0:
            start_dt = dt.datetime.strptime(entry, '%d/%m/%Y')
            end_dt = dt.datetime.strptime(emp_summary[employee][entry][0], '%d/%m/%Y')
            entry_dates = gr.entry_dates(start_dt, end_dt)
            print(entry_dates)