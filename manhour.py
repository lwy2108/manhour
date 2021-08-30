import weekly_dates as wd
import leave_parser as lp


first_day = wd.select_month_year()
wd.confirm_selection(first_day)
weeks = wd.determine_number_of_weeks(first_day)
adjusted_first_day = wd.determine_adjusted_first_day(first_day)
wk1_start, wk2_start, wk3_start, wk4_start, wk5_start = wd.determine_weekly_start_dates(adjusted_first_day, weeks)
wk1_end, wk2_end, wk3_end, wk4_end, wk5_end = wd.determine_weekly_end_dates(wk1_start, wk2_start, wk3_start, wk4_start,
                                                                            wk5_start)
print(wk1_start, wk2_start, wk3_start, wk4_start, wk5_start)
print(wk1_end, wk2_end, wk3_end, wk4_end, wk5_end)
# add loop for confirmation
# add validation / strip
# add file for leave_parser
# this file is for operations on manhour report

file, sheet_name = lp.load_file()
if sheet_name is None:
    file_sheet = file.worksheets[0]
else:
    file_sheet = file[sheet_name]
emp_count, emp_summary = lp.parse_for_employees(file_sheet)
for employee in emp_summary.keys():
    lp.parse_for_leave(employee, emp_summary, file_sheet)
    print(employee, emp_summary[employee])
