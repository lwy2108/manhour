import weekly_dates as wd
import openpyxl


def load_template():
    pass


first_day = wd.select_month_year()
wd.confirm_selection()
weeks = wd.determine_number_of_weeks()
adjusted_first_day = wd.determine_adjusted_first_day()
wk1_start, wk2_start, wk3_start, wk4_start, wk5_start = wd.determine_weekly_start_dates()
wk1_end, wk2_end, wk3_end, wk4_end, wk5_end = wd.determine_weekly_end_dates(wk1_start, wk2_start, wk3_start,wk4_start,
                                                                            wk5_start)
print(wk1_start, wk2_start, wk3_start, wk4_start, wk5_start)
print(wk1_end, wk2_end, wk3_end, wk4_end, wk5_end)
# add loop for confirmation
# add validation / strip
# add file for leave_parser
# this file is for operations on manhour report
