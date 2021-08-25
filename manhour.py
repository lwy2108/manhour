import collections

import openpyxl
import datetime as dt
from calendar import monthrange


def select_month_year():
    selected_month = input('Month (e.g. Jan): ')
    selected_year = input('Year (e.g. 2021): ')
    return dt.datetime.strptime(selected_month + selected_year, '%b%Y')


def confirm_selection():
    print('The first day is a', first_day.strftime('%A.'))  # monthrange can be used instead
    return input('Confirm (y/n)? ')


def determine_number_of_weeks():  # optimization?
    last_day = monthrange(first_day.year, first_day.month)[1]
    count = 0
    for day in range(1, last_day + 1):
        if dt.date(first_day.year, first_day.month, day).isoweekday() == 5:
            count += 1
        else:
            continue
    return count


def determine_adjusted_first_day():
    iso_first_day = first_day.isoweekday()
    if first_day.month == 1:  # Jan will always start on the 1st in line with company policy
        adjustment = dt.timedelta(days=0)
    elif iso_first_day == 1:
        adjustment = dt.timedelta(days=0)
    elif 1 < iso_first_day < 6:
        adjustment = dt.timedelta(days=(iso_first_day - 1))
    else:
        adjustment = dt.timedelta(days=(iso_first_day - 8))
    return first_day - adjustment


def determine_weekly_start_dates():
    week_delta = dt.timedelta(days=7)
    wk1 = adjusted_first_day
    if wk1.isoweekday() == 1:
        wk2 = adjusted_first_day + week_delta
    else:
        wk2 = adjusted_first_day + dt.timedelta(days=(8-wk1.isoweekday()))
    wk3 = wk2 + week_delta
    wk4 = wk3 + week_delta
    if weeks == 5:
        wk5 = wk4 + week_delta
    else:
        wk5 = None
    return wk1, wk2, wk3, wk4, wk5


def determine_weekly_end_dates(week1, week2, week3, week4, week5):
    delta = dt.timedelta(days=1)
    week_delta = dt.timedelta(days=6)
    wk1 = wk2_start - delta
    wk2 = wk2_start + week_delta
    wk3 = wk3_start + week_delta
    wk4 = wk4_start + week_delta
    try:
        if first_day.month == 12:
            wk5 = wk5_start.replace(day=31)  # Dec will always end on the 31st
        else:
            wk5 = wk5_start + week_delta
    except TypeError:
        wk5 = None
    return wk1, wk2, wk3, wk4, wk5


def load_template():
    pass


first_day = select_month_year()
# confirm_selection()
weeks = determine_number_of_weeks()
adjusted_first_day = determine_adjusted_first_day()
wk1_start, wk2_start, wk3_start, wk4_start, wk5_start = determine_weekly_start_dates()
wk1_end, wk2_end, wk3_end, wk4_end, wk5_end = determine_weekly_end_dates(wk1_start, wk2_start, wk3_start,wk4_start,
                                                                         wk5_start)
print(wk1_start, wk2_start, wk3_start, wk4_start, wk5_start)
print(wk1_end, wk2_end, wk3_end, wk4_end, wk5_end)
# add loop for confirm
# add validation or strip
