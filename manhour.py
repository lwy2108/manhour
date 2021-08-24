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


def determine_number_of_weeks():  # improve efficiency?
    last_day = monthrange(first_day.year, first_day.month)[1]
    counter = 0
    for day in range(1, last_day + 1):
        if dt.date(first_day.year, first_day.month, day).isoweekday() == 5:
            counter += 1
        else:
            continue
    return counter


# based on last Friday?
# add 7 day until into next month, then decide based on the day


def determine_weekly_start_dates():
    iso_first_day = first_day.isoweekday()
    if iso_first_day == 1:
        adjustment = dt.timedelta(days=0)
    elif 1 < iso_first_day < 6:
        adjustment = dt.timedelta(days=(iso_first_day - 1))
    # Jan should start on 1, Dec should stop on 31: 2 separate measures needed, first for Jan (force no adjustment),
    # second for Dec, force last day
    else:
        adjustment = dt.timedelta(days=(iso_first_day - 8))
    wk1 = first_day - adjustment
    print(wk1, wk1.isoweekday())
    # return wk1, wk2, wk3, wk4, wk5


def determine_weekly_end_dates():
    pass


def load_template():
    pass


first_day = select_month_year()
# confirm_selection()
weeks = determine_number_of_weeks()
print(weeks)
# add loop for confirm
# add validation or strip
