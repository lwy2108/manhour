"""for operations related to shift workers"""

import datetime as dt

# to separate shift workers from emp_summary

def separate_shift(emp_summary, matched_emp, report):

    shift_workers = {}
    shift_matched = {}

    for emp in matched_emp:
        current = report[f'A{matched_emp[emp]}'].value
        if '(SHIFT WORK)' in current:
            shift_workers[emp] = emp_summary[emp]
            del emp_summary[emp]
            shift_matched[emp] = matched_emp[emp]

    # for debugging
    # print("*************************************************")
    # print('regular', emp_summary)
    # print('shift', shift_workers)
    # print("*************************************************")

    return emp_summary,shift_workers, shift_matched

# to determine considered dates for shift workers (full 7 days a week)

def dates(start, end, weeks):

    # for debugging
    # print("***********************first day, last day**************************")
    # print(start, end)

    shift_dates = {}
    current_day = 0

    for week in range(1, weeks + 1):
        shift_dates[week] = []
        for day in range(7):
            day_to_add = start + dt.timedelta(days=current_day)
            shift_dates[week].append(day_to_add)
            current_day += 1

            # for year start/end
            if day_to_add == end:
                break

    return shift_dates