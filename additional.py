"""for operations relating to df and vep-c where shift work is the default"""

import generate_report as gr
import datetime as dt

# DF

def title_df(first_day, report_df, adjusted_first_day, last_week_end):
    month = dt.datetime.strftime(first_day, '%b').upper()
    cell = report_df['A1']
    start = dt.datetime.strftime(adjusted_first_day, '%d/%m/%Y')
    end = dt.datetime.strftime(last_week_end, '%d/%m/%Y')
    cell.value = f'MANHOUR REPORT FOR DF STAFF - {month} ({start} ~ {end})'


# VEP-C

def title_vep_c(first_day, report_vep_c, adjusted_first_day, last_week_end):
    month = dt.datetime.strftime(first_day, '%b').upper()
    cell = report_vep_c['A1']
    start = dt.datetime.strftime(adjusted_first_day, '%d/%m/%Y')
    end = dt.datetime.strftime(last_week_end, '%d/%m/%Y')
    cell.value = f'MANHOUR REPORT FOR VEP STAFF - {month} ({start} ~ {end})'


# verify line count and get first rows

def get_first_rows(report, string, weeks):

    emp_first_row = []

    if string == "df":
        max_row = report.max_row - 33
    elif string =="vep-c":
        max_row = report.max_row - 29

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

    # for debugging
    print(report['A1'].value, "max_row", max_row)
    print(string, emp_first_row)

    return emp_first_row


# fill holidays for regulars

def fill_dates_holidays(emp_first_row, report, holidays, weeks_dates, shift_dates, weeks):

    for row in emp_first_row:

        gr.fill_holidays(report, holidays, row, weeks_dates) # for regulars

        for line in range(weeks):
            report[f'D{row+line}'] = shift_dates[line+1][0]
            report[f'E{row+line}'] = shift_dates[line+1][6]


    # for debugging
    # print(shift_dates, weeks_dates)
    # print('Employee first rows --------------------------')
    # print(emp_first_row)
    # for row in emp_first_row:
    #     print(row, report[f'A{row}'].value)


# match employees

def match(emp_summary, emp_first_row, report):

    matched_emp = {}

    # matching, simple then complex (failsafe)
    for emp in emp_summary.keys():
        match, name, row = gr.report_simple_match(report, emp_first_row, emp)
        if match == 1:
            matched_emp[emp] = row
            emp_first_row.remove(row)
        else:
            match, name, row = gr.report_match_failsafe(report, emp_first_row, emp)
            if match == 1:
                matched_emp[emp] = row
                emp_first_row.remove(row)
            else:
                continue

    # for debugging
    # for row in emp_first_row:
    #     print(report[f'A{row}'].value)
    # print(len(emp_first_row))
    # print(matched_emp)

    return matched_emp


def separate_regular(report, matched_emp, emp_summary):

    regular_workers = {}
    regular_matched = {}

    for emp in matched_emp:
        current = report[f'A{matched_emp[emp]}'].value
        if '(REGULAR)' in current:
            regular_workers[emp] = emp_summary[emp]
            del emp_summary[emp]
            regular_matched[emp] = matched_emp[emp]

    # for debugging
    # print(emp_summary, regular_workers, regular_matched)

    return emp_summary, regular_workers, regular_matched

def fill_report(report, emp_summary, regular_workers, matched_emp, regular_matched, shift_dates, weeks_dates, rep_string):

    for employee in emp_summary:

        try:
            first_row = matched_emp[employee]

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
            gr.report_add_shift_entry(report, first_row, shift_dates, entry_dates, entry_type, entry_duration, rep_string)

    for employee in regular_workers:

        try:
            first_row = regular_matched[employee]
        except KeyError:
            continue
        for entry in regular_workers[employee]:
            if regular_workers[employee][entry][1] > 1.0:
                start_dt = dt.datetime.strptime(entry[0:10], '%d/%m/%Y')
                end_dt = dt.datetime.strptime(regular_workers[employee][entry][0], '%d/%m/%Y')
                entry_dates = gr.entry_dates(start_dt, end_dt)
            else:
                entry_dates = [dt.datetime.strptime(entry, '%d/%m/%Y')]
            entry_type = regular_workers[employee][entry][3]
            entry_duration = regular_workers[employee][entry][1]
            gr.report_add_entry(report, first_row, weeks_dates, entry_dates, entry_type, entry_duration, rep_string)

