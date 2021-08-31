"""
Testing phase, to run from manhour.py along with leave_parser.py after finalised
1   categorise leave entries by week, excluding 'Cancel' entries, storing the data in a new dictionary
    format: {emp1: [wk1, wk2, wk3, wk4, ...], ...}
    (default wk values should be 0, wk5 should be None if it does not exist)
    (emp should have undergone strip() and lower() to prepare for matching, remove brackets?)
    (potentially, parts of the name can be separated with an underscore and the parts sorted alphabetically)
    (matching should have a tolerance that provides accuracy with minimal failure)
    (sort into 4 or 5 weeks -- use weekly_dates.py for realistic scenarios)
    (3 possibilities: 0.5 or 1 day (handled with start date), > 2 days (handled with range inclusive, edited by name)
    (provide for PH - it's in a sep column)
2   generate report by making edits to the template (for the correct no. of weeks)
    (only start editing after creating a copy of the template)
    (fill in dates for every name)
    (fill report in a linear fashion, so that the first employee is the one first listed in the template)
    (consider generating a map of employees first, so that the coordinates are readily available)
    (report should be generated on the fly (per employee) to avoid fatal errors that cripple the program)
    (ensure correct number of lines per emp, based on 4/5 wks)
    (remove from dict after edit)
3   export to json leave_summary, with totals

def export_leave_summary():
    with open('leave_summary.json', 'w') as f:
        dump(leave_summary, f)
3   to upload sample leave_report and manhour_report, rename manhour.py to main.py
4   reorg functions into right files - pass args through a dict for start, end dates
5   provide for hired date - approved column
"""


def load_holidays():  # the right year
    pass


def remove_cancel():
    pass


def entry_dates():
    pass


def entry_duration():  # only working days, and excluding PH
    pass


def entry_match_week():
    pass


def transform_emp_names():  # upper, strip(), in keys
    pass


def load_template():  # based on no. of weeks
    pass


def report_focus_next_line():  # until approval line, variable for current line (updates after line count)
    pass


def report_verify_line_count():
    pass


def report_simple_match():  # strip()
    pass


def report_match_failsafe():  # remove symbols and match by part, etc -- perform on both sides
    pass  # do this last after analysing the fails


def report_add_entry():  # add not replace
    pass


def entry_remove_successful():  # only if successful
    pass


def append_leave_report():
    pass


def entry_print_remaining():  # for skipped or failed
    pass
