"""
Testing phase, to run from manhour.py along with leave_parser.py after finalised
1   categorise leave entries by week, excluding 'Cancel' entries, storing the data in a new dictionary
    format: {emp1: [wk1, wk2, wk3, wk4, ...], ...}
    (default wk values should be 0, wk5 should be None if it does not exist)
    (emp should have undergone strip() and lower() to prepare for matching, remove brackets?)
    (potentially, parts of the name can be separated with an underscore and the parts sorted alphabetically)
    (matching should have a tolerance that provides accuracy with minimal failure)
    (sort into 4 or 5 weeks -- use weekly_dates.py for realistic scenarios)
2   generate report by making edits to the template (for the correct no. of weeks)
    (only start editing after creating a copy of the template)
    (fill in dates for every name)
    (fill report in a linear fashion, so that the first employee is the one first listed in the template)
    (consider generating a map of employees first, so that the coordinates are readily available)
    (report should be generated on the fly (per employee) to avoid fatal errors that cripple the program)
3   export to json leave_summary, with totals

def export_leave_summary():
    with open('leave_summary.json', 'w') as f:
        dump(leave_summary, f)
3   to upload sample leave_report and manhour_report
"""


def load_template():
    pass
