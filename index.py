# Imports from openpyxl
# *** In python, imports are declared with the structure "from <module> import <class>"
from openpyxl import load_workbook, Workbook

# Define the paths
# *** In python, variables are dynamically typed and do not require a "const" or "let" keyword
existing_file_path = 'data/input/test_input.xlsx'
new_file_path = 'data/output/test_output.xlsx'

# load workbook & get active worksheet
wb = load_workbook(filename=existing_file_path)
ws = wb.active


# ----------< Tasks begin >------------

# If the first column of the first row starts with "Payout Report", copy the first row
# *** In python, if statements are declared without parentheses and with a colon at the end, code block to be executed is indented instead of within brackets
if ws['A1'].value.startswith("Payout Report"):
    # Copy the text from the first column of the first row
    payout_report = ws['A1'].value
    print("Payout Report:", payout_report)


# ----------< // Tasks End >------------

# Save the workbook to a new file
wb.save(new_file_path)

# *** In python, instead of console logs, use "print(<message>)"
print("Workbook saved to", new_file_path)