# Imports from openpyxl
from openpyxl import load_workbook, Workbook # *** In python, imports are declared with the structure "from <module> import <class>"

# Define the paths
existing_file_path = 'data/input/test_input.xlsx' # *** In python, variables are dynamically typed and do not require a "const" or "let" keyword
new_file_path = 'data/output/test_output.xlsx'

# load workbook & get active worksheet
wb = load_workbook(filename=existing_file_path)
ws = wb.active


# ----------< Tasks begin >------------
# Initialize variables
spreadsheet_title = None 

# 1. Get title
if ws['A1'].value.startswith("Payout Report"): # *** In python, if statements are declared without parentheses and with a colon at the end, code block to be executed is indented instead of within brackets
    # Copy the text from the first column of the first row
    spreadsheet_title = ws['A1'].value
else:
    print("Error: Spreadsheet title not found")
    # add other error handling code here
print("Payout Report:", spreadsheet_title)

# 2. Cut the first row
ws.delete_rows(1)

# 3. Add two empty rows to the end of the worksheet
ws.append([])
ws.append([])

# 4. Append spreadsheet title to a new row at the bottom of the worksheet
ws.append([spreadsheet_title])


# ----------< // Tasks End >------------

# Save the workbook to a new file
wb.save(new_file_path)

# *** In python, instead of console logs, use "print(<message>)"
print("Workbook saved to", new_file_path)