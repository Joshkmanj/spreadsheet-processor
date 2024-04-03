# Imports from openpyxl
from openpyxl import load_workbook, Workbook # *** In python, imports are declared with the structure "from <module> import <class>"

# Define the paths
existing_file_path = 'data/input/test_input.xlsx' # *** In python, variables are dynamically typed and do not require a "const" or "let" keyword
new_file_path = 'data/output/test_output.xlsx'

# Load worksheet
wb = load_workbook(filename=existing_file_path)
ws = wb.active

# --------------------<  T a s k s  b e g i n  >----------------------
# Variables
spreadsheet_title = None 

# 1. Get title and remove from first row
if ws['A1'].value.startswith("Payout Report"): # *** In python, if statements are declared without parentheses and with a colon at the end, code block to be executed is indented instead of within brackets
    spreadsheet_title = ws['A1'].value
else:
    print("Error: Spreadsheet title not found")
    # add other error handling code here
print("Payout Report:", spreadsheet_title)

ws.delete_rows(1)

# 2. 





# Last. Append title to the end of sheet with buffer space
ws.append([])
ws.append([])
ws.append([spreadsheet_title])

# --------------------< //  T a s k s  E n d  >----------------------

# Save the workbook to a new file
wb.save(new_file_path)

# *** In python, instead of console logs, use "print(<message>)"
print("Workbook saved to", new_file_path)