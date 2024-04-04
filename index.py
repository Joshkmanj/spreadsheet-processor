# Imports from openpyxl
from openpyxl import load_workbook, Workbook # *** In python, imports are declared with the structure "from <module> import <class>"
from openpyxl.utils import get_column_letter

# Define the paths
existing_file_path = 'data/input/test_input.xlsx' # *** In python, variables are dynamically typed and do not require a "const" or "let" keyword
new_file_path = 'data/output/test_output.xlsx'

# Load worksheet
wb = load_workbook(filename=existing_file_path)
ws = wb.active

# --------------------<  T a s k s  b e g i n  >----------------------
# Variables
spreadsheet_title = None
columns_to_keep = [{'col':'a','title':'Date'},{'col':'c','title':'Description'},{'col':'d','title':'Net Donation'},{'col':'f','title':'Stripe Fee'},{'col':'g','title':'Platform Fee'},{'col':'h','title':'Total Gross Donation'},{'col':'q','title':'Email'},{'col':'r','title':'Event'},{'col':'x','title':'Source Title'}]
# columns_to_hide = [b,e,i,j,k,l,m,n,o,p,q,s,t,u,v,w] # Not sure if I'll need this one

# 1. Get title and remove from first row
if ws['A1'].value.startswith("Payout Report"): # *** In python, if statements are declared without parentheses and with a colon at the end, code block to be executed is indented instead of within brackets
    spreadsheet_title = ws['A1'].value
else:
    print("Error: Spreadsheet title not found")
    # add other error handling code here
print("Payout Report:", spreadsheet_title)

ws.delete_rows(1)

# 2. Hide columns except those that we want to keep

max_column = ws.max_column

# Loop through each column in the worksheet
for col in range(1, max_column + 1):
    # Convert column number to letter
    column_letter = get_column_letter(col)
    print(f"Column {col} has the letter {column_letter} with value {ws[column_letter + '1'].value}")









# Last. Append title to the end of sheet with buffer space
ws.append([])
ws.append([])
ws.append([spreadsheet_title])

# --------------------< //  T a s k s  E n d  >----------------------

# Save the workbook to a new file
# wb.save(new_file_path)

# *** In python, instead of console logs, use "print(<message>)"
print("Workbook saved to", new_file_path)