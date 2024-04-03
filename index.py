# Imports from openpyxl
from openpyxl import load_workbook, Workbook

# Define the paths
existing_file_path = 'data/input/test_input.xlsx'
new_file_path = 'data/output/test_output.xlsx'

# load workbook & get active worksheet
wb = load_workbook(filename=existing_file_path)
ws = wb.active


# ----------< Tasks begin >------------

# If the first column of the first row starts with "Payout Report", copy the first row
if ws['A1'].value.startswith("Payout Report"):
    # Copy the text from the first column of the first row
    payout_report = ws['A1'].value
    print("Payout Report:", payout_report)


# ----------< // Tasks End >------------

# Save the workbook to a new file
wb.save(new_file_path)

print("Workbook saved to", new_file_path)