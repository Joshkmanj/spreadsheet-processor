# Import the load_workbook and Workbook classes from openpyxl
from openpyxl import load_workbook, Workbook

# Define the path to the existing Excel file and the path where the new file will be saved
existing_file_path = 'data/input/existing_file.xlsx'
new_file_path = 'data/output/new_file.xlsx'

# Load the existing workbook
wb = load_workbook(filename=existing_file_path)

# Get the active worksheet from the workbook
ws = wb.active

# Save the workbook to a new file
wb.save(new_file_path)

print("Workbook saved to", new_file_path)