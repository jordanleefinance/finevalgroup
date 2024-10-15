import openpyxl
from openpyxl.utils import get_column_letter

# Load the workbook and select the 'Budget to Actual' tab
file_path = 'SandBox_FFM.xlsx'  # Update this path to your file location
wb = openpyxl.load_workbook(file_path)
ws = wb['Budget to Actual']

# Step 1: Unmerge the top row (with dates)
for merged_cells in ws.merged_cells.ranges:
    if '1' in str(merged_cells):  # Assuming the date is in the top row (row 1)
        ws.unmerge_cells(str(merged_cells))

# Step 2: Find the row containing 'Forecast', 'Actual', and 'Budget'
forecast_col = None
actual_col = None
budget_col = None

for row in ws.iter_rows():
    for cell in row:
        if cell.value == 'Forecast':
            forecast_col = cell.column
        if cell.value == 'Actual':
            actual_col = cell.column
        if cell.value == 'Budget':
            budget_col = cell.column
    if forecast_col and actual_col and budget_col:
        target_row = cell.row
        break

# Step 3: Insert a column to the right of 'Forecast' and 'Actual'
ws.insert_cols(forecast_col + 1)
ws.insert_cols(actual_col + 2)  # Accounting for the shift due to the previous insertion

# Step 4: Copy formulas from the 'Actual' column to the new column
for row in range(2, ws.max_row + 1):  # Assuming data starts from row 2
    actual_cell = ws.cell(row=row, column=actual_col + 1)  # New column right of 'Actual'
    original_cell = ws.cell(row=row, column=actual_col)
    if original_cell.data_type == 'f':  # Only copy if it's a formula
        actual_cell.value = f"={original_cell.value.split('=')[1]}"  # Copy formula

# Step 5: Update the top row (with dates) with the last date of the previous month
from datetime import datetime, timedelta
# Find the last date of the previous month
today = datetime.today()
first_of_current_month = today.replace(day=1)
last_of_previous_month = first_of_current_month - timedelta(days=1)

# Set the value in the top row's newly inserted column
ws.cell(row=1, column=actual_col + 1, value=last_of_previous_month.strftime('%B %Y'))

# Step 6: Ungroup the column to the right of 'Budget'
ws.column_dimensions.group(get_column_letter(budget_col + 1), hidden=False)

# Save the workbook
wb.save('Updated_SandBox_FFM.xlsx')