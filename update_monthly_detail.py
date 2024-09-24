import openpyxl
import win32com.client
import os
from datetime import datetime, timedelta

import os

def remove_password(file_path, password, new_file_path):
    # Initialize Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Set to True if you want to see Excel

    try:
        # Open the workbook with the password
        workbook = excel.Workbooks.Open(file_path, Password=password)

        # Check if the directory for the new file exists
        new_file_dir = os.path.dirname(new_file_path)
        if not os.path.exists(new_file_dir):
            raise FileNotFoundError(f"The directory '{new_file_dir}' does not exist.")
        
        if os.path.exists(new_file_path)==False:

            # Save a new copy without password protection
            workbook.SaveAs(new_file_path, Password='', FileFormat=51)  # 51 for .xlsx
            print(f"Password removed successfully. New file saved as: {new_file_path}")
        else:
            # Save a new copy without password protection
            workbook.Save(new_file_path)  # 51 for .xlsx
            print(f"Password removed successfully. New file saved as: {new_file_path}")


    except Exception as e:
        print(f"An error occurred while removing password: {e}")

    finally:
        # Close the workbook and quit Excel
        try:
            workbook.Close(SaveChanges=True)
        except Exception as e:
            print(f"An error occurred while closing the workbook: {e}")
        excel.Quit()

def find_date_in_row(file_path, sheet_name='Monthly Detail', search_row='4:4', target_date=datetime.today().replace(day=1)-timedelta(days=1)):
    # Initialize Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Set to True if you want to see Excel
    # Open the workbook
    workbook = excel.Workbooks.Open(file_path)
        
    # Access the 'Monthly Detail' sheet
    sheet = workbook.Sheets(sheet_name)

    # Set target_date to the last day of the previous month if not provided
    if target_date is None:
        target_date = datetime.today().replace(day=1) - timedelta(days=1)

    # Convert target date to a datetime object
    if isinstance(target_date, str):
        target_date = datetime.strptime(target_date, '%m/%d/%Y')  # Adjust format as needed

    # Iterate through the specified row
    found = False
    #print(target_date)
    for cell in sheet.UsedRange.Range(search_row):
        #cell.Value = datetime.strptime(cell.Value, '%m/%d/%Y')
        #print(cell.value)
        if isinstance(cell.Value, datetime):
            if cell.Value.date() == target_date.date():  # Compare only dates
                column_letter = cell.Address.split('$')[1]  # Get the address to extract column letter
                if len(column_letter)>1:
                    pre_column_letter = column_letter[0] + chr(ord(column_letter[1]) - 1)
                    post_column_letter = column_letter[0] + chr(ord(column_letter[1]) + 1)

                # Handle special case for column 'AA'
                if column_letter == 'AA':
                    pre_column_letter = 'Z'
                else:
                    pre_column_letter = chr(ord(column_letter) - 1)
                    post_column_letter = chr(ord(column_letter) + 1)
                print(f"Found date {target_date.date()} in column {column_letter}. Values in this column:")
                found = True
                return f"{pre_column_letter}:{pre_column_letter}", f"{column_letter}:{column_letter}", f"{post_column_letter}:{post_column_letter}"

    if not found:
        print(f"Date {target_date.date()} not found in row {search_row}.")

def copy_formatting_and_formulas(file_path, target_date):
    # Initialize Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Set to True if you want to see Excel
    print(find_date_in_row(file_path, target_date=target_date))
    pre_source = str(find_date_in_row(file_path, target_date=target_date)[0])
    source = str(find_date_in_row(file_path, target_date=target_date)[1])
    target = str(find_date_in_row(file_path, target_date=target_date)[2])
    try:
        # Open the workbook
        workbook = excel.Workbooks.Open(file_path)
        
        # Access the 'Monthly Detail' sheet
        sheet = workbook.Sheets('Monthly Detail')

        # Copy formatting and formulas from column AA to AB
        pre_source_range = sheet.UsedRange.Range(pre_source)
        source_range = sheet.UsedRange.Range(source)
        target_range = sheet.UsedRange.Range(target)
        
        # Copy the source range
        source_range.Copy(target_range)
        pre_source_range.Copy(source_range)
        #pre_source_range.Borders("Z1:Z150").LineStyle=0

        # Save changes
        workbook.Save()
    except Exception as e:
            print(f"An error occurred while copying formatting: {e}")

    finally:
        # Open the workbook
        workbook = excel.Workbooks.Open(file_path)

        # Close the workbook and quit Excel
        workbook.Close(SaveChanges=True)
        excel.Quit()

    print("Formatting and formulas copied successfully from AA to AB.")
        

# Usage
if __name__ == "__main__":
    # Path to the original password-protected Excel file
    original_file_path = r'C:\Users\jorda\OneDrive\Documents\GitHub\finevalgroup\SandBox_FFM_Updated.xlsx'  # Update this with your file path
    password = "sb!"
    
    # Path to save the new unprotected Excel file
    new_file_path = r'C:\Users\jorda\OneDrive\Documents\GitHub\finevalgroup\SandBox_FFM_Unprotected.xlsx'  # Update this with your desired path

    # Check if the original file exists
    if not os.path.exists(original_file_path):
        print(f"Error: The file '{original_file_path}' was not found.")
    else:
        remove_password(original_file_path, password, new_file_path)

        # Proceed to copy formatting and formulas if the password removal was successful
        copy_formatting_and_formulas(new_file_path, '08/31/2024')