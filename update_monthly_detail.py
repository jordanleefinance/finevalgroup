import openpyxl
import win32com.client
import os

def remove_password(file_path, password, new_file_path):
    # Initialize Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Set to True if you want to see Excel

    try:
        # Open the workbook with the password
        workbook = excel.Workbooks.Open(file_path, Password=password)

        # Save a new copy without password protection
        workbook.SaveAs(new_file_path, Password="", FileFormat=51)  # 51 for .xlsx
        print(f"Password removed successfully. New file saved as: {new_file_path}")

    except Exception as e:
        print(f"An error occurred while removing password: {e}")

    finally:
        try:
            workbook.Close(SaveChanges=False)
        except: 
            pass
        excel.Quit()

def copy_formatting_and_formulas(file_path):
    try:
        # Load the workbook and select the sheet
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook['Monthly Detail']

        # Define the columns
        source_column = 'AA'
        target_column = 'AB'

        # Copy formatting and formulas from AA to AB
        for row in range(1, sheet.max_row + 1):
            source_cell = f'{source_column}{row}'
            target_cell = f'{target_column}{row}'

            # Ensure source cell exists
            if source_cell in sheet:
                # Copy the value and formula
                if sheet[source_cell].data_type == 'f':  # If the cell has a formula
                    sheet[target_cell].value = f'={source_cell}'
                else:
                    sheet[target_cell].value = sheet[source_cell].value

                # Copy formatting if the source cell has style
                if sheet[source_cell].has_style:
                    sheet[target_cell].font = sheet[source_cell].font
                    sheet[target_cell].border = sheet[source_cell].border
                    sheet[target_cell].fill = sheet[source_cell].fill
                    sheet[target_cell].number_format = sheet[source_cell].number_format
                    sheet[target_cell].protection = sheet[source_cell].protection
                    sheet[target_cell].alignment = sheet[source_cell].alignment

        # Save the workbook
        workbook.save(file_path)
        workbook.close()

        print("Formatting and formulas copied successfully from AA to AB.")

    except Exception as e:
        print(f"An error occurred while copying formatting: {e}")

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
        copy_formatting_and_formulas(new_file_path)