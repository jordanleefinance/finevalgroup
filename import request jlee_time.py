import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
import openpyxl
from datetime import datetime, timedelta

# Configuration
CLIENT_ID = 'AB11727204156Fg7jAcpAEhGthr4Ixiv9X8P3uKHbUtPhboxVn'
CLIENT_SECRET = 'jA88goCfvChG2HBf5oHTrOhSLu1CTMd6FqL2IoAG'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..dDdcI75dxCW3aFvA07E-dA.6g0ZYSbh8J-_v6QJXxCSoIKgJPGwr76hfT-TIsDxn08AAvDtNtvcoM96Mn1D_jt01KmTFr82uhuxvwzfOGgNZh7_Ia6B-r7xgV6gaTsWDU8DNrRfzfrzF1VVSELov7gNIwZG_aleSi5R44k-4mD1-r3J7oPr-ohDXDJ2T9Kprtynho-lyeCD3q2Pmkaad5UnH_0UqBcNfP1vb2kldrdlkf_obo02NSVttL-gVVZ7KCM85RddvqJyVrzdKObP6q6gX1aP6O04em6j-U0Rtt7zxC5yEEKQ8HSScUG-kmsUvbWlMmqxHvHuB7V1rmphXva0pPMF4mpgKhDaad_viWMK9eGbOwOL88pb80baMXuf9D2UuEi0yVnoLOUv-eqezLi-5dtsoFN4_ef-WMwtu40Z4H5bnxpbtBVU6bk7CTVQrtbhaCDzk6lmAwwFc9Hb5lbwcwZCX5OKmF6WQjTjZp6GKU5hQPNERaKqgGsEhIaNlJvAhqpjbtz7XOwuk58XCTYpqfSpVUpNunuOeQEBbDo6h_18_W2ZvqTifQKDwlDJlPnYo1GT-x83qUZLbaGjO4TpSwl-Tj4VS5AkqF24Eh0ZvM9VavKnp1G8hV7DIyD_cAm7zvgWLDUqH3K4sItSi0djIhV9KIconqWirUhDOw5sjoGrNtYA1V6zdf_a35M1BBKKyJGxkzj2qksVmOamOyERVer2m9DPAcCiGpYHLzMhsmZjfnCJcg9TURLyNgvaUjs.08-nypFQBfsZ2CFIxVbs_g'
COMPANY_ID = '9341452910276277'

# QuickBooks API endpoints
BASE_URL = 'https://sandbox-quickbooks.api.intuit.com/v3/company'
BALANCE_SHEET_ENDPOINT = f'{BASE_URL}/{COMPANY_ID}/reports/BalanceSheet'
PROFIT_AND_LOSS_ENDPOINT = f'{BASE_URL}/{COMPANY_ID}/reports/ProfitAndLoss'

# Set up headers
headers = {
    'Authorization': f'Bearer {ACCESS_TOKEN}',
    'Accept': 'application/json',
    'Content-Type': 'application/json'
}

# Define the date range (you can modify these dates as needed)
start_date = datetime(2024, 1, 1)
end_date = datetime.today()

# Generate a list of months within the date range
months = pd.date_range(start=start_date, end=end_date, freq='MS').strftime("%Y-%m").tolist()

# Prepare a list to hold the data
data = []

# Function to parse balance sheet rows
def parse_rows(rows, month):
    for row in rows:
        try:
            if 'ColData' in row:
                account_name = row['ColData'][0].get('value', 'N/A')
                value = row['ColData'][1].get('value', '0.00')
                data.append({
                    'Account Name': account_name,
                    'Value': value,
                    'Month': month
                })
            if 'Rows' in row:
                parse_rows(row['Rows'].get('Row', []), month)
            if 'Summary' in row:
                summary_account_name = row['Summary']['ColData'][0].get('value', 'N/A')
                summary_value = row['Summary']['ColData'][1].get('value', '0.00')
                data.append({
                    'Account Name': summary_account_name,
                    'Value': summary_value,
                    'Month': month
                })
        except (IndexError, KeyError):
            continue

# Function to parse profit and loss rows
def parse_rows2(rows, month):
    for row in rows:
        try:
            if 'ColData' in row:
                account_name = row['ColData'][0].get('value', 'N/A')
                value = row['ColData'][1].get('value', '0.00')
                data.append({
                    'Account Name': account_name,
                    'Value': value,
                    'Month': month
                })
            if 'Rows' in row:
                parse_rows2(row['Rows'].get('Row', []), month)
            if 'Summary' in row:
                summary_account_name = row['Summary']['ColData'][0].get('value', 'N/A')
                summary_value = row['Summary']['ColData'][1].get('value', '0.00')
                data.append({
                    'Account Name': summary_account_name,
                    'Value': summary_value,
                    'Month': month
                })
        except (IndexError, KeyError):
            continue

# Loop through each month and request data
for month in months:
    # Set the date range for the month
    month_start_date = f'{month}-01'
    month_end_date = f'{month}-{pd.Period(month).days_in_month}'

    # Update endpoints with date parameters
    balance_sheet_endpoint = f'{BASE_URL}/{COMPANY_ID}/reports/BalanceSheet?start_date={month_start_date}&end_date={month_end_date}'
    profit_and_loss_endpoint = f'{BASE_URL}/{COMPANY_ID}/reports/ProfitAndLoss?start_date={month_start_date}&end_date={month_end_date}'

    # Request the balance sheet and profit & loss data for the month
    response = requests.get(balance_sheet_endpoint, headers=headers)
    response2 = requests.get(profit_and_loss_endpoint, headers=headers)

    # Check if the request was successful
    if response.status_code == 200 and response2.status_code == 200:
        # Parse the JSON data
        balance_sheet_data = response.json()
        profit_and_loss_data = response2.json()
        
        # Parse the rows for balance sheet and profit and loss
        parse_rows(balance_sheet_data.get('Rows', {}).get('Row', []), month)
        parse_rows2(profit_and_loss_data.get('Rows', {}).get('Row', []), month)
    else:
        print(f"Failed to retrieve data for {month}: {response.status_code}, {response2.status_code}")
        print(response.text)
        print(response2.text)

# Convert the list to a DataFrame
df = pd.DataFrame(data)
#print(df)

# Convert 'Value' to numeric (remove commas and convert to float)
df['Value'] = pd.to_numeric(df['Value'].str.replace(',', ''), errors='coerce')

# Check for any NaN values in 'Value' column
if df['Value'].isnull().any():
    print("Warning: Some 'Value' entries could not be converted to numeric.")
    print(df[df['Value'].isnull()])  # Display rows with NaN values

# Group by 'Account Name' and 'Month' to handle duplicates
df_grouped = df.groupby(['Account Name', 'Month']).agg({'Value': 'sum'}).reset_index()

# Pivot the DataFrame
df_pivot = df_grouped.pivot(index='Account Name', columns='Month', values='Value')

# Handle missing values, e.g., fill NaNs with 0
df_pivot.fillna(0, inplace=True)

# Save to Excel
df_pivot.to_excel('pivoted_statements.xlsx')


# Display the pivoted DataFrame
#print(df_pivot)
print("Data saved to Excel files.")

import win32com.client

# Path to the Excel file
file_path = r'SandBox_FFM_Updated.xlsx'  # Update this with the correct path
password = "sb!"

# Initialize Excel application
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Set to True if you want to see Excel

try:
    # Open the workbook
    workbook = excel.Workbooks.Open('finevalgroup\SandBox_FFM_Updated.xlsx', Password=password)
    
    # Access the 'Monthly Detail' sheet
    sheet = workbook.Sheets('Monthly Detail')

    # Copy formatting and formulas from column AA to AB
    source_range = sheet.Range("AA:AA")
    target_range = sheet.Range("AB:AB")
    
    # Copy the source range
    source_range.Copy(target_range)

    # Save changes
    workbook.Save()

finally:
    # Open the workbook
    workbook = excel.Workbooks.Open(file_path, Password=password)

    # Close the workbook and quit Excel
    workbook.Close(SaveChanges=True)
    excel.Quit()

print("Formatting and formulas copied successfully from AA to AB.")