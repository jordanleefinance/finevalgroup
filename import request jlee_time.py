import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook
from datetime import datetime, timedelta

# Configuration
CLIENT_ID = 'ABMwyfUQxen1pOUKa4o7WxE2BGw2tfok9CDlAdvo8PjQivJKWc'
CLIENT_SECRET = 'oxU8wJYTywbZuFpYiEHADQbt77EFi26hK32yVdUa'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..HBoKcuTO8z3xLZZ2nHvgqw.FQHPBQagFSRtP0peDsiH5lwKhW2xyXKcfNsj3_Ghvk9VVikkxT20uBNXZP4CedAbWf8qk7AZdQa9p85GMd5iD1HQQVeBD8AruOZzKeZADRU8Av23QSSoXoRXSuvqHVmG1n4AQZOfMUfuF9CwIwJBq3re0pfIi0NrQ_y7z8EGip-KXS33UfVs02wlxXLfCLQhIRTIRdJ2QGBiUAx5-5sHibDuQ2_MlxN3qQw7uTpMN537peqT9F9NLStBnVud9Yq8b5w7T5AdpHlL_2EQJ6jN3mOX64efynIhvwDxdwpyhXgWXDVsMQPFhuGoP7EFmtdXmW-ijegSltkqQJqkcrYoO4J_ZksWF9I7SNWm6udEcbdAt9M4rf6fcTyhJyxfbeHHPfCdcudB_77lRMdrnDmX6w8QwCNixKx1Czqh3PJRC8zO6NEdJbo6Q4Jtdabt6-nt5T4rz6ngi4nRchUKRlZg0V8-nxg8x6gFkSpC_h-GpGXCHUcAvh5BRN-wKbAJHnXBMmt_UwND02LBrcQSddG5w-prfSAVWxsQq-vCnn5roByP0PQLGpJv0dYbJLtxVLSqWudHxBZyFDWRIvtVKofsV4X8WxUfTFQ5h8YMXC1-4GYXGFfoqhnyx7Io44UuUM96u9taasG5NpmGxZwSk3pss1eqaWw1eEwCNTLZYN96JO83e52LDLQfNas5mIk2mfbUr_19hw_TYwCiq21wdCl-asgXIMN_Qd7LczxHqBTqe0Q.aL0mj-rYpdW4xf72kOxt1Q'
COMPANY_ID = '9341452948993561'

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
print(df)

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
df_pivot.to_excel('pivoted_balance_sheet.xlsx')
df.to_excel('balance_sheet_time.xlsx', index=False)

# Display the pivoted DataFrame
print(df_pivot)
print("Data saved to Excel files.")