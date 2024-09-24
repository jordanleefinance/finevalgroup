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

# OAuth2 setup
oauth = OAuth2Session(client_id=CLIENT_ID)

# Set up headers
headers = {
    'Authorization': f'Bearer {ACCESS_TOKEN}',
    'Accept': 'application/json',
    'Content-Type': 'application/json'
}

# Request the balance sheet
response = oauth.get(BALANCE_SHEET_ENDPOINT, headers=headers)
response2 = oauth.get(PROFIT_AND_LOSS_ENDPOINT, headers=headers)

# Check if the request was successful
if response.status_code == 200:
    # Parse the JSON data
    balance_sheet_data = response.json()
    profit_and_loss_data = response2.json()
    
    #print(balance_sheet_data)
    #print(profit_and_loss_data)
    # Extract the relevant parts of the balance sheet
    rows = balance_sheet_data.get('Rows', {}).get('Row', [])
    rows2 = profit_and_loss_data.get('Rows', {}).get('Row', [])

# Prepare a list to hold the data
    data = []

    # Define the date range (you can modify these dates as needed)
    start_date = datetime(2024, 1, 1)
    end_date = datetime.today()

    # Generate a list of months within the date range
    months = pd.date_range(start=start_date, end=end_date, freq='MS').strftime("%Y-%m").tolist()

    # Recursive function to parse all rows
    def parse_rows(rows, month):
        for row in rows:
            try:
                # Extracting account names and values from ColData
                if 'ColData' in row:
                    # Get the account name and value
                    account_name = row['ColData'][0].get('value', 'N/A')
                    value = row['ColData'][1].get('value', '0.00')
                    data.append({
                        'Account Name': account_name,
                        'Value': value,
                        'Month': month
                    })

                # If there are nested Rows, call the function recursively
                if 'Rows' in row:
                    parse_rows(row['Rows'].get('Row', []), month)

                # If there's a Summary, extract its account name and value
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
    # Recursive function to parse all rows
    def parse_rows2(rows2, month):
        for row in rows2:
            try:
                # Extracting account names and values from ColData
                if 'ColData' in row:
                    # Get the account name and value
                    account_name = row['ColData'][0].get('value', 'N/A')
                    value = row['ColData'][1].get('value', '0.00')
                    data.append({
                        'Account Name': account_name,
                        'Value': value,
                        'Month': month
                    })

                # If there are nested Rows, call the function recursively
                if 'Rows' in row:
                    parse_rows(row['Rows'].get('Row', []), month)

                # If there's a Summary, extract its account name and value
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
    

# Start parsing from the top-level rows for each month
    for month in months:
        parse_rows(balance_sheet_data.get('Rows', {}).get('Row', []), month)
        parse_rows2(profit_and_loss_data.get('Rows', {}).get('Row', []), month)

# Convert the list to a DataFrame
    df = pd.DataFrame(data)
    print(df)

# Assuming df is your existing DataFrame created from the JSON data
# Ensure the 'Value' column is numeric (remove commas and convert to float)
    df['Value'] = df['Value'].str.replace(',', '').astype('string')

# Group by 'Account Name' and 'Month' to handle duplicates
    df_grouped = df.groupby(['Account Name', 'Month']).agg({'Value': 'sum'}).reset_index()

    # Pivot the DataFrame
    df_pivot = df_grouped.pivot(index='Account Name', columns='Month', values='Value')

    # Optional: Handle missing values, e.g., fill NaNs with 0
    df_pivot.fillna(0, inplace=True)

# Save to Excel if needed
    df_pivot.to_excel('pivoted_balance_sheet.xlsx')

# Display the pivoted DataFrame
    print(df_pivot)

# Save the DataFrame to an Excel file
    df_pivot.to_excel('balance_sheet_time.xlsx', index=True)

    print("Balance sheet saved to balance_sheet.xlsx")
else:
    print(f"Failed to retrieve balance sheet: {response.status_code}")
    print(response.text)