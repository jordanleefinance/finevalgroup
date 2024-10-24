import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
import openpyxl
from datetime import datetime, timedelta

# Configuration
CLIENT_ID = 'ABxiy7uRFMAodckBhQH7Civ7khF4z9WDRVKwKs7YKLVJUHIZZt'
CLIENT_SECRET = 'jA88goCfvChG2HBf5oHTrOhSLu1CTMd6FqL2IoAG'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..1KzFyXpxaW1yAFJkKh8eHA.MFIf8SWJoetjFN2fp4HPqUJYY7g2lohTHmzhvp4gwqxLqdJePcrR664Z1h6vmINjbz7g-ykT0jM3jCemY8b2pOlxaJtVQe_T7cCZpyI60XiPNSS-X3PnWpwP9N334-YlPtC9SV-U9DOdjumR9cC4WBCtg8ia2jYgZQwiHAQaN6SlEfgk_knejKMnFvEmFZKfO4jtKUPZsaZ2JEv_477ibnh7tf72WTuOIwJZMrNTEBZZhSMY-7HQn5gTpDodd06umT9WTDrVyR53ev3c_TIdHOuZh0kMC9RO9bhESHEgL_NA67cUl90oolpeh-Gtmu1Qa0iorgJ7IpQeK7SVxI3INUbJbLTuMG0qOwJj20XMaumn0rX5GFWPB2A99zSdpS1D5HJZRK2B9yDYE2zovkU7JKOVFRuRySYD1BnNkSUvxMWq72zXsomB5wFjCD3hNWV8Lw25CKmiG9DA5LDhWY-6XUl3Z8X48Nrb4iJgZPHqcC692xq1a-KJ0uxckWcLK__dLMQawxDLagml2RavMA-j608zUawE4cqk6ohmHlN_ouE_sUG17067-_0b8XJbaLgwlmveK3DGxW-xEvBt9RGEh7qpmySGLwc08lOqn-R7QEjJLalYwmG_h88Zpl1P7gsq9BFhO0ezG8eMm5qMxefRMyJPJDmEQhgBP1xS7ObjwbPVd5pW--h-pNNoyxEBw3nAdL-w2Un7WWUc6PBwH6cPZDAiV01MMS9fTizXqcrr1BTUZgacqWG2gwWZ0ufBsuBU.P2NfbqzjziTDHDPFbMEhCw'
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

