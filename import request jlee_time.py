import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
import openpyxl
from datetime import datetime, timedelta

# Configuration
CLIENT_ID = 'ABI1vL3lZyVenjBBSq6C1BYOcfHtgdzske95vpcqejvKtLejeN'
CLIENT_SECRET = 'B2qiYeQMGfNl68M1yEJu1bsT8G94C16EDai55O2p'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..EWDUSz9ZUmHq45dam0IbmA.gOxZJBUdWJya6czfKxB8ozTIJvT1G4eslH3v4sNz0A-k9Ex-rOZJkKYzOLqFOfMl23ZKPl616FUFbOOMpjzTVMuIqf0zG898N9YBh3jdUCO9vVqioQQV7YjC3w9_aB9GYLT52kkTgBgINptpH1nmeTZk5JVH7AYD0k2tCgXgUGn-pKiFl9NU5Yzpoq96XkAbQUnnpzshiofgc7gbJiSzPvk4VC9z0pW9ZyPuwX0ItUmKH0SwcYaqeF9BLZrsJsr5LebsvyVvBGVhlVrW77nggXzfeZfI0whN0IkLTKebD-ZEYh4YeAgJ_5k5QLkiBy5F0-ANOZmDynzggUuMuK3UgX9aZPoFTvJEJs6CF476qqDgL8ANvBzPz9TmcdYwYCb616z88CMolsMo79KlDiicHszWoLAJjyQQ7P9BPtZN4aDPfJkvvE9xcFqwZHZmFCAdtebq_MiwSgD_Kd2Uai1AfN1kSmQpJIuvtky-splPuP5PJwLYjKCwDXCBMrCnfM5LOFObQsgYwOvsUkQl9MqK9yN9EYLtf9O13gRcmHMF3ladY0jDb96v6M6ce_Xg9U9EKRrQCn5SkaUC9MjRQYzpbiKhDGFHGIHG-jjgrPtwgZ7d49PFe-jJBD_maKO44RA5e9Z_aYCBtLVm4O58j7yKDJwfkwVTigvLHPbLRRiNP4qBqLTE06f0T4qS-rR-xUV2wbNb34NyBlw20dzA3iRCPNBdunbPc0-0UgqcPjRbezI.OlGmQoHG2YEeUJ7Ivl_pIw'
COMPANY_ID = '9341453163478319'

# QuickBooks API endpoints
SAND_URL = 'https://sandbox-quickbooks.api.intuit.com/v3/company'
BASE_URL = 'https://quickbooks.api.intuit.com/'
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

