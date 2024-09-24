import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook
from datetime import datetime, timedelta

# Configuration
CLIENT_ID = 'AB11726939094pDI6Hfmsb0BdxKQxrrwtBrzqZQp2yUcg8Mn1L'
CLIENT_SECRET = 'jA88goCfvChG2HBf5oHTrOhSLu1CTMd6FqL2IoAG'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..d45qCHgqinMv_EZPDhYsjw.XrJzgoBECel4j5sotG7FSoaZOVdm1tLxZs4FGcMaiRT4xisG_OJ-5UI0RdCheI1DWxtbDIM2g1F1LqOUDtqNQAH_IAweEI5sGvVI-PqjxUTV8pkiYdfkeEcL7lSBC0lUdFTWxPenxL3B_1Z7sHHr_jqE_u1Qo18Z-SeJ_pfhUyfgTPhPznSDEESNYRJr-nOC4iX5xooqj26k2FqlgJ5IElKZ8fFZENyDyFnXrLGW3tJ-W8RN9_d7mKHn6r78dykqXPNCFQ4QuycTMkeGB0FrXjVNgk6BmRwjMS5wFKwdDXXiU1W28Ma24GI7Mpl8nUd2pSavEXb3LMCQ1vfAuQz7bzHXSJpSWd1Equtig0g8lszsgweizHTl4FG6ss13kaJqx7xfMUeNU6mgniF4wFDbbCKgNTj8lSNgGMCqee45Zz_QpDFMgOFNuJOUZ93QkQpYoLk_UMg-E9FwwS8ZbUO05sj5SgNzDZ5n2kq2kVZJWPkCxEjPrXqo4iG4dNtkEOTGqr6INseIhzWaF6lEz3QZWT1PzkfAYZc69IiaxuGOlHMInl73kKCOgf-ab-HrU3-HtdU4DNsDHyFEFXi1NPJmHr820z_1a20w2a55TJn16QOUROANXRkCbdlwpVB03tm-ZVbF6rSSfkkE6qPlT0uaoIHV4c1pLMzpO3a4ye0_zSOYRW9AZc7A0y2hwmpicaM1jCUyitzDzDO1n_AFwe85MdmSwSwpzz_PAFM4xTMq8qjHYZc7npU2Xg3c1IfMN6sr.o44y2TTbs84ju5Vw2ZdJzQ'
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
print(df)

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