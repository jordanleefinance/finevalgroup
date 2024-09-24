import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from openpyxl import Workbook
from datetime import datetime

# Configuration
CLIENT_ID = 'ABMwyfUQxen1pOUKa4o7WxE2BGw2tfok9CDlAdvo8PjQivJKWc'
CLIENT_SECRET = 'oxU8wJYTywbZuFpYiEHADQbt77EFi26hK32yVdUa'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..gXRM6nVwFN4VlOr8KV9LJQ.BhCdvKaYwzmF7SOI3YpCwwTEQi-Bjd3vc3PdDNa4WBjks0hCQjoXoGPDDflV6DZ_87BTxFLo_Xsw4KomreODsayFbMRddAumuZOZU_mPRfXoyVqseHIu1zLqJ6aV54gHVreR0NxjozMY0j_kRzAeUcmwnMsdUXNKFzfqG5LWLEiM8dDr2K19Ju4o0boDN36KXb1506rT1cCR-jM3BX4DPUZkyk-JOahhlPSl0BzSKdv66_mvY-Mm_8dNp4jtXSeT_ycQ8WEkbwja9JK1NauVygwV5Ju0125pQoL3Rubyfdlgh7NqW1pDb6Wf4Kj8WXD9rDji0XILBq5TQHYUdWmBld_nq73W5zEYR3gj5sVwfAJDRWVnaB2n8ax5YRsRFp1bhSAQ3s5HwqCAzEE6CEJgykmuntOPi_G2Dq1LA0A99kwJbltQWI347OG7R3yCZiUXsILut6EPHg8XKF4v6apuBOHAQh-DCyk_tbOYKx1OedlPVfBm5O7X6WtMVEznnq-J5KpawfLQJufiM0nM3k4_zpXoo7jEOsJu8kM_v78kW-W_7DGs1mKfXzGzNhhtCi56qbfdZ1NTtvruUT4ZMSswJaWXW7EIGUJO-h9M7PYZy5GgjLUveN_Kn9iOJLcD2MEeeLlWHJnPdnEbHLXCD5m03fUIeObPkYqdW1Xtk361Kwv0qHUwYKcTqf-HOEpwAkZ-uaXkHeHkYB3T62W9PvaT5ullDXtPKYTkWBRCN-s_-c_5QY6JCphhEkV4mi1bOG-M.Hi0XQ9UQfyBvMzd3g0RCbw'  # Replace with your actual access token
COMPANY_ID = '9341452948993561'

# QuickBooks API endpoints
BASE_URL = 'https://sandbox-quickbooks.api.intuit.com/v3/company'
BALANCE_SHEET_ENDPOINT = f'{BASE_URL}/{COMPANY_ID}/reports/BalanceSheet'

# OAuth2 setup
oauth = OAuth2Session(client_id=CLIENT_ID)

# Set up headers
headers = {
    'Authorization': f'Bearer {ACCESS_TOKEN}',
    'Accept': 'application/json',
    'Content-Type': 'application/json'
}

# Define the date range for the last 12 months
end_date = datetime.now()
start_date = end_date.replace(day=1) - pd.DateOffset(months=12)

# Request the balance sheet for the specified date range
params = {
    'start_date': start_date.strftime('%Y-%m-%d'),
    'end_date': end_date.strftime('%Y-%m-%d'),
    'summarize_columns_by': 'MONTH'  # Requesting monthly summary
}

# Request the balance sheet
response = oauth.get(BALANCE_SHEET_ENDPOINT, headers=headers, params=params)

# Check if the request was successful
if response.status_code == 200:
    # Parse the JSON data
    balance_sheet_data = response.json()
    
    # Extract the relevant parts of the balance sheet
    rows = balance_sheet_data.get('Rows', {}).get('Row', [])

    # Prepare a list to hold the data
    data = []

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

    # Generate a list of months in the specified range
    months = pd.date_range(start=start_date, end=end_date, freq='MS').strftime("%Y-%m").tolist()

    # Start parsing from the top-level rows for each month
    for month in months:
        parse_rows(rows, month)

    # Convert the list to a DataFrame
    df = pd.DataFrame(data)

    # Pivot the DataFrame to show accounts as rows and months as columns
    df_pivot = df.pivot(index='Account Name', columns='Month', values='Value').fillna('0.00')

    # Save the DataFrame to an Excel file
    df_pivot.to_excel('FinEval-JordanLee/finevalgroup/balance_sheet_monthly.xlsx')

    print("Balance sheet saved to balance_sheet_monthly.xlsx")
else:
    print(f"Failed to retrieve balance sheet: {response.status_code}")
    print(response.text)
