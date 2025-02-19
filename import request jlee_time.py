import requests
import pandas as pd
from requests.auth import HTTPBasicAuth
import openpyxl
from datetime import datetime, timedelta


# Configuration
CLIENT_ID = 'ABI1vL3lZyVenjBBSq6C1BYOcfHtgdzske95vpcqejvKtLejeN'
CLIENT_SECRET = 'B2qiYeQMGfNl68M1yEJu1bsT8G94C16EDai55O2p'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..f6TfPURW7cgd1yP_rVbsDg.2InIxggp2nQvENm2Cej-Wn_IY_TZMa8Y0sXmvVD4UAq7DT8krOvxiwjEpQxab5PMd9OCGivbgMeswYD15aXUQLgjcOgodwANBpi-bB3wTiid_-nlZmC9wpdbABmcEYsUaLwKiRxSLJu3azrffD8cuxDLipmHdopTRh7jMwUDw5muMjxmfTcbr0QIybAX2GjZ_YLZL3oWSMQ3_154-4-M-M-2OyhLkVKR9ZW9OECjU4DRg2BBNwe_9O7aNZh8q6cxJJIew1xwIst8hXS1EYJeRSdbXz6A7Z_a2XSe_oNrAezam_c6bJNgG80-ZW6lzOdBBJz7xsR-7GYSA1rsAhXn_uU1gjaangMG3j2TfYJKnW-NgWYc903FaNk6MpVF-MBRvwuv9jieMS1fPvCXaFLBeQ1AWElHGDgtGhHTkKYlCbdUxWTMpJHuhe9917AYMyBjeARS6J--BH7QAKPh2WT1sZdYt6FiWSSuev8iBeGNKTWPW5TzzkEM4ufWdiS9nAcl7OYTKWlwAaYKp_s9esQv6lXTsAbhU6gVmuAcv7TssWK95okGre8czbW_QZVgLsWJdfM38orWo0GgRi7MHSL8Dv0GCZDgjprW8H1hxGBPf_qqGQeE4x70tlY_LxuEMqPnjE0zquOW17BRMVzdQt5bFoUt2lE5uV5n-y2M3As5FKqDoaoV84idjhMa8UBkQ2P490i5egNEgU9LKXEHaYdQE2EHLjWAD7F_AOYTnyp9bwO3KSJ6kuYnEXFQ8VPfWYio.Z4e8ahbgYLzVor5pEx3Zpw'
COMPANY_ID = '9341453163478319'

# QuickBooks API endpoints
PROD_URL = 'https://quickbooks.api.intuit.com'
DEPLOY_URL = 'https://sandbox-quickbooks.api.intuit.com/v3/company'
BALANCE_SHEET_ENDPOINT = f'{PROD_URL}/{COMPANY_ID}/reports/BalanceSheet'
PROFIT_AND_LOSS_ENDPOINT = f'{PROD_URL}/{COMPANY_ID}/reports/ProfitAndLoss'

# Set up headers
headers = {
    'Authorization': f'Bearer {ACCESS_TOKEN}',
    'Accept': 'application/json',
    'Content-Type': 'application/json'
}

'https://qbo.intuit.com/app/switchCompany?companyId=9341453163478319'
'https://qbo.intuit.com/app/reportv2?token=BAL_SHEET&show_logo=false&saved_rpt_token=BAL_SHEET&date_macro=lastmonth&low_date=11/01/2024&high_date=11/30/2024&column=total&showrows=active&showcols=active&subcol_pp=&subcol_pp_chg=&subcol_pp_pct_chg=&subcol_py=&subcol_py_chg=&subcol_py_pct_chg=&subcol_py_ytd=&subcol_pct_row=&subcol_pct_col=&cash_basis=no&collapsed_rows=&edited_sections=false&divideby1000=false&hidecents=false&exceptzeros=true&negativenums=1&negativered=false&show_header_title=true&show_header_range=true&show_footer_custom_message=true&show_footer_date=true&show_footer_time=true&header_alignment=Center&footer_alignment=Center&show_header_company=true&company_name=Legacy%20Tattoo&collapse_subs=false&footer_custom_message='

# Define the date range (you can modify these dates as needed)
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 11, 30)

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
    balance_sheet_endpoint = f'{PROD_URL}/{COMPANY_ID}/reports/BalanceSheet?start_date={month_start_date}&end_date={month_end_date}'
    profit_and_loss_endpoint = f'{PROD_URL}/{COMPANY_ID}/reports/ProfitAndLoss?start_date={month_start_date}&end_date={month_end_date}'

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

# Check for any NaN values in 'Value' column
if df['Value'].isnull().any():
    print("Warning: Some 'Value' entries could not be converted to numeric.")
    print(df[df['Value'].isnull()])  # Display rows with NaN values
else:
    # Convert 'Value' to numeric (remove commas and convert to float)
    df['Value'] = pd.to_numeric(df['Value'].str.replace(',', ''), errors='coerce')

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

