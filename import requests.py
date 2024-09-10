import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook

# Configuration
CLIENT_ID = 'AB11725981363YVFLMqg7tO2OpRZyk1x2aQ60WUPAL1aA7OdKY'
CLIENT_SECRET = 'jA88goCfvChG2HBf5oHTrOhSLu1CTMd6FqL2IoAG'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..GUa4KlR8L9x4w1C-gWK6pg.ykzy1m1U8qKgpTj6e4u9lOFA2cAVK2KFidKyDeqApx1E-SyeTZOGPU3cW1NTU1ZgHpQEBK7xOYS2HY8LjVW7U7b-7l6PsAaWrsSvSu_nNe7zHmSuP4OV6B8mmDL1HZlWC0R9HQDjenCvAo51TERQ0vTvcCkPOyPrrZKVQKGC9bVNHlBvDZeYyCM7Y7yt3bn_3WOxr4t_7F1vL6HmjOMAtdiZIEzzzAYfSP4ePEMqwiuAeXPYMEj5yWJpQ8pTHczyg5zJ3JMZHOdT97GIPcb9c93Wq7Kr9qCMaiJmh5hLeGEHAuAORDXL1yv9bi4wE6BTLQx0hvc5FNHndEfBFHq8lOYDbYAvNZDXAFsywoy1TDPk_uwRwxPD_zq0RC-kDmqzE0ft8r6PInGrv8Sb79RtnXTskeMKLd6NtchSNaDQyX9A-njhlOop4wFdj3GOp0WRu2_DqcOPNbRC8xuYrEq_yjfJbQp3LM-3pWao-lJ_rNWipP_f6uGvR6rUndKlIUG3Q8wAN8XwkouCCofQZxWmhrkLUrraGhio-L5JbrZ__O73772kY3gChGey2f1KaYjc4_k_RWSKaGu6Tm-Z1PGvbCj4VEYg67QnpDqzDsYnGuNApexJJ-8wzOw2XWWj3r3PqqVNDf1c6zxhP5shc1GoQ449HnAPZ4CosExivJSoFfP4vhysp7SnvPcO7t6jbNSiXiiXD_2mDbwpivrNQhL-C9Tzhk8IcGduJptePZt8vZWxQ75zMmFR04LfedSWfXDk.vqSqN8ktfQsYK0oYVDYqgg'
COMPANY_ID = '9341452910276277'

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

# Request the balance sheet
response = oauth.get(BALANCE_SHEET_ENDPOINT, headers=headers)

# Check if the request was successful
if response.status_code == 200:
    # Parse the JSON data
    balance_sheet_data = response.json()

    # Extract the relevant parts of the balance sheet
    rows = balance_sheet_data.get('Rows', {}).get('Row', [])
    
    # Prepare a list to hold the data
    data = []
    '''def parse_balance_sheet(data):
        balance_sheet = data.get('Rows', {}).get('Row', [])
        headers = []
        values = []

        for row in balance_sheet:
            cells = row.get('Cells', [])
            if not headers:
                headers = [cell.get('value') for cell in cells]
            else:
                values.append([cell.get('value') for cell in cells])

        return headers, values'''
    

    for row in rows:

        '''if 'Summary' in row:
            data.append({
                'Label': row['Summary']['ColData'][0]['value'],
                'Amount': row['Summary']['ColData'][1]['value']
            })'''

        # Append the extracted data to the list
        data.append(row)

    # Convert the list to a DataFrame
    #df = pd.DataFrame(values, columns=headers)
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    df.to_excel('balance_sheet.xlsx', index=False)
    
    print("Balance sheet saved to balance_sheet.xlsx")
else:
    print(f"Failed to retrieve balance sheet: {response.status_code}")
    print(response.text)
