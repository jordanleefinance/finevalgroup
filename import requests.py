import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook

# Configuration
CLIENT_ID = 'AB11725981363YVFLMqg7tO2OpRZyk1x2aQ60WUPAL1aA7OdKY'
CLIENT_SECRET = 'jA88goCfvChG2HBf5oHTrOhSLu1CTMd6FqL2IoAG'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..C_suFv2uL2IoLBIpH735GQ.Q6a61mhj5h2C8ZbyGbQQGvhm_EeZ0ShKfvdywf60IdNd3rfb6-m6j_Wk628Bm0wrNz0vVo1kH8R18j2UUHwFpkscTttuktms9X6evzAZOCzobJWcKLy9S3HSNhLzQeVp2xAQrxjY4ZUImTO7IyFUJxQGUjyZFR3YYVD4_EPO-Iea-avYWPW9MoHnLceGcM0K_sPW7wuxHfD71tWe0jpU6jXZgBKwgIryQyW4-kZ8gccLhrovstwf5YAQOyEzmTlVIUem1lGGcJZ0ML6Wo_yUGr0urlwQ41355CrGVnmYPtB5bhpIyxuRsx17iBdZ6-GbKj7mGoiPLOJxweGboejilohxWeNKO2MlEoBvjh-GQ_MR76XB28XXGaUP87DjCkYM-qXyaZZGJOUSPmGuQiG9ZC4J7aWZweg_8vZd_RhCEwFBNbELb0CQ6s77NhJ2BOnZT_vjOvJ3EbLjv46eTJrnLrk9NrUkVTLmBcd-816__FpJQDY4tv5PrQ5rNzza0dARk9YA_W2lIMPm24hpVC1cLY0ADpC6I-7bC_lzxC-wFP7ic6eZoxCkJj75Q2yX6DIvrrIupZCZmW4TQ7ASbB0T4h7K0WIDYhhkp0L7x_a98PcG4baGzJyY1VyKbBMMUS_ZumoDt7gve_4tUEZwXt1pWhz2dkd9q8LjftEYenIxof8yZ3hsvH9vWlviH0gf8aqXJ1hBJOYADl7YfufUyMX0OA4Vf7ToMtsoLVtkvx8Un-MwtzeN9SxnSDBnJyhirPTq.vuzPrIG5OitbknCjXBAVYA'
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
        

    # Loop through rows of data
    for row in rows:
        # Check if 'Rows' exists within the current row
        if 'Rows' in row:

            temp = list(row['Rows']['Row'][0].items())
            res = [idx for idx, key in enumerate(temp)]
            #print(res)
            #print(row['Rows']['Row'])
            #print(row['Rows']['Row'][0]['Rows']['Row'][0]['Rows']['Row'][0])
            # Parse the key data necessary for financial forecasting
            i = 0
            for key in row['Rows']['Row'][0]['Rows']['Row']:
                
                try:
                    account_name = key['Rows']['Row'][1]['ColData'][0].get('value', 'N/A')
                    value = key['Rows']['Row'][1]['ColData'][1].get('value', '0.00')
                    print(account_name)
                    print(value)
                    data.append({
                        'Account Name': account_name,
                        'Value': value
                    })
                except IndexError:
                    print("IndexError")
                    continue
                except KeyError:
                    print("KeyError")
                    continue

                i+=1
                print(data)
                #print(key['Summary']['ColData'][0]['value'])
                #print(i)
                #print(row['Rows']['Row'][key])

            '''# Loop through sub-rows in 'Rows'
            for sub_row in row['Rows']['Row']:
                if 'ColData' in sub_row:
                    account_name = sub_row['Header']['ColData'][0]['value'] if sub_row['Header']['ColData'][0]['value'] else 'N/A'
                    value = sub_row['Header']['ColData'][1]['value'] if sub_row['Header']['ColData'][1]['value'] else '0.00'
                    print(account_name)
                    
                    # Add the parsed data into the list
                    data.append({
                        'Account Name': account_name,
                        'Value': value
                    })'''
            
    # Convert the list to a DataFrame
    #df = pd.DataFrame(values, columns=headers)
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    df.to_excel('balance_sheet.xlsx', index=False)
    
    print("Balance sheet saved to balance_sheet.xlsx")
else:
    print(f"Failed to retrieve balance sheet: {response.status_code}")
    print(response.text)
