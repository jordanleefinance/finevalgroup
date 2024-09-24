import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook

# Configuration
CLIENT_ID = 'AB11726782003qxZxMYCyRD5QoRExBReqO29Auko4oz28EyTZo'
CLIENT_SECRET = 'oxU8wJYTywbZuFpYiEHADQbt77EFi26hK32yVdUa'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..-PzXhP9V0mfg0lFt_xGbYg.6ihlJPDiCs0qq_IEaDyVIYs6We0zC8YlknlzwIgudowj6rn1e7GsPDwKUU64Wa27giNt9NFq4LWBn81BKp9aW6ENofPLDzYrgH9sGaypu-SsUDy3v4rVY52puKNFtzZFgn815EXcHKPyodcgxviyL8PT_cNVpq_OsVJkDrkOUr7XpquUgWhxiU6htS-JBXF-bhMwG3xmP-otbpJXrYHYOBNVNafsuj9IWBxrYMvdCzeTtZWDK3Y-5H9866qmDfPjWdYnKq_8GVI672IQL3kdKM9oFAlrcSO1ebwoFGqVvnjr-_r_wIqJmj0lcXa2fxqwsYKmoVp8j0-jNWBOH8QWvRG5IUi64P61sW-1f9lLIu3UjpgOFjYzer0eEH8Zy2qR3aar0hnECmpWkT0pNNSE4WQUhKvs7EbnwtQe8auQDO21PPXuH1Ztyi9HQhXzjnhNZ6A57pvdgqT6_Kn_Bf1uamCyn3CawcD2oQhWUZVEiWIWiHEGUBrysfDz1QEyS5JNs1BHwYAsIfi1q-dFV-sLhbcOi8Krf15EU07ZmlsKmNP8jC1HsEkM0-3Di92CPbI1ugSnvbOy2_e-Qs0ZVtlAivacVqClnbDXt2cs5I9iO_8-54IkssDBV1LWDumWVGZTiNYE8MlvFrHdk5coDnDCbr6GV5qPZXBWGOPApfrQo7cpDl7-Myhp9wRtldFnDTe-s0EQ73z_3m1JBIAGP1PsWWKmrLhB9RKBloWFW8aHefFmKCfWLpV8NUGraD5GeMrc.G0GAbgZdseoIdEZ0p0NCxw'
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
    
    #print(balance_sheet_data)
    # Extract the relevant parts of the balance sheet
    rows = balance_sheet_data.get('Rows', {}).get('Row', [])
    
    # Prepare a list to hold the data
    data = []
     

    # Loop through rows of data

    def parse_rows(rows):
        for row in rows:
            try:
                # Check if the row contains 'ColData' for account name and value
                if 'ColData' in row:
                    account_name = row['ColData'][0].get('value', 'N/A')
                    value = row['ColData'][1].get('value', '0.00')
                    data.append({
                        'Account Name': account_name,
                        'Value': value
                    })
            # If there are nested Rows, call the function recursively
                if 'Rows' in row:
                    parse_rows(row['Rows'].get('Row', []))
            except (IndexError, KeyError):
                continue
    """for row in rows:
        
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
                    #print("IndexError")
                    continue
                except KeyError:
                    #print("KeyError")
                    continue 

                i+=1
                print(data)
                #print(key['Summary']['ColData'][0]['value'])
                #print(i)
                #print(row['Rows']['Row'][key])

            """
    # Start parsing from the top-level rows
    top_level_rows = balance_sheet_data.get('Rows', {}).get('Row', [])
    parse_rows(top_level_rows)

# Convert the list to a DataFrame
    df = pd.DataFrame(data)


    # Save the DataFrame to an Excel file
    df.to_excel('balance_sheet.xlsx', index=False)
    
    print("Balance sheet saved to balance_sheet.xlsx")
else:
    print(f"Failed to retrieve balance sheet: {response.status_code}")
    print(response.text)
