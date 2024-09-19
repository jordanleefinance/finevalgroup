import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook

# Configuration
CLIENT_ID = 'ABMwyfUQxen1pOUKa4o7WxE2BGw2tfok9CDlAdvo8PjQivJKWc'
CLIENT_SECRET = 'oxU8wJYTywbZuFpYiEHADQbt77EFi26hK32yVdUa'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..YDFPSN1k8Vh_FJbmjbxFbg.TEuQm8arOaHsK1sxbdBhYw4XZFuze5L7-AvW1WODjkUpvZmVKO1d7fYFXo3lK2yudE5eN5noqJ5aCQgWCMe4cHznZOdm0FhgsOpaPZ8UEmhPBWr7yQ1L_Wlkx3UEBAJuKJA6NO6LXnZ7dKLvqbLSNcCl61GrBD8G8NwBeB2KP4NTInvQUrc0XIam1jJRxNxSCmdkS87uSw6VD9nWh8YJO8iBVBsAbShk-87F8V_Evgj17kD0EUdvvf8LY5mfOsUm_eREqClwtTyk7TguJLn5tJqdKoGNs2s0t6yFZHcSbiZXxDOa8v633qL64YHSiVvFWb61bS_RXBBkiiRFlV16v6q2Mwchi0H1BdICicN0PHh6Ra43MoHO2tS8CsgdU5xCuJoBZR4Jgk3nK3UrXpm6pTTybp8jWZkZJz57L4z_klx53vUBdFZhLfEJM7yVGaceCBku7bW6Fkzl1fw0A__fMsl5W1qyVt5pa7MtCDJwWt0E68J8c50sYId3nsO2bMY27G4iSFCt6rWHnjkoD0BuGn_1jcti7A7Rcd_TlQXNB5y_VAAQSMW0cg3YBJxmTwLiTE-R8qmGjwK1EEgrnEwcSkPTof0_P2PuScO0tdOE0SaXTI57cbEKr59Jda5ddtbRvAG_ZfD3HuJZDSNujSNP4dC8jcu6qMwnHCgtkUUnoNui7Krj7dHJKOfiygzpHaLhdtnR6tM3HxAa91L94tQMeunkczDO4IAzZeqmR5nBkzU.drAXGQskSUthX_ODMVRUOg'
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
    df.to_excel('FinEval-JordanLee/finevalgroup/balance_sheet_salam.xlsx', index=False)
    
    print("Balance sheet saved to balance_sheet_salam.xlsx")
else:
    print(f"Failed to retrieve balance sheet: {response.status_code}")
    print(response.text)
