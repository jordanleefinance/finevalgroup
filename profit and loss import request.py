import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook

# Configuration
CLIENT_ID = 'AB11726783993klv3O6o5TO22HTFdoLHMIRLzJubVpksoQNJJh'
CLIENT_SECRET = 'oxU8wJYTywbZuFpYiEHADQbt77EFi26hK32yVdUa'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..gXRM6nVwFN4VlOr8KV9LJQ.BhCdvKaYwzmF7SOI3YpCwwTEQi-Bjd3vc3PdDNa4WBjks0hCQjoXoGPDDflV6DZ_87BTxFLo_Xsw4KomreODsayFbMRddAumuZOZU_mPRfXoyVqseHIu1zLqJ6aV54gHVreR0NxjozMY0j_kRzAeUcmwnMsdUXNKFzfqG5LWLEiM8dDr2K19Ju4o0boDN36KXb1506rT1cCR-jM3BX4DPUZkyk-JOahhlPSl0BzSKdv66_mvY-Mm_8dNp4jtXSeT_ycQ8WEkbwja9JK1NauVygwV5Ju0125pQoL3Rubyfdlgh7NqW1pDb6Wf4Kj8WXD9rDji0XILBq5TQHYUdWmBld_nq73W5zEYR3gj5sVwfAJDRWVnaB2n8ax5YRsRFp1bhSAQ3s5HwqCAzEE6CEJgykmuntOPi_G2Dq1LA0A99kwJbltQWI347OG7R3yCZiUXsILut6EPHg8XKF4v6apuBOHAQh-DCyk_tbOYKx1OedlPVfBm5O7X6WtMVEznnq-J5KpawfLQJufiM0nM3k4_zpXoo7jEOsJu8kM_v78kW-W_7DGs1mKfXzGzNhhtCi56qbfdZ1NTtvruUT4ZMSswJaWXW7EIGUJO-h9M7PYZy5GgjLUveN_Kn9iOJLcD2MEeeLlWHJnPdnEbHLXCD5m03fUIeObPkYqdW1Xtk361Kwv0qHUwYKcTqf-HOEpwAkZ-uaXkHeHkYB3T62W9PvaT5ullDXtPKYTkWBRCN-s_-c_5QY6JCphhEkV4mi1bOG-M.Hi0XQ9UQfyBvMzd3g0RCbw'
COMPANY_ID = '9341452948993561'

# QuickBooks API endpoints
BASE_URL = 'https://sandbox-quickbooks.api.intuit.com/v3/company'
PROFIT_LOSS_ENDPOINT = f'{BASE_URL}/{COMPANY_ID}/reports/ProfitandLoss'

# OAuth2 setup
oauth = OAuth2Session(client_id=CLIENT_ID)

# Set up headers
headers = {
    'Authorization': f'Bearer {ACCESS_TOKEN}',
    'Accept': 'application/json',
    'Content-Type': 'application/json'
}

# Request the balance sheet
response = oauth.get(PROFIT_LOSS_ENDPOINT, headers=headers)

# Check if the request was successful
if response.status_code == 200:
    # Parse the JSON data
    profit_loss_data = response.json()
    
    #print(balance_sheet_data)
    # Extract the relevant parts of the balance sheet
    rows = profit_loss_data.get('Rows', {}).get('Row', [])
    
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
    top_level_rows = profit_loss_data.get('Rows', {}).get('Row', [])
    parse_rows(top_level_rows)

# Convert the list to a DataFrame
    df = pd.DataFrame(data)


    # Save the DataFrame to an Excel file
    df.to_excel('FinEval-JordanLee/finevalgroup/profit_loss_salam.xlsx', index=False)
    
    print("Profit Loss saved to balance_sheet_salam.xlsx")
else:
    print(f"Failed to retrieve profit loss sheet: {response.status_code}")
    print(response.text)