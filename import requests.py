import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook

# Configuration
CLIENT_ID = 'ABxiy7uRFMAodckBhQH7Civ7khF4z9WDRVKwKs7YKLVJUHIZZt'
CLIENT_SECRET = 'jA88goCfvChG2HBf5oHTrOhSLu1CTMd6FqL2IoAG'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..S30MQ9oC4FLfQhrrXNxrPg.w2MBqCXnNVg7_P4uBwVxnEa1cK3BLvCax5R3attuJNm8Ein-p-MUCgBkdcp9CO7BIBNtNHqdHtqHDQ_h685S7-_G-72u7qng_O-vEVhZqeSBXfxamI2TUnmqyKgSVvTln47ei1d9JtSAVB2soEr0OugvM1yvuIdhOgopvOln4v-mvcejiz7IlX3Wg3jXo6k1mGjPhf5FYRxyFpko1jZy7XtIMqHd18uKeY9Xop-SPtOn4cg6BFJFOqeZ_bdcCjcCQ6iqISgTaeuoAJUtDIs_cjEeQSKjOBnvepKpJcm9znFuJv1o3OVNBFYjRyvvKtPN29MYjcz1zd9aNC5yxMkpJjeqs3d51vkEsdAxX26y--4PQyCoBAtqSmDGZzixuXUi_tNfmsYMWeZxN5jh2x3mAx46-YMbx0WvS4AaK-p2L3BIG6ruMFYvyYyJWQa5kGN0TqaOJ5yMZKNPZaNIBm0xx7SyEyKTojwUpjNKS8uj3JtCdhGs9MQqs5MfEecXE0Ljh8hO9KwTuEiPXeYDfxkgG8pezxAxOrCCdvMi8TbXZx-Y29WyEpykA27RyuCNhgFK3EmTqoZDJfnrmJ4NdN3b887EkOkYO1M0nPHZtWZcdAmMyHSQcHMJt6fr_vkNayYpKitr1qaaa7SIxd605Wg_aW8YTo3R5GmYVo5v43G93hGfGWl23Scon8u-4z0AiAZt1TqX-rxTOL7PnjiZlWA6tR8Uh3SPJnrcGtpWqvIKA_6pPC3aHvWkiOLCjfnTvG9V.ev8sUL8eQVE0iJLoi6upww'
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

    for row in rows:
        if 'Summary' in row:
            data.append({
                'Label': row['Summary']['ColData'][0]['value'],
                'Amount': row['Summary']['ColData'][1]['value']
            })

    # Convert the list to a DataFrame
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    df.to_excel('balance_sheet.xlsx', index=False)
    
    print("Balance sheet saved to balance_sheet.xlsx")
else:
    print(f"Failed to retrieve balance sheet: {response.status_code}")
    print(response.text)
