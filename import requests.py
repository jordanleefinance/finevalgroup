import requests
import pandas as pd
from requests_oauthlib import OAuth2Session
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook

# Configuration
CLIENT_ID = 'ABxiy7uRFMAodckBhQH7Civ7khF4z9WDRVKwKs7YKLVJUHIZZt'
CLIENT_SECRET = 'jA88goCfvChG2HBf5oHTrOhSLu1CTMd6FqL2IoAG'
ACCESS_TOKEN = 'eyJlbmMiOiJBMTI4Q0JDLUhTMjU2IiwiYWxnIjoiZGlyIn0..Fs0FejquBVx4cEuJiZamGA.RVvPWFYjIlAyWSLyGINZxmHiNJIK6QfmsQ_xcPgi4IpBzVkGehRU7emAJuJKB3v_kDTjY9AetD6avfVx1sYa3s54wKqWcI_sP2BNS0vlbXoP0Elc1kZOrR4h8NfELaif5tXmU9g_1Eq2K_43Bm0ruu9qf9KX_3vnFRxv_u5NTy30TZqCrxqUSm7DJ4a0U92otmNNK9NmPETYdCipCH3-BlCH-tz8barFS2l8zTVnRUNA--KtexxDNZ1RgSrspwar-edfhPuo6xtxAxb4C8nILpU8nyKenQtRXMKuB2Xwq1H8il-Urpw9I3UOlJ473ubltaooLjBfw5mWICaLJU9lEBoLdaLwJR8pvP_qN9EP8yM2kcQFehfiD1ATHr3JTgUKb23dA79LliBx6fG49f3EuoELyXDxifVnQFNwxwzVFMQMHnMD50kHBS5ZLrJDhlqprLAxBh7LaKb3vcCiZjhHlFQ66fHbP8dxHPwM-aBuXx1Ok-HPOQypwF8IShVzicQ8ViVNQcFQKN6kIugBipJ82OxO4CYDgMgj-warfqDAgBZ4b7ijDJ93yvjk9QbNaXkuoAJqc-AbmycFpOM_38cbLNtVKoG47SuBUUMXrVhYVDReraRYRFisf7swgFdKHulHzZZhPYyi1O496OUj0bg58PIgJ_Pj8mTACHl3wrAiMwi7kZoJvDB27DFMEVHwZ-zY3ugpVC97ydoHfgeOOOmJASaSzCR5o2F5ziEecyYutrU.Bos_fpDPJTNjE2k8_Di4aw'
COMPANY_ID = '9341452910276277'

# QuickBooks API endpoints
BASE_URL = 'https://sandbox-quickbooks.api.intuit.com'
BALANCE_SHEET_ENDPOINT = 'reports/BalanceSheet'

# OAuth2 setup
oauth = OAuth2Session(client_id=CLIENT_ID)