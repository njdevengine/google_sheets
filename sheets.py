#https://towardsdatascience.com/how-to-integrate-google-sheets-and-jupyter-notebooks-c469309aacea
#get email to share sheet with from project-name@lead-automation-1234567.iam.gserviceaccount.com

import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pandas.io.json import json_normalize

## Connect to our service account
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name("project-name-123456-1a1abcf44dd3.json", scope)
gc = gspread.authorize(credentials)

##Get candidate data sheet from Google Drive
spreadsheet_key = 'get this key from url of google sheet'
book = gc.open_by_key(spreadsheet_key)
worksheet = book.worksheet("Broker Leads")
table = worksheet.get_all_values()

##Convert table data into a dataframe
df = pd.DataFrame(table[1:], columns=table[0])

# ## Save the data back to a new sheet in the dataframe
# from df2gspread import df2gspread as d2g

# wks_name = 'Jupyter Manipulated Data'

# d2g.upload(scores_df, spreadsheet_key, wks_name, credentials=credentials, row_names=True)
