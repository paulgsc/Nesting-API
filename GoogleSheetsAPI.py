#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import os
from pathlib import Path
from credsTest.Google import Create_Service
import pandas as pd
import numpy as np
import re


# In[2]:


df = pd.read_excel('myexport.xlsx',sheet_name='export',index_col=False)
df.replace(np.nan,'',inplace=True)

# Convert all date columns to string type
for col in  df.select_dtypes(include=['datetime64']).columns.tolist():
    df[col] = df[col].astype(str)


# In[3]:


FOLDER_PATH = Path.cwd()
FILE_NAME = 'client_secret_file.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']



# In[4]:



CLIENT_SECRET_FILE = FOLDER_PATH / FILE_NAME
# CLIENT_SECRET_FILE = os.path.join(FOLDER_PATH, 'Client_Secret.json')

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

def SheetsNew():
    """
    To specify Google Sheets file basic settings and as well as configure default worksheets
    """
    sheet_body = {
        'properties': {
            'title': 'ApiSheetsNew',
            'locale': 'en_US', # optional
            'timeZone': 'America/Los_Angeles'
            }
        ,
        'sheets': [
            {
                'properties': {
                    'title': 'Data'
                }
            },
            {
                'properties': {
                    'title': 'Pivot'
                }
            }
        ]
    }

    sheets_file2 = service.spreadsheets().create(body=sheet_body).execute()
    return {'Url':sheets_file2['spreadsheetUrl'],'gsheetId':sheets_file2['spreadsheetId'],'sheet_names':sheets_file2['sheets']}



# In[11]:

def writeDataToSheetDf(worksheet_name,gsheetId,df):
    cell_range_insert = 'A1'
    
    # Replace Null values
    df.replace(np.nan,'',inplace=True)

    # Convert all date columns to string type
    for col in  df.select_dtypes(include=['datetime64']).columns.tolist():
        df[col] = df[col].astype(str)

    response_date = service.spreadsheets().values().append(
        spreadsheetId=gsheetId,
        valueInputOption='RAW',
        range=worksheet_name,
        body=dict(
            majorDimension = 'ROWS',
            values = df.T.reset_index().T.values.tolist()
        )
    ).execute()



def writeDataToSheet(worksheet_name,gsheetId,file_name,sheet_name):
    cell_range_insert = 'A1'
    # read file in path
    file_string = str(file_name)
    if 'xls' not in file_string:
    	df = pd.read_csv(file_name)
    else:
    	df = pd.read_excel(file_name,sheet_name,index_col=False)

    # Replace Null values
    df.replace(np.nan,'',inplace=True)

    # Convert all date columns to string type
    for col in  df.select_dtypes(include=['datetime64']).columns.tolist():
        df[col] = df[col].astype(str)

    response_date = service.spreadsheets().values().append(
        spreadsheetId=gsheetId,
        valueInputOption='RAW',
        range=worksheet_name,
        body=dict(
            majorDimension = 'ROWS',
            values = df.T.reset_index().T.values.tolist()
        )
    ).execute()



# In[8]:


# method to create new sheets  tabs

def add_sheets(gsheetId, sheet_name):
    spreadsheets = service.spreadsheets()
    try:
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                        'tabColor': {
                            'red': 0.44,
                            'green': 0.99,
                            'blue': 0.50
                        }
                    }
                }
            }]
        }

        response = spreadsheets.batchUpdate(
            spreadsheetId=gsheetId,
            body=request_body
        ).execute()

        return response
    except Exception as e:
        print(e)

def clearSheets(gsheetId,worksheet_name):
    service.spreadsheets().values().clear(
        spreadsheetId = gsheetId,
        range=worksheet_name
    ).execute()

def clearSheetsRange(gsheetId,worksheet_name,cell_range):    
    service.spreadsheets().values().clear(
        spreadsheetId = gsheetId,
        range=worksheet_name + '/s' + cell_range
    ).execute()

def getLatestFileName (path: Path, pattern: str = "*"):
    files = path.rglob(pattern)
    return max(files,key=lambda x: x.stat().st_ctime)



