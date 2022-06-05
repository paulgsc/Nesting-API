#!/usr/bin/env python
# coding: utf-8

# In[7]:


"""

# List of api functions

# write data to sheets using
# writeDataToSheet(worksheet_name,gsheetId,file_name,sheet_name)
# writeDataToSheetDf(worksheet_name,gsheetId,df)

SheetsNew()
 returns a dict of the following
 dict = {'Url':sheets_file2['spreadsheetUrl'],
 'gsheetId':sheets_file2['spreadsheetId'],'sheet_names':sheets_file2['sheets']}


requires work sheet name, file name path, and name of sheet 
define values below

worksheet_name : this is the name of the sheet you are writing to

file_name: this is the path file name of the file you are reading to then import into google sheets

df : use data frame 

sheet_name : this is the name of the sheet you are reading from file_name

gsheetId : this is the sheet api id of the google sheet you are writing to

# worksheet_name = 'Sheet3'
sheet_name = 0
# gsheetId = '1DAsPko2wpbdUiXEcULRg0z8N5K7X4blahhr-Ci45Vk0'

# read google sheets
# main(gsheetID,SAMPLE_RANGE_NAME)

gsheetID : this is the sheet api id of the google sheet you are reading from
SAMPLE_RANGE_NAME : this is the range you are reading

# SAMPLE_RANGE_NAME = 'A1:B4'

# clear sheets
# clearSheetsRange(gsheetId,worksheet_name,[cell_range])
# clearSheets(gsheetId,worksheet_name)

cell_range : define range to clear
"""

# read latest file in downloads
# getLatestFileName (path: Path, pattern: str = "*")
import pandas as pd
import numpy as np
from pathlib import Path
from credsTest.Google import Create_Service
from credsTest.GoogleSheetsAPI import *
from credsTest.readSheets import main
import re

def createSewingPivot(pattern):

    """
    Createing nest file
    """

    # read export
    # first define path
    location = 'Downloads'
    file_path = Path.cwd() / location

    # next assign path and define pattern of file name
    downloaded_file = Path.cwd() / 'Downloads'
    pattern = '*'+pattern+'*'

    # use funciton to get latest file
    myFile = getLatestFileName(downloaded_file,pattern)

    # next read the contents of the file must be either csv or excel file
    if 'xls' not in str(myFile):
        nest_file = pd.read_csv(myFile,index_col=False)
    else:
        nest_file = pd.read_excel(myFile,sheet_name=0,index_col=False)




    """
    modify nest file

    """
    def createCoversData(nest_file):
        # Mirror excel countif to generate sliced so line qty to produce
        tempList = nest_file['Sale Order Line/Product/Display Name'].fillna('blanks').tolist()
        bondi_chicory = []
        covers_rename = []
        for x in range(0,len(tempList)):
            if (re.search('(?i)(?:mst1 bondi|mst1 chicory)',tempList[x])!=None):
                bondi_chicory.append((re.search('(?i)(?:mst1 bondi|mst1 chicory)',tempList[x])).group(0))
            else:
                bondi_chicory.append('other')
        nest_file['BONDI|CHICORY'] = bondi_chicory

        covers_rename = nest_file['Product/Display Name'].fillna('blanks').tolist()
        for x in range(0,len(covers_rename)):
            covers_rename[x] = re.sub(r'(?i)(?:\[.+\]\s|\sv2$)','',covers_rename[x])

        nest_file['Covers'] = covers_rename
        rename = nest_file['Sale Order Line/Product/Display Name'].fillna('blanks').tolist()
        rename2 = nest_file['Product/Display Name'].fillna('blanks').tolist()
        for x in range(0,len(rename)):
            rename[x] = re.sub(r'(?i)(.*\]\s)','',rename[x])
            rename2[x] = re.sub(r'(?i)(.*\]\s)','',rename2[x])
        to_assign = {'Product Name':rename,'Component Name':rename2}
        nest_file = nest_file.assign(**to_assign)

        # use numpy to create countif of so quantity
        so_count = []
        list1 = nest_file['Sale Order Line/Qty to Produce'].tolist()
        tempList = nest_file['Sale Order Line/ID'].tolist()
        for x in tempList:
            so_count.append(tempList.count(x))
        ar1 = np.array(list1)
        ar2 = np.array(so_count)
        ar3 = ar1 / ar2
        nest_file['SO Line Product Qty'] = ar3
        ar4 =  nest_file['Quantity To Be Produced'].tolist()
        nest_file['Qty per SO qty'] = ar4 / ar1

        """
        sewing wo data detailed table

        """
        # creating column for fabric number

        # read the fabric color and number for mapping
        fabric_color = main('1tOkainSd6Q_cdwyPMpAyvs3ze2fCU20yhM1yp4VEHRc','Sheet1!A2:B').dropna()
        color_att = (nest_file['Sale Order Line/Product Attributes'].astype(str)+', ').str.extract(r'Color:\s(.+?),\s')
        color_att.columns = ['color']
        kwargs = {'Fabric Color': color_att['color']}
        nest_file = nest_file.assign(**kwargs) 
        nest_file = nest_file.merge(fabric_color,left_on='Fabric Color',right_on='CABA NAME',sort=False).drop('CABA NAME',axis=1)


        # s = []
        # for x in range(0,len(colorList)):
        #     s.append(re.sub(',.*','',test[x].split('Color: ')[1]))
        # d = {'Color': s}
        # color_att = pd.DataFrame(data=d)

        # add target week column
        df_week = nest_file['Sale Order Line/Commitment Date']
        df_week = pd.to_datetime(df_week, infer_datetime_format=True)  
        df_week = df_week.dt.tz_localize('UTC')
        a=df_week.dt.strftime('%W').fillna(0)
        a = np.array(a)
        a = a.astype(int)+1
        a = pd.DataFrame({'target week': a.tolist()})
        nest_file['So Target Week'] = a
          


        # rename covers to give chicory accent pillow size
        accentAttr = nest_file['Sale Order Line/Product Attributes'].fillna('blanks').tolist()
        for x in range(0,len(accentAttr)):
            accentAttr[x] =  re.sub(r'(?i)(?:.*\spillow\soptions\W\s|\sw/insert)','',accentAttr[x])
        accentAttr = pd.DataFrame({'attr':accentAttr}) 
        df4 = pd.DataFrame({'prod':nest_file['Covers']})
        df = nest_file['Covers'].str.cat(accentAttr['attr'],sep=' ')
        df1  = pd.DataFrame({'attrProduct':df})
        df5 = df4.where(~(df4['prod'].str.match('(.*for chicory.*)',case=False)),df1['attrProduct'],axis=0)
        nest_file['Covers'] =  df5
        return nest_file


    # In[ ]:


    """
    create new sheet and get sheet properties

    """
    
    nest_file = createCoversData(nest_file)
    dict = SheetsNew()
    gsheetId = dict['gsheetId']
    sheetProperties = dict['sheet_names']
    Url = dict['Url']
    sheet_names = [sheetProperties[0]['properties']['title'],sheetProperties[1]['properties']['title']]
    sheetIds = [sheetProperties[0]['properties']['sheetId'],sheetProperties[1]['properties']['sheetId']]



    # In[ ]:


    """
    Insert nest file data to new sheet
    """
    worksheet_name = sheet_names[0]
    # run this to write data to google sheets
    clearSheets(gsheetId,worksheet_name)
    writeDataToSheetDf(worksheet_name,gsheetId,nest_file)


    """
    create pivot table using data inserted to sheets

    """

    # PivotTable JSON Template
    request_body = {
        'requests': [
            {
                'updateCells': {
                    'rows': {
                        'values': [
                            {
                                'pivotTable': {
                                    # Data Source
                                    'source': {
                                        'sheetId': sheetIds[0],
                                        'startRowIndex': 0,
                                        'startColumnIndex': 0,
                                        'endRowIndex': len(nest_file),
                                        'endColumnIndex': len(nest_file.columns) # base index is 1
                                    },

                                    # Rows Field(s)
                                    'rows': [
                                        # row field #1
                                        {
                                            'sourceColumnOffset': nest_file.columns.get_loc('Odoo Fabric Number'),
                                            'showTotals': True, # display subtotals
                                            'sortOrder': 'ASCENDING',
                                            'repeatHeadings': False,
                                            'label': 'Fabric #',
                                         },

                                        {
                                            'sourceColumnOffset': nest_file.columns.get_loc('Lot/Serial Number/Lot/Serial Number'),
                                            'showTotals': True, # display subtotals
                                            'sortOrder': 'ASCENDING',
                                            'repeatHeadings': False,
    #                                         'label': 'Product List',
                                        },

                                        {
                                            'sourceColumnOffset': nest_file.columns.get_loc('Sale Order Line/Product/Display Name'),
                                            'showTotals': False, # display subtotals
                                            'sortOrder': 'ASCENDING',
                                            'repeatHeadings': False,
    #                                         'label': 'Product List',
                                        },

                                        {
                                            'sourceColumnOffset': nest_file.columns.get_loc('Sale Order Line/Qty to Produce'),
                                            'showTotals': False, # display subtotals
                                            'sortOrder': 'ASCENDING',
                                            'repeatHeadings': False,
    #                                         'label': 'Product List',
                                        }

    #                                     {
    #                                         'sourceColumnOffset': 16,
    #                                         'showTotals': False, # display subtotals
    #                                         'sortOrder': 'ASCENDING',
    #                                         'repeatHeadings': False,
    # #                                         'label': 'Product List',
    #                                     }                   

                                    ],

    #                                 Columns Field(s)
                                    'columns': [
                                        # column field #1
                                        {
                                            'sourceColumnOffset': nest_file.columns.get_loc('Covers'),
                                            'sortOrder': 'ASCENDING', 
                                            'showTotals': True
                                        }
                                    ],

                                    'criteria': {
                                        nest_file.columns.get_loc('Operation/Display Name'): {
                                            'visibleValues': [
                                                'Sewing QC/Prep'
                                            ]
                                        },

                                        nest_file.columns.get_loc('BONDI|CHICORY'): {
                                            'visibleValues': [
                                                'MST1 Bondi', 'MST1 Chicory'
                                            ]
                                        },
                                           nest_file.columns.get_loc('Assigned to/Display Name'): {
                                            'visibleValues': [
                                                'FALSE'
                                            ]
                                        },

    #                                     11: {
    #                                         'condition': {
    #                                             'type': 'NUMBER_BETWEEN',
    #                                             'values': [
    #                                                 {
    #                                                     'userEnteredValue': '10000'
    #                                                 },
    #                                                 {
    #                                                     'userEnteredValue': '100000'
    #                                                 }
    #                                             ]
    #                                         },
    #                                         'visibleByDefault': True
    #                                     }
                                    },

                                    # Values Field(s)
                                    'values': [
                                        # value field #1
                                        {
                                            'sourceColumnOffset': nest_file.columns.get_loc('Quantity To Be Produced'),
                                            'summarizeFunction': 'SUM',
                                            'name': 'Covers Qty:'
                                        }
                                    ],

                                    'valueLayout': 'HORIZONTAL'
                                }
                            }
                        ]
                    },

                    'start': {
                        'sheetId': sheetIds[1],
                        'rowIndex': 0, # 4th row
                        'columnIndex': 0 # 3rd column
                    },
                    'fields': 'pivotTable'
                }
            }
        ]
    }

    response = service.spreadsheets().batchUpdate(
        spreadsheetId=gsheetId,
        body=request_body
    ).execute()
    return dict;


