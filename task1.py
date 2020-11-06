# -*- coding: utf-8 -*-
"""
Created on Fri Nov  6 08:23:56 2020

@author: NITHIN KUMAR
"""
#create, write,open CSV file
import csv      
#Read Excel data
import pandas
#get all sheet name
import xlrd

def get_Header(file,sh):
    header=[]
    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_name(sh)
    sheet.cell_value(0, 0)
    for i in range(sheet.ncols):
        header.append(sheet.cell_value(0, i))    
    return header

def get_AllSheet(file):
    xls = xlrd.open_workbook(file, on_demand=True)
    l=list(xls.sheet_names())
    return l

def get_sheet(file,sheetList):
    sheets=[]
    for i in sheetList:
        if file in str(i):
            sheets.append(i)
    return sheets
    
    
def Task_One(file,outputfile):
    result=[]   #all rows

    CK_FILES=['../data.xls','../data_1.xls']
    for CK_FILE in CK_FILES:
        #get sheet
        sheetList=get_AllSheet(CK_FILE)
        #get sheet on condition
        sheets=get_sheet(file,sheetList)
                
        #create header
        if len(sheets) > 0:
            header=get_Header(CK_FILE,sheets[0])
        else:
            break            
        #retrive data
        for i in sheets:
            excel_data_df = pandas.read_excel(CK_FILE, sheet_name=i, usecols=header)
            s= excel_data_df.to_dict(orient='record')   #retriving excel data and store data in dictonary format
            result.append(s)
            
    #add data
    file = open(outputfile, 'w', newline ='') 
    with file: 
        # identifying header   
        writer = csv.DictWriter(file, fieldnames = header) 
        writer.writeheader()
        for j in result:
            writer.writerows(j)   #adding all rows
    
        


Task_One("Detail_67_","detail.csv")
Task_One("DetailVol_67_","detailVol.csv")
Task_One("DetailTemp_67_","detailTemp.csv") 