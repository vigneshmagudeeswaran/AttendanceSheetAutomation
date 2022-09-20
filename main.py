import pandas as pd
import os
import openpyxl
import csv
from csv import DictWriter

col_lable =['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AB','AC','AD','AE','AF','AG'] #used to get the columns
len_col_lable = len(col_lable)
cwd = os.getcwd()
file = 'aug.xlsx'
data = pd.ExcelFile(file)
sheetlist = data.sheet_names  # this returns the all the sheets in the excel file ['Sheet1']
df = data.parse('Sheet1')
ps = openpyxl.load_workbook('aug.xlsx')
sheet = ps['Sheet1']
n = 13
def requireddata(row):
    name =[]
    name.append(sheet[col_lable[2] + str(row)].value)
    day =[]
    working_hrs =[]
    date =[]
    status =[]
    for i in col_lable[2:]:
        dayA = sheet[i + str(row + 1)].value
        if dayA != None:
            day.append(dayA)
        dateA = sheet[i + str(row + 2)].value
        statusA = sheet[i + str(row + 10)].value
        if dateA != None:
            dateA = sheet[i + str(row + 2)].value
            date.append(dateA)
            #statusA = sheet[i + str(row + 10)].value
            #if statusA != None:
            working_hrsA = sheet[i + str(row + 5)].value
            working_hrs.append(working_hrsA)
            if statusA not in ['P','LT']:
                status.append(statusA)
            else:
                status.append(working_hrsA)
        #df = pd.DataFrame.from_dict(datadict,orient='index',columns=['name', 'working Hrs', 'date', 'status'])
        #print(df)
        #df.to_csv('dict2.csv',mode='a')
        #print(df.to_excel('dict1.xlsx'))
    #datadict = {name:mydict}
    print('length of working hrs: ',len(working_hrs))
    print('length of status: ',len(status))
    print('length of day: ', len(day))
    print('length of date: ', len(date))
    daydate =[day,date]
    # working_hrsStatus =[status]
    coldata= []
    for i in range(len(day)):
        coldata.append(None)
    # df = pd.DataFrame(data=[name],index=['name'],columns =[None])
    # df.to_csv('dict6.csv', mode='a')
    # df = pd.DataFrame(data=[date],index=['date'],columns =coldata)
    # df.to_csv('dict6.csv', mode='a')
    if row == 3:
        df = pd.DataFrame(data=[status],index=[name],columns =daydate)
    else:
        df = pd.DataFrame(data=[status], index=[name], columns=coldata)
    print(df.columns)
    df.to_csv('dict6.csv', mode='a')
    # # df = pd.DataFrame(data=daydate)
    # # df.to_csv('dict6.csv', mode='a')
    # df = pd.DataFrame(data =[working_hrsStatus,date],index=['status'])
    # df.to_csv('dict6.csv', mode='a')
def forallemp(total_employees):# This function used to find the each employees first row
    for j in range(0,total_employees):
        row = 3+j*n
        requireddata(row)
totalemployees = int(input('Please enter the total Employees: '))
forallemp(totalemployees)

