import pandas as pd
import openpyxl

col_lable =['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AB','AC','AD','AE','AF','AG'] #used to get the columns
len_col_lable = len(col_lable)
inputFileName =input('Enter Input file Name: ') + '.xlsx'
ps = openpyxl.load_workbook(inputFileName)
sheet = ps['Sheet1']
n = 13 # Frequency between one employee name to other Employee name
outputFileName =input('Enter output file Name: ') + '.csv'
def requireddata(row,outputFileName):
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
            working_hrsA = sheet[i + str(row + 5)].value
            working_hrs.append(working_hrsA)
            if statusA not in ['P','LT']:
                status.append(statusA)
            else:
                status.append(working_hrsA)
    daydate =[day,date]
    coldata= []
    for i in range(len(day)):
        coldata.append(None)
    if row == 3:
        df = pd.DataFrame(data=[status],index=[name],columns =daydate)
        df.to_csv(outputFileName)
    else:
        df = pd.DataFrame(data=[status], index=[name], columns=coldata)
        df.to_csv(outputFileName, mode='a')
def forallemp(total_employees,outputFileName):# This function used to find the each employees first row and based on that row start excel modification
    for j in range(0,total_employees):
        row = 3+j*n
        requireddata(row,outputFileName)
totalemployees = int(input('Please enter the total Employees: '))
forallemp(totalemployees,outputFileName)

