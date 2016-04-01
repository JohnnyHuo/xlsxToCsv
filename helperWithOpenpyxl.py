#!/usr/bin/python


# $ pip install openpyxl

import sys
import xlrd

from datetime import datetime
# try:
#     fileName=input('Enter the .xlsx file name with double quote, for example, "input.xlsx": ')
# except(AttributeError,NameError):
#     print('Incorrect input file name')
#     raise

# i = datetime.now()
date=datetime.now().strftime('%Y-%m-%d')
outputName='CompromisedAccount'+date+'.csv'

s=set()

for arg in sys.argv:
    if arg == 'xlsHelper.py':
        continue
    try:
        wb = xlrd.open_workbook(arg)
    except(IOError):
        print("File NOT exist, double check your input file name")
        sys.exit(0)
    try:
        sh = wb.sheet_by_name('ATO')
        emails = sh.col_values(2)
        loginuids=sh.col_values(3)
    except(NameError):
        print("ATO sheet not exist")
        sys.exit(0)
    for i in range(1,len(emails)):
        s.add(emails[i])
        s.add(loginuids[i])




with open(outputName, 'w') as f:
    for val in s:
         f.write(val+',' + date +','+'N'+'\n')

