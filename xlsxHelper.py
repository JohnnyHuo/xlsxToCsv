#!/usr/bin/env python
import sys
import xlrd
from datetime import datetime
# try:
#     fileName=input('Enter the .xlsx file name with double quote, for example, "input.xlsx": ')
# except(AttributeError,NameError):
#     print('Incorrect input file name')
#     raise

# i = datetime.now()

# print(sys.argv, len(sys.argv))
#
# date=datetime.now().strftime('%Y-%m-%d')
# outputName='CompromisedAccount'+date+'.csv'
#

# For old format, no customer_id contained
# s=set()
# j=0
# for arg in sys.argv:
#     if j > 0:
#         try:
#             print 'Processing ' + arg
#             wb = xlrd.open_workbook(arg)
#         except(IOError):
#             print("File NOT exist, double check your input file name")
#             sys.exit(0)
#         try:
#             sh = wb.sheet_by_name('ATO')
#             # emails = sh.col_values(2)
#             # loginuids=sh.col_values(3)
#         except(NameError):
#             print("ATO sheet not exist")
#             sys.exit(0)
#         # for i in range(1,len(emails)):
#         #     s.add(emails[i])
#         #     s.add(loginuids[i])
#     j+=1
# with open(outputName, 'w') as f:
#     for val in s:
#          f.write(val+',' + date +','+'N'+'\n')

# print(sys.argv, len(sys.argv))

date=datetime.now().strftime('%Y-%m-%d')
outputName='CompromisedAccount'+date+'.csv'
print('\n')
print('Output file name: ' + outputName)

j=0
dict = None
with open(outputName, 'w') as f:
    for arg in sys.argv:
        if j > 0:
            try:
                print 'Processing ' + arg
                wb = xlrd.open_workbook(arg)
            except(IOError):
                print("File NOT exist, double check your input file name")
                sys.exit(0)
            try:
                sh = wb.sheet_by_name('ATO')
                for cur_row in range(sh.nrows):
                    customer_id = sh.row(cur_row)[1].value
                    if customer_id == 'omscid':
                        continue
                    customer_id = str(int(customer_id)) #omscid
                    email = str(sh.row(cur_row)[2].value)
                    loginuid = str(sh.row(cur_row)[3].value)
                    if email != loginuid:
                        f.write(loginuid +',' + date + ',' + 'N' + ',' + customer_id +'\n')
                    f.write(email +',' + date + ',' + 'N' + ',' + customer_id +'\n')
            except(NameError):
                print("ATO sheet not exist")
                sys.exit(0)
        j+=1



# xhuo2@L-SB835ULG8W-M:~/PycharmProjects/CompromisedAccountLoader/xlsxToCsv$ ./xlsxHelper.py "2016-03-11-03-SuccessfullyCheckedAccounts.xlsx" "2016-03-14-01-SuccessfullyCheckedAccounts.xlsx" "2016-03-14-01-SuccessfullyCheckedAccounts_2.xlsx" "2016-03-15-01-SuccessfullyCheckedAccounts.xlsx"
# Traceback (most recent call last):
#   File "./xlsxHelper.py", line 23, in <module>
#     wb = xlrd.open_workbook(arg)
#   File "/Library/Python/2.7/site-packages/xlrd/__init__.py", line 441, in open_workbook
#     ragged_rows=ragged_rows,
#   File "/Library/Python/2.7/site-packages/xlrd/book.py", line 91, in open_workbook_xls
#     biff_version = bk.getbof(XL_WORKBOOK_GLOBALS)
#   File "/Library/Python/2.7/site-packages/xlrd/book.py", line 1230, in getbof
#     bof_error('Expected BOF record; found %r' % self.mem[savpos:savpos+8])
#   File "/Library/Python/2.7/site-packages/xlrd/book.py", line 1224, in bof_error
#     raise XLRDError('Unsupported format, or corrupt file: ' + msg)
# xlrd.biffh.XLRDError: Unsupported format, or corrupt file: Expected BOF record; found '#!/usr/b'
