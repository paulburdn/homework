import os
import win32api
import win32com.client
from datetime import date, timedelta
import time

drctions = ['C:\\Users\\*******\\*****\*****\\', 'Super****n ',\
            'C:\\Users\\********\\*****\\****\\', 'Dom****ts ']

def Manage_Dates():
    today = date.today()
    d1 = str(today)
    dmm=d1[5:7]
    dmmint = int(dmm)
    dyyint = int (d1[:2])
    if dmmint > 6:
        dyyint = dyyint +1
    dyy = str(dyyint)
    return dyy

n=0
filenmeset = []

'''
Create list of portfolio excel files, including full path, ready to run 
macros for uploading data from commsec
'''

for i in range(2):
    dyy = Manage_Dates()
    sep =''
    if n == 0:
        filepth = sep.join([drctions[n], dyy, '\\'])
        filenme = sep.join([drctions[n+1], dyy, '.xlsm'])
        filepthnme = sep.join([filepth, filenme])
        filenmeset.append(filepthnme)
    else:
        filepth = sep.join([drctions[n+1], dyy, '\\'])
        filenme = sep.join([drctions[n+2], dyy, '.xlsm'])
        filepthnme = sep.join([filepth, filenme])
        filenmeset.append(filepthnme)
    n+=1


def main():
    n=0
    for i in range(2):
        if os.path.exists(filenmeset[n]):
            xl=win32com.client.Dispatch('Excel.Application')
            try:
                xl.Workbooks.Open(Filename=filenmeset[n])
                xl.Visible = True
            except:
                print('File already open')
            xl.Application.Run("Module1.Load**********ings")
            xl.Application.Run("Module2.Update******ion")
            xl.ActiveWorkbook.Save()
        print (filenmeset[i])
        time.sleep(5)
        n+=1

if __name__ == '__main__':
    main()
