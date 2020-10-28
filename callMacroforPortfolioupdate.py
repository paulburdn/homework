import os
import win32api
import win32com.client
from datetime import date, timedelta
import time

flpth='C:\\Users\\Public\\Documents\\Reference\\'
sep=''
flnme = sep.join([flpth, 'inptdtaPth.txt'])
filehandle = open(flnme, 'r')
drctionsPth=[]
while True:
    line = filehandle.readline()
    line= line.rstrip('\n')
    if not line:
        break
    drctionsPth.append(line)

flnme = sep.join([flpth, 'drctionsExcFle.txt'])
filehandle = open(flnme, 'r')
drctionsExcFle=[]
while True:
    line = filehandle.readline()
    line= line.rstrip('\n')
    if not line:
        break
    drctionsExcFle.append(line)        

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

filenmeset = []

'''
Create list of portfolio excel files, including full path, ready to run 
macros for uploading data from commsec
'''
dyy = Manage_Dates()
for i in range(0, len(drctionsPth)):
    filepth = sep.join([drctionsPth[i], dyy, '\\'])
    filenme = sep.join([drctionsExcFle[i], ' ', dyy, '.xlsm'])
    filepthnme = sep.join([filepth, filenme])
    filenmeset.append(filepthnme) 

def main():
    for i in range(0, len(drctionsPth)):
        if os.path.exists(filenmeset[i]):
            xl=win32com.client.Dispatch('Excel.Application')
            try:
                xl.Workbooks.Open(Filename=filenmeset[i])
                xl.Visible = True
            except:
                print('File already open')
            xl.Application.Run("Module1.Load_********gs")
            xl.Application.Run("Module2.Update_*******ion")
            xl.ActiveWorkbook.Save()
        time.sleep(5)

if __name__ == '__main__':
    main()
