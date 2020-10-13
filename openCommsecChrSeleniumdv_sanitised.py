from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from lxml import html
import requests
import sys, getopt
import os


try:
    ttldt= sys.argv[1]
    lgnid = ttldt[len(ttldt)-16:len(ttldt)-8]
    pword = ttldt[:len(ttldt)-16]
    StDt =  ttldt[len(ttldt)-8:]
except:
    lgnid = input("Enter Username: ")
    StDt = input("Enter Start Date - ddmmyyyy:")

url ='https://www.commsec.com.au/'
dwnldpth = 'C:\\Users\Business\\Downloads'

chromeOptions = webdriver.ChromeOptions()
prefs = ({"download.default_directory" : dwnldpth, 
         "download.prompt_for_download": False})
chromeOptions.add_experimental_option("prefs",prefs)

driver = webdriver.Chrome()
driver.get(url)

time.sleep(5)
driver.find_element_by_id("txt-clientId").clear()
driver.find_element_by_id("txt-clientId").send_keys(lgnid)


try:
    driver.find_element_by_id("password-field").send_keys(pword) #Commsec
except:
    time.sleep(10)
    
driver.find_element_by_id("btn-login").click() #Commsec
time.sleep(5)

#Holdings Download
#------------------

urlshrt='https://www2.commsec.com.au/portfolio/holdings********'

dwnldxpth='/html/body/portfolio-root/div/portfolio-holdings\
                    /root-page-container/div/div[5]/equity-holdings-table\
                    /div/div/div/div[1]/div[2]/download-button/with-tooltip\
                    /div/button/with-spinner/div'
sep=""

#Kearnvest Holdings Download
urlhld = '*********8p3c39FxsJ**'

def dwnldact():
    url = sep.join([urlshrt, urlhld])
    driver.get(url)
    time.sleep(5)
    try:
        driver.find_element_by_xpath(dwnldxpth).click()
        time.sleep(5)
    except:
        print ('Element not found')

dwnldact()

#Domaintell Holdings Download
urlhld = '*****KSueoAVk85SV0xswu*****'

dwnldact()


#Acquisitions Stage
#------------------

url ='https://www2.commsec.com.au/Private/MyPortfolio/Confirmations/Confirmations.aspx'
prtflacc = 'ctl00_BodyPlaceHolder_ConfirmationsView1_accountList_ddlAccounts_field'
ordrprd = 'ctl00_BodyPlaceHolder_ConfirmationsView1_calendarFrom_field'
srch = 'ctl00_BodyPlaceHolder_ConfirmationsView1_btnSearch_implementation_field'
csvdwnld = 'ctl00_BodyPlaceHolder_ConfirmationsView1_gdvwConfirmationDetails\
_Underlying_TopPagerRow_btnDownloadCSV_implementation'
driver.get(url)
time.sleep(5)

#Kearnvest Share Acquisition confirmations ConfirmationsDetails.csv
driver.find_element_by_id(prtflacc).send_keys ("2")
driver.find_element_by_id(ordrprd).clear()
driver.find_element_by_id(ordrprd).send_keys (StDt)
driver.find_element_by_id(srch).click()
time.sleep(5)
#Line 62

try:
    driver.find_element_by_id(csvdwnld).click()
    time.sleep(5)
except:
    print ('No confirmations apply')

try:
    dwnldfletmp=os.path.join(dwnldpth, "ConfirmationDetails.csv")
    dwnldfle =os.path.join(dwnldpth, "ConfirmationDetails2*.csv")
    os.rename(dwnldfletmp, dwnldfle) 
except:
    print ('File does not exists')

#Domaintell Share Acquisition confirmations ConfirmationsDetails (1).csv
driver.find_element_by_id(prtflacc).send_keys ("3")
driver.find_element_by_id(ordrprd).clear()
driver.find_element_by_id(ordrprd).send_keys (StDt)
driver.find_element_by_id(srch).click()
time.sleep(5)

try:
    driver.find_element_by_id(csvdwnld).click()
    time.sleep(5)
except:
    print ('No confirmations apply')

try:
    dwnldfletmp=os.path.join(dwnldpth, "ConfirmationDetails.csv")
    dwnldfle =os.path.join(dwnldpth, "ConfirmationDetails3*.csv")
    print (dwnldfle)
    os.rename(dwnldfletmp, dwnldfle) 
except:
    print ('File does not exists')

url='https://www2.commsec.com.au/'
driver.get(url)

logoutxpth= '/html/body'
'''
xpath when not in selenium / auto mode:
logoutxpth= "/html/body/form/div[3]/commsec-header//header/div[1]/nav[1]/div/div[2]/a[2]/span"   
old element id seemingly no longer applicable:
driver.find_element_by_id("ctl00_ucClientBox_lbLogout").click()
'''

driver.find_element_by_xpath(logoutxpth).click()
time.sleep(3)
            
driver.quit()

import callMacroforPortfolioupdate # import update_Plan
if __name__ == '__main__':
    callMacroforPortfolioupdate.main()
