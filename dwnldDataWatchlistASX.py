from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import re
import lxml
from lxml import etree
import requests
import time
import sys, getopt
import win32com.client
import csv
import codecs

tgtfile = open('C:\\Users\\Business\\Kearnvest\\shares\
\\19\\Watchlist.csv', 'r')
try:
    os.remove('C:\\Users\\Business\\Kearnvest\\shares\
\\19\\WtchlstRrdr.csv')
except:
    pass
rsltfl = open('C:\\Users\\Business\\Kearnvest\\shares\
\\19\\WtchlstRrdr.csv', 'w+')
csvwriter = csv.writer(rsltfl)
 
url ='https://www.asx.com.au/'
dwnldpth = 'C:\\Users\Business\\Kearnvest\\shares\\19'
Library_head = []
Library_head.append("Company")
Library_head.append("Price")
Library_head.append("Last Dividend"
Library_head.append("Yield)
Library_head.append("52 Week Low")
Library_head.append("52 Week High")
Library_head.append("P/E")
csvwriter.writerow(Library_head)

nr=0
for l, i in enumerate(tgtfile):
    datanw=[]
    datanw1=[]
    
    if l<1:
        l=l+1
    else:
        data = i.split(",")
        if data[0] == '': # limit list temporarily
            break
        #print(data[1])
        asx = data[1]
        url = 'https://www.asx.com.au/asx/share-price-research/company/' + str(asx)
        print(url)
        driver = webdriver.Firefox()
        driver.get(url)
        shrprc_source = driver.page_source
        soup = BeautifulSoup(shrprc_source,"lxml")
        count=0
        for tag in soup.find_all('span', class_='ng-binding'):
            if count < 41:
##                print (tag.text)
                datatmp = tag.text
                datanw.append(datatmp)
                count = count +1
##        print (datanw)
        datanw1.append(datanw[0])
        datanw1.append(datanw[1])
        datanw1.append(datanw[7])
        datanw1.append(datanw[8])
        datanw1.append(datanw[9])
        datanw1.append(datanw[10])
        datanw1.append(datanw[40])
        print (datanw1)        
##        count = 0
##        for tag in soup.find_all('span', class_='low'):
##            if count < 2:
##                datatmp = tag.text
##                datatmp = datatmp.split()
##                datatmp = datatmp[1]
##                datanw.append(datatmp)               
##                count = count +1
##        count = 0
##        for tag in soup.find_all('span', class_='high'):
##            if count < 2:
##                datatmp = tag.text
##                datatmp = datatmp.split()
##                datatmp = datatmp[1]
##                datanw.append(datatmp)               
##                count = count +1
        csvwriter.writerow(datanw1)
        driver.close()
        
rsltfl.close()
