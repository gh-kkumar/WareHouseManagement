import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
import OrderManagement as OM
import os
#import win32com.client as comclt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from pathlib import Path
from selenium.webdriver.common import keys


FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
Sheet = 'Portal Urls'
Url = str(RWDE.ReadData(FilePath, Sheet, 2, 3))

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('disable-notifications')
driver = webdriver.Chrome(executable_path = str(Path().resolve()) + '\Browser\chromedriver_win32\chromedriver', options=chrome_options)
driver.maximize_window()
driver.get(Url)
#print(driver.title)

FilePath = str(Path().resolve()) + '\Excel Files\WareHouseManagement.xlsx'
Seconds = 1

#1. This is for SFDC Login

Sheet = 'Login Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)

for RowIndex in range(2, RowCount + 1):

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#username', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 2))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#password', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 3))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#rememberUn', 60)
    Element.click()

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#Login', 60)
    Element.click()

    Sheet1 = 'Ware House page Data'
    RowCount1 = RWDE.RowCount(FilePath, Sheet1)
    for Rowindex1 in range(2, RowCount1 + 1):

        # WayBillScan Div
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Way Bill Scan"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        # WayBill BarCode TextBox
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Enter Waybill barcode…"]', 60)
        if(str(RWDE.ReadData(FilePath, Sheet1, Rowindex1, 2)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet1, Rowindex1, 2))
        else:
            Element.send_keys('')

        # Package Received
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Package Received"]', 60)
        driver.execute_script('arguments[0].click();', Element)

        # Main Message
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[6]/div/div/div/div//span', 60)
        print(Element.text)
        if(str(RWDE.ReadData(FilePath, Sheet1, Rowindex1, 3)) == Element.text):
            RWDE.WriteData(FilePath, Sheet1, Rowindex1, 4, Element.text)
            RWDE.WriteData(FilePath, Sheet1, Rowindex1, 5, 'Passed')
        else:
            RWDE.WriteData(FilePath, Sheet1, Rowindex1, 4, Element.text)
            RWDE.WriteData(FilePath, Sheet1, Rowindex1, 5, 'Failed')

        time.sleep(3)

        # Account Icon
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button/div/span//span', 60)
        driver.execute_script('arguments[0].click();', Element)

        # Log Out Link
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[. = "Log Out"]', 60)
        driver.execute_script('arguments[0].click();', Element)






