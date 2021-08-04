import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
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


def Place_Order(WebDriver, FileName, Sheet, RowIndex, RowCount, Stime, Stime1):
    # Order Management Screen
    # Draw Type
    time.sleep(Stime1)
    if (RWDE.ReadData(FileName, Sheet, RowIndex, 32) != 'None'):
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//div[3]/div/lightning-combobox//input[@placeholder = "Select an Option"]', 60)
        WebDriver.execute_script('arguments[0].click();', Element)

        # Draw Type Element
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span/span[. = "' + str(RWDE.ReadData(FileName, Sheet, RowIndex, 31)) + '"]', 60)
        Element.click()

    # ICD10 Radio Button
    time.sleep(Stime1)
    Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[2]/label/span[1]', 60)
    WebDriver.execute_script('arguments[0].click();', Element)

    # Other ICD Code
    #time.sleep(Stime1)
    #Element = BEP.WebElement(webdriver, EC.presence_of_element_located, By.XPATH, '//div[4]//input', 60)
    #WebDriver.execute_script('arguments[0].click();', Element)

    ## GHPositiveResult
    #if (str(RWDE.ReadData(FileName, Sheet, RowIndex, 33)) != 'None'):
    #    time.sleep(Stime)
    #    Element = BEP.WebElement(webdriver, EC.presence_of_element_located, By.XPATH, '//div[4]//input[@placeholder = "Select an Option"]', 60)
    #    WebDriver.execute_script('arguments[0].click();', Element)

    #    # GHPositiveResult Element
    #    Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span/span[. = "' + str(RWDE.ReadData(FileName, Sheet, RowIndex, 33)) + '"]', 60)
    #    WebDriver.execute_script('arguments[0].click();', Element)

    # ShareResult CheckBox
    if (str(RWDE.ReadData(FileName, Sheet, RowIndex, 34)) == 'Yes'):
        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//lightning-input/div/span//span[1]', 60)
        Element.click()

        # SecondaryRecipient
        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//div[6]//input', 60)
        if (str(RWDE.ReadData(FileName, Sheet, RowIndex, 35)) != 'None'):
            Element.send_keys(str(RWDE.ReadData(FileName, Sheet, RowIndex, 35)))
        else:
            Element.send_keys('')

        # SecondaryRecipientFax
        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//div[7]//input', 60)
        if (str(RWDE.ReadData(FileName, Sheet, RowIndex, 36)) != 'None'):
            Element.send_keys(str(RWDE.ReadData(FileName, Sheet, RowIndex, 36)))
        else:
            Element.send_keys('')

    # Next Button
    time.sleep(Stime)
    Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//button[. = "Next"]', 60)
    Element.click()

    # Submit Button
    time.sleep(Stime)
    Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//button[. = "Submit"]', 60)
    Element.click()

    if (RowIndex == RowCount):
        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//img', 60)
        Element.click()

        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[. = "Log Out"]', 60)
        Element.click()
    else:
        time.sleep(3)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//button[. = "New Patient"]', 60)
        WebDriver.execute_script('arguments[0].click()', Element)

def Skip_From_Flow(WebDriver, RowIndex, RowCount, Stime):
    time.sleep(Stime)
    Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//header//*[name()="svg"]', 60)
    Element.click()
    if (RowIndex == RowCount):
        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//img', 60)
        Element.click()

        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//span[. = "Log Out"]', 60)
        Element.click()
    else:
        time.sleep(Stime)
        Element = BEP.WebElement(WebDriver, EC.presence_of_element_located, By.XPATH, '//button[. = "New Patient"]', 60)
        WebDriver.execute_script('arguments[0].click()', Element)