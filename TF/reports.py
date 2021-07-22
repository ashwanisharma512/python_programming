import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from datetime import datetime
from openpyxl import Workbook,load_workbook
import glob
import re
import pickle

url_home = 'http://10.125.64.86/sevakendra_BRPL/'

url_report = 'http://10.125.64.86/sevakendra_BRPL/Report/frmTFProcessRPT.aspx'

# div = 

op = Options()
#op.add_experimental_option("excludeSwitches", ["enable-automation"])
#op.add_experimental_option('useAutomationExtension', False)
# op.add_argument('--headless')

driver = webdriver.Chrome(executable_path="chromedriver.exe",options=op)
driver.maximize_window()
driver.get(url_home)

driver.find_element_by_id('MainContent_UserName').send_keys('41018328')
driver.find_element_by_id('MainContent_Password').send_keys('12345678')
driver.find_element_by_id('MainContent_LoginButton').click()

driver.get(url_AMPS)

from_date = '21-Jan-2021'

Select(driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_ddlRequestType"]')).select_by_value('U01')
for radio in div_list:
    driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_chkDivisionList_'+str(radio)+'"]').click()
# driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_chkDivisionList_1"]').click()
driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_rbd_CaseType_1"]').click()
time.sleep(5)
driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_chkView"]').click()
jscript = '$("#ContentPlaceHolder1_txtFromDate").val("'+ from_date +'")'

driver.execute_script(jscript)

# driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_rbdActionReq_0"]').click()
# time.sleep(20)
#driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_txtFromDate"]').clear().send_keys('21-Jan-2021')