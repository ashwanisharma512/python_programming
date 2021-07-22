import tkinter as tk
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from datetime import datetime,timedelta
import time

# Variables Decelaration
chrome_path = r'D:\Process Improvement Project\python_programming\chromedriver.exe'
base_url1 = '10.125.64.87'
base_url2 = '125.22.84.50:7860'
url_ioms = 'http://{}/ioms'
ioms_username = '41018328'
ioms_password1 = 'india@123'
ioms_password2 = 'pass#123'

base_url = base_url2
ioms_password = ioms_password2

def init_driver():
    global dr
    dr = webdriver.Chrome(executable_path=chrome_path)
    dr.maximize_window()
    login_ioms()

def login_ioms():
    dr.get(url_ioms.format(base_url))
    dr.find_element_by_name('UserID').send_keys(ioms_username)
    dr.find_element_by_name('Password').send_keys(ioms_password)
    dr.find_element_by_xpath('//*[@id="login"]/form/table/tbody/tr[3]/td[3]/button').click()

def imgloader():
    imgloader = dr.find_element_by_id('imageLoader')
    while True:
        if 'none' in imgloader.get_attribute('style'):
            break

def grab_data(tablename):
    data = []
    script1 = '$("#bdfrom").val("{} 00:00")'.format((datetime.today() -timedelta(1)).strftime('%d-%m-%Y'))
    dr.execute_script(script1)
    script2 = '$("#bdto").val("{} 00:00")'.format(datetime.today().strftime('%d-%m-%Y'))
    dr.execute_script(script2)
    dr.find_element_by_id('showbd').click()
    time.sleep(1)
    imgloader()
    head_coulmn_counts = len(dr.find_elements_by_xpath(f'//*[@id="{tablename}"]/thead/tr/th'))
    keys = []
    for x in range(1,head_coulmn_counts+1):
        keys.append(dr.find_element_by_xpath(f'//*[@id="{tablename}"]/thead/tr/th[{x}]').text)
    body_row_count = len(dr.find_elements_by_xpath(f'//*[@id="{tablename}"]/tbody/tr'))
    for row in range(1,body_row_count+1):
        d = {}
        for col in range(1,head_coulmn_counts+1):
            d[keys[col-1]] = dr.find_element_by_xpath(f'//*[@id="{tablename}"]/tbody/tr[{row}]/td[{col}]').text
        data.append(d)
    return data,keys

def breakdown_report():
    global bd_data,bd_keys
    dr.get('http://{}/IOMS/Reports/breakdowntmis'.format(base_url))
    bd_data,bd_keys = grab_data('gridtable')

def planshutdown_report():
    global psd_data,psd_keys
    dr.get('http://{}/IOMS/Reports/pshutmis'.format(base_url))
    psd_data,psd_keys = grab_data('gridT')

def emergencyshutdown_report():
    global esd_data,esd_keys
    dr.get('http://{}/IOMS/Reports/eshutmis'.format(base_url))
    esd_data,esd_keys = grab_data('gridT')

breakdown_report()
planshutdown_report()
emergencyshutdown_report()



# //*[@id="gridT"]/tbody/tr[1]/td[2]