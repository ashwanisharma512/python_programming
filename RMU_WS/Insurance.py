from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
import tkinter as tk
from datetime import datetime
import random
from glob import glob
import time
import pyautogui as pygui

chrome_path = r'D:\Process Improvement Project\python_programming\chromedriver.exe'
base_url = 'http://10.125.64.86/Insurance_claim'
dr = ''


def init_driver():
    global dr
    dr = webdriver.Chrome(executable_path=chrome_path)
    dr.maximize_window()

def input_box(id,text):
    elm = dr.find_element_by_id(id)
    value = elm.get_attribute('value')
    if value == '':
        elm.send_keys(text)
    else:
        tt = '\b'*len(value)
        elm.send_keys(tt)
        elm.send_keys(text)

def login_insurance_user():
    dr.get(base_url)
    input_box('ContentPlaceHolder1_txtuser','USER')
    input_box('ContentPlaceHolder1_txtpass','12345678')
    dr.find_element_by_id('ContentPlaceHolder1_LoginButton').click()

def login_insurance_admin():
    dr.get(base_url)
    input_box('ContentPlaceHolder1_txtuser','ADMIN')
    input_box('ContentPlaceHolder1_txtpass','12345678')
    dr.find_element_by_id('ContentPlaceHolder1_LoginButton').click()

def select_sub_div():
    ids = ['ContentPlaceHolder1_ddlCirle','ContentPlaceHolder1_ddlDivision','ContentPlaceHolder1_ddlSubDivision']
    for id in ids:
        sel = Select(dr.find_element_by_id(id))
        nos = len(sel.options) - 1
        index = random.randint(1,nos)
        sel.select_by_index(index)
        time.sleep(1)

def emp_detail():
    details = {
        'ContentPlaceHolder1_txtuser' : 'Test User',
        'ContentPlaceHolder1_txtempid' : '41012345',
        'ContentPlaceHolder1_txtmobile' : '9876543210',
        'ContentPlaceHolder1_txtEmailID' : 'testuser@relianceada.com'
    }
    for key,val in details.items():
        input_box(key,val)

def asset_detail():
    selected_value = (Select(dr.find_element_by_id('ContentPlaceHolder1_ddlAssetType')).all_selected_options)[0].get_attribute('value')
    details = {
        'ContentPlaceHolder1_txtassetQTY' : random.randint(1,10),
        'ContentPlaceHolder1_txtassetCodeSAP' : '2000{}'.format(random.randint(10000,99999)),
        'ContentPlaceHolder1_txtAccessories' : 'Accessories Name',
        'ContentPlaceHolder1_txtAccessoriesDT' : 'Functional Location',
        'ContentPlaceHolder1_txtAccessorySAP_Code' : '2000{}'.format(random.randint(10000,99999)),
        'ContentPlaceHolder1_txtAccessoryQty' : random.randint(1,10),
        'ContentPlaceHolder1_txtAccessoryCost' : random.randint(100,10000)
    }
    if selected_value != '0':
        for key,val in details.items():
            input_box(key,val)

def remark():
    input_box('ContentPlaceHolder1_txtremarks','Test Remarks')

def upload_document():
    details = {
	    "6" : r'D:\Process Improvement Project\python_programming\RMU_WS\Test_file\6. MRS.pdf', #"Copy of MRS where damaged/stolen asset /part replaced",
	    "2" : r'D:\Process Improvement Project\python_programming\RMU_WS\Test_file\2. Fire Report.pdf', #"In case of FIRE, FIRE report",
	    "5" : r'D:\Process Improvement Project\python_programming\RMU_WS\Test_file\5. SAP rate.pdf', #"SAP rate document for the BOQ verification",
        "1" : r'D:\Process Improvement Project\python_programming\RMU_WS\Test_file\1. e-FIR.pdf', #"Scan copy of e-FIR copy u/s 379 of IPC (online) with complete details",
	    "7" : r'D:\Process Improvement Project\python_programming\RMU_WS\Test_file\7. IR.pdf', #"Signed &amp; stamped copy Incident report",
	    "4" : r'D:\Process Improvement Project\python_programming\RMU_WS\Test_file\4. boq.pdf', #"Signed &amp; stamped copy of BOQ of damaged asset or its accessries",
	    "3" : r'D:\Process Improvement Project\python_programming\RMU_WS\Test_file\3. digi.JPG', #"Supporting Pictures (no limitation) (before &amp; after the incident where asset replaced)\n",
    }
    file_id = 'ContentPlaceHolder1_FileUpload1'
    for key,val in details.items():
        sel = Select(dr.find_element_by_id('ContentPlaceHolder1_ddldoctype'))
        opts = sel.options
        valuelist = [op.get_attribute('value') for op in opts ]
        print(f'key :{key} \t val : {val} \t {valuelist}')
        if key in valuelist:
            sel.select_by_value(key)
            time.sleep(2)
            dr.find_element_by_id(file_id).send_keys(val)
            time.sleep(2)
            dr.find_element_by_id('ContentPlaceHolder1_btnimage').click()
            time.sleep(2)

window = tk.Tk()
window.title("RMU assistant")
window.geometry('1250x35')

btn_details = [
    {
        'text' : 'Open Browser',
        'function' : init_driver,

    },
    {
        'text' : 'Login USER',
        'function' : login_insurance_user
    },
    {
        'text' : 'Login ADMIN',
        'function' : login_insurance_admin
    },
    {
        'text' : 'Select SubDiv',
        'function' : select_sub_div
    },
    {
        'text' : 'Emp Detail',
        'function' : emp_detail
    },
    {
        'text' : 'Asset Details',
        'function' : asset_detail
    },
    {
        'text' : 'Remark',
        'function' : remark
    },
    {
        'text' : 'Upload Document',
        'function' : upload_document
    }
]

for c,bt in enumerate(btn_details):
    btn = tk.Button(window,text=bt['text'],bd=5,command=bt['function'],width=int(len(bt['text'])))
    btn.grid(row=1,column=c)

window.attributes('-topmost',1)

# window.mainloop()

