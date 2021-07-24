import pickle
import time
from tkinter.constants import ACTIVE, DISABLED
from urllib.parse import DefragResult
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from PIL import Image, ImageDraw
from datetime import datetime
from openpyxl import Workbook,load_workbook
import glob
import os
import tkinter as tk
from datetime import datetime
import requests
import threading
import json
import win32com.client as win32
import pyautogui

# Variables
chrome_path = r'D:\Process Improvement Project\python_programming\chromedriver.exe'
url_home = 'http://10.125.64.86/ODS/Login.html'
url_ioms = 'http://10.125.64.81/ioms'
url_ioms_status = 'http://10.125.64.81/IOMS/DOMS/Status'
ioms_username = '41018328'
ioms_password = 'pass#123'
url_ioms_report = 'http://10.125.64.81/IOMS/Reports/{}'
ioms_bd = 'breakdowntmis'
ioms_psd = 'pshutmis'
ioms_em = 'eshutmis'
ioms_intbd = 'internalmis'

dts_table_xpath = '//*[@id="GraphicLayer_DT_layer"]'
dts_xpath = '//*[@id="GraphicLayer_DT_layer"]/image[1]'
key_xpath = '//*[@id="ODSMap_root"]/div[3]/div[1]/div[2]/div/table/tbody/tr[{}]/th[1]'
value_xpath = '//*[@id="ODSMap_root"]/div[3]/div[1]/div[2]/div/table/tbody/tr[{}]/td'
planned_xpath = '//*[@id="popover205995"]/div[2]/div/div[1]/input'
unplanned_xpath = '//*[@id="popover205995"]/div[2]/div/div[2]/input'
amr_xpath = '//*[@id="popover205995"]/div[2]/div/div[3]/input'

error,info,keys = [],[],[]
ioms_window = ''
scada_window = ''
timestamp = ''
outageid = ''
grid_name = ''
feeder_name = ''
tripping_date = ''
ioms_d = ''
ods_data,scadalive_data,ioms_data,sanket_data = [],[],[],[]
excel = ''
row = 1
wb = ''
ws = ''
timestamp = ''

def init_driver():
    global driver1
    driver1 = webdriver.Chrome(executable_path=chrome_path)
    driver1.maximize_window()
    login_ODS()

def login_ODS():
    global timestamp
    driver1.get(url_home)
    time.sleep(5)
    driver1.find_element_by_id('EmpID').send_keys('DASHBOARD')
    driver1.find_element_by_id('Password').send_keys('bses@1234')
    driver1.find_element_by_xpath('//*[@id="LoginForm"]/div[3]/div[2]/button').click()
    timestamp = datetime.now().strftime('%d%m%Y_%H%M')

def read_table():
    d={}
    tr_count = len(driver1.find_elements_by_xpath('//*[@id="ODSMap_root"]/div[3]/div[1]/div[2]/div/table/tbody/tr'))
    if tr_count > 0:
        for n in range(1,tr_count+1):
            key = driver1.find_element_by_xpath(key_xpath.format(n)).text
            value = driver1.find_element_by_xpath(value_xpath.format(n)).text
            d[key] = value
    # d['screenshot'] = screenshot(d['Outage ID'],driver1)
    return d

def screenshot(outageid,driver):
    check_directory(timestamp)
    path = r'D:\Process Improvement Project\python_programming\Dashboard_ODS\screenshots\{}\{}.png'
    exsisting_files = glob.glob(path.format(timestamp,'*'))
    max_ = 0
    for file in exsisting_files:
        try:
            no =int(file.split('_')[-1][:-4])
        except:
            no = 0
        if no > max_:
            max_ = no
    filename = outageid + 'screenshot_{}'.format(max_+1)
    driver.get_screenshot_as_file(path.format(timestamp,filename))
    watermark(path.format(timestamp,filename))
    return filename

def check_directory(timestamp):
    path = r'D:\Process Improvement Project\python_programming\Dashboard_ODS\screenshots\{}'.format(timestamp)
    if not os.path.exists(path):
        os.mkdir(path)

def watermark(img):
    my_img = Image.open(img)
    x,y = my_img.size
    img_edit = ImageDraw.Draw(my_img)
    text = 'Screenshot Time :\n' + str(datetime.now().strftime('%H:%M:%S - %d/%m/%Y'))
    img_edit.text((x-150,y-50),text,fill=(0,0,0))
    my_img.save(img)

def scada_live():
    d = read_table()
    gid = ''
    with open(r'D:\Process Improvement Project\python_programming\Dashboard_ODS\grid_name.dat','rb') as f:
        gridlist = pickle.load(f)
    for g in gridlist:
        if g['gname'] == d['Grid Name']:
            gid = g['gid']
            break
    url = f'http://10.125.64.86/Scada_Live/Base/GETFeederLoadByGrid?gridId={gid}&Type=All'
    try:
        response = requests.get(url,20)
        data = json.loads(response.json())
        for dt in data:
            if dt['FEEDERNAME'] == d['Feeder Name']:
                print(dt)
                scadamsg(dt)
                # scada_window_msg(dt)
                break
    except Exception as e:
        print(e.args)
        scadamsg('Some Error Occured')
        # scada_window_msg('Some Error Occured')

# def imgloader():
#     img = driver2.find_element_by_id('imageLoader')
#     while True:
#         if 'none' in img.get_attribute('style'):
#             break

def refresh():
    driver1.refresh()

def scadamsg(d):
    if type(d) == str:
        msg = d
    else:
        msg =f"### SCADA LIVE DATA ###\\nGRID NAME : {d['GRIDNAME']}\\nFeeder Name : {d['FEEDERNAME']}\\nTime : {d['Y_SYSTIME']}\\nBreaker Status : {d['EVENT']}\\nAutoTrip : {d['AUTOTRIP']}\\nR-Y-B Current : {d['R']},{d['Y']},{d['B']}\\nR-Y-B Voltage : {d['RY']},{d['BR']},{d['YB']}"
    driver1.execute_script(f'alert("{msg}")')
    time.sleep(1)
    window_screenshot('window_')

def scada_window_msg(d):
    popup = tk.Tk()
    popup.title('SCADA LIVE DATA')
    popup.geometry('1000x100')
    if type(d) != str:
        scada_data = {
            'GRID NAME' : d['GRIDNAME'],
            'Feeder Name' : d['FEEDERNAME'],
            'Time' : d['Y_SYSTIME'],
            'Breaker Status' : d['EVENT'],
            'AutoTrip' : d['AUTOTRIP'],
            'Over Current' : d['O_C'],
            'Earth Fault' : d['E_F'],
            'R-Y-B Current' : f"{d['R']}, {d['Y']}, {d['B']}",
            'R-Y-B Voltage' : f"{d['RY']}, {d['BR']}, {d['YB']}"
        }
    else:
        scada_data = {
            'Error Message' : d
        }
    for col,tp in enumerate(scada_data.items()):
        lbl1 = tk.Label(popup,text=tp[0])
        lbl2 = tk.Label(popup,text=tp[1])
        lbl1.grid(row=1,column=col)
        lbl2.grid(row=2,column=col)
    popup.attributes('-topmost',1)
    # popup.mainloop()
    time.sleep(1)
    window_screenshot('window_')

def window_screenshot(outageid):
    check_directory(timestamp)
    path = r'D:\Process Improvement Project\python_programming\Dashboard_ODS\screenshots\{}\{}.png'
    exsisting_files = glob.glob(path.format(timestamp,'*'))
    max_ = 0
    for file in exsisting_files:
        try:
            no =int(file.split('_')[-1][:-4])
        except:
            no = 0
        if no > max_:
            max_ = no
    filename = outageid + 'screenshot_{}'.format(max_+1)
    myscreenshot = pyautogui.screenshot()
    myscreenshot.save(path.format(timestamp,filename))
    # driver.get_screenshot_as_file(path.format(timestamp,filename))
    watermark(path.format(timestamp,filename))

# d = {'PTR': 'VASANT VIHAR_33/11KV POWER TRF NO-1(20MVA)', 'PTRID': 'DLDLHIRKPR3005XXXXPT0003', 'GRIDNAME': '33 kV VASANT VIHAR GRID', 'GRIDID': 'DLDLHIRKPR3005', 'FEEDERID': 'VAVI_11KV_401274', 'FEEDERNAME': '11kV FDR O/G VASANT ENCLAVE S/S NO.-2', 'EXTERNALSCADAID': 'VAVI11B11131', 'Y_SYSTIME': '15/07/21 13:58:35', 'SWITCH_ID': '401274', 'EVENT': 'ON  ', 'O_C': '--', 'E_F': '--', 'AUTOTRIP': '--', 'PW': '-0.9867', 'THD': '1.1678', 'RAP': '0.20', 'R': '68.27', 'Y': '72.89', 'B': '68.42', 'RY': '0', 'BR': '0', 'YB': '0', 'AP': '1.27', 'Y_LSTPKLOAD_Day': '84.90', 'Y_LSTPKLOAD_Day_DT': '14-Jul-21 12:03:35 AM', 'Y_LSTPKLOAD_Month': '124.12', 'Y_LSTPKLOAD_Month_DT': '01-Jul-21 04:13:39 PM'}


def quit():
    driver1.quit()
    window.destroy()

window = tk.Tk()
window.title("ODS assistant")
window.geometry('1300x35')

functions = [
    {
        'fun' : init_driver,
        'text' : 'Open Browser'
    },
    {
        'fun' : login_ODS,
        'text' : 'Login ODS'
    },
    {
        'fun' : refresh,
        'text' : 'Refresh ODS'
    },
    {
        'fun' : scada_live,
        'text' : 'Get Scada Live Status'
    },
    {
        'fun' : quit,
        'text' : 'Quit'
    }
]

for col,fun in enumerate(functions):
    btn = tk.Button(window,text=fun['text'],bd=5,command=fun['fun'],width=len(fun['text']))
    btn.grid(row=1,column=col)

window.attributes('-topmost',1)

window.mainloop()


        