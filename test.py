# import glob
# import time
# import requests
# path = r'D:\Process Improvement Project\python_programming\Downloads'
# while True:
#     files = glob.glob(path+'\\*.*')
#     print(len(files))
#     time.sleep(100)
# url = 'http://10.125.66.12:6080/arcgis/rest/services/GIS/OUTAGE_2/MapServer/4/query?f=json&where=FEEDERID%20IN(%27NGLW_11KV_301666%27)&returnGeometry=true&spatialRel=esriSpatialRelIntersects&outFields=*&outSR=102100'

# res = requests.get(url)

# from openpyxl import load_workbook

# # files = glob.glob(r'D:\Reports\Dash Board Report\2018\2018 06 Jun\MIS Report\*.xls')

# file = r'D:\Reports\Dash Board Report\2021\04 Apr 2021\daily tripping details\Tripping details 01.04.2021.xlsx'

# wb = load_workbook(file)

# ws = wb.active
# for a in range(1,ws.max_row+1):
#     for b in range(1,ws.max_column+1):
#         print(ws.cell(a,b).number_format,end=' ')
#     print()
# import glob

# emails = glob.glob(r'D:\Reports\Dash Board Report\2021\05 May 2021\emails\*.eml')

# for mail in emails:
#     if 'EHV TRIPPING' in mail:
#         print(mail,'Tripping')
#     elif 'Dash board report' in mail:
#         print(mail,'Dash Board')

# # from datetime import datetime
# import glob

# # file_path = r'D:\Process Improvement Project\python_programming\Dashboard_ODS\excel_reports\{}.xlsx'.format(datetime.now().strftime('%d%m%Y_%H%M'))

# # print((file_path))

# path = r'D:\Process Improvement Project\python_programming\Dashboard_ODS\screenshots\{}.png'
# exsisting_files = glob.glob(path.format('*'))
# max_ = 0
# for file in exsisting_files:
#     try:
#         no =int(file.split('_')[-1][:-4])
#     except:
#         no = 0
#     if no > max_:
#         max_ = no
# filename = 'screenshot_{}'.format(max_+1)

# print(path.format(filename))

# dic = {
#     'a' : 3223,
#     'b' : 'fsdfsfs',
#     'c' : 'fsfsdfsf'
# }

# if 'b' in dic.keys():
#     print('True')
# else:
#     print('False')

# from datetime import datetime
# from PIL import ImageDraw,Image
# import time

# start = time.time()
# time.sleep(2)
# total_time = time.time() - start

# print(round(total_time,3))

# p = r'D:\Process Improvement Project\python_programming\Dashboard_ODS\screenshots\16052021_1704\main_screenshot_1.png'
# def watermark(img):
#     my_img = Image.open(img)
#     x,y = my_img.size
#     img_edit = ImageDraw.Draw(my_img)
#     text = 'Screenshot Time :\n {}'.format(datetime.now().strftime('%H:%M:%s - %d/%m/%Y'))
#     img_edit.text((x-150,y-50),text,fill=(0,0,0))
#     my_img.save(img)
# from base64 import decode
# import pickle

# txt = ''
# gridlist = []
# with open('Dashboard_ODS\grid_name.txt','r') as f:
#     txt = f.read()
# li = txt.split('</option>')

# for l in li[1:-1]:
#     l2 = l.split('>')
#     try:
#         d = {
#             'gid' : l2[0][-15:-1],
#             'gname' : l2[1]

#         }
#         gridlist.append(d)
#     except:
#         pass
# with open('Dashboard_ODS\grid_name.dat','rb') as f:
#     gridlist = pickle.load(f)
# print(gridlist)
# for grid in gridlist:
#     print(grid)

#  li_old = pickle.load(f)

# dic = {
#     'x' : 3,
#     'y' : ['fdsfdsf']
# }
# print(dic)

# for n in range(5):
#     if type(dic['x']) != list:
#         dic['x'] = [dic['x']]
#     dic['x'].append(n)

# print(dic)

# t = ','.join('{}'.format(v) for v in dic['y'])
# u = str(dic['x'])

# print(t)
# print(u)

# filelist = []

# path = r'D:\Reports\Dash Board Report\{y}\{m} {y}\MIS Report'
# path2 = r'E:\mis 20-21\{}'

# years = [2020,2021]
# months = ['04 Apr','05 May','06 Jun','07 July','08 Aug','09 Sep','10 Oct','11 Nov','12 Dec','01 Jan','02 Feb','03 Mar']
# for month in months:
#     if months.index(month) <= 8:
#         year = '2020'
#     else:
#         year = '2021'
#     pathx = path.format(y=year,m=month)
#     print(pathx)
# #     fl = glob(pathx+'\\*')
# #     for f in fl:
# #         filelist.append(f)

# # for f in filelist:
# #     fname = f.split('\\')[-1]
# #     shutil.copy(f,path2.format(fname))

# import pickle
# from socket import timeout
# import time
# from tkinter.constants import ACTIVE, DISABLED
# from selenium import webdriver
# from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.chrome.options import Options
# from PIL import Image, ImageDraw
# from datetime import datetime
# from openpyxl import Workbook,load_workbook
# import glob
# import os
# import tkinter as tk
# from datetime import datetime
# import requests
# import threading
# import json

# chrome_path = r'D:\Process Improvement Project\python_programming\chromedriver.exe'

# url_ioms = 'http://10.125.64.81/IOMS/'
# ioms_username = '41018328'
# ioms_password = 'pass#123'

# driver = webdriver.Chrome(executable_path=chrome_path)
# driver.maximize_window()

# driver.get(url_ioms)
# driver.find_element_by_name('UserID').send_keys(ioms_username)
# driver.find_element_by_name('Password').send_keys(ioms_password)
# driver.find_element_by_xpath('//*[@id="login"]/form/table/tbody/tr[3]/td[3]/button').click()


# url_Req = 'http://10.125.64.81/IOMS/Reports/getbreakdownmisreport'

# payload = {"Company":"BRPL","circle":"SOUTH 2","Division":"NZD","voltage":"ALL","from":"15-05-2021 01:00","to":"25-05-2021 02:00","Reason":"All"}

# res = requests.get(url=url_Req,params=payload,timeout=20)

# html = json.loads(res.json())


# with open("temp.html",'w') as f:
#     f.write(res.content)

# from datetime import datetime 

# d = datetime.today().strftime('%d-%m-%Y')

# print(d)

# import os
# from glob import glob

# filelist = glob(r"C:\Users\BRPL\Downloads\*.*")

# fileinfo = os.stat(filelist[1])

# def latestfile():
#     filelist = glob(r"C:\Users\BRPL\Downloads\*.*")
#     latestfile = ''
#     latesttime = 99999999999
#     for file in filelist:
#         time = os.stat(file).st_mtime
#         if time < latesttime:
#             latestfile = file
#     return latestfile

# with open('Winium-Test\mail_body.txt','r') as fr:
#     txt_lines = fr.readlines

# import win32com.client as win32
# from datetime import datetime, timedelta
# import os.path

# def new_method(wsx):
#     offset = 0
#     date = ''
#     if wsx.Cells(1,1).Value == None:
#         date = datetime.strptime(wsx.Cells(1,2).Value[-10:],'%d.%m.%Y').strftime('%d-%b-%Y')
#         offset = 9
#     else:
#         date = datetime.strptime(wsx.Cells(1,1).Value[-10:],'%d.%m.%Y').strftime('%d-%b-%Y')
#         offset = 8
#     d = {
#         'date1' : dt,
#         'date2' : date,
#         'GENERATION': '',
#         'BRPL' : '',
#         'TRANSCO' : '',
#         'NDPL' : ''
#     }
#     max_col = 20 #wsx.UsedRange.Columns.Count
#     max_row = wsx.UsedRange.Rows.Count
#     for row in range(1,max_row):
#         for col in range(1,max_col):
#             if type(wsx.Cells(row,col).Value) == str:
#                 if 'BRPL' in wsx.Cells(row,col).Value:
#                     d['GENERATION'] = wsx.Cells(row-1,offset).Value
#                 elif 'TRANSCO' in wsx.Cells(row,col).Value:
#                     d['BRPL'] = wsx.Cells(row-1,offset).Value
#                 elif 'NDPL' in wsx.Cells(row,col).Value:
#                     d['TRANSCO'] = wsx.Cells(row-1,offset).Value
#                 elif 'GRAND TOTAL' in wsx.Cells(row,col).Value:
#                     d['NDPL'] = wsx.Cells(row-1,offset).Value
#                     return d

# excel = win32.gencache.EnsureDispatch('Excel.Application')
# excel.Visible = False
# pathx = r'D:\Reports\Dash Board Report\{yyyy}\{mm} {mmm} {yyyy}\load shedding\BRPL LOAD SHEDDING DETAIL WITH REASON {dd}.{mm}.{yyyy}.xls'
# dt = datetime(2020,3,31)
# mu_list = []
# file_not_find = []
# while dt < datetime.now():
#     dt = dt + timedelta(1)
#     yyyy = dt.strftime('%Y')
#     mmm = dt.strftime('%b')
#     mm = dt.strftime('%m')
#     dd = dt.strftime('%d')
#     path = pathx.format(yyyy=yyyy,mm=mm,mmm=mmm,dd=dd)
#     print(path,os.path.exists(path))
#     if os.path.exists(path):
#         wbx = excel.Workbooks.Open(path)
#         wsx = wbx.ActiveSheet
#         d = new_method(wsx)
#         wbx.Close()
#     else:
#         file_not_find.append(path)
#         d = {
#         'date1' : dt,
#         'date2' : '',
#         'GENERATION': '',
#         'BRPL' : '',
#         'TRANSCO' : '',
#         'NDPL' : ''
#         }
#     mu_list.append(d)
# excel.Visible = True
# wb = excel.Workbooks.Add()
# ws = wb.Worksheets.Add()
# ws.Name = 'Muloss'
# ws.Range("A1:E1").Value = ['Date1','Date2','GENERATION','BRPL','TRANSCO','NDPL']
# for n,dic in enumerate(mu_list,start=2):
# #     ws.Range('A{n}:E{n}'.format(n=n)).Value = [dic.values()]


# # import os
# # import pyautogui
# # winium_path = r'D:\Process Improvement Project\python_programming\Winium.Desktop.Driver.exe'
# # pyautogui.hotkey('win','r')
# # pyautogui.typewrite(winium_path)
# # pyautogui.hotkey('enter')
# # # os.system(winium_path)

# # import time
# # import pyautogui

# # time.sleep(5)

# # pyautogui.hotkey('ctrl','home')
# # pyautogui.hotkey(['down']*6)
# # d = {}

# # d['GRIDNAME'] = 'xyz'
# # d['FEEDERNAME'] = 'abc'
# # d['Y_SYSTIME'] = 'XX.XX.XXXX'
# # d['EVENT'] = 'ON'
# # d['AUTOTRIP'] = 'NO'
# # d['R'] = 12
# # d['Y'] = 14
# # d['B'] = 16
# # d['RY'] = 11
# # d['BR'] = 11
# # d['YB'] = 11

# # msg = f'''GRID NAME : {d['GRIDNAME']}
# # Feeder Name : {d['FEEDERNAME']}
# # Time : {d['Y_SYSTIME']}
# # Breaker Status : {d['EVENT']}
# # AutoTrip : {d['AUTOTRIP']}
# # R-Y-B Current : {d['R']},{d['Y']},{d['B']}
# # R-Y-B Voltage : {d['RY']},{d['BR']},{d['YB']}'''

# # print(msg)

# def type1():
#     print('print Type 1')

# def type2():
#     print('print Type 2')

# def type3():
#     print('print Type 3')

# def type4():
#     print('print Type 4')

# def type5():
#     print('print Type 5')

# def type6():
#     print('Type 6')



# switchcase = {
#     str : type1,
#     float : type2,
#     int : type3,


# }

# i = 'sfsfsdf'
# print(type(i))
# print()
# switchcase[type(i)]()


# type_list = []
# key_list = ['End\nTime (In Date Hrs)','Start\nTime (In Date Hrs)','Time (In Date Hrs)']
# for data in company_data:
#     for key in key_list:
#         if type(data[key]) not in type_list:
#             type_list.append(type(data[key]))

import requests

payload = {
    "Company":"BRPL",
    "circle":"",
    "Division":"",
    "voltage":"ALL",
    "from":"01-04-2021 00:00",
    "to":"22-07-2021 00:00",
    "Reason":"All"
    }
cookies = {
    'ASP.NET_SessionId' :'sxw43mph50slbzymfkltewcg;', 
    '.ASPXAUTH' : '958286E1AEB9A45982E78620DAD8E0EA5153A07B840E066F9FFB47733DCA3B351D302991768AAD8B301952DEB0789C7D43101C5BAA18A1AE4391ADE39C35F8FFA8F4C9C3DEE7CD253B9CFC06640EEAC5831744BC1C8ED10D5EC4BF7D9188A5FC8525E373A0029B02534561BCF349B6B18E022CD931457AB3CA95C2EB5A95F3D57E40D6960409CDB6F03DBE9836C198F9'
}

url = 'http://10.125.64.81/IOMS/Reports/getbreakdownmisreport'

res = requests.get(url=url)



