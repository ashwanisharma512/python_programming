lotus_note_address = r'C:\Program Files (x86)\IBM\Notes\notes.exe'
mail_name = 'BRPLCEOCell Monitoring - Mail'
body_text_temp = '''Dear All,

Please find the attached updated EHV tripping Summary and  ATR (from 1st April 2019 till {till_date} ).

The status of pending cases are highlighted in colors and would be discussed in scheduled Con-call on Monday i.e <b> on {concall_date}. <b> 

 _attach_screenshot_ 

 _attach_file_ 

Note:- <b> Total Tripping events in the week were {this_week_tripping} nos. comparatively to {last_year_tripping} nos. last year same period.
         There were {more_than_1hr_count} instance in last week when load was affected for around 1 Hr. <b> 

--
Regards,
Monitoring Team,
CEO Cell,
BRPL, Delhi
'''

cc_address = 'Chandra K Mohan/REL/RelianceADA@INFOCOMM, Rajesh M Bansal/REL/RelianceADA@INFOCOMM, Sanjay Ku Garg/REL/RelianceADA@INFOCOMM, Ashwani Sharma/REL/RelianceADA@INFOCOMM'
subject_temp = 'EHV Tripping Summary & ATR till {till_date}'
try_count = 0
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pyautogui as pygui
import time
import win32com.client as win32
from datetime import datetime,timedelta

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks("EHV Tripping weekly format.xlsm")
ws = wb.Worksheets('Jun 21')

lotus_password = 'jun@2021'
pygui.hotkey('win','r')
pygui.typewrite('D:\Process Improvement Project\python_programming\Winium.Desktop.Driver.exe')
pygui.hotkey('enter')

while True:
    try:
        try_count +=1
        driver = webdriver.Remote(
            command_executor='http://localhost:9999',
            desired_capabilities={
            "debugConnectToRunningApp": 'false',
            "keyboardSimulator": '0',
            "app": lotus_note_address
            })
        break
    except:
        time.sleep(10)
    finally:
        if try_count > 5:
            print('Unable to start server')
            break

def last_week_count(lastdate):
    sum = 0
    for col in range(7):
        sum += ws.Cells(22,2+col).Value
    return int(sum)

def till_Date():
    today = datetime.now()
    for n in range(4):
        d = today - timedelta(n)
        if d.strftime('%A') == 'Thursday':
            break
    return d,post_date(d),post_date(d+timedelta(4))

def post_date(d):
    day = int(d.strftime('%d'))
    month = '{}'.format(d.strftime('%b'))
    year = '{}'.format(d.strftime('%Y'))
    if day%10 == 1:
        txt = '{}st'.format(day)
    elif day%10 == 2:
        txt = '{}nd'.format(day)
    elif day%10 == 3:
        txt = '{}rd'.format(day)
    else:
        txt = '{}th'.format(day)
    return '{} {} {}'.format(txt,month,year)

lastdate,till_date,concall_date = till_Date()
this_week_tripping = last_week_count(lastdate)
last_year_tripping = '****'
more_than_1hr_count = '****'

body_text = body_text_temp.format(till_date=till_date,concall_date=concall_date,this_week_tripping=this_week_tripping,last_year_tripping=last_year_tripping,more_than_1hr_count=more_than_1hr_count)
subject = subject_temp.format(till_date=till_date)

def write_mail():
    mail_body = driver.find_element_by_class_name('NotesRichText')
    mail_body.send_keys('') 
    time.sleep(2)
    words = body_text.split(' ')
    for word in words:
        if word == '<b>':
            pygui.hotkey('ctrl','b')
        elif word == '_attach_screenshot_':
            attach_screenshot()
        elif word == '_attach_file_':
            attach_file()
        else:
            pygui.typewrite('{} '.format(word))

def check_login_screen():
    password_screen = driver.find_elements_by_class_name('IRIS.password') 
    if len(password_screen)> 0:
        password_screen[0].send_keys(lotus_password)
        driver.find_element_by_name('Log In').click()

def attach_screenshot():
    print('screenshot')
    pass

def attach_file():
    print('attachment')
    pass

def enter_cc_subject():
    elms = driver.find_elements_by_class_name('IRIS.tedit')   # 0. To, 1.CC, 2. BCC, 3. Subject
    elms[1].send_keys(cc_address)
    elms[3].send_keys(subject)


check_login_screen()
time.sleep(5)
mail = driver.find_element_by_name(mail_name)
mail.click()
time.sleep(5)
new_mail = driver.find_element_by_name('New')
new_mail.click()
time.sleep(5)

enter_cc_subject()
write_mail()


