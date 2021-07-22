body_text_temp = '''Dear All,

Please find the attached updated EHV tripping Summary and  ATR (from 1st April 2019 till {till_date} ).

The status of pending cases are highlighted in colors and would be discussed in scheduled Con-call on Monday i.e <b> on {concall_date}. <b> 

 _attach_screen_shot_ 

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
file_folder_path = 'D:\Reports\EHV Tripping weekly meeting\{year}\{month}'

import pyautogui as pygui
from datetime import datetime,timedelta
import time
import clipboard
import win32com.client as win32

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

def write_mail():
    time.sleep(1)
    words = body_text.split(' ')
    for word in words:
        if word == '<b>':
            pygui.hotkey('ctrl','b')
        elif word == '_attach_screen_shot_':
            attach_screenshot()
        elif  word == '_attach_file_':
            pass
            # attach_file()            
        else:
            pygui.typewrite('{} '.format(word))

def attach_screenshot():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks('NEW EHV Tripping weekly format.xlsx')
    ws = wb.ActiveSheet
    ws.Range("A3:S33").CopyPicture()
    pygui.hotkey('ctrl','v')
    
def attach_file():
    today = datetime.now()
    year = today.strftime('%Y')
    month = today.strftime('%b')
    month_ = today.strftime('%m')
    mx = '{} {} {}'.format(month_,month,year[2:])
    path = file_folder_path.format(year=year,month=mx)
    pygui.hotkey('win','r')
    pygui.typewrite(path)
    time.sleep(1)
    pygui.hotkey('enter')
    time.sleep(3)
    pygui.hotkey('end')
    time.sleep(3)
    pygui.hotkey('ctrl','c')
    pygui.hotkey('alt','tab','tab')
    # pygui.hotkey('ctrl','home')
    # pygui.hotkey('down')
    # pygui.hotkey('down')
    # pygui.hotkey('down')
    # pygui.hotkey('down')
    # pygui.hotkey('down')
    pygui.hotkey('ctrl','v')

def enter_cc_subject():
    time.sleep(5)
    pygui.hotkey('tab')
    clipboard.copy(cc_address)
    pygui.hotkey('ctrl','v')
    pygui.hotkey('tab')
    pygui.hotkey('tab')
    pygui.typewrite(subject)
    pygui.hotkey('tab')

lastdate,till_date,concall_date = till_Date()
this_week_tripping = '23'
last_year_tripping = '24'
more_than_1hr_count = 'ONE'

body_text = body_text_temp.format(till_date=till_date,concall_date=concall_date,this_week_tripping=this_week_tripping,last_year_tripping=last_year_tripping,more_than_1hr_count=more_than_1hr_count)
subject = subject_temp.format(till_date=till_date)

time.sleep(5)
enter_cc_subject()
write_mail()