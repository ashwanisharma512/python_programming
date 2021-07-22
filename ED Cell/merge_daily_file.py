from glob import glob
from datetime import datetime,timedelta
from pickle import encode_long
import win32com.client as win32
import shutil
import email
import os

def colnum_string(n):
    string = ''
    while n > 0:
        n,rem = divmod(n-1,26)
        string = chr(65 + rem) + string
    return string

def sheet_name():
    pass

excel = win32.Dispatch('Excel.Application')

bypl_folder = r'D:\Process Improvement Project\ED CELL\Reports from BYPL\{}\{}'

folder_name = (datetime.today()-timedelta(1)).strftime('%d %b %Y')

file_list = glob(bypl_folder.format(folder_name,'*.xlsx'))

def merge_files():
    global new_wb
    new_wb = excel.Workbooks.Add()
    for file in file_list:
        wb = excel.Workbooks.Open(file)
        ws = wb.ActiveSheet
        colno = colnum_string(ws.UsedRange.columns.count)
        rowno = ws.UsedRange.rows.count
        new_ws = new_wb.Worksheets.Add()
        if ws.name not in [sh.Name for sh in new_wb.Sheets]:
            new_ws.Name = ws.name
        ws.Range(f'A1:{colno}{rowno}').Copy(Destination = new_ws.Range(f'A1:{colno}{rowno}'))
        wb.Close()
    new_file_name = bypl_folder.format(folder_name,'Cumulative_{}.xlsx'.format((datetime.today()-timedelta(1)).strftime('%d_%m_%Y')))
    new_wb.SaveAs(new_file_name)

merge_files()

def modify_cumulative_file():
    new_ws = new_wb.Worksheets('OUTAGE')


def reading_mails():
    mail_folder_path = r'D:\Process Improvement Project\ED CELL\Outage Reports\emails\{}'
    mail_list = glob(mail_folder_path.format('*.eml'))
    #seperate mails
    for mail in mail_list:
        if 'BRPL' in mail:
            filename = mail.split('\\')[-1]
            shutil.move(mail,mail_folder_path.format('BRPL')+'\\'+filename)
        if 'BYPL' in mail:
            filename = mail.split('\\')[-1]
            shutil.move(mail,mail_folder_path.format('BYPL')+'\\'+filename)
    mail_list = glob(mail_folder_path.format('*.eml'))
    for mail in mail_list:
        with open(mail,'rb') as f:
            em_ = f.read()
        em = email.message_from_bytes(em_)
    
# BSES YAMUNA POWER LTD. OUTAGE REPORT FOR 11.07.2021.xlsx

def collect_data():
    global data
    key = ['Sr. No.', 'Date', 'Division', 'Start Time', 'End Time', 'Duration', 'Main Reason', 'Details', 'Action', 'Main Area affected', 'Remarks']    
    data = []
    check_bd_catagory = {
    'Breakdown less than one hours' : 'BD',
    'Breakdown More than one hours' : 'BD',
    'Emergency Shutdown less than one hours' : 'ESD',
    'Emergency Shutdown more than one hours' : 'ESD',
    'DTL Tripping less than one hours' : 'DTL',
    'DTL Tripping More than one hours' : 'DTL',
    }
    for file in relative_files:
        wb = excel.Workbooks.Open(file)
        ws = wb.ActiveSheet
        catagory = ''
        for row in range(1,ws.UsedRange.Rows.Count+1):
            d = {}
            if ws.Cells(row,1).Value in check_bd_catagory.keys():
                catagory = check_bd_catagory[ws.Cells(row,1).Value]
            d['catagory'] = catagory
            if type(ws.Cells(row,1).Value ) == float:
                for col,k in enumerate(key,start=1):
                    d[k] = ws.Cells(row,col).Value
                data.append(d)
        wb.Close()

def find_related_file():
    global relative_files
    relative_files = []
    base_path = r'D:\Process Improvement Project\ED CELL\Outage Reports\Reports from BYPL'
    relative_file_name = 'BSES YAMUNA POWER LTD. OUTAGE REPORT FOR'
    for root,dir,file in os.walk(base_path):
        for f in file:
            if '.xls' in f:
                if relative_file_name in f:
                    relative_files.append(os.path.join(root,f))

def cumulative_file():
    base_path = r'D:\Process Improvement Project\ED CELL\Outage Reports\Reports from BYPL\{}'
    wb = excel.Workbooks.Add()
    ws = wb.ActiveSheet
    for col,key in enumerate(data[0].keys(),start=1):
        ws.Cells(1,col).Value = key
    for row,d in enumerate(data,start=2):
        for col,val in enumerate(d.values(),start=1):
            ws.Cells(row,col).Value = val
    wb.Saveas(base_path.format('Cumulative till date.xlsx'))

def date_change():
    ws = wb.ActiveSheet
    rowno = ws.UsedRange.Rows.Count
    for row in range(2,rowno+1):
        dt = ws.Cells(row,3).Value
        new_dt = '{}-{}-{}'.format(dt[2],dt[4,6],dt[-4:])
        ws.Cells(row,3).Value = new_dt
