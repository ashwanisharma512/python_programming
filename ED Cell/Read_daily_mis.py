import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
# import win32com.client.constants as cc
import glob,re
from datetime import datetime, time,timedelta
from module1 import time_converter,value_converstion
import tkinter as tk

company = 'BRPL'
main_file = r"D:\Process Improvement Project\ED CELL\Outage Reports\Report as per our format\Updated_Summary.xlsx"

excel.Visible = True

def change_company():
    global company
    if company == 'BRPL':
        company = 'BYPL'        
    else:
        company = 'BRPL'
    lbl.config(text=company)

def file_list():
    global all_files
    folder_path = r'D:\Process Improvement Project\ED CELL\Outage Reports\Report as per our format\{}\{}'
    all_files = glob.glob(folder_path.format(company,'*.xlsx'))
    remove_updated_files()

def remove_updated_files():
    global files
    global column
    try:
        wb_main = excel.Workbooks.Open(main_file)
    except:
        wb_main = excel.Workbooks(main_file.split('\\')[-1])
    ws_main = wb_main.Worksheets('python_backup')
    if company == 'BRPL':
        column = 'A'
    elif company == 'BYPL':
        column = 'B'
    rowno = ws_main.Range(f'{column}9999').End(3).Row
    # Already_updated_file = ws_main.Range(f'{column}2:{column}{rowno}').Value
    Already_updated_file = [ ws_main.Range(f'{column}{r}').Value for r in range(2,rowno+1) ]
    files = [ file for file in all_files if file not in Already_updated_file ]

def read_files():
    global company_data
    file_list()
    company_data = []
    excel.DisplayAlerts = False
    for file in files:
        try:
            wb = excel.Workbooks.Open(file)
        except:
            wb = excel.Workbooks(file.split('\\')[-1])
        ws = wb.ActiveSheet
        rowno = ws.UsedRange.Rows.Count
        colno = ws.UsedRange.Columns.Count
        date = datetime.strptime(re.findall('\d\d.\d\d.\d\d\d\d',file)[0],'%d.%m.%Y').strftime('%d-%b-%Y')
        # keys = [ ws.Cells(2,col).Value if ws.Cells(3,col).Value == None else ws.Cells(3,col).Value for col in range(1,colno+1) ]
        for row in range(4,rowno+1):
            if company == 'BYPL':
                if ws.Cells(row,1).Value == None:
                    continue
            d = {}
            d['Date'] = date
            for col in range(1,colno+1):
                if ws.Cells(3,col).Value == None:
                    key = ws.Cells(2,col).Value
                else:
                    key = ws.Cells(3,col).Value
                val = ws.Cells(row,col).Value
                d[key] = val
            company_data.append(d)
        wb.Close()
        excel.DisplayAlerts = True
    # excel.Quit()

def correct_time():
    for data in company_data:
        try:
            data['End\nTime (In Date Hrs)'],data['Start\nTime (In Date Hrs)'],data['Time (In Date Hrs)'] = time_converter(
                data['Date'],data['End\nTime (In Date Hrs)'],data['Start\nTime (In Date Hrs)'],data['Time (In Date Hrs)'])
            if isinstance(data['End\nTime (In Date Hrs)'], datetime) and isinstance(data['Start\nTime (In Date Hrs)'],datetime):
                data['Duration'] = data['End\nTime (In Date Hrs)'] - data['Start\nTime (In Date Hrs)']
            else:
                data['Duration'] = ''
            if isinstance(data['Start\nTime (In Date Hrs)'], datetime) and isinstance(data['Time (In Date Hrs)'],datetime):
                data['Load Effected Duration'] = data['Time (In Date Hrs)'] - data['Start\nTime (In Date Hrs)']
            else:
                data['Load Effected Duration'] = ''
        except Exception as e:
            print(e.args)
            print(data['Date'],data['Sr. No.'])

def all_in_one_file():
    keys = []
    for data in company_data:
        for key in data.keys():
            if key not in keys:
                keys.append(key)
    if len(keys) > 0:
        try:
            wb_main = excel.Workbooks.Open(main_file)
        except:
            wb_main = excel.Workbooks(main_file.split('\\')[-1])
        ws_main = wb_main.Worksheets(company)
        rowno = ws_main.Range('A9999').End(3).Row
        if rowno == 1:
            for col,key in enumerate(keys,start=1):
                ws_main.Cells(1,col).Value = key
        colno = ws_main.Range('A1').End(2).Column
        for row,data in enumerate(company_data,start=rowno+1):
            for col in range(1,colno+1):
                if ws_main.Cells(1,col).Value in data.keys():
                    ws_main.Cells(row,col).Value = value_converstion(data[ws_main.Cells(1,col).Value])
        ws_main = wb_main.Worksheets('python_backup')
        if company == 'BRPL':
            column = 'A'
        elif company == 'BYPL':
            column = 'B'
        rowno = ws_main.Range(f'{column}9999').End(3).Row +1
        for row,file in enumerate(files,start=0):
            ws_main.Range(f'{column}{rowno+row}').Value = file          
        wb_main.Save()

btn_list =[ {
    'text' : 'Change Company',
    'function' : change_company
},
{
    'text' : 'List files',
    'function' : file_list
},
{
    'text' : 'Read Files',
    'function' : read_files
},
{
    'text' : 'Correct time',
    'function' : correct_time
},
{
    'text' : 'Write_file',
    'function' : all_in_one_file
}
]

window = tk.Tk()
window.title('Report Assistant')
window.geometry('1000x40')

for col,btn in enumerate(btn_list):
    bt = tk.Button(text=btn['text'],command=btn['function'],width=len(btn['text']))
    bt.grid(row=1,column=col)

lbl = tk.Label(text=company)
lbl.grid(row=1,column=col+1)

window.attributes('-topmost',1)
# window.mainloop()