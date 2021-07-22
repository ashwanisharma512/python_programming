import win32com.client as win32
import glob,re
from datetime import datetime, time,timedelta
from module1 import time_converter,value_converstion
import tkinter as tk

company = 'BRPL'

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

def change_company():
    global company
    if company == 'BRPL':
        company = 'BYPL'        
    else:
        company = 'BRPL'
    lbl.config(text=company)

def file_list():
    global files
    folder_path = r'D:\Process Improvement Project\ED CELL\Outage Reports\Report as per our format\{}\{}'
    files = glob.glob(folder_path.format(company,'*.xlsx'))

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
    # excel.Quit()

def correct_time():
    for data in company_data:
        # print('{} , {} , {}'.format(data['End\nTime (In Date Hrs)'],data['Start\nTime (In Date Hrs)'],data['Time (In Date Hrs)']),end=' ')
        data['End\nTime (In Date Hrs)'],data['Start\nTime (In Date Hrs)'],data['Time (In Date Hrs)'] = time_converter(
            data['Date'],data['End\nTime (In Date Hrs)'],data['Start\nTime (In Date Hrs)'],data['Time (In Date Hrs)'])
        # print(' --{}, {} , {} , {}'.format(data['Date'],data['End\nTime (In Date Hrs)'],data['Start\nTime (In Date Hrs)'],data['Time (In Date Hrs)']),end=' ')
        # input()
        if isinstance(data['End\nTime (In Date Hrs)'], datetime) and isinstance(data['Start\nTime (In Date Hrs)'],datetime):
            data['Duration'] = data['End\nTime (In Date Hrs)'] - data['Start\nTime (In Date Hrs)']
        else:
            data['Duration'] = ''
        if isinstance(data['Start\nTime (In Date Hrs)'], datetime) and isinstance(data['Time (In Date Hrs)'],datetime):
            data['Load Effected Duration'] = data['Time (In Date Hrs)'] - data['Start\nTime (In Date Hrs)']
        else:
            data['Load Effected Duration'] = ''

def all_in_one_file():
    keys = []
    for data in company_data:
        for key in data.keys():
            if key not in keys:
                keys.append(key)
    if len(keys) > 0:
        wbnew = excel.Workbooks.Add()
        ws = wbnew.ActiveSheet
        for col,key in enumerate(keys,start=1):
            ws.Cells(1,col).Value = key
        for row,data in enumerate(company_data,start=2):
            for col,key in enumerate(keys,start=1):
                ws.Cells(row,col).Value = value_converstion(data[key])
        wbnew.SaveAs(f'{company}.xlsx')

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