from functools import partial
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox
from functools import partial

filepath = r"D:\Process Improvement Project\Training Programs\{}"
main_file = 'BRPL_Training_Programs.xlsm'
data_base = "Manpower Data-May'21.xlsx"

xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.Visible = True
try:
    wb = xl.Workbooks(main_file)
except:
    wb = xl.Workbooks.Open(filepath.format(main_file))

try:
    wb2 = xl.Workbooks(data_base)
except:
    wb2 = xl.Workbooks.Open(filepath.format(data_base))

ws2 = wb2.ActiveSheet
max_row = ws2.UsedRange.Rows.Count

def col_no(key):
    link = {
        'A' : 1, 'B' : 2,'C': 3, 'D' : 4, 'E' : 5, 'F' : 6, 'G' : 7
    }
    if key in link.keys():
        return link[key]
    alertbox('Select column A to G with text')
    return None

def alertbox(txt):
    messagebox.showwarning('Warning',txt)

def copy_data(row,rowno,ws):
    short = wb.Worksheets('short')
    max_row_short = short.UsedRange.Rows.Count
    for n in range(2,8):
        val = ws2.Cells(row,n).Value
        if n == 3:
            val = str(val).title()
        if n == 7:
            val = str(val).lower()
        if n == 4:
            for rw in range(2,max_row_short+1):
                if val == short.Cells(rw,1).Value:
                    if short.Cells(rw,2).Value != None:
                        val = short.Cells(rw,2).Value
        if n == 5:
            for rw in range(2,max_row_short+1):
                if val == short.Cells(rw,4).Value:
                    if short.Cells(rw,5).Value != None:
                        val = short.Cells(rw,5).Value
            pass
        ws.Cells(rowno,n).Value = val
        
def get_emp_data():
    ws = wb.ActiveSheet
    if type(xl.ActiveCell.Value) == str:
        colno,rowno = col_no(xl.ActiveCell.Address.split('$')[1]),xl.ActiveCell.Address.split('$')[-1]
        if colno:
            data = []
            searchItem = (ws.Cells(rowno,colno).Value).upper()
            for row in range(2,max_row+1):
                if searchItem in (ws2.Cells(row,colno).Value).upper():
                    data.append(row)
            if len(data) > 1:
                data_table = tk.Tk()
                data_table.title('Select emplyoee')
                a= 0
                for r in data:
                    for c in range(1,8):
                        txt = ws2.Cells(r,c).Value
                        if type(txt) == float:
                            txt = int(txt)
                        lbl = tk.Label(data_table,text=txt)
                        lbl.grid(row=a,column=c-1)
                    btn = tk.Button(data_table,text='Select',command=partial(copy_data,r,rowno,ws))
                    btn.grid(row=a,column=7)
                    a +=1
                data_table.geometry('1000x{}'.format(a*28))
                data_table.attributes('-topmost',1)
                data_table.mainloop()
            elif len(data) == 1:
                copy_data(data[0],rowno,ws)
            else:
                alertbox('No match found')
    elif type(xl.ActiveCell.Value) == float:
        colno,rowno = col_no(xl.ActiveCell.Address.split('$')[1]),xl.ActiveCell.Address.split('$')[-1]
        if colno:
            searchItem = ws.Cells(rowno,colno).Value
            for row in range(2,max_row+1):
                if searchItem == ws2.Cells(row,colno).Value:
                    copy_data(row,rowno,ws)
                    break

btns_details =[ 
    # ( 'By numerical', get_emp_data_by_number ),
    ( 'Get Data', get_emp_data)
]

window = tk.Tk()
window.title("Training Assistant")
window.geometry('500x35')

for c,btns_data in enumerate(btns_details):
    btn = tk.Button(window,text=btns_data[0],bd=5,command=btns_data[1],width=15)
    btn.grid(row=1,column=c)

window.attributes('-topmost',1)
window.mainloop()