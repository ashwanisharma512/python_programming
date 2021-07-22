import email
import glob
from os import mkdir
import re
import os.path
from datetime import datetime,timedelta
import PyPDF2
from tkinter.constants import ACTIVE, DISABLED
import tkinter as tk
import win32com.client as win32

weekly_file_path = r'D:\Reports\EHV Tripping weekly meeting\{}'
weekly_file = "NEW EHV Tripping weekly format.xlsx"

def last_week():
    today = datetime.now()
    for n in range(4):
        d = today - timedelta(n)
        if d.strftime('%A') == 'Thursday':
            break
    return d

def email_paths():
    last_date = last_week()
    path = r'D:\Reports\Dash Board Report\{year}\{month_} {month} {year}\emails\{p}\*{day} {month_} {year}*'
    dashboard_emails = []
    tripping_emails = []
    for n in range(7):
        dt = last_date - timedelta(n)
        year = dt.strftime('%Y')
        month = dt.strftime('%b')
        month_ = dt.strftime('%m')
        day = dt.strftime('%d')
        p1 = path.format(year=year,month=month,day=day,month_=month_,p='dashboard')
        p2 = path.format(year=year,month=month,day=day,month_=month_,p='tripping')
        dashboard_emails.append(p1)
        tripping_emails.append(p2)
    return dashboard_emails,tripping_emails

def email_path():
    dashboard_emails = []
    tripping_emails = []
    x,y = email_paths()
    for p in x:
        try:
            dashboard_emails.append(glob.glob(p)[0])
        except:
            pass
    for q in y:
        try:
            tripping_emails.append(glob.glob(q)[0])
        except:
            pass
    return dashboard_emails,tripping_emails

def extract_mails():
    global last_7pdf
    global last_7trips
    dashboard_emails,tripping_emails = email_path()
    last_7pdf = []
    last_7trips = []
    for mail in dashboard_emails+tripping_emails:
        with open(mail,'rb') as f:
            em_ = f.read()
        em = email.message_from_bytes(em_)
        for part in em.walk():
            fileName = part.get_filename()
            if bool(fileName):
                ext = fileName.split('.')[-1]
                if ext == 'pdf':
                    dt_ = re.findall('\d\d.\d\d.\d\d\d\d',fileName)
                    if len(dt_) > 0:
                        dt_y = dt_[0][-4:]
                        dt_m = dt_[0][3:5]
                        dt_d = dt_[0][:2]
                        fileName = dt_y +' '+dt_m+' '+dt_d+'.'+ext
                        pathx = mail.split('emails')[0]+'pdfs\\'+fileName
                        last_7pdf.append(pathx)
                elif ext == 'xlsx':
                    pathx = mail.split('emails')[0]+'daily tripping details\\'+fileName
                    last_7trips.append(pathx)
                if ext in ['pdf','xlsx']:
                    path_y = pathx.split(fileName)[0]
                    if not os.path.exists(path_y):
                        mkdir(path_y)
                    if not os.path.exists(pathx):
                        with open(pathx,'wb') as f:
                            f.write(part.get_payload(decode=True))
                        print(fileName, 'File Saved..')
                    else:
                        print(fileName, 'File already exsist..')

def getAttachments(reader):
      catalog = reader.trailer["/Root"]
      fileNames = catalog['/Names']['/EmbeddedFiles']['/Names']
      attachments = {}
      if bool(fileNames):
            for f in fileNames:
                if isinstance(f, str):
                    dataIndex = fileNames.index(f) + 1
                    fDict = fileNames[dataIndex].getObject()
                    name = fDict['/F']
                    fData = fDict['/EF']['/F'].getData()
                    attachments[name] = fData
      return attachments

def extract_pdf():
    global loadshedding_list
    loadshedding_list = []
    for file_path in last_7pdf:
        try:
            handler = open(file_path, 'rb')
            reader = PyPDF2.PdfFileReader(handler)
            dictionary = getAttachments(reader)
            folder_path = file_path.split('pdfs')[0]
            for fName, fData in dictionary.items():    
                if re.search('BRPL CABLE REPORT',fName):
                    path = 'Cable Report'
                    if not os.path.exists(folder_path+path):
                        os.mkdir(folder_path+path)
                elif re.search('LOAD SHEDDING',fName):
                    path = 'load shedding'
                    loadshedding_list.append(folder_path + path +'\\' +fName)
                    if not os.path.exists(folder_path+path):
                        os.mkdir(folder_path+path)
                elif re.search('MIS REPORT',fName):
                    path = 'MIS Report'
                    if not os.path.exists(folder_path+path):
                        os.mkdir(folder_path+path)
                elif re.search('ASSET ISSUES',fName):
                    path = 'Asset Issue'
                    if not os.path.exists(folder_path+path):
                        os.mkdir(folder_path+path)
                elif re.search('Leakage',fName):
                    path = 'DC Leakage'
                    if not os.path.exists(folder_path+path):
                        os.mkdir(folder_path+path)
                else:
                    path = 'Extra'
                if not os.path.exists(folder_path+path):
                    os.mkdir(folder_path+path)
                if os.path.exists(folder_path + path +'\\' +fName):
                    print("File already Exsist : ",fName)
                else:
                        with open(folder_path + path +'\\' +fName, 'wb') as outfile:                                              
                            outfile.write(fData)
                        print("File extracted :",fName)
        except:
            print("Error in file :",file_path)

def open_weekly_excel():
    global excel,wb,ws
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        wb = excel.Workbooks(weekly_file)
    except:
        wb = excel.Workbooks.Open(weekly_file_path.format(weekly_file))
    ws = wb.ActiveSheet

def copy_mu_loss():
    global mu_list
    Mus_Name = ['GENERATION','BRPL','TRANSCO','NDPL']
    mu_list = []
    for loadshed in loadshedding_list:
        d = {}
        nos = 0
        wbx = excel.Workbooks.Open(loadshed)
        wsx = wbx.ActiveSheet
        data = wsx.UsedRange.Value
        for row in data:
            values = [ x for x in row if x != None]
        for val in values :
            if 'LOAD SHEDDING' in val:
                d['date'] = datetime.strptime(val[0][-10:],'%d.%m.%y').strftime('%d-%m-%Y')
            elif 'TOTAL' in val:
                d[Mus_Name[nos]] = val[1]
                nos +=1
        mu_list.append(d)
        # wbx.Close()
    ws.Range("W3:X20").Clear
    pass


window = tk.Tk()
window.title("EHV assistant")
window.geometry('180x500')

btn_extract_mail = tk.Button(window,text='Extract Mails',bd=5,command=extract_mails,width=40)
btn_extract_mail.pack(padx=5,pady=5)

btn_extract_pdf = tk.Button(window,text='Extract PDFs',bd=5,command=extract_pdf,width=40)
btn_extract_pdf.pack(padx=5,pady=5)

btn_open_xl = tk.Button(window,text='Open Excel',bd=5,command=open_weekly_excel,width=40)
btn_open_xl.pack(padx=5,pady=5)

btn_copy_mu = tk.Button(window,text='Copy MUs',bd=5,command=copy_mu_loss,width=40)
btn_copy_mu.pack(padx=5,pady=5)

window.attributes('-topmost',1)

window.mainloop()