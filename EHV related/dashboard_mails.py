import email
import glob
from os import mkdir
import re
import os.path
from datetime import datetime,timedelta

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
        year = '{}'.format(dt.strftime('%Y'))
        month = '{}'.format(dt.strftime('%b'))
        month_ = '{}'.format(dt.strftime('%m'))
        day = '{}'.format(dt.strftime('%d'))
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
        dashboard_emails.append(glob.glob(p)[0])
    for q in y:
        tripping_emails.append(glob.glob(q)[0])
    return dashboard_emails,tripping_emails

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

with open(r'D:\Process Improvement Project\python_programming\EHV related\template\list.txt','w') as f:
    f.write('{}'.format(last_7pdf+last_7trips))


