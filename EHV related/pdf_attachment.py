import PyPDF2
import glob
import re
import os.path
from os import path
from datetime import datetime,timedelta

def last_Date():
    today = datetime.now()
    for n in range(4):
        d = today - timedelta(n)
        if d.strftime('%A') == 'Thursday':
            break
    return d

def pdf_paths():
    last_date = last_Date()
    path = r'D:\Reports\Dash Board Report\{year}\{month_} {month} {year}\pdfs\{year} {month_} {day}.pdf'
    li = []
    for n in range(7):
        dt = last_date - timedelta(n)
        year = '{}'.format(dt.strftime('%Y'))
        month = '{}'.format(dt.strftime('%b'))
        month_ = '{}'.format(dt.strftime('%m'))
        day = '{}'.format(dt.strftime('%d'))
        p1 = path.format(year=year,month=month,day=day,month_=month_)
        li.append(p1)
    return li

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

file_list = pdf_paths()

for file_path in file_list:
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
            
