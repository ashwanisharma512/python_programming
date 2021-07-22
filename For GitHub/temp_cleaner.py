import tempfile
import os
temp_path = tempfile.gettempdir()
dlist = []
flist = []
for root,dir,file in os.walk(temp_path):
    if len(dir) > 0:
        for d in dir:
            dlist.append(os.path.join(root,d))
    if len(file) > 0:
        for f in file:
            flist.append(os.path.join(root,f))

# remove file
imp_file = []
ext_imp = ['.xlsx','.xls','.xlsm','.docx','.pdf',]

for file in flist:
    if file.split('.')[-1] in ext_imp:
        imp_file.append(file)
    else:
        try:
            os.remove(file)
        except Exception as e:
            print(e.args)

            


