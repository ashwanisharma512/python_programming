from _typeshed import WriteableBuffer
import win32com.client as win32



excel = ''
filepath = ''
filename = ''


def ioms_display():
    global excel
    global wb
    global ws
    if excel == '':
        excel = win32.gencache.EnsureDispatch('Excel.Application')
    
    pass

def ioms_update():
    pass

def ioms_create():
    pass

def ioms_grab_data():
    url_ioms = 'http://10.125.64.81/IOMS'
    ioms_username = '41018328'
    ioms_password = 'pass#123'
    url_ioms_report = 'http://10.125.64.81/IOMS/Reports/{}'
    ioms_bd = 'breakdowntmis'
    ioms_psd = 'pshutmis'
    ioms_em = 'eshutmis'
    ioms_intbd = 'internalmis'

    pass
