#<class 'str'>, <class 'NoneType'>, <class 'pywintypes.datetime'>, <class 'float'>
from datetime import datetime, timedelta
import re
import pytz
def time_str(dt,t):
    hypen_date = re.findall('\d\d-\d\d-\d\d\d\d \d\d:\d\d:\d\d',t)
    slash_date = re.findall('\d\d/\d\d/\d\d\d\d \d\d:\d\d:\d\d',t)
    if t.upper() in ['WIP','PENDING']:
        return 'Pending'
    if len(hypen_date) > 0:
        return datetime.strptime(hypen_date[0],'%d-%m-%Y %H:%M:%S')
    if len(slash_date) > 0 :
        return datetime.strptime(slash_date[0],'%d/%m/%Y %H:%M:%S')
    return 'Str Converstion Error'

def time_nonetype(dt,t):
    return ''

def time_datetime(dt,t):
    return t

def time_float(dt,t):
    # time_converter
    tmp_d = datetime.strptime(dt,'%d-%b-%Y')
    # tmp_d.tzinfo = pytz.timezone('Asia/Kolkata')
    if t > 0:
        # nothing new
        return tmp_d + timedelta(t+0.229166666666667)
    elif t == 0:
        return tmp_d + timedelta(1+0.229166666666667)


switchcase = {
    'str' : time_str,
    'float' : time_float,
    'pywintypes.datetime' : time_datetime,
    'NoneType' : time_nonetype,
    'datetime.datetime' : time_datetime
}

def time_converter(dt,*times):
    returnvalue = []
    for t in times:
        key = str(type(t)).split("'")[1]
        if key in switchcase.keys():
            returnvalue.append(switchcase[key](dt,t))
        else:
            print(key,dt,t)
            returnvalue.append('Converstion Error')
    return returnvalue

def value_converstion(val):
    if isinstance(val,timedelta):
        mx,_ = divmod(val.seconds,60)
        hrs,mins = divmod(mx,60)
        return f'{hrs}:{mins}'
    else:
        return val




