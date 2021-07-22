from win32com.client import Dispatch, constants

xl = Dispatch('Excel.Application')

data = {}

xl.Range(row=2,column=5).value


def row_col(x,y):
    alphabet = 'abcdefghijklmnopqrstuvwxyz'
    if x > 26:
        p1 = int(x/26)
        p2 = x%26 - 1
        col = '{}{}'.format(alphabet[p1],alphabet[p2]).upper()
    else:
        col = '{}'.format(alphabet[x]).upper()
    return '{}{}'.format(col,y)