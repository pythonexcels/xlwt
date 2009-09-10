

import sys
import re
from xlwt import Workbook, easyxf

def doxl():
    '''Read raw account number and name strings, separate the data and
       write to an excel spreadsheet.  Properly capitalize the account
       names and mark cells with no account number as 99999 with red fill
       '''
    try:
        fp = open("hspdata.txt")
    except:
        print 'Failed to open hspdata.txt'
        sys.exit(1)
    lines = fp.readlines()

    nameandnum = re.compile(r'(\d+)\s*(.*)\s*')
    wb = Workbook()
    wsraw = wb.add_sheet('Raw Data')
    ws = wb.add_sheet('Account List')
    ws.write(0,0,'Account Number')
    ws.write(0,1,'Account Name')
    ws.col(0).width = len('Account Number') * 256
    ws.col(1).width = max([len(l) for l in lines]) * 256
    r = 1
    
    for line in lines:
        wsraw.write(r,0,line.strip())
        m = nameandnum.match(line)
        if m:
            ws.write(r,0,int(m.group(1)))
            ws.write(r,1,' '.join([w.capitalize() for w in m.group(2).split()]))
        else:
            ws.write(r,0,99999,easyxf('pattern: pattern solid, fore_colour red;'))
            ws.write(r,1,' '.join([w.capitalize() for w in line.split()]))
        r += 1
    wb.save('accounts.xls')
    print 'Wrote accounts.xls'

if __name__ == "__main__":
    doxl()
