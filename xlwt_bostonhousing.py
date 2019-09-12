#
# xlwt_bostonhousing.py
#
import sys
# from urllib2 import urlopen
from xlwt import Workbook, easyxf, Formula

def doxl():
    '''Read the boston_corrected.txt file based on
       Harrison, David, and Daniel L. Rubinfeld, "Hedonic Housing Prices
       and the Demand for Clean Air," Journal of Environmental Economics
       and Management, Volume 5, (1978), write to an excel spreadsheet .
       '''
    URL = 'http://lib.stat.cmu.edu/datasets/boston_corrected.txt'
    try:
        # For Python 3.0 and later
        from urllib.request import urlopen
    except ImportError:
        # Fall back to Python 2's urllib2
        from urllib2 import urlopen

    try:
        fp = urlopen(URL)
    except:
        print ('Failed to download %s' % URL)
        sys.exit(1)
    lines = fp.readlines()

    wb = Workbook()
    ws = wb.add_sheet('Housing Data')
    ulstyle = easyxf('font: underline single')
    r = 0
    for line in lines:
        tokens = line.decode('cp1250').strip().split('\t')
        if len(tokens) != 21:
            continue
        for c,t in enumerate(tokens):
            for dtype in (int,float):
                try:
                    t = dtype(t)
                except:
                    pass
                else:
                    break
            ws.write(r,c+1,t)
        if r == 0:
            hdr = tokens
            ws.write(r,0,'MAPLINK')
        else:
            d = dict(zip(hdr,tokens))
            link = 'HYPERLINK("http://maps.google.com/maps?q=%s,+%s+(Observation+%s)&hl=en&ie=UTF8&z=14&iwloc=A";"MAP")' % (d['LAT'],d['LON'],d['OBS.'])
            ws.write(r,0,Formula(link),ulstyle)

        r += 1
    wb.save('bostonhousing.xls')
    print ('Wrote bostonhousing.xls')

if __name__ == "__main__":
    doxl()
