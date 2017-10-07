from openpyxl import load_workbook
from time import strftime
import code
import os.path
import re
import sys

mylog = os.path.sep.join(['workspace', 'mylog.log'])

inputtype = 'xlsx'
inputfile = os.path.sep.join(['workspace', 'cache', 'orcl-2017.xlsx'])
debug = False

# Factory to return input adapter.
class InputFactory(object):
    @classmethod
    def Get(cls, type):
        if type == 'xlsx':
            return XlsxInput()
        elif type == 'url':
            return UrlInput()
        else:
            Error('Unknown Input type: {0}'.format(type))


class XlsxInput(object):
    def __init__(self):
        pass

    def Load(self, filename):
        if not os.path.isfile(filename):
            Error('Input file does not exist: {0}'.format(filename))
        wb = load_workbook(filename = filename)
        return wb


class Sheet(object):
    def __init__(self, name, sheet):
        self.name = name
        self.sheet = sheet
        self.kvp = {}


def PrepareDirs():
    if not os.path.exists('workspace'):
        os.makedirs('workspace')
    if not os.path.exists(os.path.sep.join(['workspace', 'cache'])):
        os.makedirs(os.path.sep.join(['workspace', 'cache']))

def ParseArgs():
    global debug
    ptr = 1
    while ptr < len(sys.argv):
        if sys.argv[ptr] == '--debug':
            debug = True
        else:
            print 'Invalid argument: {0}'.format(sys.argv[ptr])
            exit(1)
        ptr += 1

def Process(wb):
    # Locate sheets for the 3 statements.

    # Sheet name is limited to 31 characters. Use
    # the A1 element of each sheet as its full name.
    fullname2sheet = {}
    for name in wb.get_sheet_names():
        sheet = wb[name]
        fullname = sheet['A1'].value.lower()
        fullname2sheet[fullname] = sheet

    # Locate consolidated balance sheet.
    cbsnames = []
    for fullname in fullname2sheet.keys():
        if re.match('^.*consolidated balance sheet.*$', fullname) and not re.match('^.*parenthetical.*$', fullname):
            cbsnames.append(fullname)
    if len(cbsnames) == 0:
        Error('We cannot find the consolidated balance sheet.')
    elif len(cbsnames) > 1:
        Error('We find more than one consolidated balance sheets: {0}'.format(', '.join(cbsnames)))
    cbs = Sheet(cbsnames[0], fullname2sheet[cbsnames[0]])

    # Parse consolidated balance sheet.
    for row in cbs.sheet.iter_rows(row_offset=1):
        if row[0].value == None:
            continue
        rowkey = row[0].value.lower()

        ### TODO: remove $, ' '; convert e.g. '9,821' to float
        # Assets.
        if re.match('^.*total current asset.*$', rowkey):
            cbs.kvp['total current assets'] = row[1].value
        elif re.match('^.*total non-current asset.*$', rowkey):
            cbs.kvp['total non-current assets'] = row[1].value
        elif re.match('^.*total asset.*$', rowkey):
            cbs.kvp['total assets'] = row[1].value
        # Liabilities.
        elif re.match('^$', rowkey):
            pass
        elif re.match('^$', rowkey):
            pass
        # Equity.
        elif re.match('^.*total equity.*$', rowkey) is not None:
            cbs.kvp['total equity'] = row[1].value
        # Liabilities and equity.
        elif re.match('^.*total liabilities and equity.*$', rowkey):
            cbs.kvp['total liabilities and equity'] = row[1].value
    print cbs.kvp
    code.interact(local=locals())

def Render(output):
    pass

def Log(msg):
    fid = open(mylog, 'a')
    fid.write('[{0}] {1}\n'.format(strftime('%Y-%m-%d %H:%M:%S'), msg))
    fid.close()

def LogDebug(msg):
    if debug == True:
        Log('DEBUG: ' + msg)

def Error(msg):
    Log('ERROR: ' + msg)
    exit(1)


if __name__ == '__main__':
    PrepareDirs()
    ParseArgs()
    wb = InputFactory.Get(inputtype).Load(inputfile)
    output = Process(wb)
    Render(output)
