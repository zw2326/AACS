import code
import FileSelector
import os
import pandas as pd
import re

class Statement(object):
    ''' Instance represents a collection of supported sheets for a single financial statement. '''

    # Supported targets:
    # CBS-consolidated balance sheets
    # CSI-consolidated statements of income
    # CSCF-consolidated statements of cash flows
    supported = ['CBS', 'CSI', 'CSCF']

    def __init__(self, inputFile):
        # Map target (e.g. 'CBS') to {'sheet': sheet (DataFrame), 'origName': original name in the Excel (String)}.
        self.sheets = {x: None for x in Statement.supported}
        self.inputFile = inputFile
        xl = pd.ExcelFile(inputFile)

        # Map name to sheet.
        for origName in xl.sheet_names:
            sheet = xl.parse(origName)

            # Sheet name in Excel is truncated to 31 characters.
            # Use the A1 element of the sheet as its real name.
            realName = sheet.columns[0].lower()
            # Compare real name against each supported target.
            for target in Statement.supported:
                if self.__IsMatch__(target, realName):
                    if self.sheets[target] is not None:
                        raise Exception('Multiple {0} sheets found: {1}, {2}'.format(target, self.sheets[target]['origName'], origName))
                    self.sheets[target] = {'sheet': sheet, 'origName': origName}

        # Check if all supported targets are found.
        for target in self.sheets.keys():
            if self.sheets[target] is None:
                raise Exception('{0} sheet is not found'.format(target))

    # Given a supported target, check if the sheet is the target.
    def __IsMatch__(self, target, sheetName):
        if target == 'CBS':
            return re.match('^.*consolidated balance sheet.*$', sheetName) and not re.match('^.*parenthetical.*$', sheetName)
        elif target == 'CSI':
            return re.match('^.*consolidated statements of income.*$', sheetName) and not re.match('.*comprehensive.*$', sheetName)
        elif target == 'CSCF':
            return re.match('^.*consolidated statements of cash flows.*$', sheetName)
        else:
            raise Exception('Unsupported target: {0}'.format(target))

    def GetSheet(self, target):
        if target not in self.sheets.keys():
            raise Exception('Unsupported target {0}'.format(target))
        return self.sheets[target]['sheet']

    def GetOrigName(self, target):
        if target not in self.sheets.keys():
            raise Exception('Unsupported target {0}'.format(target))
        return self.sheets[target]['origName']


class StatementParser(object):
    @classmethod
    def Parse(cls, inputFiles, outputDir):
        # For each sheet, manage major and minor indices separately.
        all = {target: {'major': pd.DataFrame(), 'minor': pd.DataFrame()} for target in Statement.supported}
        for inputFile in inputFiles: # Parse each input file.
            s = Statement(inputFile)
            for target in Statement.supported: # Parse each target in the file.
                StatementParser.__ParseTarget__(target, s.GetSheet(target), all)

    @classmethod
    # Given a sheet, parse the sheet using the target syntax.
    def __ParseTarget__(cls, target, sheet, ret):
        if target == 'CBS':
            StatementParser.__ParseCBS__(sheet, ret)
        elif target == 'CSI':
            pass
        elif target == 'CSCF':
            pass

    @classmethod
    def __ParseCBS__(cls, sheet, ret):
        code.interact(local = locals())
        for row in cbssheet.sheet.iter_rows(row_offset=1):
            if row[0].value == None:
                continue
            rowkey = row[0].value.lower()

            ### TODO: remove $, ' '; convert e.g. '9,821' to float
            ### TODO: save original row key
            ### TODO: do self check
            # Assets.
            if re.match('^.*total current assets.*$', rowkey):
                cbssheet.k2v['total current assets'] = row[1].value
                cbssheet.k2k['total current assets'] = rowkey
            elif re.match('^.*total non-current assets.*$', rowkey):
                cbssheet.k2v['total non-current assets'] = row[1].value
                cbssheet.k2k['total non-current assets'] = rowkey
            elif re.match('^.*total assets.*$', rowkey):
                cbssheet.k2v['total assets'] = row[1].value
                cbssheet.k2k['total assets'] = rowkey
            # Liabilities.
            elif re.match('^.*total current liabilities.*$', rowkey):
                cbssheet.k2v['total current liabilities'] = row[1].value
                cbssheet.k2k['total current liabilities'] = rowkey
            elif re.match('^.*total non-current liabilities.*$', rowkey):
                cbssheet.k2v['total non-current liabilities'] = row[1].value
                cbssheet.k2k['total non-current liabilities'] = rowkey
            elif re.match('^.*total liabilities.*$', rowkey) and not re.match('^.*equity.*$', rowkey):
                cbssheet.k2v['total liabilities'] = row[1].value
                cbssheet.k2k['total liabilities'] = rowkey
            # Equity.
            elif re.match('^.*total.*equity.*$', rowkey) and not re.match('^.*liabilities.*$', rowkey):
                cbssheet.k2v['total equity'] = row[1].value
                cbssheet.k2k['total equity'] = rowkey
            # Liabilities and equity.
            elif re.match('^.*total liabilities and.*equity.*$', rowkey):
                cbssheet.k2v['total liabilities and equity'] = row[1].value
                cbssheet.k2k['total liabilities and equity'] = rowkey


def ParseArgs():
    ret = {}
    return ret


if __name__ == '__main__':
    args = ParseArgs()
    StatementParser.Parse([r'C:\Users\mypc\Desktop\AACS\workspace\cache\GOOG\statement\GOOG-2015-4-10K.xlsx'], 'aa')
    '''
    if args['inputDir'] != None: # Load files from input dir.
        args['inputFiles'] = ...
    if args['filter'] != None: # Do filtering.
        args['inputFiles'] = ...
    StatementParser.Parse(args['inputFiles'], args['outputDir'])
    '''