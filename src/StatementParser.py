import code
import math
import numpy as np
import pandas as pd
import re
import sys

class Schema(object):
    ''' This class documents the set of supported sheets, and the mapping criteria for a number of well-known indices on each sheet. '''

    # For each supported sheet, two components shall be provided:
    # namePattern - the matching and non-matching pattern for the sheet name
    # indexPattern - a list matching & nonmatching patterns to map indices.
    #
    # Supported sheets:
    # CBS - consolidated balance sheets
    # CSI - consolidated statements of income
    # CSCF - consolidated statements of cash flows
    supported = {
        'CBS': {
            'namePattern': ('^.*consolidated balance sheet.*$', '^.*parenthetical.*$'),
            'indexPattern': {
                'total current assets'         : ('^.*total current assets.*$'         , None),
                'total non-current assets'     : ('^.*total non-current assets.*$'     , None),
                'total assets'                 : ('^.*total assets.*$'                 , None),
                'total current liabilities'    : ('^.*total current liabilities.*$'    , None),
                'total non-current liabilities': ('^.*total non-current liabilities.*$', None),
                'total liabilities'            : ('^.*total liabilities.*$'            , '^.*equity.*$'),
                'total equity'                 : ('^.*total.*equity.*$'                , '^.*liabilities.*$'),
                'total liabilities and equity' : ('^.*total liabilities and.*equity.*$', None)
            }
        },
        'CSI': {
            'namePattern': ('^.*(consolidated statements of (earnings|income|operations)|consolidated income statements).*$', '^.*parenthetical.*$'),
            'indexPattern': {}
        },
        # 'CSCF': {}
    }


class Statement(object):
    ''' Instance represents a collection of supported sheets for a single financial statement file. '''

    def __init__(self, inputFile):
        # Map sheet name (e.g. 'CBS') to {'sheet': sheet object (DataFrame), 'origName': original sheet name in the file (String)}.
        self.sheets = {x: None for x in Schema.supported.keys()}
        self.inputFile = inputFile
        xl = pd.ExcelFile(inputFile)

        # Locate all supported sheets in the input file.
        for origName in xl.sheet_names:
            sheet = xl.parse(origName)

            # Sheet name in Excel is truncated to 31 characters. Use the A1 element of the sheet as its real name.
            realName = sheet.columns[0].lower()
            # Compare real name against each supported sheet.
            for target in Schema.supported.keys():
                matchPattern, nonmatchPattern = Schema.supported[target]['namePattern']
                if re.match(matchPattern, realName) and ((not re.match(nonmatchPattern, realName)) if nonmatchPattern is not None else True):
                    if self.sheets[target] is not None:
                        raise Exception('Multiple {0} sheets found: {1}, {2}'.format(target, self.sheets[target]['origName'], origName))
                    self.sheets[target] = {'sheet': sheet, 'origName': origName}
                    break

        # Check if all supported sheets are found.
        # TODO: report all sheets that are not found?
        for target in self.sheets.keys():
            if self.sheets[target] is None:
                raise Exception('{0} sheet is not found'.format(target))

    def GetSheet(self, target):
        if target not in self.sheets.keys():
            raise Exception('Unsupported target {0}'.format(target))
        return self.sheets[target]['sheet']

    def GetOrigName(self, target):
        if target not in self.sheets.keys():
            raise Exception('Unsupported target {0}'.format(target))
        return self.sheets[target]['origName']


class StatementParser(object):
    ''' Class to parse the financial statement files. '''

    debug = True

    @classmethod
    def Parse(cls, inputFiles):
        # Get all column names corresponding to all input files.
        inputFile2ColumnName = StatementParser.__GetColumnNames__(inputFiles)

        # For each supported sheet, maintain a DataFrame for all indices. Each column in the DataFrame corresponds to values from one input file.
        ret = {target: pd.DataFrame(columns=inputFile2ColumnName.keys(), index=Schema.supported[target]['indexPattern'].keys()) for target in Schema.supported.keys()}

        for inputFile in inputFiles: # Parse each input file.
            columnName = inputFile2ColumnName[inputFile]
            s = Statement(inputFile)
            for target in Schema.supported.keys(): # Parse each supported sheet in the file.
                StatementParser.__ParseTarget__(target, s.GetSheet(target), columnName, ret)

        return ret
        # code.interact(local = locals())

    @classmethod
    # Determine the column names to be added to DataFrames for all input files.
    # TODO: extract and return the dates as the column names
    def __GetColumnNames__(cls, inputFiles):
        return {x: x for x in inputFiles}

    @classmethod
    # Given a sheet, parse the sheet using the target syntax.
    def __ParseTarget__(cls, target, sheet, column, ret):
        if target in ['CBS', 'CSI', 'CSCF']:
            StatementParser.__ParseCCC__(target, sheet, column, ret)

    @classmethod
    # Method to parse CBS, CSI and CSCF.
    def __ParseCCC__(cls, target, sheet, column, ret):
        for rownum, row in sheet.iterrows():
            if not np.issubdtype(type(row.iloc[1]), np.number) or math.isnan(row.iloc[1]): # Ignore rows containing non-numeric or NaN value.
                continue

            index = row.iloc[0].lower()
            value = float(row.iloc[1])

            # Compare row index against each supported well-known index for this sheet.
            isMajor = False
            for indexTarget, pattern in Schema.supported[target]['indexPattern'].items():
                matchPattern, nonmatchPattern = pattern
                if re.match(matchPattern, index) and ((not re.match(nonmatchPattern, index)) if nonmatchPattern is not None else True):
                    # Check if there is a duplicate. If yes, check if values are the same.
                    if not math.isnan(ret[target].loc[indexTarget, column]) and ret[target].loc[indexTarget, column] != value:
                        raise Exception("Multiple values for {0} found: {1}, {2}".format(indexTarget, ret[target].loc[indexTarget, column], value))
                    ret[target].loc[indexTarget, column] = value
                    if StatementParser.debug == True:
                        print "{0} mapping (line={1}, value={2}):\n        {3}\n     -> {4}".format(target, rownum, value, row.iloc[0].encode('utf-8'), indexTarget)
                    isMajor = True
                    break
            if not isMajor: # Not a well-known index.
                ret[target].loc[index, column] = value
                if StatementParser.debug == True:
                    print "{0} (line={1}, value={2}):\n        {3}".format(target, rownum, value, index.encode('utf-8'))


def ParseArgs():
    ret = {}
    return ret


if __name__ == '__main__':
    args = ParseArgs()
    if sys.platform.lower().startswith('win'):
        StatementParser.Parse([r'C:\Users\mypc\Desktop\AACS\workspace\cache\GOOG\statement\GOOG-2015-4-10K.xlsx', r'C:\Users\mypc\Desktop\AACS\workspace\cache\GOOG\statement\GOOG-2016-1-10Q.xlsx'])
    else:
        StatementParser.Parse([r'/Users/a2326/Git/AACS/workspace/cache/GOOG/statement/GOOG-2015-4-10K.xlsx'])
    '''
    if args['inputDir'] != None: # Load files from input dir.
        args['inputFiles'] = ...
    StatementParser.Parse(args['inputFiles'])
    '''