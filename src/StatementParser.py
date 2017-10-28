import code
import FileSelector
import math
import os
import pandas as pd
import re
import sys

class Schema(object):
    ''' This class documents the set of supported sheets, and the matching criteria for their major indices. '''

    # For each supported sheet, two components must be provided:
    # pattern - the matching and non-matching pattern for the sheet name
    # major - a list of major indices in the sheet, each associated with its matching and non-matching pattern.
    #
    # Supported sheets:
    # CBS - consolidated balance sheets
    # CSI - consolidated statements of income
    # CSCF - consolidated statements of cash flows
    supported = {
        'CBS': {
            'pattern': ('^.*consolidated balance sheet.*$', '^.*parenthetical.*$'),
            'major': {
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
        'CSI': {},
        'CSCF': {}
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
                matchPattern, nonmatchPattern = Schema.supported[target]['pattern']
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
    @classmethod
    def Parse(cls, inputFiles, outputDir):
        # For each supported sheet, maintain a major and a minor index DataFrame.
        # Each column in either of the DataFrame corresponds to values from a input file.
        ret = {target: {'major': pd.DataFrame(index=Schema.supported[target]['major'].keys()), 'minor': pd.DataFrame()}
            for target in Schema.supported.keys()}
        for inputFile in inputFiles: # Parse each input file.
            s = Statement(inputFile)
            for target in Schema.supported.keys(): # Parse each supported sheet in the file.
                StatementParser.__ParseTarget__(target, s.GetSheet(target), ret)

    @classmethod
    # Given a sheet, parse the sheet using the target syntax.
    def __ParseTarget__(cls, target, sheet, ret):
        if target in ['CBS', 'CSI', 'CSCF']:
            StatementParser.__ParseCCC__(target, sheet, ret)

    @classmethod
    def __ParseCCC__(cls, target, sheet, ret):
        # TODO: Add a new column to both major and minor index collection for the new input file.
        ret[target]['major'].add(pd.Series([], index=ret[target]['major'].index))
        for rownum, row in sheet.iterrows():
            if math.isnan(row.iloc[1]): # Ignore rows containing NaN value.
                continue

            index = row.iloc[0].lower()
            value = float(row.iloc[1])

            # Compare row index against each supported major index for this sheet.
            isMatched = False
            for indexTarget, pattern in Schema.supported[target]['major'].items():
                matchPattern, nonmatchPattern = pattern
                if re.match(matchPattern, index) and ((not re.match(nonmatchPattern, index)) if nonmatchPattern is not None else True):
                    if ret.loc[indexTarget][-1] is not None: # TODO
                        raise Exception("Multiple values for {0} found: {1}, {2}".format(indexTarget, ret.loc[indexTarget][-1], value))
                    ret[target]['major'].loc[indexTarget][-1] = value
                    isMatched = True
                    break
            if not isMatched: # A minor index.
                if index in ret[target]['minor'].index: # TODO
                    ret[target]['minor'][index][-1] = value
                else:
                    ret[target]['minor'].append(index, value) # TODO: add index and add value to -1 position!
 
            code.interact(local = locals())


def ParseArgs():
    ret = {}
    return ret


if __name__ == '__main__':
    args = ParseArgs()
    if sys.platform.lower().startswith('win'):
        StatementParser.Parse([r'C:\Users\mypc\Desktop\AACS\workspace\cache\GOOG\statement\GOOG-2015-4-10K.xlsx'], 'aa')
    else:
        StatementParser.Parse([r'/Users/a2326/Git/AACS/workspace/cache/GOOG/statement/GOOG-2015-4-10K.xlsx'], 'aa')
    '''
    if args['inputDir'] != None: # Load files from input dir.
        args['inputFiles'] = ...
    if args['filter'] != None: # Do filtering.
        args['inputFiles'] = ...
    StatementParser.Parse(args['inputFiles'], args['outputDir'])
    '''