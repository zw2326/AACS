import pandas as pd
import code

class FileLoader(object):
    syncTypes = {'no': 1, 'latest': 2, 'full': 3}
    syncLimit = '2010'
    urlTemplate = ''

    def __init__(self):
        self.syncType = FileLoader.syncTypes['no']

    # Get all files that satisfy a given filter for a symbol.
    def GetFiles(self, symbol, filter=None):
        allFiles = self.__LoadAllCandidates__(symbol)
        return self.__Filter__(allFiles, filter)

    # Get all files for a symbol.
    def __LoadAllCandidates__(self, symbol):
        hashFnTuples = self.__LoadCachePage__(symbol)
        if self.syncType == FileLoader.syncTypes['no']:
            # No sync, return all cached files for this symbol. No op.
            pass
        elif self.syncType == FileLoader.syncTypes['latest']:
            # Only download files newer than the latest cached file.
            hashFnTuples = self.__DownloadLatest__(symbol, hashFnTuples)
        else:
            # Disregard cached files. Download all from scratch.
            hashFnTuples = self.__DownloadAll__(symbol)
        return map(lambda x: x[1], hashFnTuples)

    # Filter files.
    def __Filter__(self, allFiles, filter):
        ret = allFiles
        return allFiles

    # Download files newer than the latest cached one and update cache page.
    def __DownloadLatest__(self, symbol, hashFnTuples):
        latest = hashFnTuples[-1]
        url = FileLoader.urlTemplate.format(...)
        html = urllib.open(url).read()
        tuples = self.__ExtractLink__(html)
        ### TODO: get next page as well if all first page files are new!
        for tuple in tuples.reverse():
            if tuple[1] <= latest[1]:
                # Older than most recent cache file or limit.
                continue
            urllib.download(tuple[3], 'XXXfilename')
            hashFnTuples.append((tuple[0], 'XXXfilename'))
        self.__UpdateCachePage__(symbol, hashFnTuples)
        return hashFnTuples

    # Download all files and update cache page.
    def __DownloadAll__(self):
        pass

    # Extract file type, date, link and hash from HTML page.
    def __ExtractLink__(self, html):
        ret = [] # [(type, date, link, hash)]
        return ret

    # Update cache page.
    def __UpdateCachePage__(self, symbol, hashFnTuples):
        pass




if __name__ == '__main__':
    xl = pd.ExcelFile(u'workspace\cache\GOOG-2017-2-10Q.xlsx')
    code.interact(local = locals())