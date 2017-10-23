import pandas as pd
import code

if __name__ == '__main__':
    xl = pd.ExcelFile(u'workspace\cache\GOOG-2017-2-10Q.xlsx')
    code.interact(local = locals())