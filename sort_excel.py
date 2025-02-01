# -*- coding: utf-8 -*-
"""
Created on Sat Dec 25 17:39:56 2021

@author: JEEVAN
"""

import pandas as pd
xl = pd.ExcelFile("sports_test.xlsx")
df = xl.parse("Sheet")
df = df.sort_values(by=["trunk_no"])
df=df.sort_values(by=["total"],ascending=False)
writer = pd.ExcelWriter('sports_test_sorted.xlsx')
df.to_excel(writer,sheet_name='Sheet',columns=['trunk_no','name','university','category',
                                               'attempt1','attempt2','attempt3','max1',
                                               'attempt21','attempt22','attempt23','max2','total',
                                               'next_weight'],
            index=False)
writer.save()
writer.close()
