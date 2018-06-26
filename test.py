# -*- coding: utf8 -*-
import pandas as pd
import math
import glob

extensions = ['*.xls','*.xlsx','*.xlsm']
filenames = []
for extension in extensions:
    filenames.extend(glob.iglob('/Users/sunny/Documents/UROP/PDS/worldbank/**/' + extension, recursive=True))

result = open('result.txt', 'w+', encoding='utf-8')

for filename in filenames:

    #result = open('result.txt', 'w+', encoding='utf-8')
    result.write(filename)
    print(filename)
    
    workbook = pd.ExcelFile(filename)
    for sheet in workbook.sheet_names:

        data = pd.read_excel(filename, sheet)

        stringCount = 0
        numberCount = 0
        count = 0
        for column in data:
            for item in data[column]:
                if count >= 100:
                    break
                elif type(item) == str:
                    stringCount += 1
                    count += 1
                elif type(item) == float or type(item) == int:
                    if not math.isnan(item):
                        numberCount += 1
                        count += 1
                else:
                    numberCount += 1
                    count += 1
                            
        if stringCount > numberCount:
            line = 'Useless ' + sheet
        else:
            line = 'Useful ' + sheet
        result.write(line)
    #result.close()

result.close()
