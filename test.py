# -*- coding: utf8 -*-
import pandas as pd
import math
import glob
from distutils.util import strtobool

extensions = ['*.xls','*.xlsx','*.xlsm']
filenames = []
folder = input('(Example) /Users/sunny/Documents/UROP/\nSearch Directory:\n')
recur = input('Search dirctory recursively? (True/False))\n')
if bool(strtobool(recur)):
    folder = folder + '**/'
for extension in extensions:
    filenames.extend(glob.iglob(folder + extension, recursive = bool(strtobool(recur))))

result = open('result.txt', 'w+', encoding='utf-8')

for filename in filenames:

    result.write(filename + '\n')
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

        if stringCount >= numberCount or count == 0:
            line = 'Useless ' + sheet + '\n'
        else:
            line = 'Useful ' + sheet + '\n'
        result.write(line)

result.close()
