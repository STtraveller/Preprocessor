# -*- coding: utf8 -*-


# import the modules

import pandas as pd
import math
import glob
from distutils.util import strtobool


# get the file list to be searched

extensions = ['*.xls','*.xlsx','*.xlsm']
filenames = []
# folder = input('(Example) /Users/sunny/Documents/UROP/\nSearch Directory:\n')
# if folder[-1] != '/':
#     folder = folder + '/'
folder = './sample spreadsheets/'
recur = input('Search directory recursively? (True/False))\n')
if bool(strtobool(recur)):
    folder = folder + '**/'
for extension in extensions:
    filenames.extend(glob.iglob(folder + extension, recursive = bool(strtobool(recur))))
filenames.sort()


# clean up/ creaste the output text file

result = open('result.txt', 'w+', encoding='utf-8')
result.close()


# start reading files

for filename in filenames:

    result = open('result.txt', 'a+', encoding='utf-8')
    result.write(filename + '\n')
    print(filename)


# main algorithm

    try:
        workbook = pd.ExcelFile(filename)
        for sheet in workbook.sheet_names:

            data = pd.read_excel(filename, sheet)

            stringCount = 0
            numberCount = 0
            count = 0

            for column in data:

                if count >= 100:
                    break

                columnSearch = 0

                for item in data[column]:

                    if count >= 100 or columnSearch >= 10:
                        break
                    elif type(item) == str:
                        stringCount += 1
                        count += 1
                        columnSearch += 1
                    elif type(item) == float or type(item) == int:
                        if not math.isnan(item):
                            numberCount += 1
                            count += 1
                            columnSearch += 1
                    else:
                        numberCount += 1
                        count += 1
                        columnSearch += 1

            if count == 0:
                line = 'Empty ' + sheet + '\n'
            elif stringCount >= numberCount:
                line = 'Useless ' + sheet + '\n'
            else:
                line = 'Useful ' + sheet + '\n'
            result.write(line)
        result.close()

# Error handling

    except:
        print(filename + " has error.\n")
        result.write(filename + " has error.\n")
        result.close()
