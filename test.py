# -*- coding: utf8 -*-


# import the modules

import pandas as pd
import numpy as np
import math
import glob
import random
import os
import shutil
from distutils.util import strtobool


# get the file list to be searched

def getFiles(scope, number):
    extensions = ['*.xls','*.xlsx','*.xlsm']
    filenames = []
    if scope[-1] != '/':
        scope = scope + '/'
    for file in os.listdir(scope):
        if os.path.isdir(scope + file):
            folder = scope + file + '/**/'
            temp = []
            for extension in extensions:
                temp.extend(glob.iglob(folder + extension, recursive = True))
            if number <= 0:
                filenames.extend(temp)
            elif len(temp) < number:
                filenames.extend(random.sample(temp, len(temp)))
            else:
                filenames.extend(random.sample(temp, number))
    filenames.sort()
    return filenames


def getTestList(filename, searchingDir, number):
    things = open(filename, 'w+')
    for item in getFiles(searchingDir, number):
        things.write(item + '\n')
    things.close()


def copyFileToFolder(filelist):
    try:
        shutil.rmtree('./test objects')
    except FileNotFoundError:
        pass
    finally:
        os.mkdir('test objects')
        for file in filelist:
            shutil.copy2(file,'./test objects')

def readList(filename):
    file = open(filename, 'r', encoding='utf-8')
    filelist = file.readlines()
    file.close()
    filelist = [item[:-1] for item in filelist]
    return filelist

def iamlazy():
    copyFileToFolder(readList('things.csv'))

def lazyMain():
    getTestList('things.csv', '/Volumes/SSD/SpreadsheetCorpus', 15)
    iamlazy()
    main()

def main():

    filenames = []
    testObjectDirectory = './test objects/'
    for file in os.listdir(testObjectDirectory):
        filenames.append(testObjectDirectory + file)
    filenames.sort()
    totalSearch = len(filenames)
    searchCount = 0

    # clean up / create the output text file

    result = open('result.csv', 'w+', encoding='utf-8')
    result.write('full path, filename, sheet, processed result, checked result' + '\n')
    result.close()

    # start reading files

    for filename in filenames:

        result = open('result.csv', 'a+', encoding='utf-8')
        print(filename)
        searchCount += 1

    # main algorithm

        try:
            workbook = pd.ExcelFile(filename)
            path, file = os.path.split(filename)
            comma = ', '
            for sheet in workbook.sheet_names:

                data = pd.read_excel(filename, sheet, header = None)

                stringCount = 0
                numberCount = 0
                zeroCount = 0
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
                                if item == 0:
                                    zeroCount += 1
                        else:
                            numberCount += 1
                            count += 1
                            columnSearch += 1

                if count == 0:
                    line = comma.join([filename, file, sheet, 'Empty\n'])
                elif stringCount >= numberCount or zeroCount >= numberCount / 2:
                    line = comma.join([filename, file, sheet, 'Junk\n'])
                else:
                    line = comma.join([filename, file, sheet, 'Useful\n'])
                result.write(line)
            result.close()

# Error handling

        except:
            print(filename + " has error.\n")
            result.write(filename + " has error.\n")
        result.close()

# A fun progress bar

        print('Progress: %.2f' % (searchCount / totalSearch * 100) + '%   ({}/{})'.format(searchCount, totalSearch))
