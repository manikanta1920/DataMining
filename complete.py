#python -m pip install -U pip setuptools
from itertools import combinations

import xlrd as myexcel
import copy
import math
import csv

data = []
num_cols = 0
num_rows = 0
fileName = ""

def read_excel():
    global num_cols, num_rows, fileName
    fileName = input("enter file name with extension:")
    print("reading file....")
    try:
        with myexcel.open_workbook(fileName) as book:
            first_sheet = book.sheet_by_index(0)
            num_cols = first_sheet.ncols
            num_rows = first_sheet.nrows
            for row in range(first_sheet.nrows):
                cells = first_sheet.row_slice(row, 0, num_cols)
                rr = []
                for cell in cells:
                    rr.append(cell.value)
                data.append(rr)
            print("success to read file!")
    except IOError:
        print("Can't open file")

def has_missing(row):
    """
    check if row contains missing value
    :param row: one record data
    :return: if contains missing value then return True, else return False

    example : row = ['0.34','','0.56','-0.5','','0.12'] return True
              row = ['0.34','0.1','0.56','-0.5','0.43','0.12'] return False
    """
    for d in row:
        if d == '':
            return True
    return False

def is_float(str):
    """
    Check if data is float data
    :param str: data
    :return: if data is float data then return True, else False

    example :   data = '2.334' return True
                data = '2.3fg34' return False
                data = 'fg2.334' return False
                data = '12' return True
    """
    try:
      width = float(str)
      return  True
    except ValueError:
      return False

def complete_imputation():
    global num_cols, num_rows
    k = 0
    while True:
        k = int(input("Enter k (k > 0) = "))
        if(k > 0):
            break

    while True:
        cn_Type = int(input("Choose Categorical (0) or Numerical (1): "))
        if cn_Type == 0 or cn_Type == 1:
            break

    print("complete_imputation (calculating ...)")
    comData = copy.deepcopy(data)
    missingData = []
    observedData = []
    for row in comData: # split M (missing data set) and C (completed data set ) from input data set
        if(has_missing(row)):
            missingData.append(row)
        else:
            observedData.append(row)
    calcnt = 0
    totalCnt = len(missingData)
    for mr in missingData:
        calcnt += 1
        print("%s/%s..." % (calcnt, totalCnt))
        minRows = []
        minDDs = []
        for obr in observedData:
            dd = 0
            for ii in range(len(mr)):
                if(mr[ii] == ''):
                    continue
                # calculate d(xi,xj)
                #if(is_float(mr[ii])):
                if cn_Type == 1:
                    dd += math.pow((mr[ii] - obr[ii]), 2)
                else:
                    if(mr[ii] == obr[ii]):
                        dd = dd + 0
                    else:
                        dd = dd + 1
            # update Ki
            if len(minRows) < k:
                minRows.append(obr)
                minDDs.append(dd)
            else:
                for ii in range(k):
                    if(dd >= minDDs[ii]):
                        continue
                    minDDs[ii] = dd
                    minRows[ii] = obr
                    break
        #imputate missing values
        for ii in range(len(mr)):
            if mr[ii] == '':
                #if(is_float(minRows[0][ii])):
                if cn_Type == 1:
                    dd = 0
                    for jj in range(len(minRows)):
                        dd += minRows[jj][ii]
                    mr[ii] = dd / k
                else:
                    maxID = 0
                    maxCnt = 0
                    for jj in range(len(minRows)):
                        cnt = 0
                        for kk in range(len(minRows)):
                            if(minRows[jj][ii] == minRows[kk][ii]):
                                cnt += 1
                        if(cnt > maxCnt):
                            maxCnt = cnt
                            maxID = jj
                    mr[ii] = minRows[maxID][ii]
    outFileNmae = fileName +  "(complete).csv"
    with open(outFileNmae, mode = 'w') as csvFile:
        csv_writer = csv.writer(csvFile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        for row in comData:
            csv_writer.writerow(row)
    print("completed")
    csvFile.close()

read_excel()
complete_imputation()
