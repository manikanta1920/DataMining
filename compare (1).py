import xlrd as myexcel
import copy
import math
import csv

imputed_data = []
oriData = []

def read_ori_data():
    fileName = input("enter original data file name with extension:")
    print("reading original file....")
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
                oriData.append(rr)
            print("success to read original data file!")
    except IOError:
        print("Can't open original file")

def read_imputed_data():
    filename = input("Enter imputed data file name with extension :")
    try:
        with open(filename, 'r') as csvfile:
            csvdata = csv.reader(csvfile)
            for x in csvdata:
                if len(x) > 0 :
                    imputed_data.append(x)
            csvfile.close()
            print("success to read imputed data file!")
    except IOError:
        print("Can't open imputed file")

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

def compare():
    rowCount1 = len(oriData)
    rowCount2 = len(imputed_data)
    while True:
        cn_Type = int(input("Choose Categorical (0) or Numerical (1): "))
        if cn_Type == 0 or cn_Type == 1:
            break

    if rowCount1 != rowCount2:
        print("Invalid Data ....")
        return
    originalAe = 0
    imputedAe = 0
    diffAe = 0

    print("calculating...")
    for ii in range(rowCount1):
        oriRow = oriData[ii]
        impRow = imputed_data[ii]
        colCount1 = len(oriRow)
        colCount2 = len(impRow)
        if(colCount1 != colCount2):
            print("Invalid Data ....")
            return
        for jj in range(colCount1):
            oriValue = oriRow[jj]
            impValue = impRow[jj]
            #if(is_float(oriValue)):
            if cn_Type == 1 :
                oriValue = float(oriValue)
                impValue = float(impValue)
                originalAe += (oriValue * oriValue)
                imputedAe += (impValue * impValue)
                diffAe += math.pow((impValue - oriValue), 2)
            else:
                originalAe += 1
                imputedAe += 1
                if str(oriValue) != str(impValue):
                    diffAe += 1

    oAE = math.sqrt(originalAe)
    iAE = math.sqrt(imputedAe)
    nrms = math.sqrt(diffAe) / math.sqrt(originalAe)
    if cn_Type == 0:
        print("Original Data AE : %s" % oAE)
        print("Imputed Data AE : %s" % iAE)
    else :
        print("NRMS : %s" % nrms)

read_ori_data()
read_imputed_data()
compare()