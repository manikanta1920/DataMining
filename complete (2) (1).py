from itertools import combinations
import os
import xlrd as myexcel
from xlwt import Workbook
import xlsxwriter
import copy
import math



#filelist=[{"fname" : "Wine_AE_1.xlsx" , "kvalue" : 2 , "cat" : 1 },{"fname" : "Wine_AE_5.xlsx" , "kvalue" : 2 , "cat" : 1 },{"fname" : "Wine_AE_10.xlsx" , "kvalue" : 2 , "cat" : 1 },{"fname" : "Wine_AE_20.xlsx" , "kvalue" : 2 , "cat" : 1 }]
filelist=[f for f in os.listdir(".") if f.endswith('.xlsx')]
for listitem in filelist:
    data = []
    num_cols = 0
    num_rows = 0
    fileName = ""
    def read_excel():
        global num_cols, num_rows, fileName
        fileName = listitem
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
            k = 2
            if(k > 0):
                break

        while True:
            cn_Type = 0
            if cn_Type == 0 or cn_Type == 1 or cn_Type == 2:
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
                    if (cn_Type == 1) or (cn_Type == 2 and is_float(mr[ii])):
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
                    if cn_Type == 1 or (cn_Type == 2 and is_float(mr[ii])):
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

        outFileNmae = "imputed(complete)" + fileName
        saveExcel(outFileNmae, comData)

    def saveExcel(fileName, comData):
        #create workbook
        workbook = xlsxwriter.Workbook(fileName)
        worksheet = workbook.add_worksheet()
        worksheet.write('A1', 'Hello world')
        # add_sheet is used to create sheet.
        for row in range(len(comData)):
            oneData = comData[row]
            for col in range(len(oneData)):
                val = oneData[col]
                worksheet.write(row, col, val)
        workbook.close()
        print("finished")
    read_excel()
    complete_imputation()
