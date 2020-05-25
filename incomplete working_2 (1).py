#python -m pip install -U pip setuptools
from itertools import combinations
import os
import xlrd as myexcel
import copy
import math
import xlsxwriter
data = []
num_cols = 0
num_rows = 0
fileName = ""
#filelist=[{"fname" : "Wine_AE_1.xlsx" , "kvalue" : 2 , "cat" : 1 },{"fname" : "Wine_AE_5.xlsx" , "kvalue" : 2 , "cat" : 1 },{"fname" : "Wine_AE_10.xlsx" , "kvalue" : 2 , "cat" : 1 },{"fname" : "Wine_AE_20.xlsx" , "kvalue" : 2 , "cat" : 1 }]
filelist=[f for f in os.listdir(".") if f.endswith('.xlsx')]
for listitem in filelist:
    data = []
    num_cols = 0
    num_rows = 0
    fileName = ""

    def read_excel():
        """
        read data from xlsx file
        :return:
        """
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

    def get_observed_ids(row):
        """
        get index list of observed values from one record data
        :param row: one record data
        :return: index list of observed values

        example : row = ['0.34','','0.56','-0.5','','0.12']
                return [0,2,3,5]
        """
        tt = []
        for i in range(len(row)):
            if(row[i] == ''):
                continue
            tt.append(i)
        return tt

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

    def incomplete_imputation():
        global num_cols, num_rows
        k = 0
        while True:
            k = 2
            if(k > 0):
                break

        while True:
            cn_Type = 2 
            if cn_Type == 0 or cn_Type == 1 or cn_Type == 2:
                break

        print("incomplete_imputation (calculating ...)")
        comData = copy.deepcopy(data)
        finalData = copy.deepcopy(comData)
        missingData = []
        finalMisData = []
        # get M (missing data set) from input data set
        for row in finalData:
            if(has_missing(row)):
                finalMisData.append(row)

        for row in comData:
            if(has_missing(row)):
                missingData.append(row)

        calcnt = 0
        totalCnt = len(missingData)
        for mr in missingData:
            calcnt += 1
            print("%s/%s..." % (calcnt, totalCnt))
            observedIDs = get_observed_ids(mr)
            for jj in range(len(mr)):
                mc = mr[jj]
                if(mc != ''):
                    continue
                minRows = []
                minDDs = []
                oids = copy.deepcopy(observedIDs)
                oids.append(jj)
                for obr in comData:
                    oidsTemp = get_observed_ids(obr)
                    result = all(elem in oidsTemp for elem in oids)
                    if result == False:
                        continue
                    dd = 0
                    for ii in range(len(mr)):
                        if(mr[ii] == ''):
                            continue
                        # calculate d(xi,xj)
                        # if(is_float(mr[ii])):
                        if cn_Type == 1 or (cn_Type == 2 and is_float(mr[ii])):
                            dd += math.pow((mr[ii] - obr[ii]), 2)
                        else:
                            if(mr[ii] == obr[ii]):
                                dd = dd + 0
                            else:
                                dd = dd + 1
                    # update Ki,mj
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
                # imputate missing values
                #if (is_float(mr[jj])):
                if cn_Type == 1 or (cn_Type == 2 and is_float(mr[ii])):
                    dd = 0
                    for rr in minRows:
                        dd += rr[jj]
                    #mr[jj] = dd / k
                    finalMisData[calcnt - 1][jj] = dd / k
                else:
                    maxID = 0
                    maxCnt = 0
                    for tt in range(len(minRows)):
                        cnt = 0
                        for kk in range(len(minRows)):
                            if (minRows[tt][jj] == minRows[kk][jj]):
                                cnt += 1
                        if (cnt > maxCnt):
                            maxCnt = cnt
                            maxID = tt
                    #mr[jj] = minRows[maxID][jj]
                    finalMisData[calcnt - 1][jj] = minRows[maxID][jj]
        outFileNmae = "imputed(incomplete)" + fileName
        saveExcel(outFileNmae, finalData)

    def saveExcel(fileName, comData):
        # create workbook
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
    incomplete_imputation()
