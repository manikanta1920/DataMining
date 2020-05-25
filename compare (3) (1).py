import xlrd as myexcel
import os
import copy
import math
import xlsxwriter
imputed_data = []
oriData = []

new_file=open("report.txt",mode="w")

filelist=[f for f in os.listdir(".") if f.startswith('imp')]
for listitem in filelist:
    
    def read_ori_data():
        fileName = "DERM.xlsx"
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
        fileName = listitem
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
                    imputed_data.append(rr)
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
            cn_Type = 1
            if cn_Type == 0 or cn_Type == 1 or cn_Type == 2:
                break

        if rowCount1 != rowCount2:
            print("Invalid Data ....")
            return
        originalAf = 0
        diffAf = 0
        totalVE = 0
        totalAE = 0

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
                if cn_Type == 1 or (cn_Type == 2 and is_float(oriValue)):
                    oriValue = float(oriValue)
                    impValue = float(impValue)
                    originalAf += (oriValue * oriValue)
                    diffAf += math.pow((impValue - oriValue), 2)
                elif cn_Type == 0:
                    totalVE += 1
                    if str(oriValue) != str(impValue):
                        totalAE += 1
        nrms = 0
        if(originalAf > 0):
            nrms = math.sqrt(diffAf) / math.sqrt(originalAf)
        ae = 0
        if(totalVE > 0):
            ae = totalAE / totalVE
        if cn_Type == 0:
            print("AE : %s" % ae)
            new_file.write( listitem + "AE : %s" % ae + "\n")
        elif cn_Type == 1 :
            print("NRMS : %s" % nrms)
            new_file.write( listitem + "NRMS : %s" % nrms + "\n")
        else:
            print("NRMS : %s" % nrms)
            new_file.write( listitem + "NRMS : %s" % nrms + "\n")
            print("AE : %s" % ae)
            new_file.write( listitem + "AE : %s" % ae + "\n")

    read_ori_data()
    read_imputed_data()
    compare()
    
new_file.close()
