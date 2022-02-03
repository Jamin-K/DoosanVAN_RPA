# 신규생성 : 2022.02.03 김재민
# 개요 : doosanReleasePlan 엑셀 세팅
# input : PlanFilePath, 품번, 납기일, 수량
# range(C5[3,5], C130[3,130]) find itemNumber

from openpyxl import load_workbook
import pandas as pd

def startWriteCell(planFilePath, itemNumber, releaseDate) :
    wb = load_workbook(planFilePath)
    ws = wb.active

    for rowItem in range(5,133) :
        tempItem = ws.cell(row=rowItem, column=3).value
        if tempItem == itemNumber :
            for columnDate in range(4, 38) :
                tempDate = ws.cell(row=4, column=columnDate).value
                if tempDate == releaseDate :
                    print('write Cell!! %s %s' %(itemNumber,releaseDate))


    wb.close()