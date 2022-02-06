# 신규생성 : 2022.02.03 김재민
# 개요 : doosanReleasePlan 엑셀 세팅
# 수정 : 2022.02.04 김재민 : 함수 input값 추가

from openpyxl import load_workbook
import findValueLocation
import pandas as pd

def startWriteCell(filePath, rowFr, rowTo, fixColumn, columnFr, columnTo, fixRow, findValue1, findValue2) :
    wb = load_workbook(filePath)
    ws = wb.active

    itemNumber = findValue1
    releaseDate = findValue2

    findValueLocation.startFindValue(filePath, rowFr, rowTo, fixColumn, columnFr, columnTo, fixRow, itemNumber, releaseDate)

    # for rowItem in range(5,133) : # 찾고자 하는 row의 범위
    #     tempItem = ws.cell(row=rowItem, column=3).value
    #     if tempItem == itemNumber :
    #         for columnDate in range(4, 38) : # 찾고자 하는 column의 범위
    #             tempDate = ws.cell(row=4, column=columnDate).value
    #             if tempDate == releaseDate :
    #                 print('write Cell!! %s %s' %(itemNumber,releaseDate))


    wb.close()