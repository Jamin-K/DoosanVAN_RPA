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

    wb.close()


