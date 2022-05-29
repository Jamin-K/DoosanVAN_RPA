# 신규생성 : 2022.05.29 김재민
# 개요 : 완료데이터 엑셀 저장 후 해당 파일을 가지고 납품처가 군산공장일 경우 1일에 보내는 수량이 10%==0 이 되도록
# 묶음 처리하는 스크립트
# 수정 : yyyy.mm.dd 김재민 :

import pandas as pd
from openpyxl import load_workbook
import datetime


def Start(defaultPath, todayDate) :
    #input
    # (1) defaultPath = 'C:/Users/KJM/Desktop/DSVAN'
    # (2) todayDate  = 오늘날짜
    #Output : errorFlag

    print('군산 10개 묶음 START!!')

    tempList = []
    errorFlag = 0 # 0일때 error false, 1일때 error true
    filePath = defaultPath + todayDate + '/완료데이터/ReleasePlan.xlsx'

    releaseWorkBook = load_workbook(filePath)
    releaseWorkSheet = releaseWorkBook.active

    # row 범위 탐색 START
    endOfRow = len(releaseWorkSheet['B'])
    startRow = 0
    rowCount = 0

    for i in range(1, endOfRow):
        if (releaseWorkSheet.cell(i, 2).value == '군산공장' and startRow == 0):
            startRow = i
            continue

        if (releaseWorkSheet.cell(i, 2).value == '군산공장' and startRow != 0):
            rowCount = rowCount + 1
            continue
        i = i + 1

    rowFr = startRow
    rowTo = startRow + rowCount + 1
    # row 범위 탐색 END

    # cell 범위 탐색 START
    endOfCol = 15
    startCell = 0
    colCount = 0

    datetimeToday = datetime.datetime.strptime(todayDate, '%Y%m%d')

    tempDate = datetime.datetime.strftime(datetimeToday + datetime.timedelta(days=0), '%m/%d')
    for i in range(6, endOfCol):
        if(releaseWorkSheet.cell(4, i).value == tempDate):
            startCell=i
        else:
            colCount = colCount + 1
    colFr = startCell

    startCell = 0
    tempDate = datetime.datetime.strftime(datetimeToday + datetime.timedelta(days=5), '%m/%d')

    for i in range(colFr, endOfCol):
        if (releaseWorkSheet.cell(4, i).value == tempDate):
            startCell = i
            break
    colTo = startCell

    print('rowFr : ', rowFr)
    print('rowTo : ', rowTo)
    print('colFr : ', colFr)
    print('colTo : ', colTo)

    # cell 범위 탐색 END










