# 신규생성 : 2022.02.03 김재민
# 개요 : doosanReleasePlan 엑셀 세팅
# 수정 : 2022.02.04 김재민 : 함수 input값 추가
#       2022.02.19 김재민 : 클래스 생성, 함수 호출 및 getter 접근 #001
#       2022.02.19 김재민 : 얻은 좌표에 Write OrderCount 및 엑셀 저장 #002
#       2022.03.02 김재민 : ReleasePlan에 쓰기 실패한 데이터 리스트 기록 in writeFaieldlist.xlsx #003

from openpyxl import load_workbook
import findValueLocation
from findValueLocation import Coordinate
import pandas as pd
import CheckWorkDays


def startWriteCell(filePath, rowFr, rowTo, fixColumn, columnFr, columnTo, fixRow, findValue1, findValue2, orderCount,
                   orderNumber, semiOrderNumber) :
    wb = load_workbook(filePath)
    ws = wb.active
    cordinate = Coordinate() #001
    itemNumber = findValue1

    # 공휴일 및 주말 체크 함수 START
    findValue2 = CheckWorkDays.checkHolidays(findValue2)
    # 공휴일 및 주말 체크 함수 END

    releaseDate = findValue2 #yyyy/mm/dd 형태로 받아옴

    cordinate.startFindValue(filePath, rowFr, rowTo, fixColumn, columnFr, columnTo, fixRow, itemNumber, releaseDate, ws) #001
    if(cordinate.getRow() != 0 and cordinate.getCol() != 0) : #001
        print('Write Cordinate!! : %d %d' % (cordinate.getRow(), cordinate.getCol())) #001
        print('Write Cordinate of OrderCount : %d' % orderCount) #001
        ws.cell(cordinate.getRow(), cordinate.getCol(), orderCount) #002
    elif(cordinate.getRow() == 0 and cordinate.getCol() == 0) : #003
        print('Write Failed 데이터 엑셀 기록')


    wb.save(filePath) #002
    wb.close()


