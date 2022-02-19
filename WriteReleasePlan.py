# 신규생성 : 2022.02.03 김재민
# 개요 : doosanReleasePlan 엑셀 세팅
# 수정 : 2022.02.04 김재민 : 함수 input값 추가
#       2022.02.19 김재민 : 클래스 생성, 함수 호출 및 getter 접근 #001
#       2022.02.19 김재민 : 얻은 좌표에 Write OrderCount 및 엑셀 저장 #002

from openpyxl import load_workbook
import findValueLocation
from findValueLocation import Coordinate
import pandas as pd

def startWriteCell(filePath, rowFr, rowTo, fixColumn, columnFr, columnTo, fixRow, findValue1, findValue2, orderCount) :
    wb = load_workbook(filePath)
    ws = wb.active
    cordinate = Coordinate() #001
    itemNumber = findValue1
    releaseDate = findValue2

    cordinate.startFindValue(filePath, rowFr, rowTo, fixColumn, columnFr, columnTo, fixRow, itemNumber, releaseDate, ws) #001
    if(cordinate.getRow() != 0 and cordinate.getCol() != 0) : #001
        print('Write Cordinate!! : %d %d' % (cordinate.getRow(), cordinate.getCol())) #001
        print('Write Cordinate of OrderCount : %d' % orderCount) #001
        ws.cell(cordinate.getRow(), cordinate.getCol(), orderCount) #002


    #wb.save(filePath) #002
    wb.close()


