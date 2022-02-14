# 신규생성 : 2022.02.04 김재민
# 개요 : 가로 세로 2개의 값을 입력받아 해당 셀의 위치를 반환하는 함수
# 수정 : 2022.02.14 김재민 : 엑셀시트에 해당 품명이 존재 여부를 체크하는 로직 추가 #001

from openpyxl import load_workbook
import time

# 변수선언 START
checkFindValue = False #001
checkFindValueItem = False #001
# 변수선언 END

def startFindValue(filePath, rowFr, rowTo, fixColumn, columnFr, columnTo, fixRow, findValue1, findValue2) :
    wb = load_workbook(filePath)
    ws = wb.active
    global checkFindValue
    global checkFindValueItem

    for rowIndex in range(rowFr, rowTo) :
        rowTempData = ws.cell(row=rowIndex, column=fixColumn).value
        if rowTempData == findValue1 :
            checkFindValueItem = True #001

            for columnIndex in range(columnFr, columnTo) :
                columnTempData = ws.cell(row=fixRow, column=columnIndex).value
                if columnTempData == findValue2 :
                    checkFindValue = True #001
                    print('Success Find Value!! %d %d' % (rowIndex, columnIndex))
                    break
                else :
                    checkFindValue = False #001
            break
        else :
            checkFindValueItem = False #001
    if(checkFindValue == False and checkFindValueItem == False) : #001
        print('Failed Find Value!!(품목없음)%s %s' %(findValue1, findValue2))

    elif(checkFindValue == False and checkFindValueItem == True) : #001
        print('Failed Find Value!!(날짜경과)%s %s' % (findValue1, findValue2))

    checkFindValue = False #001
    checkFindValueItem = False #001



    wb.close()

    # 1000INCHEON 직송 -> 한양정밀
    # 1000INCHEON 일반 -> 인천
    # 1100CKD         -> CKD
    # 1130INCHOEN     -> 인천
    # 6000ANSAN       -> ?
    # 1000JISINCHEON  -> 인천
    # 1111JISGUNSAN   -> 군산
    # Failed Item send mail
