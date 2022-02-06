# 신규생성 : 2022.02.04 김재민
# 개요 : 가로 세로 2개의 값을 입력받아 해당 셀의 위치를 반환하는 함수

from openpyxl import load_workbook

# 변수선언 START
checkFindValue = False
# 변수선언 END

def startFindValue(filePath, rowFr, rowTo, fixColumn, columnFr, columnTo, fixRow, findValue1, findValue2) :
    wb = load_workbook(filePath)
    ws = wb.active

    for rowIndex in range(rowFr, rowTo) : # 행 범위 탐색
        rowTempData = ws.cell(row=rowIndex, column=fixColumn).value
        if rowTempData == findValue1 :
            for columnIndex in range(columnFr, columnTo) : # 열 범위 탐색
                columnTempData = ws.cell(row=fixRow, column=columnIndex).value
                if columnTempData == findValue2 :
                    checkFindValue = True
                    break
                else :
                    checkFindValue = False
            if(checkFindValue == True) :
                print('Success Find Value!! %d %d' % (rowIndex, columnIndex))
            else :
                print('Failed Find Value!! %s %s' % (findValue1, findValue2))

    wb.close()

    # 1000INCHEON 직송 -> 한양정밀
    # 1000INCHEON 일반 -> 인천
    # 1100CKD         -> CKD
    # 1130INCHOEN     -> 인천
    # 6000ANSAN       -> ?
    # 1000JISINCHEON  -> 인천
    # 1111JISGUNSAN   -> 군산
    # Failed Item send mail