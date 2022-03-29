# 신규생성 : 2022.02.03 김재민
# 개요 : doosanReleasePlan 엑셀 파일 초기 세팅
# 수정 : 2022.03.29 김재민 : 세팅이 끝난 doosanReleasePlan+todayDate를 완료데이터/ReleasePlan.xlsx 로 복사 및 이름변경 #001

from openpyxl import load_workbook
from datetime import datetime, timedelta
import shutil

def setStartPlanFile(planFilePath, todayDate, defaultPath) :
    # input - defaultPath : 'C:/Users/KJM/Desktop/DSVAN'

    print('SetExcel START!!')
    # today -2 ~ today + 32 날짜 세팅 START
    wb = load_workbook(planFilePath)
    ws = wb.active
    rows = 4
    columns = 6
    for i in range(-2,32) :
        date = datetime.now() + timedelta(days=i)# + timedelta(days=-20)
        date = date.strftime('%m') + '/' + date.strftime('%d')
        ws.cell(row=rows, column=columns, value=date)
        columns = columns + 1
    wb.save(planFilePath)
    wb.close()
    # today -2 ~ today + 32 날짜 세팅 END

    # 001 START
    shutil.copyfile(planFilePath, defaultPath+todayDate+'/완료데이터/ReleasePlan.xlsx')
    # 001 END





