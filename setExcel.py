# 신규생성 : 2022.02.03 김재민
# 개요 : doosanReleasePlan 엑셀 파일 초기 세팅

from openpyxl import load_workbook
import pandas as pd
from datetime import datetime, timedelta

def setStartPlanFile(planFilePath) :
    # today -2 ~ today + 32 날짜 세팅 START
    wb = load_workbook(planFilePath)
    ws = wb.active
    rows = 4
    columns = 6
    for i in range(-2,32) :
        date = datetime.now() + timedelta(days=i) + timedelta(days=-20)
        date = date.strftime('%m') + '/' + date.strftime('%d')
        ws.cell(row=rows, column=columns, value=date)
        columns = columns + 1
    wb.save(planFilePath)
    wb.close()
    # today -2 ~ today + 32 날짜 세팅 END




