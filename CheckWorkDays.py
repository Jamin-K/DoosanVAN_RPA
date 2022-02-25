# 신규생성 : 2022.02.25 김재민
# 개요 : input으로 날짜를 받아와서 그 날짜가 공휴일에도 속하지 않고, 주말에도 속하지 않으면 해당 날짜를 그대로 반환.
#       그 날짜가 둘 중 하나라도 속하면 날짜를 바꿔 바꾼 날짜를 반환
#       이 함수에서 납기일-1 로직도 수행
# 수정 :
import datetime

import pandas as pd
import datetime as dt
import numpy as np


def checkHolidays(fullDate):     #fullDate는 yyyy/mm/dd형태

    # 변수선언 START
    holidayList = []
    workdayList = []
    workdayFlag = False
    filePath = 'C:/Users/KJM/Desktop/DSVAN20220214/' # 추후 holiday.xlsx 를 DSVAN20220214 상위폴더로 이동
    fileNameHoliday = 'holiday.xlsx'
    fileNameWorkday = 'workday.xlsx'

    # 변수선언 END

    # 납기일보다 -1day 로직 수행 START -> 납기일보다 하루 빠르게 출발 시켜야함
    fulldt = dt.datetime.strptime(fullDate, '%Y/%m/%d')
    fulldt = fulldt + datetime.timedelta(days=-1)
    # 납기일보다 -1day 로직 수행 END

    # DataFrame 기본 옵션 세팅 START
    pd.set_option('display.max_seq_items', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    # DataFrame 기본 옵션 세팅 END

    # DataFrame 초기화 START
    holidayDataFrame = pd.read_excel(filePath + fileNameHoliday, header=None, usecols=[0, 1, 2, 3])
    workdayDataFrame = pd.read_excel(filePath + fileNameWorkday, usecols=[0, 1])
    # DataFrame 초기화 END

    # 공휴일 or 국경일 리스트 만들기 START
    for i in range(len(holidayDataFrame)) :
        tempDate = str(holidayDataFrame.iloc[i, 2]) # 비교날짜 tempDate에 담기, tempdate는 yyyymmdd 형태로 넘어옴
        tempDate = tempDate[0:4] + '/' + tempDate[4:6] + '/' + tempDate[6:]# fullDate와 같은 형태로 가공 -> yyyy/mm/dd
        tempdt = dt.datetime.strptime(tempDate, '%Y/%m/%d')
        holidayList.append(tempdt)
    # 공휴일 or 국경일 리스트 만들기 END

    # 특근일자 리스트 만들기 START
    for i in range(len(workdayDataFrame)) :
        tempDate = str(workdayDataFrame.iloc[i, 1])
        tempDate = tempDate[0:4] + '/' + tempDate[4:6] + '/' + tempDate[6:]# fullDate와 같은 형태로 가공 -> yyyy/mm/dd
        tempdt = dt.datetime.strptime(tempDate, '%Y/%m/%d')
        workdayList.append(tempdt)
    # 특근일자 리스트 만들기 END

    print('input : %s' %fulldt)
    print(holidayList)
    print(workdayList)

    holidayList.reverse() # 뒤에서부터 탐색

    # 특근일자 검색 START
    for item in workdayList:
        if(item == fulldt) :
            print('해당날짜는 특근일입니다. 공휴일 및 주말 check skip!!')
            workdayFlag = True
        else :
            continue
    # 특근일자 검색 END

    if(workdayFlag == False) :
        # 공휴일 여부 탐색 FIRST START
        for item in holidayList:
            if (item == fulldt):
                print('해당날짜(%s)는 공휴일입니다. -1day 실시!!' %fulldt)
                fulldt = fulldt + datetime.timedelta(days=-1)
        # 공휴일 여부 탐색 FIRST END

        # 주말 여부 탐색(납기일이 토,일,월이면 전주 금요일에 납기) START
        while (True):
            if (fulldt.weekday() == 5 or fulldt.weekday() == 6 or fulldt.weekday() == 0):
                print('해당날짜(%s)는 토 or 일 or 월요일입니다. -1day 실시!!' %fulldt)
                fulldt = fulldt + datetime.timedelta(days=-1)
                continue
            break
        # 주말 여부 탐색(납기일이 토,일,월이면 전주 금요일에 납기) START

        # 공휴일 여부 탐색 SECOND START
        for item in holidayList:
            if (item == fulldt):
                print('해당날짜(%s)는 공휴일입니다. -1day 실시!!' %fulldt)
                fulldt = fulldt + datetime.timedelta(days=-1)
        # 공휴일 여부 탐색 SECOND END


    # # 공휴일 여부 탐색 FIRST START
    # for item in holidayList:
    #     if (item == fulldt):
    #         print('해당날짜(%s)는 공휴일입니다. -1day 실시!!' % fulldt)
    #         fulldt = fulldt + datetime.timedelta(days=-1)
    # # 공휴일 여부 탐색 FIRST END
    #
    # # 주말 여부 탐색(납기일이 토,일,월이면 전주 금요일에 납기) START
    # while (True):
    #     if (fulldt.weekday() == 5 or fulldt.weekday() == 6 or fulldt.weekday() == 0):
    #         print('해당날짜(%s)는 토 or 일 or 월요일입니다. -1day 실시!!' % fulldt)
    #         fulldt = fulldt + datetime.timedelta(days=-1)
    #         continue
    #     break
    # # 주말 여부 탐색(납기일이 토,일,월이면 전주 금요일에 납기) START
    #
    # # 공휴일 여부 탐색 SECOND START
    # for item in holidayList:
    #     if (item == fulldt):
    #         print('해당날짜(%s)는 공휴일입니다. -1day 실시!!' % fulldt)
    #         fulldt = fulldt + datetime.timedelta(days=-1)
    # # 공휴일 여부 탐색 SECOND END


    return fulldt.strftime('%Y/%m/%d')

    # 공휴일 or 국경일 여부 탐색 END




