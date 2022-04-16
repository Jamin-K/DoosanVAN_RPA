# 신규생성 : 2022.01.31 김재민
# 개요 : 두산 VAN 데이터 추출 파이썬 스크립트

# 개발 메모
# 1000INCHEON 직송 -> 한양정밀
# 1000INCHEON 일반 -> 인천
# 1100CKD         -> CKD
# 1130INCHOEN     -> 인천
# 6000ANSAN       -> ?
# 1000JISINCHEON  -> 인천
# 1111JISGUNSAN   -> 군산

import os
import datetime
from tkinter import messagebox
import ExcelfileType1, ExcelfileType2, ExcelfileType3, setExcel
import setHolidays
import AddFailedData
import CreatedFailedExcelFile
import time
from openpyxl import load_workbook
from tkinter import *
import PresentDataProcessing
# RPA에서 수행해야 할 내용
# 1. 각 폴더 생성(DSVAN+todayDate, FailedData, 수행예정데이터, 완료데이터), holiday.xlsx, workday.xlsx 복사
# 2. 확장자 변경(xls -> xlsx) 후 .xls 파일 삭제
# 3. d-1일자의 doosanReleasePlan+d-1Date 를 d+day 폴더에 복붙하고 todayDate로 이름 변경


# 필요로직
# 1. 실패데이터에 대한 중복 체크 필요.(발주항번과 발주번호를 기준으로) -> 중복체크 로직 구현 및 테스트완료

# GUI 생성 START
# root = Tk()
# def okClick():
#     name = txt.get()
#     messagebox.showinfo("이름", name)
#     startTime = time.time()
#
#     todayDate = datetime.datetime.now().strftime('%Y%m%d')
#
#     # 공휴일 및 국경일 Excel Write START
#     # 01월 01일에 실행
#     if (todayDate[4:9] == '0101'):
#         setHolidays.startSetHoliday(todayDate[0:4])
#         setHolidays.startSetHoliday(str(int(todayDate[0:4]) + 1))
#     # 공휴일 및 국경일 Excel Write END
#
#     # 오늘 날짜 추출 START
#     todayDate = datetime.datetime.now().strftime('%Y%m%d')
#     print('수행날짜 : %s' % todayDate)
#     print(datetime.datetime.now().weekday())
#     todayDate = '20220214'  # TestCode
#     # 오늘 날짜 추출 END
#
#     # 폴더 파일 리스트 추출 START
#     path = 'C:/Users/KJM/Desktop/DSVAN' + todayDate
#     file_list = os.listdir(path)
#     # print(file_list)
#     # 폴더 파일 리스트 추출 END
#
#     # 출고계획 엑셀 파일 set 함수 호출 START
#     # setExcel.setStartPlanFile(path+'/'+'doosanReleasePlan'+todayDate+'.xlsx')
#     # setExcel.setStartPlanFile(path+'/'+'doosanReleasePlan20220118.xlsx') #TestCode
#     # 출고계획 엑셀 파일 set 함수 호출 END
#     print(file_list)
#
#     # 파일 이름에 따른 엑셀 데이터 추출 함수 호출 START
#     for fileName in file_list:
#         if '1000DIrINCHEON' + todayDate in fileName:
#             ExcelfileType1.getStartData(path + '/' + fileName)
#         elif '1000INCHEON' + todayDate in fileName:
#             ExcelfileType1.getStartData(path + '/' + fileName)
#         elif '1100CKD' + todayDate in fileName:
#             ExcelfileType1.getStartData(path + '/' + fileName)
#         elif '1130INCHOEN' + todayDate in fileName:
#             ExcelfileType1.getStartData(path + '/' + fileName)
#         elif '6000ANSAN' + todayDate in fileName:
#             ExcelfileType2.getStartData(path + '/' + fileName)
#         elif '1000JISINCHEON' + todayDate in fileName:
#             ExcelfileType3.getStartData(path + '/' + fileName)
#         elif '1111JISGUNSAN' + todayDate in fileName:
#             ExcelfileType3.getStartData(path + '/' + fileName)
#         else:
#             print('파일 분류 에러 : %s' % fileName)
#
#     # 파일 이름에 따른 엑셀 데이터 추출 함수 호출 END
#
#     print('코드 수행 시간 :', time.time() - startTime)
#
# lbl = Label(root, text="이름")
# lbl.grid(row=0, column=0)
# txt = Entry(root)
# txt.grid(row=0, column=1)
#
# # 버튼 클릭 이벤트와 핸들러 정의
# btn = Button(root, text="OK", command=okClick)
#
# btn.grid(row=1, column=1)
#
# root.mainloop()
# GUI 생성 END


startTime = time.time()

todayDate = datetime.datetime.now().strftime('%Y%m%d')

# 공휴일 및 국경일 Excel Write START
# 01월 01일에 실행
if(todayDate[4:9] == '0101') :
    setHolidays.startSetHoliday(todayDate[0:4])
    setHolidays.startSetHoliday(str(int(todayDate[0:4])+1))
# 공휴일 및 국경일 Excel Write END

# 오늘 날짜 추출 START
print('수행날짜 : %s' %todayDate)
todayDate = '20220401' #TestCode
# 오늘 날짜 추출 END

# D-1 날짜 추출 START
tempDate = datetime.datetime.strptime(todayDate, '%Y%m%d')
tempDate = tempDate + datetime.timedelta(days=-1)
oneDaysAgoDate = datetime.datetime.strftime(tempDate, '%Y%m%d')
# D-1 날짜 추출 END

# 폴더 파일 리스트 추출 START
defaultPath = 'C:/Users/KJM/Desktop/DSVAN'
path = defaultPath+todayDate
file_list = os.listdir(path)
#print(file_list)
# 폴더 파일 리스트 추출 END

# 실패한 데이터를 담을 excel 파일 만들기 START
failedFileName = 'FailedDataList.xlsx'
CreatedFailedExcelFile.start(path, failedFileName)
wbFailedListExcel = load_workbook(path + '/FailedData/' + failedFileName)
# 실패한 데이터를 담을 excel 파일 만들기 END

# 출고계획 엑셀 파일 set 함수 호출 START
#setExcel.setStartPlanFile(path+'/'+'doosanReleasePlan'+todayDate+'.xlsx', todayDate, defaultPath) 수정필요
setExcel.setStartPlanFile('C:/Users/KJM/Desktop/ReleasePlan.xlsx', todayDate, oneDaysAgoDate ,defaultPath) #TestCode
#setExcel.setStartPlanFile(path+'/'+'doosanReleasePlan20220118.xlsx') #TestCode
# 출고계획 엑셀 파일 set 함수 호출 END
print(file_list)

# 여기서 Present WorkBook과 WorkSheet 선언

# 파일 이름에 따른 엑셀 데이터 추출 함수 호출 START
# 아래에 수행예정 데이터 전처리 실시
for fileName in file_list :
    if '1000DirINCHEON'+todayDate in fileName :
        PresentDataProcessing.startProcessing(fileName, path)
        AddFailedData.addFailedDataStart(path, fileName, todayDate)
        ExcelfileType1.getStartData(path, fileName, wbFailedListExcel, todayDate)

    elif '1000INCHEON'+todayDate in fileName :
        PresentDataProcessing.startProcessing(fileName, path)
        AddFailedData.addFailedDataStart(path, fileName, todayDate)
        ExcelfileType1.getStartData(path, fileName, wbFailedListExcel, todayDate)

    elif '1100CKD'+todayDate in fileName :
        PresentDataProcessing.startProcessing(fileName, path)
        AddFailedData.addFailedDataStart(path, fileName, todayDate)
        ExcelfileType1.getStartData(path, fileName, wbFailedListExcel, todayDate)

    elif '1130INCHEON'+todayDate in fileName :
        PresentDataProcessing.startProcessing(fileName, path)
        AddFailedData.addFailedDataStart(path, fileName, todayDate)
        ExcelfileType1.getStartData(path, fileName, wbFailedListExcel, todayDate)

    elif '6000ANSAN'+todayDate in fileName :
        PresentDataProcessing.startProcessing(fileName, path)
        AddFailedData.addFailedDataStart(path, fileName, todayDate)
        ExcelfileType2.getStartData(path, fileName, wbFailedListExcel, todayDate)

    elif '1000JISINCHEON'+todayDate in fileName :
        PresentDataProcessing.startProcessing(fileName, path)
        AddFailedData.addFailedDataStart(path, fileName, todayDate)
        ExcelfileType3.getStartData(path, fileName, wbFailedListExcel, todayDate)

    elif '1111JISGUNSAN'+todayDate in fileName :
        PresentDataProcessing.startProcessing(fileName, path)
        AddFailedData.addFailedDataStart(path, fileName, todayDate)
        ExcelfileType3.getStartData(path, fileName, wbFailedListExcel, todayDate)

    else :
        print('파일 분류 에러 : %s' %fileName)

# 파일 이름에 따른 엑셀 데이터 추출 함수 호출 END

# workbook, worksheet 저장
wbFailedListExcel.save(path + '/FailedData/' + failedFileName)
wbFailedListExcel.close()
print('코드 수행 시간 :', time.time() - startTime)