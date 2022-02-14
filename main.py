# 신규생성 : 2022.01.31 김재민
# 개요 : 두산 VAN 데이터 추출 파이썬 스크립트

import os
import datetime
import ExcelfileType1, ExcelfileType2, ExcelfileType3, setExcel
import time



# 오늘 날짜 추출 START
todayDate = datetime.datetime.now().strftime('%Y%m%d')
print('수행날짜 : %s' %todayDate)
#todayDate = '20220131' #TestCode
# 오늘 날짜 추출 END

# 폴더 파일 리스트 추출 START
path = 'C:/Users/KJM/Desktop/DSVAN'+todayDate
file_list = os.listdir(path)
#print(file_list)
# 폴더 파일 리스트 추출 END

# 출고계획 엑셀 파일 set 함수 호출 START
setExcel.setStartPlanFile(path+'/'+'doosanReleasePlan'+todayDate+'.xlsx')
#setExcel.setStartPlanFile(path+'/'+'doosanReleasePlan20220118.xlsx') #TestCode
# 출고계획 엑셀 파일 set 함수 호출 END
print(file_list)

# 파일 이름에 따른 엑셀 데이터 추출 함수 호출 START
for fileName in file_list :
    if '1000DIrINCHEON'+todayDate in fileName :
        ExcelfileType1.getStartData(path+'/'+fileName)
    elif '1000INCHEON'+todayDate in fileName :
        ExcelfileType1.getStartData(path+'/'+fileName)
    elif '1100CKD'+todayDate in fileName :
        ExcelfileType1.getStartData(path+'/'+fileName)
    elif '1130INCHOEN'+todayDate in fileName :
        ExcelfileType1.getStartData(path+'/'+fileName)
    elif '6000ANSAN'+todayDate in fileName :
        ExcelfileType2.getStartData(path + '/' + fileName)
    elif '1000JISINCHEON'+todayDate in fileName :
        ExcelfileType3.getStartData(path + '/' + fileName)
    elif '1111JISGUNSAN'+todayDate in fileName :
        ExcelfileType3.getStartData(path + '/' + fileName)
    else :
        print('파일 분류 에러 : %s' %fileName)

# 파일 이름에 따른 엑셀 데이터 추출 함수 호출 END
