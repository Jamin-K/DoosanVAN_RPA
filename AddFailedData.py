# 신규생성 : 2022.03.03 김재민
# 개요 : 전날에 실패한 엑셀 데이터와 현재 수행할 데이터를 합치는 로직

from openpyxl import load_workbook
import pandas as pd
import datetime

def addFailedDataStart(path, todayFileName, todayDate) :
    print('-----------------------------------------------------')
    # 변수 선언 START
    print(path)
    path = path[0:path.find('DSVAN'+todayDate) + 5] #----> 날짜 데이터를 더해서 사용
    print(path)

    tempDate = datetime.datetime.strptime(todayDate, '%Y%m%d')
    tempDate = tempDate + datetime.timedelta(days=-1)
    oneDaysAgoDate = datetime.datetime.strftime(tempDate, '%Y%m%d')
    pastPath = path + oneDaysAgoDate
    presentPath = path + todayDate
    findIndex = todayFileName.find('doosan') + 6
    sheetName = todayFileName[findIndex:]
    sheetName = sheetName[0:-13]
    sheetName = sheetName.strip()
    faileFileName = path + oneDaysAgoDate + '/FailedData/FailedDataList.xlsx'
    # 변수 선언 END

    print('파일 합치기 실행!!')
    print('파일이름 : %s' % todayFileName[todayFileName.find('doosan'):])
    print('시트이름 : %s' % sheetName)
    print('실행날짜 : %s' % todayDate)
    print('하루 전 날짜 : %s' % oneDaysAgoDate)
    print('현재경로 : %s' % presentPath)
    print('하루전경로 : %s' % pastPath)
    print(todayFileName)

    # DataFrame 기본 옵션 세팅 START
    pd.set_option('display.max_seq_items', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    # DataFrame 기본 옵션 세팅 END

    # 하루 전 실패 데이터를 시트 이름에 따라 DataFrame으로 가져오기 START
    failedExcelDateFrame = pd.read_excel(faileFileName, dtype={'발주번호':str,
                                                         '발주항번':str},
                                   sheet_name=sheetName)
    # 하루 전 실패 데이터를 시트 이름에 따라 DataFrame으로 가져오기 END

    # 실패 엑셀 파일의 해당 시트에 데이터가 존재할때만 수행 예정인 파일과 데이터 합치기 수행
    if(failedExcelDateFrame.empty != True):
        print('%s 시트 데이터 존재' %sheetName)
        # 실패 데이터와 수행 예정 데이터 합치기 START
        presentExcelDataFrame = pd.read_excel(presentPath + '/수행예정데이터/' + sheetName + '.xlsx')
        presentExcelDataFrame.drop(presentExcelDataFrame.columns[0], axis=1, inplace=True) # dataframe 인덱스 행 제거
        print(presentExcelDataFrame)

        # 실패 데이터와 수행 예정 데이터 합치기 END
    else :
        print('데이터 합치기 skip!!')

    print('-----------------------------------------------------')


def appendToExcel(path, df, sheetName):
    # 함수개요 : dataframe과 엑셀 파일의 데이터를 append 하여 결과를 엑셀로 저장
    # input - path : dataframe과 append 할 결과가 나올 엑셀 파일
    # input - df : path의 엑셀 파일과 더할 dataframe
    # input - sheetName : 데이터가 있는 시트 이름
    # 함수동작 : path의 데이터를 dataframe으로 받아와서 합칠 df 를 concat 하여 엑셀파일에 다시 저장
    tempDataFrame = pd.read_excel(path)
    tempDataFrame.drop(tempDataFrame.columns[0], axis=1, inplace=True)






