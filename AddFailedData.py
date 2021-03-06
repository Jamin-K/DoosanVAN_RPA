# 신규생성 : 2022.03.03 김재민
# 개요 : 전날에 실패한 엑셀 데이터와 현재 수행할 데이터를 합치는 로직
# 수정 : 2022.04.16 김재민 : 실패데이터 중복 체크 후 중복 제거 로직 추가 #001
#       2022.07.06 김재민 : 1000INCHEON 파일 수행안함. 1130DirINCHEON 파일 수행 #002

from openpyxl import load_workbook
import pandas as pd
import datetime

def addFailedDataStart(path, todayFileName, todayDate) :
    # input - path : C:/Users/KJM/Desktop/DSVAN20220214
    # input - todayFileName : doosan1100CKD20220214.xlsx
    # input - todayDate : 20220214
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

    # 002 START
    errorFlag = False
    try:
        failedExcelDateFrame = pd.read_excel(faileFileName, dtype={'발주번호':str,
                                                                   '발주항번':str},
                                             sheet_name=sheetName)
    except:
        print('Failed 데이터 파일 없음')
        errorFlag = True
        # failedExcelDateFrame = pd.DataFrame({
        #                         '발주번호': [''],
        #                         '발주항번': [''],
        #                         '품명': [''],
        #                         '날짜': [''],
        #                         '발주수량': [''],
        #                         'JIS': [''],
        #                         'Category': ['']
        #                                     })
        # failedExcelDateFrame['발주수량'] = pd.to_numeric(failedExcelDateFrame['발주수량'])

    # 001 END
    # failedExcelDateFrame = pd.read_excel(faileFileName, dtype={'발주번호':str,
    #                                                      '발주항번':str},
    #                                sheet_name=sheetName)
    # 하루 전 실패 데이터를 시트 이름에 따라 DataFrame으로 가져오기 END

    # 실패 엑셀 파일의 해당 시트에 데이터가 존재할때만 수행 예정인 파일과 데이터 합치기 수행

    if(errorFlag == False):

        if(failedExcelDateFrame.empty != True):
            print('%s 시트 데이터 존재' % sheetName)
            appendToExcel(presentPath + '/수행예정데이터/' + sheetName + '.xlsx', failedExcelDateFrame, sheetName)
    else:
        print('데이터 합치기 skip!!')

    print('---------------------------------------------------------------')

    print('-----------------------------------------------------')

    # 실패데이터 포맷
    # 발주번호, 발주항번, 품명, 날짜, 발주수량, JIS, Category

def appendToExcel(path, df, sheetName):
    # 함수개요 : dataframe과 엑셀 파일의 데이터를 append 하여 결과를 엑셀로 저장
    # input - path : dataframe과 append 할 결과가 나올 엑셀 파일
    # input - df : path의 엑셀 파일과 더할 dataframe
    # input - sheetName : 데이터가 있는 시트 이름
    # 함수동작 : path의 데이터를 dataframe으로 받아와서 합칠 df 를 concat 하여 엑셀파일에 다시 저장

    print('appendToExcel function call!!')

    # 시트이름에 따라 필요없는 컬럼 Drop 하고 열 순서 변경하는 로직 START
    if(sheetName == '1000DirINCHEON' or sheetName == '1100CKD' or sheetName == '1130INCHEON' or sheetName == '1130DirINCHEON') : #002
        df.rename(columns={'품명' : '품번',
                           '발주수량' : '납품잔량',
                           '날짜' : '납기일자'}, inplace=True)
        df.drop(df.columns[6], axis=1, inplace=True)
        df = df[['JIS', '발주번호', '발주항번', '품번', '납품잔량', '납기일자']]
        #print('컬럼 리스트')
        #print(list(df.columns))
        print('fileType1 작업 수행')
    elif(sheetName == '6000ANSAN') :
        df.rename(columns={'품명': '품번',
                           '발주수량': '납품잔량',
                           '날짜': '납기일자'}, inplace=True)
        df.drop(df.columns[6], axis=1, inplace=True)
        df.drop(df.columns[5], axis=1, inplace=True)
        df = df[['품번', '납품잔량', '납기일자', '발주번호', '발주항번']]
        #print('컬럼 리스트')
        #print(list(df.columns))
        print('fileType2 작업 수행')
    elif(sheetName == '1000JISINCHEON' or sheetName == '1111JISGUNSAN') :
        df.rename(columns={'품명': '품번',
                           '발주수량': '요청수량',
                           '날짜': '납기일(Actual)'}, inplace=True)
        df.drop(df.columns[5], axis=1, inplace=True)
        df = df[['발주번호', '발주항번', '품번', 'Category', '납기일(Actual)', '요청수량']]
        #print('컬럼 리스트')
        #print(list(df.columns))
        print('fileType3 작업 수행')
    else :
        print('sheetName Error!! : %s' %sheetName)
    # 시트이름에 따라 필요없는 컬럼 Drop 하고 열 순서 변경하는 로직 END

    # 데이터프레임 합치기 START
    tempDataFrame = pd.read_excel(path, dtype={'발주번호':str, '발주항번':str})
    tempDataFrame.drop(tempDataFrame.columns[0], axis=1, inplace=True)
    addDataFrame = pd.concat([tempDataFrame, df], axis=0, ignore_index=True)
    # 데이터프레임 합치기 END

    # 합친 파일 수행예정데이터 폴더에 작성 START
    addDataFrame = addDataFrame.drop_duplicates(subset=['발주번호', '발주항번'], inplace=False) #001
    addDataFrame.dropna(axis=0)
    addDataFrame.to_excel(path, header=True)
    # 합친 파일 수행예정데이터 폴더에 작성 END









