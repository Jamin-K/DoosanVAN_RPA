# 신규생성 : 2022.01.31 김재민
# 개요 : 1000DirINCHEON, 1000INCHEON, 1100CKD, 1130INCHOEN 파일에 대한 데이터 추출 함수
# 수정 : 2022.02.03 김재민 : checkValue 값 체크로직 삭제
#       2022.02.03 김재민 : 발주계획 엑셀 파일 write 함수 call 부분 추가 #001
#       2022.02.03 김재민 : 품번, 납기일 전역변수 추가 및 납기일 Split #002
#       2022.02.04 김재민 : startWriteCell() 함수 호출을 위한 변수선언 및 함수 호출 #003
#       2022.02.13 김재민 : 데이터가 1개일떄, 2개일때 함수 call 로직 추가

import datetime
import pandas as pd
import numpy as np
import openpyxl
import WriteReleasePlan
from openpyxl import load_workbook

# 변수선언 START
todayDate = datetime.datetime.now().strftime('%Y%m%d')
itemNumber = None; # 품번 #002
releaseDate = None; # 납기일 #002
rowFr = None
rowTo = None
fixColumn = 3
columnFr = 6
columnTo = 40
fixRow = 4
#fileDirPath = 'C:/Users/KJM/Desktop/DSVAN'+todayDate+'/'
#releaseFileName = filrDirPath + 'doosanReleasePlan' + todayDate + '.xlsx'
fileDirPath = 'C:/Users/KJM/Desktop/DSVAN20220214/' #TestCode
releaseFileName = fileDirPath + 'doosanReleasePlan20220214.xlsx' #TestCode
orderNumber = None
semiOrderNumber = None
# 변수선언 END

# DataFrame 기본 옵션 세팅 START
pd.set_option('display.max_seq_items', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
# DataFrame 기본 옵션 세팅 END

def getStartData(fileName, wbFailedListExcel) :
    if '1000INCHEON' in fileName :
        print('1000INCHEON 파일 시작')
        rowFr = 11
        rowTo = 50
    elif '1000DirINCHEON' in fileName : # 한양정밀
        print('1000DirINCHOEN 파일 시작')
        rowFr = 5
        rowTo = 7
    elif '1100CKD' in fileName :
        print('1100CKD 파일 시작')
        rowFr = 7
        rowTo = 11
    elif '1130INCHEON' in fileName :
        print('1130INCHOEN 파일 시작')
        rowFr = 11
        rowTo = 50
    else :
        print('파일 분류 에러 : ExcelfileType1')

    print('START : %s' %fileName)
    print('▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼')
    excelDataFrame = pd.read_excel(fileName, usecols=[10, 16, 17, 19, 31, 38],
                                   dtype={'발주번호':str,
                                          '발주항번':str})
    # excelDataFrame = pd.read_excel(fileName, names=['JIS', '발주번호', '발주항번', '품번', '납품잔량', '납기일자'],
    #                                dtype={'발주번호': str,
    #                                       '발주항번': str})
    # JIS, 발주번호, 품번, 납품잔량, 납기일자, 발주항번 필드 추출

    # JIS 값이 Y인 컬럼 DROP ROW START
    for index, row in excelDataFrame.iterrows() :
        if(row['JIS'] == 'Y') :
            excelDataFrame.drop(index, inplace=True)
    # JIS 값이 Y인 컬럼 DROP ROW END

    # 발주번호 필드 str로 형변환 START
    excelDataFrame['발주번호'] = excelDataFrame['발주번호'].astype(str)
    excelDataFrame['발주항번'] = excelDataFrame['발주항번'].astype(str)

    # 발주번호 필드 str로 형변환 END

    # dataframe index 재선언 START
    print('전체 인덱스 개수 : %d' % len(excelDataFrame))
    print('-----------------------------------------------------')
    newIdxArr = []
    for i in range(len(excelDataFrame)):
        newIdxArr.append((i))
    excelDataFrame.set_index(keys=[newIdxArr], inplace=True)
    # dataframe index 재선언 END

    orderCount = 0
    checkValue = False

    # 데이터 추출 로직 START
    print('-----------------------------------------------------')
    # 데이터가 1개 있을 때 실행되는 로직 START
    if (len(excelDataFrame) == 1):
        orderCount = excelDataFrame.iloc[0, 4]
        print('납품수량 합계 : %d' % orderCount)
        print('발주번호 : %s' % excelDataFrame.iloc[0, 1])
        print('발주항번 : %s' % excelDataFrame.iloc[0, 2])
        print('품번 : %s' % excelDataFrame.iloc[0, 3])
        print('납기일 : %s' % excelDataFrame.iloc[0, 5])
        print('납품잔량 : %d' % excelDataFrame.iloc[0, 4])
        itemNumber = excelDataFrame.iloc[0, 3]
        releaseDate = excelDataFrame.iloc[0, 5]
        orderNumber = excelDataFrame.iloc[0, 1]
        semiOrderNumber = excelDataFrame.iloc[0, 2]
        print('ExcelWrite Function Call')
        WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                        fixColumn, columnFr, columnTo,
                                        fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                        , semiOrderNumber, wbFailedListExcel, fileName)  # 003
    # 데이터가 1개 있을 때 실행되는 로직 END

    # 데이터가 2개 있을 때 실행되는 로직 START
    elif (len(excelDataFrame) == 2):
        if (excelDataFrame.iloc[0, 2] == excelDataFrame.iloc[1, 2]):
            if (excelDataFrame.iloc[0, 4] == excelDataFrame.iloc[1, 4]):
                # 동일품번, 동일납기
                orderCount = excelDataFrame.iloc[0, 4] + excelDataFrame.iloc[1, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[0, 1])
                print('발주항번 : %s' % excelDataFrame.iloc[0, 2])
                print('품번 : %s' % excelDataFrame.iloc[0, 3])
                print('납기일 : %s' % excelDataFrame.iloc[0, 5])
                print('납품잔량 : %d' % excelDataFrame.iloc[0, 4])
                itemNumber = excelDataFrame.iloc[0, 3]
                releaseDate = excelDataFrame.iloc[0, 5]
                orderNumber = excelDataFrame.iloc[0, 1]
                semiOrderNumber = excelDataFrame.iloc[0, 2]
                # releaseDate = excelDataFrame.iloc[0, 4][5:10]
                print('ExcelWrite Function Call')
                WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                                fixColumn, columnFr, columnTo,
                                                fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                                , semiOrderNumber, wbFailedListExcel, fileName)  # 003
                print('-----------------------------------------------------')
            else:
                # 동일품번, 다른납기
                orderCount = excelDataFrame.iloc[0, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[0, 1])
                print('발주항번 : %s' % excelDataFrame.iloc[0, 2])
                print('품번 : %s' % excelDataFrame.iloc[0, 3])
                print('납기일 : %s' % excelDataFrame.iloc[0, 5])
                print('납품잔량 : %d' % excelDataFrame.iloc[0, 4])
                itemNumber = excelDataFrame.iloc[0, 3]
                releaseDate = excelDataFrame.iloc[0, 5]
                orderNumber = excelDataFrame.iloc[0, 1]
                semiOrderNumber = excelDataFrame.iloc[0, 2]
                print('ExcelWrite Function Call')
                WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                                fixColumn, columnFr, columnTo,
                                                fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                                , semiOrderNumber, wbFailedListExcel, fileName)  # 003
                print('-----------------------------------------------------')
                orderCount = excelDataFrame.iloc[1, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[1, 1])
                print('발주항번 : %s' % excelDataFrame.iloc[1, 2])
                print('품번 : %s' % excelDataFrame.iloc[1, 3])
                print('납기일 : %s' % excelDataFrame.iloc[1, 5])
                print('납품잔량 : %d' % excelDataFrame.iloc[1, 4])
                itemNumber = excelDataFrame.iloc[1, 3]
                releaseDate = excelDataFrame.iloc[1, 5]
                orderNumber = excelDataFrame.iloc[1, 1]
                semiOrderNumber = excelDataFrame.iloc[1, 2]
                print('ExcelWrite Function Call')
                WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                                fixColumn, columnFr, columnTo,
                                                fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                                , semiOrderNumber, wbFailedListExcel, fileName)  # 003
        else:
            # 다른품번
            orderCount = excelDataFrame.iloc[0, 3]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % excelDataFrame.iloc[0, 1])
            print('발주항번 : %s' % excelDataFrame.iloc[0, 2])
            print('품번 : %s' % excelDataFrame.iloc[0, 3])
            print('납기일 : %s' % excelDataFrame.iloc[0, 5])
            print('납품잔량 : %d' % excelDataFrame.iloc[0, 4])
            itemNumber = excelDataFrame.iloc[0, 3]
            releaseDate = excelDataFrame.iloc[0, 5]
            orderNumber = excelDataFrame.iloc[0, 1]
            semiOrderNumber = excelDataFrame.iloc[0, 2]
            print('ExcelWrite Function Call')
            WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                            fixColumn, columnFr, columnTo,
                                            fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                            , semiOrderNumber, wbFailedListExcel, fileName)  # 003
            print('-----------------------------------------------------')
            orderCount = excelDataFrame.iloc[1, 3]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % excelDataFrame.iloc[1, 1])
            print('발주항번 : %s' % excelDataFrame.iloc[1, 2])
            print('품번 : %s' % excelDataFrame.iloc[1, 3])
            print('납기일 : %s' % excelDataFrame.iloc[1, 5])
            print('납품잔량 : %d' % excelDataFrame.iloc[1, 4])
            itemNumber = excelDataFrame.iloc[1, 3]
            releaseDate = excelDataFrame.iloc[1, 5]
            orderNumber = excelDataFrame.iloc[1, 1]
            semiOrderNumber = excelDataFrame.iloc[1, 2]
            print('ExcelWrite Function Call')
            WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                            fixColumn, columnFr, columnTo,
                                            fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                            , semiOrderNumber, wbFailedListExcel, fileName)  # 003
    # 데이터가 2개 있을 때 실행되는 로직 END

    # 데이터가 3개 이상 있을 때 실행되는 로직 start
    else:
        for i in range(len(excelDataFrame) - 1):
            print('-----------------------------------------------------')
            if (excelDataFrame.iloc[i, 3] == excelDataFrame.iloc[i + 1, 3]):
                if (excelDataFrame.iloc[i, 5] == excelDataFrame.iloc[i + 1, 5]):
                    if (i >= len(excelDataFrame) - 2):
                        #checkValue = True
                        orderCount = orderCount + excelDataFrame.iloc[i + 1, 4]
                        print('납품수량 합계 : %d' % orderCount)
                        print('발주번호 : %s' % excelDataFrame.iloc[i, 1])
                        print('발주항번 : %s' % excelDataFrame.iloc[i, 2])
                        print('품번 : %s' % excelDataFrame.iloc[i, 3])
                        print('납기일 : %s' % excelDataFrame.iloc[i, 5])
                        print('납품잔량 : %d' % excelDataFrame.iloc[i, 4])
                        print('-----------------------------------------------------')

                    # 원래 동일품번 동일납기 로직 수행
                    #checkValue = True
                    orderCount = orderCount + excelDataFrame.iloc[i + 1, 4]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % excelDataFrame.iloc[i+1, 1])
                    print('발주항번 : %s' % excelDataFrame.iloc[i+1, 2])
                    print('품번 : %s' % excelDataFrame.iloc[i+1, 3])
                    print('납기일 : %s' % excelDataFrame.iloc[i+1, 5])
                    print('납품잔량 : %d' % excelDataFrame.iloc[i+1, 4])
                    itemNumber = excelDataFrame.iloc[i+1, 3]
                    releaseDate = excelDataFrame.iloc[i+1, 5]
                    orderNumber = excelDataFrame.iloc[i+1, 1]
                    semiOrderNumber = excelDataFrame.iloc[i+1, 2]
                    if (i == len(excelDataFrame) - 2):
                        print('ExcelWrite Function Call') #001
                        WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                                        fixColumn, columnFr, columnTo,
                                                        fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                                        , semiOrderNumber, wbFailedListExcel, fileName)  # 003

                    print('-----------------------------------------------------')

                else:
                    # 원래 동일품번 다른납기 로직 수행
                    # orderCount 기록 후 0으로 초기화
                    # 1 orderCount 기록
                    # if (checkValue == True):
                    #     checkValue = False
                    #     orderCount = orderCount + excelDataFrame.iloc[i, 3]
                    # else:
                    #     orderCount = orderCount + excelDataFrame.iloc[i, 3]
                    orderCount = orderCount + excelDataFrame.iloc[i, 4]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % excelDataFrame.iloc[i, 1])
                    print('발주항번 : %s' % excelDataFrame.iloc[i, 2])
                    print('품번 : %s' % excelDataFrame.iloc[i, 3])
                    print('납기일 : %s' % excelDataFrame.iloc[i, 5])
                    print('납품잔량 : %d' % excelDataFrame.iloc[i, 4])
                    itemNumber = excelDataFrame.iloc[i, 3]
                    releaseDate = excelDataFrame.iloc[i, 5]
                    orderNumber = excelDataFrame.iloc[i, 1]
                    semiOrderNumber = excelDataFrame.iloc[i, 2]
                    #if (i == len(excelDataFrame) - 2):
                    print('ExcelWrite Function Call')  # 001
                    WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                                    fixColumn, columnFr, columnTo,
                                                    fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                                    , semiOrderNumber, wbFailedListExcel, fileName)  # 003

                    print('-----------------------------------------------------')
                    # 2 orderCount = 0 초기화
                    orderCount = 0
                    # 마지막 index 수행
                    if (i == len(excelDataFrame) - 2):
                        print('마지막 index 실행')
                        orderCount = orderCount + excelDataFrame.iloc[i + 1, 4]
                        print('납품수량 합계 : %d' % orderCount)
                        print('발주번호 : %s' % excelDataFrame.iloc[i+1, 1])
                        print('발주항번 : %s' % excelDataFrame.iloc[i+1, 2])
                        print('품번 : %s' % excelDataFrame.iloc[i+1, 3])
                        print('납기일 : %s' % excelDataFrame.iloc[i+1, 5])
                        print('납품잔량 : %d' % excelDataFrame.iloc[i+1, 4])
                        itemNumber = excelDataFrame.iloc[i+1, 3]
                        releaseDate = excelDataFrame.iloc[i+1, 5]
                        orderNumber = excelDataFrame.iloc[i + 1, 1]
                        semiOrderNumber = excelDataFrame.iloc[i + 1, 2]
                        #if (i == len(excelDataFrame) - 2):
                        print('ExcelWrite Function Call')  # 001
                        WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                                        fixColumn, columnFr, columnTo,
                                                        fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                                        , semiOrderNumber, wbFailedListExcel, fileName)  # 003

                        print('-----------------------------------------------------')
                        orderCount = 0

            else:
                # 다른 품번
                # orderCount 기록 후 0으로 초기화
                # 1 orderCount 기록
                # if (checkValue == True):
                #     checkValue = False
                #     orderCount = orderCount + excelDataFrame.iloc[i, 3]
                # else:
                #     orderCount = orderCount + excelDataFrame.iloc[i, 3]
                orderCount = orderCount + excelDataFrame.iloc[i, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[i, 1])
                print('발주항번 : %s' % excelDataFrame.iloc[i, 2])
                print('품번 : %s' % excelDataFrame.iloc[i, 3])
                print('납기일 : %s' % excelDataFrame.iloc[i, 5])
                print('납품잔량 : %d' % excelDataFrame.iloc[i, 4])
                itemNumber = excelDataFrame.iloc[i, 3]
                releaseDate = excelDataFrame.iloc[i, 5]
                orderNumber = excelDataFrame.iloc[i, 1]
                semiOrderNumber = excelDataFrame.iloc[i, 2]
                print('ExcelWrite Function Call')  # 001
                WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                                fixColumn, columnFr, columnTo,
                                                fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                                , semiOrderNumber, wbFailedListExcel, fileName)  # 003

                print('-----------------------------------------------------')
                # 2 orderCount = 0 초기화
                orderCount = 0
                # 마지막 index 수행
                if (i == len(excelDataFrame) - 2):
                    print('마지막 index 실행')
                    orderCount = orderCount + excelDataFrame.iloc[i + 1, 4]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % excelDataFrame.iloc[i+1, 1])
                    print('발주항번 : %s' % excelDataFrame.iloc[i+1, 2])
                    print('품번 : %s' % excelDataFrame.iloc[i+1, 3])
                    print('납기일 : %s' % excelDataFrame.iloc[i+1, 5])
                    print('납품잔량 : %d' % excelDataFrame.iloc[i+1, 4])
                    itemNumber = excelDataFrame.iloc[i+1, 3]
                    releaseDate = excelDataFrame.iloc[i+1, 5]
                    orderNumber = excelDataFrame.iloc[i + 1, 1]
                    semiOrderNumber = excelDataFrame.iloc[i + 1, 2]
                    print('ExcelWrite Function Call')  # 001
                    WriteReleasePlan.startWriteCell(releaseFileName, rowFr, rowTo,
                                                    fixColumn, columnFr, columnTo,
                                                    fixRow, itemNumber, releaseDate, orderCount, orderNumber
                                                    , semiOrderNumber, wbFailedListExcel, fileName)  # 003

                    print('-----------------------------------------------------')
                    orderCount = 0

    # 데이터가 3개 이상 있을 때 실행되는 로직 END
    # 데이터 추출 로직 END

    print('▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲')
