# 신규생성 : 2022.01.31 김재민
# 개요 : 1000DirINCHEON, 1000INCHEON, 1100CKD, 1130INCHOEN 파일에 대한 데이터 추출 함수
# 수정 : 2022.02.03 김재민 : checkValue 값 체크로직 삭제
#       2022.02.03 김재민 : 발주계획 엑셀 파일 write 함수 call 부분 추가 #001


import pandas as pd
import numpy as np
import openpyxl
import WriteReleasePlan

# 변수선언 START
itemNumber = None; # 품번 #002
releaseDate = None; # 납기일 #002
rowFr = None
rowTo = None
fixColumn = 3
columnFr = 4
columnTo = 38
fixRow = 4
fileDirPath = 'C:/Users/KJM/Desktop/DSVAN20220131/'
# 변수선언 END

# DataFrame 기본 옵션 세팅 START
pd.set_option('display.max_seq_items', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
# DataFrame 기본 옵션 세팅 END

def getStartData(fileName) :
    if '1000INCHEON' in fileName :
        rowFr = 10
        rowTo = 50
    elif '1000DirINCHEON' in fileName : # 한양정밀
        rowFr = 5
        rowTo = 7
    elif '1100CKD' in fileName :
        rowFr = 7
        rowTo = 11
    elif '1130INCHOEN' in fileName :
        rowFr = 10
        rowTo = 50
    else :
        print('파일 분류 에러 : ExcelfileType1')

    print('START : %s' %fileName)
    print('▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼')
    excelDataFrame = pd.read_excel(fileName, usecols=[10, 16, 19, 31, 38]) #JIS, 발주번호, 품번, 납품잔량, 납기일자 필드 추출

    # JIS 값이 Y인 컬럼 DROP ROW START
    for index, row in excelDataFrame.iterrows() :
        if(row['JIS'] == 'Y') :
            excelDataFrame.drop(index, inplace=True)
    # JIS 값이 Y인 컬럼 DROP ROW END

    # 발주번호 필드 str로 형변환 START
    excelDataFrame['발주번호'] = excelDataFrame['발주번호'].astype(str)
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
        orderCount = excelDataFrame.iloc[0, 3]
        print('납품수량 합계 : %d' % orderCount)
        print('발주번호 : %s' % excelDataFrame.iloc[0, 1])
        print('품번 : %s' % excelDataFrame.iloc[0, 2])
        print('납기일 : %s' % excelDataFrame.iloc[0, 4])
        print('납품잔량 : %d' % excelDataFrame.iloc[0, 3])
    # 데이터가 1개 있을 때 실행되는 로직 END

    # 데이터가 2개 있을 때 실행되는 로직 START
    elif (len(excelDataFrame) == 2):
        if (excelDataFrame.iloc[0, 2] == excelDataFrame.iloc[1, 2]):
            if (excelDataFrame.iloc[0, 4] == excelDataFrame.iloc[1, 4]):
                # 동일품번, 동일납기
                orderCount = excelDataFrame.iloc[0, 3] + excelDataFrame.iloc[1, 3]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[0, 1])
                print('품번 : %s' % excelDataFrame.iloc[0, 2])
                print('납기일 : %s' % excelDataFrame.iloc[0, 4])
                print('납품잔량 : %d' % excelDataFrame.iloc[0, 3])
                print('-----------------------------------------------------')
            else:
                # 동일품번, 다른납기
                orderCount = excelDataFrame.iloc[0, 3]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[0, 1])
                print('품번 : %s' % excelDataFrame.iloc[0, 2])
                print('납기일 : %s' % excelDataFrame.iloc[0, 4])
                print('납품잔량 : %d' % excelDataFrame.iloc[0, 3])
                print('-----------------------------------------------------')
                orderCount = excelDataFrame.iloc[1, 3]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[1, 1])
                print('품번 : %s' % excelDataFrame.iloc[1, 2])
                print('납기일 : %s' % excelDataFrame.iloc[1, 4])
                print('납품잔량 : %d' % excelDataFrame.iloc[1, 3])

        else:
            # 다른품번
            orderCount = excelDataFrame.iloc[0, 3]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % excelDataFrame.iloc[0, 1])
            print('품번 : %s' % excelDataFrame.iloc[0, 2])
            print('납기일 : %s' % excelDataFrame.iloc[0, 4])
            print('납품잔량 : %d' % excelDataFrame.iloc[0, 3])
            print('-----------------------------------------------------')
            orderCount = excelDataFrame.iloc[1, 3]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % excelDataFrame.iloc[1, 1])
            print('품번 : %s' % excelDataFrame.iloc[1, 2])
            print('납기일 : %s' % excelDataFrame.iloc[1, 4])
            print('납품잔량 : %d' % excelDataFrame.iloc[1, 3])
    # 데이터가 2개 있을 때 실행되는 로직 END

    # 데이터가 3개 이상 있을 때 실행되는 로직 start
    else:
        for i in range(len(excelDataFrame) - 1):
            print('-----------------------------------------------------')
            if (excelDataFrame.iloc[i, 2] == excelDataFrame.iloc[i + 1, 2]):
                if (excelDataFrame.iloc[i, 4] == excelDataFrame.iloc[i + 1, 4]):
                    if (i >= len(excelDataFrame) - 2):
                        #checkValue = True
                        orderCount = orderCount + excelDataFrame.iloc[i + 1, 3]
                        print('납품수량 합계 : %d' % orderCount)
                        print('발주번호 : %s' % excelDataFrame.iloc[i, 1])  # 발주번호
                        print('품번 : %s' % excelDataFrame.iloc[i, 2])  # 품번
                        print('납기일 : %s' % excelDataFrame.iloc[i, 4])  # 납기일
                        print('요청수량 : %d' % excelDataFrame.iloc[i, 3])  # 요청수량 INTEGER
                        print('-----------------------------------------------------')

                    # 원래 동일품번 동일납기 로직 수행
                    #checkValue = True
                    orderCount = orderCount + excelDataFrame.iloc[i + 1, 3]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % excelDataFrame.iloc[i + 1, 1])  # 발주번호
                    print('품번 : %s' % excelDataFrame.iloc[i + 1, 2])  # 품번
                    print('납기일 : %s' % excelDataFrame.iloc[i + 1, 4])  # 납기일
                    print('요청수량 : %d' % excelDataFrame.iloc[i + 1, 3])  # 요청수량 INTEGER
                    itemNumber = excelDataFrame.iloc[i+1, 2] #002
                    releaseDate = excelDataFrame.iloc[i+1, 4][5:10] #002
                    if (i == len(excelDataFrame) - 2):
                        print('ExcelWrite Function Call') #001
                        WriteReleasePlan.startWriteCell(fileDirPath + 'doosanReleasePlan20220118.xlsx', rowFr, rowTo,
                                                        fixColumn, columnFr, columnTo,
                                                        fixRow, itemNumber, releaseDate)  # 003

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
                    orderCount = orderCount + excelDataFrame.iloc[i, 3]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % excelDataFrame.iloc[i, 1])  # 발주번호
                    print('품번 : %s' % excelDataFrame.iloc[i, 2])  # 품번
                    print('납기일 : %s' % excelDataFrame.iloc[i, 4])  # 납기일
                    print('요청수량 : %d' % excelDataFrame.iloc[i, 3])  # 요청수량 INTEGER
                    itemNumber = excelDataFrame.iloc[i, 2]  # 002
                    releaseDate = excelDataFrame.iloc[i, 4][5:10]  # 002
                    #if (i == len(excelDataFrame) - 2):
                    print('ExcelWrite Function Call')  # 001
                    WriteReleasePlan.startWriteCell(fileDirPath + 'doosanReleasePlan20220118.xlsx', rowFr, rowTo,
                                                        fixColumn, columnFr, columnTo,
                                                        fixRow, itemNumber, releaseDate)  # 003

                    print('-----------------------------------------------------')
                    # 2 orderCount = 0 초기화
                    orderCount = 0
                    # 마지막 index 수행
                    if (i == len(excelDataFrame) - 2):
                        print('마지막 index 실행')
                        orderCount = orderCount + excelDataFrame.iloc[i + 1, 3]
                        print('납품수량 합계 : %d' % orderCount)
                        print('발주번호 : %s' % excelDataFrame.iloc[i + 1, 1])  # 발주번호
                        print('품번 : %s' % excelDataFrame.iloc[i + 1, 2])  # 품번
                        print('납기일 : %s' % excelDataFrame.iloc[i + 1, 4])  # 납기일
                        print('요청수량 : %d' % excelDataFrame.iloc[i + 1, 3])  # 요청수량 INTEGER
                        itemNumber = excelDataFrame.iloc[i + 1, 2]  # 002
                        releaseDate = excelDataFrame.iloc[i + 1, 4][5:10]  # 002
                        #if (i == len(excelDataFrame) - 2):
                        print('ExcelWrite Function Call')  # 001
                        WriteReleasePlan.startWriteCell(fileDirPath + 'doosanReleasePlan20220118.xlsx', rowFr,
                                                            rowTo,
                                                            fixColumn, columnFr, columnTo,
                                                            fixRow, itemNumber, releaseDate)  # 003

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
                orderCount = orderCount + excelDataFrame.iloc[i, 3]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[i, 1])  # 발주번호
                print('품번 : %s' % excelDataFrame.iloc[i, 2])  # 품번
                print('납기일 : %s' % excelDataFrame.iloc[i, 4])  # 납기일
                print('요청수량 : %d' % excelDataFrame.iloc[i, 3])  # 요청수량 INTEGER
                itemNumber = excelDataFrame.iloc[i, 2]  # 002
                releaseDate = excelDataFrame.iloc[i, 4][5:10]  # 002
                print('ExcelWrite Function Call')  # 001
                WriteReleasePlan.startWriteCell(fileDirPath + 'doosanReleasePlan20220118.xlsx', rowFr, rowTo,
                                                    fixColumn, columnFr, columnTo,
                                                    fixRow, itemNumber, releaseDate)  # 003

                print('-----------------------------------------------------')
                # 2 orderCount = 0 초기화
                orderCount = 0
                # 마지막 index 수행
                if (i == len(excelDataFrame) - 2):
                    print('마지막 index 실행')
                    orderCount = orderCount + excelDataFrame.iloc[i + 1, 3]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % excelDataFrame.iloc[i + 1, 1])  # 발주번호
                    print('품번 : %s' % excelDataFrame.iloc[i + 1, 2])  # 품번
                    print('납기일 : %s' % excelDataFrame.iloc[i + 1, 4])  # 납기일
                    print('요청수량 : %d' % excelDataFrame.iloc[i + 1, 3])  # 요청수량 INTEGER
                    itemNumber = excelDataFrame.iloc[i + 1, 2]  # 002
                    releaseDate = excelDataFrame.iloc[i + 1, 4][5:10]  # 002
                    #if (i == len(excelDataFrame) - 2):
                    print('ExcelWrite Function Call')  # 001
                    WriteReleasePlan.startWriteCell(fileDirPath + 'doosanReleasePlan20220118.xlsx', rowFr, rowTo,
                                                        fixColumn, columnFr, columnTo,
                                                        fixRow, itemNumber, releaseDate)  # 003

                    print('-----------------------------------------------------')
                    orderCount = 0

    # 데이터가 3개 이상 있을 때 실행되는 로직 END
    # 데이터 추출 로직 END

    print('▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲')
