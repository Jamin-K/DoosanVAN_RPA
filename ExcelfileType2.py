# 신규생성 : 2022.01.31 김재민
# 개요 : 6000ANSAN 파일에 대한 데이터 추출 함수

import pandas as pd
import numpy as np
import openpyxl

# 변수선언 START

# 변수선언 END

# DataFrame 기본 옵션 세팅 START
pd.set_option('display.max_seq_items', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
# DataFrame 기본 옵션 세팅 END

def getStartData(fileName) :
    print('START : %s' %fileName)
    print('▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼')
    excelDataFrame = pd.read_excel(fileName, usecols=[4, 9, 12]) # 품번, 납품잔량, 납기일자

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
    # 데이터가 1개일 때 실행되는 로직 START
    if (len(excelDataFrame) == 1) :
        orderCount = excelDataFrame.iloc[0, 1]
        print('납품잔량 합계 : %d' % orderCount)
        print('품번 : %s' % excelDataFrame.iloc[0, 0])
        print('납기일자 : %s' % excelDataFrame.iloc[0, 2])
    # 데이터가 1개일 때 실행되는 로직 END

    # 데이터가 2개일 때 실행되는 로직 START
    elif (len(excelDataFrame) == 2):
        if (excelDataFrame.iloc[0, 0] == excelDataFrame.iloc[1, 0]):
            if (excelDataFrame.iloc[0, 2] == excelDataFrame.iloc[1, 2]):
                # 동일품번 동일납기
                orderCount = excelDataFrame.iloc[0, 1] + excelDataFrame.iloc[1, 1]
                print('납품수량 합계 : %d' % orderCount)
                print('품번 : %s' % excelDataFrame.iloc[0, 0])
                print('납기일자 : %s' % excelDataFrame.iloc[0, 2])

            else:
                # 동일품번 다른납기
                orderCount = excelDataFrame.iloc[0, 1]
                print('납품잔량 합계 : %d' % orderCount)
                print('품번 : %s' % excelDataFrame.iloc[0, 0])
                print('납기일자 : %s' % excelDataFrame.iloc[0, 2])
                print('-----------------------------------------------------')
                orderCount = excelDataFrame.iloc[1, 1]
                print('납품잔량 합계 : %d' % orderCount)
                print('품번 : %s' % excelDataFrame.iloc[1, 0])
                print('납기일자 : %s' % excelDataFrame.iloc[1, 2])
        else :
            # 다른품번
            orderCount = excelDataFrame.iloc[0, 1]
            print('납품잔량 합계 : %d' % orderCount)
            print('품번 : %s' % excelDataFrame.iloc[0, 0])
            print('납기일자 : %s' % excelDataFrame.iloc[0, 2])
            print('-----------------------------------------------------')
            orderCount = excelDataFrame.iloc[1,1]
            print('납품잔량 합계 : %d' % orderCount)
            print('품번 : %s' % excelDataFrame.iloc[1, 0])
            print('납기일자 : %s' % excelDataFrame.iloc[1, 2])
    # 데이터가 2개일 때 실행되는 로직 END

    # 데이터가 3개 이상일 때 실행되는 로직 START
    else:
        for i in range(len(excelDataFrame) - 1):
            print('-----------------------------------------------------')
            if (excelDataFrame.iloc[i, 0] == excelDataFrame.iloc[i + 1, 0]):
                if (excelDataFrame.iloc[i, 2] == excelDataFrame.iloc[i + 1, 2]):
                    if (i >= len(excelDataFrame) - 2):
                        checkValue = True
                        orderCount = orderCount + excelDataFrame.iloc[i + 1, 1]
                        print('납품수량 합계 : %d' % orderCount)
                        print('품번 : %s' % excelDataFrame.iloc[i, 0])
                        print('납품잔량 : %d' % orderCount)
                        print('납기일자 : %s' % excelDataFrame.iloc[i, 2])
                        print('-----------------------------------------------------')

                    # 원래 동일품번 동일납기 로직 수행
                    checkValue = True
                    orderCount = orderCount + excelDataFrame.iloc[i + 1, 1]
                    print('납품수량 합계 : %d' % orderCount)
                    print('품번 : %s' % excelDataFrame.iloc[i + 1, 0])
                    print('납품잔량 : %d' % excelDataFrame.iloc[i + 1, 0])
                    print('납기일자 : %s' % excelDataFrame.iloc[i + 1, 2])
                    print('-----------------------------------------------------')

                else:
                    # 원래 동일품번 다른납기 로직 수행
                    # orderCount 기록 후 0으로 초기화
                    # 1 orderCount 기록
                    if (checkValue == True):
                        checkValue = False
                        orderCount = orderCount + excelDataFrame.iloc[i, 1]
                    else:
                        orderCount = orderCount + excelDataFrame.iloc[i, 1]
                    print('납품수량 합계 : %d' % orderCount)
                    print('품번 : %s' % excelDataFrame.iloc[i, 0])
                    print('납품잔량 : %d' % excelDataFrame.iloc[i, 1])
                    print('납기일자 : %s' % excelDataFrame.iloc[i, 2])
                    print('-----------------------------------------------------')
                    # 2 orderCount = 0 초기화
                    orderCount = 0
                    # 마지막 index 수행
                    if (i == len(excelDataFrame) - 2):
                        print('마지막 index 실행')
                        orderCount = orderCount + excelDataFrame.iloc[i + 1, 1]
                        print('납품수량 합계 : %d' % orderCount)
                        print('품번 : %s' % excelDataFrame.iloc[i + 1, 0])
                        print('납품잔량 : %d' % excelDataFrame.iloc[i + 1, 1])
                        print('납기일자 : %s' % excelDataFrame.iloc[i + 1, 2])
                        orderCount = 0

            else:
                # 다른 품번
                # orderCount 기록 후 0으로 초기화
                # 1 orderCount 기록
                if (checkValue == True):
                    checkValue = False
                    orderCount = orderCount + excelDataFrame.iloc[i, 1]
                else:
                    orderCount = orderCount + excelDataFrame.iloc[i, 1]
                print('납품수량 합계 : %d' % orderCount)
                print('품번 : %s' % excelDataFrame.iloc[i, 0])
                print('납품잔량 : %d' % excelDataFrame.iloc[i, 1])
                print('납기일자 : %s' % excelDataFrame.iloc[i, 2])
                print('-----------------------------------------------------')
                # 2 orderCount = 0 초기화
                orderCount = 0
                # 마지막 index 수행
                if (i == len(excelDataFrame) - 2):
                    print('마지막 index 실행')
                    orderCount = orderCount + excelDataFrame.iloc[i + 1, 1]
                    print('납품수량 합계 : %d' % orderCount)
                    print('품번 : %s' % excelDataFrame.iloc[i + 1, 0])
                    print('납품잔량 : %d' % excelDataFrame.iloc[i + 1, 1])
                    print('납기일자 : %s' % excelDataFrame.iloc[i + 1, 2])
                    print('-----------------------------------------------------')
                    orderCount = 0
    # 데이터가 3개 이상일 때 실행되는 로직 END
    # 데이터 추출 로직 END

    print('▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲')