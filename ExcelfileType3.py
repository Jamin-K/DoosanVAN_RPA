# 신규생성 : 2022.01.31 김재민
# 개요 : 1000JISINCHEON, 1111JISGUNSAN 파일에 대한 데이터 추출 함수

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

def getStartData(fileName):
    print('START : %s' % fileName)
    print('▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼')
    excelDataFrame = pd.read_excel(fileName, usecols=[6, 9, 15, 29, 33])  # 발주번호, 품번, Category, 납기일, 요청수량

    # 발주번호가 NULL값인 행 제거 START
    excelDataFrame['발주번호'].replace('', np.nan, inplace=True)
    excelDataFrame.dropna(subset=['발주번호'], inplace=True)
    # 발주번호가 NULL값인 행 제거 END

    # Category 컬럼이 Q이거나 R이 아닌 행 제거 START
    for index, row in excelDataFrame.iterrows():
        if (row['Category'] != 'Q' and row['Category'] != 'R'):
            excelDataFrame.drop(index, inplace=True)
        else:
            continue
    # Category 컬럼이 Q이거나 R이 아닌 행 제거 END

    # 발주번호 str로 형변환 START
    excelDataFrame['발주번호'] = excelDataFrame['발주번호'].astype(str)  # 차후 소수점 자르는 로직 필요
    # 발주번호 str로 형변환 END

    # dataframe index 재선언 START
    print('전체 인덱스 개수 : %d' % len(excelDataFrame))
    print('-----------------------------------------------------')
    newIdxArr = []
    for i in range(len(excelDataFrame)):
        newIdxArr.append(i)
    excelDataFrame.set_index(keys=[newIdxArr], inplace=True)
    print('-----------------------------------------------------')
    # dataframe index 재선언 END

    orderCount = 0
    checkValue = False

    # 데이터 추출 로직 START
    # 데이터가 1개 있을 때 실행되는 로직 START
    if (len(excelDataFrame) == 1):
        # print('데이터가 1개, 예외처리 필요')
        orderCount = excelDataFrame.iloc[i, 4]
        print('납품수량 합계 : %d' % orderCount)
        print('발주번호 : %s' % excelDataFrame.iloc[i, 0])  # 발주번호
        print('품번 : %s' % excelDataFrame.iloc[i, 1])  # 품번
        print('Category : %s' % excelDataFrame.iloc[i, 2])  # Category
        print('납기일 : %s' % excelDataFrame.iloc[i, 3])  # 납기일
        print('요청수량 : %d' % excelDataFrame.iloc[i, 4])  # 요청수량 INTEGER
    # 데이터가 1개 있을 때 실행되는 로직 END

    # 데이터가 2개 있을 떄 실행되는 로직 START
    elif (len(excelDataFrame) == 2):
        if (excelDataFrame.iloc[0, 1] == excelDataFrame.iloc[1, 1]):
            if (excelDataFrame.iloc[0, 3] == excelDataFrame.iloc[1, 3]):
                # 동일품번 동일납기
                orderCount = excelDataFrame.iloc[0, 4] + excelDataFrame.iloc[1, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[0, 0])  # 발주번호
                print('품번 : %s' % excelDataFrame.iloc[0, 1])  # 품번
                print('Category : %s' % excelDataFrame.iloc[0, 2])  # Category
                print('납기일 : %s' % excelDataFrame.iloc[0, 3])  # 납기일
                print('요청수량 : %d' % excelDataFrame.iloc[0, 4])  # 요청수량 INTEGER
                print('-----------------------------------------------------')
            else:
                # 동일품번 다른납기
                orderCount = excelDataFrame.iloc[0, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[0, 0])  # 발주번호
                print('품번 : %s' % excelDataFrame.iloc[0, 1])  # 품번
                print('Category : %s' % excelDataFrame.iloc[0, 2])  # Category
                print('납기일 : %s' % excelDataFrame.iloc[0, 3])  # 납기일
                print('요청수량 : %d' % excelDataFrame.iloc[0, 4])  # 요청수량 INTEGER
                print('-----------------------------------------------------')
                orderCount = excelDataFrame.iloc[1, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[1, 0])  # 발주번호
                print('품번 : %s' % excelDataFrame.iloc[1, 1])  # 품번
                print('Category : %s' % excelDataFrame.iloc[1, 2])  # Category
                print('납기일 : %s' % excelDataFrame.iloc[1, 3])  # 납기일
                print('요청수량 : %d' % excelDataFrame.iloc[1, 4])  # 요청수량 INTEGER
                print('-----------------------------------------------------')
        else:
            # 다른품번
            orderCount = excelDataFrame.iloc[0, 4]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % excelDataFrame.iloc[0, 0])  # 발주번호
            print('품번 : %s' % excelDataFrame.iloc[0, 1])  # 품번
            print('Category : %s' % excelDataFrame.iloc[0, 2])  # Category
            print('납기일 : %s' % excelDataFrame.iloc[0, 3])  # 납기일
            print('요청수량 : %d' % excelDataFrame.iloc[0, 4])  # 요청수량 INTEGER
            print('-----------------------------------------------------')
            orderCount = excelDataFrame.iloc[1, 4]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % excelDataFrame.iloc[1, 0])  # 발주번호
            print('품번 : %s' % excelDataFrame.iloc[1, 1])  # 품번
            print('Category : %s' % excelDataFrame.iloc[1, 2])  # Category
            print('납기일 : %s' % excelDataFrame.iloc[1, 3])  # 납기일
            print('요청수량 : %d' % excelDataFrame.iloc[1, 4])  # 요청수량 INTEGER
            print('-----------------------------------------------------')
    # 데이터가 2개 있을 때 실행되는 로직 END

    # 데이터가 3개 이상 있을 때 실행되는 로직 START
    for i in range(len(excelDataFrame) - 1):
        print('-----------------------------------------------------')
        if (excelDataFrame.iloc[i, 1] == excelDataFrame.iloc[i + 1, 1]):
            if (excelDataFrame.iloc[i, 3] == excelDataFrame.iloc[i + 1, 3]):
                if (i >= len(excelDataFrame) - 2):
                    checkValue = True
                    orderCount = orderCount + excelDataFrame.iloc[i + 1, 4]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % excelDataFrame.iloc[i, 0])  # 발주번호
                    print('품번 : %s' % excelDataFrame.iloc[i, 1])  # 품번
                    print('Category : %s' % excelDataFrame.iloc[i, 2])  # Category
                    print('납기일 : %s' % excelDataFrame.iloc[i, 3])  # 납기일
                    print('요청수량 : %d' % excelDataFrame.iloc[i, 4])  # 요청수량 INTEGER
                    print('-----------------------------------------------------')

                # 원래 동일품번 동일납기 로직 수행
                checkValue = True
                orderCount = orderCount + excelDataFrame.iloc[i + 1, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[i + 1, 0])  # 발주번호
                print('품번 : %s' % excelDataFrame.iloc[i + 1, 1])  # 품번
                print('Category : %s' % excelDataFrame.iloc[i + 1, 2])  # Category
                print('납기일 : %s' % excelDataFrame.iloc[i + 1, 3])  # 납기일
                print('요청수량 : %d' % excelDataFrame.iloc[i + 1, 4])  # 요청수량 INTEGER
                print('-----------------------------------------------------')

            else:
                # 원래 동일품번 다른납기 로직 수행
                # orderCount 기록 후 0으로 초기화
                # 1 orderCount 기록
                if (checkValue == True):
                    checkValue = False
                    orderCount = orderCount + excelDataFrame.iloc[i, 4]
                else:
                    orderCount = orderCount + excelDataFrame.iloc[i, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[i, 0])  # 발주번호
                print('품번 : %s' % excelDataFrame.iloc[i, 1])  # 품번
                print('Category : %s' % excelDataFrame.iloc[i, 2])  # Category
                print('납기일 : %s' % excelDataFrame.iloc[i, 3])  # 납기일
                print('요청수량 : %d' % excelDataFrame.iloc[i, 4])  # 요청수량 INTEGER
                print('-----------------------------------------------------')
                # 2 orderCount = 0 초기화
                orderCount = 0
                # 마지막 index 수행
                if (i == len(excelDataFrame) - 2):
                    print('마지막 index 실행')
                    orderCount = orderCount + excelDataFrame.iloc[i + 1, 4]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % excelDataFrame.iloc[i + 1, 0])  # 발주번호
                    print('품번 : %s' % excelDataFrame.iloc[i + 1, 1])  # 품번
                    print('Category : %s' % excelDataFrame.iloc[i + 1, 2])  # Category
                    print('납기일 : %s' % excelDataFrame.iloc[i + 1, 3])  # 납기일
                    print('요청수량 : %d' % excelDataFrame.iloc[i + 1, 4])  # 요청수량 INTEGER
                    print('-----------------------------------------------------')
                    orderCount = 0

        else:
            # 다른 품번
            # orderCount 기록 후 0으로 초기화
            # 1 orderCount 기록
            if (checkValue == True):
                checkValue = False
                orderCount = orderCount + excelDataFrame.iloc[i, 4]
            else:
                orderCount = orderCount + excelDataFrame.iloc[i, 4]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % excelDataFrame.iloc[i, 0])  # 발주번호
            print('품번 : %s' % excelDataFrame.iloc[i, 1])  # 품번
            print('Category : %s' % excelDataFrame.iloc[i, 2])  # Category
            print('납기일 : %s' % excelDataFrame.iloc[i, 3])  # 납기일
            print('요청수량 : %d' % excelDataFrame.iloc[i, 4])  # 요청수량 INTEGER
            print('-----------------------------------------------------')
            # 2 orderCount = 0 초기화
            orderCount = 0
            # 마지막 index 수행
            if (i == len(excelDataFrame) - 2):
                print('마지막 index 실행')
                orderCount = orderCount + excelDataFrame.iloc[i + 1, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % excelDataFrame.iloc[i + 1, 0])  # 발주번호
                print('품번 : %s' % excelDataFrame.iloc[i + 1, 1])  # 품번
                print('Category : %s' % excelDataFrame.iloc[i + 1, 2])  # Category
                print('납기일 : %s' % excelDataFrame.iloc[i + 1, 3])  # 납기일
                print('요청수량 : %d' % excelDataFrame.iloc[i + 1, 4])  # 요청수량 INTEGER
                print('-----------------------------------------------------')
                orderCount = 0
    # 데이터가 3개 이상 있을 때 실행되는 로직 END
    # 데이터 추출 로직 END

    print('▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲')
