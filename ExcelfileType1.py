# 신규생성 : 2022.01.31 김재민
# 개요 : 1000DirINCHEON, 1000INCHEON, 1100CKD, 1130INCHOEN 파일에 대한 데이터 추출 함수
# 수정 : 2022.02.03 김재민 : checkValue 값 체크로직 삭제
#       2022.02.03 김재민 : 발주계획 엑셀 파일 write 함수 call 부분 추가 #001
#       2022.02.03 김재민 : 품번, 납기일 전역변수 추가 및 납기일 Split #002
#       2022.02.04 김재민 : startWriteCell() 함수 호출을 위한 변수선언 및 함수 호출 #003
#       2022.02.13 김재민 : 데이터가 1개일떄, 2개일때 함수 call 로직 추가
#       2022.03.24 김재민 : 납품처 탐색 범위 하드코딩에서 다이나믹으로 변경 #004
#       2022.07.06 김재민 : 1000INCHEON 파일 수행안함. 1130DirINCHEON 파일 수행 #005

import datetime
import pandas as pd
import WriteReleasePlan
from openpyxl import load_workbook






def getStartData(path, fileName, wbFailedListExcel, todayDate) :
    # input - path : 'C:/Users/KJM/Desktop/DSVAN'+todayDate
    # input - fileName : 1000INCHOEN.xlsx
    # input - wbFailedListExcel : load_workbook(실패한 데이터를 작성할 엑셀)

    # excelDataFrame = pd.read_excel(path+'/'+fileName, usecols=[10, 16, 17, 19, 31, 38],
    #                                dtype={'발주번호': str,
    #                                       '발주항번': str})

    # 변수선언 START
    itemNumber = None;  # 품번 #002
    releaseDate = None;  # 납기일 #002
    rowFr = None
    rowTo = None
    pastRowFr = None #005
    pastRowTo = None #005
    fixColumn = 3
    fixRow = 4
    columnFr = 6
    columnTo = 40
    fileDirPath = 'C:/Users/KJM/Desktop/DSVAN'+todayDate+'/'
    #releaseFileName = fileDirPath + 'doosanReleasePlan' + todayDate + '.xlsx'
    releaseFileName = fileDirPath + '/완료데이터/ReleasePlan.xlsx'
    # fileDirPath = 'C:/Users/KJM/Desktop/DSVAN20220214/'  # TestCode
    # releaseFileName = fileDirPath + 'doosanReleasePlan20220214.xlsx'  # TestCode
    orderNumber = None
    semiOrderNumber = None
    # 변수선언 END

    # DataFrame 기본 옵션 세팅 START
    pd.set_option('display.max_seq_items', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    # DataFrame 기본 옵션 세팅 END

    releaseWorkBook = load_workbook(fileDirPath + 'doosanReleasePlan' + todayDate + '.xlsx')
    releaseWorkSheet = releaseWorkBook.active

    # if '1000INCHEON' in fileName : #005
    #     print('1000INCHEON 파일 시작')
    #     excelDataFrame = pd.read_excel(fileDirPath + '/수행예정데이터/1000INCHEON.xlsx',
    #                                    dtype={'발주번호': str,
    #                                           '발주항번': str})
    #
    #     # 004 START
    #     endOfRow = len(releaseWorkSheet['B'])
    #     startRow = 0
    #     rowCount = 0
    #
    #     for i in range(1, endOfRow) :
    #         if(releaseWorkSheet.cell(i, 2).value == '인천공장' and startRow == 0) :
    #             startRow = i
    #             continue
    #
    #         if(releaseWorkSheet.cell(i, 2).value == '인천공장' and startRow != 0) :
    #             rowCount = rowCount + 1
    #             continue
    #         i = i + 1
    #
    #     rowFr = startRow
    #     rowTo = startRow + rowCount + 1
    #     # 004 END
    #
    #     excelDataFrame.drop(excelDataFrame.columns[0], axis=1, inplace=True)

    if '1130DirINCHEON' in fileName : #005
        print('1130DirINCHEON 파일 시작')
        excelDataFrame = pd.read_excel(fileDirPath + '/수행예정데이터/1130DirINCHEON.xlsx',
                                       dtype={'발주번호': str,
                                              '발주항번': str})

        # 004 START
        endOfRow = len(releaseWorkSheet['B'])
        startRow = 0
        rowCount = 0

        for i in range(1, endOfRow) : #세진크랑크
            if(releaseWorkSheet.cell(i, 2).value == '세진크랑크' and startRow == 0) :
                startRow = i
                continue

            if(releaseWorkSheet.cell(i, 2).value == '세진크랑크' and startRow != 0) :
                rowCount = rowCount + 1
                continue
            i = i + 1

        rowFr = startRow
        rowTo = startRow + rowCount + 1
        # 004 END

        excelDataFrame.drop(excelDataFrame.columns[0], axis=1, inplace=True)


    elif '1000DirINCHEON' in fileName : # 한양정밀
        print('1000DirINCHOEN 파일 시작')
        excelDataFrame = pd.read_excel(fileDirPath + '/수행예정데이터/1000DirINCHEON.xlsx',
                                       dtype={'발주번호': str,
                                              '발주항번': str})

        # 004 START
        endOfRow = len(releaseWorkSheet['B'])
        startRow = 0
        rowCount = 0

        for i in range(1, endOfRow):
            if (releaseWorkSheet.cell(i, 2).value == '한양정밀' and startRow == 0):
                startRow = i
                continue

            if (releaseWorkSheet.cell(i, 2).value == '한양정밀' and startRow != 0):
                rowCount = rowCount + 1
                continue
            i = i + 1

        rowFr = startRow
        rowTo = startRow + rowCount + 1
        # 004 END

        excelDataFrame.drop(excelDataFrame.columns[0], axis=1, inplace=True)
        # rowFr = 5
        # rowTo = 7

    elif '1100CKD' in fileName :
        print('1100CKD 파일 시작')
        excelDataFrame = pd.read_excel(fileDirPath + '/수행예정데이터/1100CKD.xlsx',
                                       dtype={'발주번호': str,
                                              '발주항번': str})

        # 004 START
        endOfRow = len(releaseWorkSheet['B'])
        startRow = 0
        rowCount = 0

        for i in range(1, endOfRow):
            if (releaseWorkSheet.cell(i, 2).value == 'CKD' and startRow == 0):
                startRow = i
                continue

            if (releaseWorkSheet.cell(i, 2).value == 'CKD' and startRow != 0):
                rowCount = rowCount + 1
                continue
            i = i + 1

        rowFr = startRow
        rowTo = startRow + rowCount + 1
        # 004 END

        excelDataFrame.drop(excelDataFrame.columns[0], axis=1, inplace=True)
        # rowFr = 7
        # rowTo = 11

    elif '1130INCHEON' in fileName :
        print('1130INCHOEN 파일 시작')
        excelDataFrame = pd.read_excel(fileDirPath + '/수행예정데이터/1130INCHEON.xlsx',
                                       dtype={'발주번호': str,
                                              '발주항번': str})

        # 004 START
        endOfRow = len(releaseWorkSheet['B'])
        startRow = 0
        rowCount = 0

        for i in range(1, endOfRow):
            if (releaseWorkSheet.cell(i, 2).value == '인천공장' and startRow == 0):
                startRow = i
                continue

            if (releaseWorkSheet.cell(i, 2).value == '인천공장' and startRow != 0):
                rowCount = rowCount + 1
                continue
            i = i + 1

        rowFr = startRow
        rowTo = startRow + rowCount + 1
        # 004 END

        excelDataFrame.drop(excelDataFrame.columns[0], axis=1, inplace=True)
        # rowFr = 11
        # rowTo = 50

    else :
        print('파일 분류 에러 : ExcelfileType1')

    releaseWorkBook.close()
    print('START : %s' %fileName)
    print('rowFr : %d' %rowFr)
    print('rowTo : %d' %rowTo)
    print('endOfRow : %d' %endOfRow)
    print('▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼')

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
                                        , semiOrderNumber, wbFailedListExcel, fileName, todayDate, jisData='N')  # 003
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
                                                , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                                jisData='N')  # 003
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
                                                , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                                jisData='N')  # 003
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
                                                , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                                jisData='N')  # 003
        else:
            # 다른품번
            #orderCount = excelDataFrame.iloc[0, 3]
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
                                            , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                            jisData='N')  # 003
            print('-----------------------------------------------------')
            #orderCount = excelDataFrame.iloc[1, 3]
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
                                            , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                            jisData='N')  # 003
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
                                                        , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                                        jisData='N')  # 003

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
                                                    , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                                    jisData='N')  # 003

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
                                                        , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                                        jisData='N')  # 003

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
                                                , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                                jisData='N')  # 003

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
                                                    , semiOrderNumber, wbFailedListExcel, fileName, todayDate,
                                                    jisData='N')  # 003

                    print('-----------------------------------------------------')
                    orderCount = 0

    # 데이터가 3개 이상 있을 때 실행되는 로직 END
    # 데이터 추출 로직 END

    print('▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲')
