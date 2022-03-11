# 신규생성 : 2022.03.11 김재민
# 개요 : VAN 엑셀 파일을 읽어와 데이터 가공 후 엑셀로 저장
# 수정 :

from openpyxl import load_workbook
import pandas as pd

# 변수선언 START
processingFileName = None
# 변수선언 END
def startProcessing(vanFileName, filePath) :
    # filePath : 'C:/Users/KJM/Desktop/DSVAN'+todayDate

    # DataFrame 기본 옵션 세팅 START
    pd.set_option('display.max_seq_items', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    # DataFrame 기본 옵션 세팅 END

    if '1000DirINCHEON' in vanFileName :
        print('전처리 시작')
        presentFilePath = filePath + '/수행예정데이터/' + '1000DirINCHEON.xlsx'
        tempDateFrame = pd.read_excel(filePath + '/' + vanFileName, usecols=[10, 16, 17, 19, 31, 38],
                                      dtype={'발주번호':str,
                                             '발주항번':str})
        tempDateFrame.to_excel(presentFilePath, header=True)

    elif '1000INCHEON' in vanFileName :
        print('전처리 시작')
        presentFilePath = filePath + '/수행예정데이터/' + '1000INCHEON.xlsx'
        tempDateFrame = pd.read_excel(filePath + '/' + vanFileName, usecols=[10, 16, 17, 19, 31, 38],
                                      dtype={'발주번호': str,
                                             '발주항번': str})
        tempDateFrame.to_excel(presentFilePath, header=True)

    elif '1100CKD' in vanFileName :
        print('전처리 시작')
        presentFilePath = filePath + '/수행예정데이터/' + '1100CKD.xlsx'
        tempDateFrame = pd.read_excel(filePath + '/' + vanFileName, usecols=[10, 16, 17, 19, 31, 38],
                                      dtype={'발주번호': str,
                                             '발주항번': str})
        tempDateFrame.to_excel(presentFilePath, header=True)

    elif '1130INCHEON' in vanFileName :
        print('전처리 시작')
        presentFilePath = filePath + '/수행예정데이터/' + '1130INCHEON.xlsx'
        tempDateFrame = pd.read_excel(filePath + '/' + vanFileName, usecols=[10, 16, 17, 19, 31, 38],
                                      dtype={'발주번호': str,
                                             '발주항번': str})
        tempDateFrame.to_excel(presentFilePath, header=True)

    elif '6000ANSAN' in vanFileName :
        print('전처리 시작')
        presentFilePath = filePath + '/수행예정데이터/' + '6000ANSAN.xlsx'
        tempDateFrame = pd.read_excel(filePath + '/' + vanFileName, usecols=[4, 9, 12, 45, 46],
                                      dtype={'발주번호': str,
                                             '발주항번': str})
        tempDateFrame.to_excel(presentFilePath, header=True)

    elif '1000JISINCHEON' in vanFileName :
        print('전처리 시작')
        presentFilePath = filePath + '/수행예정데이터/' + '1000JISINCHEON.xlsx'
        tempDateFrame = pd.read_excel(filePath + '/' + vanFileName, usecols=[6, 8, 9, 15, 29, 33],
                                      dtype={'발주번호': str,
                                             '발주항번': str})
        tempDateFrame.to_excel(presentFilePath, header=True)

    elif '1111JISGUNSAN' in vanFileName :
        print('전처리 시작')
        presentFilePath = filePath + '/수행예정데이터/' + '1111JISGUNSAN.xlsx'
        tempDateFrame = pd.read_excel(filePath + '/' + vanFileName, usecols=[6, 8, 9, 15, 29, 33],
                                      dtype={'발주번호': str,
                                             '발주항번': str})
        tempDateFrame.to_excel(presentFilePath, header=True)

    else :
        print('전처리 조건 없음!!')

