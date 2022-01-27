# 범위 주석 처리 및 해제 : ctrl and /
# 사용라이브러리, pandas, numpy, xlrd
# 제한사항. RPA를 통해 두산 VAN에서 xls 파일을 받은 다음 xlsx로 변환해야함

from openpyxl import load_workbook
import pandas as pd
import numpy as np

# wb = load_workbook(filename='C:/Users/KJM/Desktop/업무문서/개발/20220108_RPA_두산 VAN/doosanJIS1000CE220111_복사.xlsx')
#
# sheet_range = wb['Sheet1']
#
# testdata1 = sheet_range['C9'].value
# testdata2 = sheet_range['C4'].value
#
# print(testdata1)
# print(testdata2)
#
# if(testdata1 == testdata2) :
#     print('일치')
#
# else:
#     print('불일치')

#---------------------------JIS 파일 데이터 추출 START---------------------------

# 필수세팅값 start
pd.set_option('display.max_seq_items', None) #dataframe 을 생략없이 모두 표현하기
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
# 필수세팅값 end

#dataframe = pd.read_excel('C:/Users/KJM/Desktop/업무문서/개발/20220108_RPA_두산 VAN/doosanJIS1000CE220111_복사.xlsx', engine='openpyxl')

# 파일경로 지정
dataframe = pd.read_excel('C:/Users/KJM/Desktop/doosanJIS1111GUNSAN20220117.xls', usecols=[6, 9, 15, 29, 33]) #품번, 납기일, 요청수량 추출 완료
#dataframe = pd.read_excel('C:/Users/KJM/Desktop/업무문서/개발/20220108_RPA_두산 VAN/doosanJIS1000CE220111_복사.xlsx', usecols=[6, 9, 15, 29, 33]) #품번, 납기일, 요청수량 추출 완료
#dataframe = pd.read_excel('C:/Users/KJM/Desktop/doosanJIS1000INCHEON20220117.xls', usecols=[6, 9, 15, 29, 33]) #품번, 납기일, 요청수량 추출 완료

# 발주번호가 Null값인 행 제거 start, 인덱스는 밀리지 않고 원본 인덱스 그대로 유지
dataframe['발주번호'].replace('', np.nan, inplace=True)
dataframe.dropna(subset=['발주번호'], inplace=True)
# 발주번호가 Null값인 행 제거 end

# 카테고리가 Q이거나 R이 아닌 행 삭제 start
for index1, row in dataframe.iterrows():
    if(row['Category'] != 'Q' and row['Category'] != 'R') :
        dataframe.drop(index1, inplace=True)
    else :
        continue

# 발주번호 str로 형변환 start
dataframe['발주번호'] = dataframe['발주번호'].astype(str) #차후 소수점 자르는 로직 필요
# 발주번호 str로 형변환 end
print(dataframe)
print('-----------------------------------------------------')
# 카테고리가 Q이거나 R이 아닌 행 삭제 end


# dataframe index 순서대로 재선언 start
print('전체 인덱스 개수 : %d'%len(dataframe))
print('-----------------------------------------------------')
newIdxArr = []
for i in range(len(dataframe)):
    newIdxArr.append(i)
dataframe.set_index(keys=[newIdxArr], inplace=True)
print(dataframe)
print('-----------------------------------------------------')
# dataframe index 순서대로 재선언 end

# 품번, 납기일, 요청수량 전처리 start
orderCount = 0
checkValue = False
# 데이터가 1개 있을 때 실행되는 로직 start
if(len(dataframe) == 1) :
    #print('데이터가 1개, 예외처리 필요')
    orderCount = dataframe.iloc[i,4]
    print('납품수량 합계 : %d' % orderCount)
    print('발주번호 : %s' % dataframe.iloc[i, 0])  # 발주번호
    print('품번 : %s' % dataframe.iloc[i, 1])  # 품번
    print('Category : %s' % dataframe.iloc[i, 2])  # Category
    print('납기일 : %s' % dataframe.iloc[i, 3])  # 납기일
    print('요청수량 : %d' % dataframe.iloc[i, 4])  # 요청수량 INTEGER
# 데이터가 1개 있을 때 실행되는 로직 end

# 데이터가 2개 있을 떄 실행되는 로직 start
elif(len(dataframe) == 2) :
    if(dataframe.iloc[0,1] == dataframe.iloc[1,1]) :
        if(dataframe.iloc[0,3] == dataframe.iloc[1,3]) :
            # 동일품번 동일납기
            orderCount = dataframe.iloc[0,4] + dataframe.iloc[1,4]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[0, 0])  # 발주번호
            print('품번 : %s' % dataframe.iloc[0, 1])  # 품번
            print('Category : %s' % dataframe.iloc[0, 2])  # Category
            print('납기일 : %s' % dataframe.iloc[0, 3])  # 납기일
            print('요청수량 : %d' % dataframe.iloc[0, 4])  # 요청수량 INTEGER
            print('-----------------------------------------------------')
        else :
            # 동일품번 다른납기
            orderCount = dataframe.iloc[0,4]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[0, 0])  # 발주번호
            print('품번 : %s' % dataframe.iloc[0, 1])  # 품번
            print('Category : %s' % dataframe.iloc[0, 2])  # Category
            print('납기일 : %s' % dataframe.iloc[0, 3])  # 납기일
            print('요청수량 : %d' % dataframe.iloc[0, 4])  # 요청수량 INTEGER
            print('-----------------------------------------------------')
            orderCount = dataframe.iloc[1, 4]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[1, 0])  # 발주번호
            print('품번 : %s' % dataframe.iloc[1, 1])  # 품번
            print('Category : %s' % dataframe.iloc[1, 2])  # Category
            print('납기일 : %s' % dataframe.iloc[1, 3])  # 납기일
            print('요청수량 : %d' % dataframe.iloc[1, 4])  # 요청수량 INTEGER
            print('-----------------------------------------------------')
    else :
        # 다른품번
        orderCount = dataframe.iloc[0, 4]
        print('납품수량 합계 : %d' % orderCount)
        print('발주번호 : %s' % dataframe.iloc[0, 0])  # 발주번호
        print('품번 : %s' % dataframe.iloc[0, 1])  # 품번
        print('Category : %s' % dataframe.iloc[0, 2])  # Category
        print('납기일 : %s' % dataframe.iloc[0, 3])  # 납기일
        print('요청수량 : %d' % dataframe.iloc[0, 4])  # 요청수량 INTEGER
        print('-----------------------------------------------------')
        orderCount = dataframe.iloc[1, 4]
        print('납품수량 합계 : %d' % orderCount)
        print('발주번호 : %s' % dataframe.iloc[1, 0])  # 발주번호
        print('품번 : %s' % dataframe.iloc[1, 1])  # 품번
        print('Category : %s' % dataframe.iloc[1, 2])  # Category
        print('납기일 : %s' % dataframe.iloc[1, 3])  # 납기일
        print('요청수량 : %d' % dataframe.iloc[1, 4])  # 요청수량 INTEGER
        print('-----------------------------------------------------')

# 데이터가 2개 있을 때 실행되는 로직 end

# 데이터가 3개 이상 있을 때 실행되는 로직 start
for i in range(len(dataframe) - 1) :
    print('-----------------------------------------------------')
    if (dataframe.iloc[i, 1] == dataframe.iloc[i + 1, 1]):
        if (dataframe.iloc[i, 3] == dataframe.iloc[i + 1, 3]):
            if(i >= len(dataframe) - 2) :
                checkValue = True
                orderCount = orderCount + dataframe.iloc[i + 1, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % dataframe.iloc[i, 0])  # 발주번호
                print('품번 : %s' % dataframe.iloc[i, 1])  # 품번
                print('Category : %s' % dataframe.iloc[i, 2])  # Category
                print('납기일 : %s' % dataframe.iloc[i, 3])  # 납기일
                print('요청수량 : %d' % dataframe.iloc[i, 4])  # 요청수량 INTEGER
                print('-----------------------------------------------------')

            # 원래 동일품번 동일납기 로직 수행
            checkValue = True
            orderCount = orderCount + dataframe.iloc[i+1, 4]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[i+1, 0])  # 발주번호
            print('품번 : %s' % dataframe.iloc[i+1, 1])  # 품번
            print('Category : %s' % dataframe.iloc[i+1, 2])  # Category
            print('납기일 : %s' % dataframe.iloc[i+1, 3])  # 납기일
            print('요청수량 : %d' % dataframe.iloc[i+1, 4])  # 요청수량 INTEGER
            print('-----------------------------------------------------')

        else :
            # 원래 동일품번 다른납기 로직 수행
            # orderCount 기록 후 0으로 초기화
            #1 orderCount 기록
            if(checkValue == True) :
                checkValue = False
                orderCount = orderCount + dataframe.iloc[i,4]
            else :
                orderCount = orderCount + dataframe.iloc[i,4]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[i, 0])  # 발주번호
            print('품번 : %s' % dataframe.iloc[i, 1])  # 품번
            print('Category : %s' % dataframe.iloc[i, 2])  # Category
            print('납기일 : %s' % dataframe.iloc[i, 3])  # 납기일
            print('요청수량 : %d' % dataframe.iloc[i, 4])  # 요청수량 INTEGER
            print('-----------------------------------------------------')
            #2 orderCount = 0 초기화
            orderCount = 0
            #마지막 index 수행
            if(i == len(dataframe) - 2) :
                print('마지막 index 실행')
                orderCount = orderCount + dataframe.iloc[i + 1, 4]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % dataframe.iloc[i + 1, 0])  # 발주번호
                print('품번 : %s' % dataframe.iloc[i + 1, 1])  # 품번
                print('Category : %s' % dataframe.iloc[i + 1, 2])  # Category
                print('납기일 : %s' % dataframe.iloc[i + 1, 3])  # 납기일
                print('요청수량 : %d' % dataframe.iloc[i + 1, 4])  # 요청수량 INTEGER
                print('-----------------------------------------------------')
                orderCount = 0

    else :
        # 다른 품번
        # orderCount 기록 후 0으로 초기화
        #1 orderCount 기록
        if(checkValue == True) :
            checkValue = False
            orderCount = orderCount + dataframe.iloc[i,4]
        else :
            orderCount = orderCount + dataframe.iloc[i,4]
        print('납품수량 합계 : %d' %orderCount)
        print('발주번호 : %s' % dataframe.iloc[i, 0])  # 발주번호
        print('품번 : %s' % dataframe.iloc[i, 1])  # 품번
        print('Category : %s' % dataframe.iloc[i, 2])  # Category
        print('납기일 : %s' % dataframe.iloc[i, 3])  # 납기일
        print('요청수량 : %d' % dataframe.iloc[i, 4])  # 요청수량 INTEGER
        print('-----------------------------------------------------')
        #2 orderCount = 0 초기화
        orderCount = 0
        #마지막 index 수행
        if(i == len(dataframe) - 2) :
            print('마지막 index 실행')
            orderCount = orderCount + dataframe.iloc[i + 1, 4]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[i + 1, 0])  # 발주번호
            print('품번 : %s' % dataframe.iloc[i + 1, 1])  # 품번
            print('Category : %s' % dataframe.iloc[i + 1, 2])  # Category
            print('납기일 : %s' % dataframe.iloc[i + 1, 3])  # 납기일
            print('요청수량 : %d' % dataframe.iloc[i + 1, 4])  # 요청수량 INTEGER
            print('-----------------------------------------------------')
            orderCount = 0
# 데이터가 3개 이상 있을 때 실행되는 로직 end
# ---------------------------JIS 파일 데이터 추출 END---------------------------

#
#----------------------------------------------------------------------------------------------------------------------------
# ---------------------------1000INCHEONDir, 1000INCHEON, 1100CKD, 1130INCHEON 파일 데이터 추출 START---------------------------
pd.set_option('display.max_seq_items', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# JIS, 발주번호, 품번, 납품잔량, 납기일자 추출 start
#dataframe = pd.read_excel('C:/Users/KJM/Desktop/doosan1000INCHEON20220117.xls', usecols=[10, 16, 19, 31, 38])
dataframe = pd.read_excel('C:/Users/KJM/Desktop/VAN_20220120/doosan1100CKD20220120.xlsx', usecols=[10, 16, 19, 31, 38])

# JIS, 발주번호, 품번, 납품잔량, 납기일자 추출 end

# JIS 값이 N인 행 drop start
for index, row in dataframe.iterrows():
    if(row['JIS'] == 'Y') :
        dataframe.drop(index, inplace=True)
    else :
        continue
# JIS 값이 N인 행 drop end

# 발주번호 필드 str로 형변환 start
dataframe['발주번호'] = dataframe['발주번호'].astype(str)
# 발주번호 필드 str로 형변환 end

# dataframe index 재선언 start
print('전체 인덱스 개수 : %d' %len(dataframe))
print('-----------------------------------------------------')
newIdxArr = []
for i in range(len(dataframe)) :
    newIdxArr.append((i))
dataframe.set_index(keys=[newIdxArr], inplace=True)
# dataframe index 재선언 end

#풉번, 납기일, 납품잔량 전처리 start
orderCount = 0
checkValue = False

print(dataframe)
print('-----------------------------------------------------')
# 데이터가 1개 있을 때 실행되는 로직 start
if (len(dataframe) == 1):
    orderCount = dataframe.iloc[0, 3]
    print('납품수량 합계 : %d' % orderCount)
    print('발주번호 : %s' % dataframe.iloc[0, 1])
    print('품번 : %s' % dataframe.iloc[0, 2])
    print('납기일 : %s' % dataframe.iloc[0, 4])
    print('납품잔량 : %d' % dataframe.iloc[0, 3])
# 데이터가 1개 있을 때 실행되는 로직 end

# 데이터가 2개 있을 때 실행되는 로직 start
elif(len(dataframe) == 2) :
    if(dataframe.iloc[0,2] == dataframe.iloc[1,2]) :
        if(dataframe.iloc[0,4] == dataframe.iloc[1,4]) :
            # 동일품번, 동일납기
            orderCount = dataframe.iloc[0,3] + dataframe.iloc[1,3]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[0, 1])
            print('품번 : %s' % dataframe.iloc[0, 2])
            print('납기일 : %s' % dataframe.iloc[0, 4])
            print('납품잔량 : %d' % dataframe.iloc[0, 3])
            print('-----------------------------------------------------')
        else :
            # 동일품번, 다른납기
            orderCount = dataframe.iloc[0, 3]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[0, 1])
            print('품번 : %s' % dataframe.iloc[0, 2])
            print('납기일 : %s' % dataframe.iloc[0, 4])
            print('납품잔량 : %d' % dataframe.iloc[0, 3])
            print('-----------------------------------------------------')
            orderCount = dataframe.iloc[1, 3]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[1, 1])
            print('품번 : %s' % dataframe.iloc[1, 2])
            print('납기일 : %s' % dataframe.iloc[1, 4])
            print('납품잔량 : %d' % dataframe.iloc[1, 3])

    else :
        # 다른품번
        orderCount = dataframe.iloc[0, 3]
        print('납품수량 합계 : %d' % orderCount)
        print('발주번호 : %s' % dataframe.iloc[0, 1])
        print('품번 : %s' % dataframe.iloc[0, 2])
        print('납기일 : %s' % dataframe.iloc[0, 4])
        print('납품잔량 : %d' % dataframe.iloc[0, 3])
        print('-----------------------------------------------------')
        orderCount = dataframe.iloc[1, 3]
        print('납품수량 합계 : %d' % orderCount)
        print('발주번호 : %s' % dataframe.iloc[1, 1])
        print('품번 : %s' % dataframe.iloc[1, 2])
        print('납기일 : %s' % dataframe.iloc[1, 4])
        print('납품잔량 : %d' % dataframe.iloc[1, 3])
# 데이터가 2개 있을 때 실행되는 로직 end

# 데이터가 3개 이상 있을 때 실행되는 로직 start
else :
    for i in range(len(dataframe) - 1):
        print('-----------------------------------------------------')
        if (dataframe.iloc[i, 2] == dataframe.iloc[i + 1, 2]):
            if (dataframe.iloc[i, 4] == dataframe.iloc[i + 1, 4]):
                if (i >= len(dataframe) - 2):
                    checkValue = True
                    orderCount = orderCount + dataframe.iloc[i + 1, 3]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % dataframe.iloc[i, 1])  # 발주번호
                    print('품번 : %s' % dataframe.iloc[i, 2])  # 품번
                    print('납기일 : %s' % dataframe.iloc[i, 4])  # 납기일
                    print('요청수량 : %d' % dataframe.iloc[i, 3])  # 요청수량 INTEGER
                    print('-----------------------------------------------------')

                # 원래 동일품번 동일납기 로직 수행
                checkValue = True
                orderCount = orderCount + dataframe.iloc[i + 1, 3]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % dataframe.iloc[i + 1, 1])  # 발주번호
                print('품번 : %s' % dataframe.iloc[i + 1, 2])  # 품번
                print('납기일 : %s' % dataframe.iloc[i + 1, 4])  # 납기일
                print('요청수량 : %d' % dataframe.iloc[i + 1, 3])  # 요청수량 INTEGER
                print('-----------------------------------------------------')

            else:
                # 원래 동일품번 다른납기 로직 수행
                # orderCount 기록 후 0으로 초기화
                # 1 orderCount 기록
                if (checkValue == True):
                    checkValue = False
                    orderCount = orderCount + dataframe.iloc[i, 3]
                else:
                    orderCount = orderCount + dataframe.iloc[i, 3]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % dataframe.iloc[i, 1])  # 발주번호
                print('품번 : %s' % dataframe.iloc[i, 2])  # 품번
                print('납기일 : %s' % dataframe.iloc[i, 4])  # 납기일
                print('요청수량 : %d' % dataframe.iloc[i, 3])  # 요청수량 INTEGER
                print('-----------------------------------------------------')
                # 2 orderCount = 0 초기화
                orderCount = 0
                # 마지막 index 수행
                if (i == len(dataframe) - 2):
                    print('마지막 index 실행')
                    orderCount = orderCount + dataframe.iloc[i + 1, 3]
                    print('납품수량 합계 : %d' % orderCount)
                    print('발주번호 : %s' % dataframe.iloc[i + 1, 1])  # 발주번호
                    print('품번 : %s' % dataframe.iloc[i + 1, 2])  # 품번
                    print('납기일 : %s' % dataframe.iloc[i + 1, 4])  # 납기일
                    print('요청수량 : %d' % dataframe.iloc[i + 1, 3])  # 요청수량 INTEGER
                    print('-----------------------------------------------------')
                    orderCount = 0

        else:
            # 다른 품번
            # orderCount 기록 후 0으로 초기화
            # 1 orderCount 기록
            if (checkValue == True):
                checkValue = False
                orderCount = orderCount + dataframe.iloc[i, 3]
            else:
                orderCount = orderCount + dataframe.iloc[i, 3]
            print('납품수량 합계 : %d' % orderCount)
            print('발주번호 : %s' % dataframe.iloc[i, 1])  # 발주번호
            print('품번 : %s' % dataframe.iloc[i, 2])  # 품번
            print('납기일 : %s' % dataframe.iloc[i, 4])  # 납기일
            print('요청수량 : %d' % dataframe.iloc[i, 3])  # 요청수량 INTEGER
            print('-----------------------------------------------------')
            # 2 orderCount = 0 초기화
            orderCount = 0
            # 마지막 index 수행
            if (i == len(dataframe) - 2):
                print('마지막 index 실행')
                orderCount = orderCount + dataframe.iloc[i + 1, 3]
                print('납품수량 합계 : %d' % orderCount)
                print('발주번호 : %s' % dataframe.iloc[i + 1, 1])  # 발주번호
                print('품번 : %s' % dataframe.iloc[i + 1, 2])  # 품번
                print('납기일 : %s' % dataframe.iloc[i + 1, 4])  # 납기일
                print('요청수량 : %d' % dataframe.iloc[i + 1, 3])  # 요청수량 INTEGER
                print('-----------------------------------------------------')
                orderCount = 0

# 데이터가 3개 이상 있을 때 실행되는 로직 end

#품번, 납기일, 납품잔량 전처리 end

# ---------------------------1000INCHEONDir, 1000INCHEON, 1100CKD, 1130INCHEON 파일 데이터 추출 END ----------------------------
#----------------------------------------------------------------------------------------------------------------------------

# ------------------------------------ 6000ANSAN 파일 데이터 추출 START -------------------------------------
print('-----------------------------------------------------')
pd.set_option('display.max_seq_items', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# 품번, 납품잔량, 납기일 추출 START
dataframe = pd.read_excel('C:/Users/KJM/Desktop/VAN_20220120/doosan6000ANSAN20220122.xlsx', usecols=[4, 9, 12])
# 품번, 납품잔량, 납기일 추출 END

print('전체 인덱스 개수 : %d' %(len(dataframe)))
print('-----------------------------------------------------')
print(dataframe)

# 품번, 납품잔량, 납기일 전처리 START
orderCount = 0
checkValue = False
# 데이터가 1개일 때 실행되는 로직 START
if(len(dataframe) == 1) :
    orderCount = dataframe.iloc[0, 1]
    print('납품잔량 합계 : %d' % orderCount)
    print('품번 : %s' %dataframe.iloc[0, 0])
    print('납기일자 : %s' % dataframe.iloc[0, 2])
# 데이터가 1개일 때 실행되는 로직 END

# 데이터가 2개일 때 실행되는 로직 START
elif(len(dataframe) == 2) :
    if(dataframe.iloc[0,0] == dataframe.iloc[1,0]) :
        if(dataframe.iloc[0,2] == dataframe.iloc[1,2]) :
            # 동일품번 동일납기
            orderCount = dataframe.iloc[0,1] + dataframe.iloc[1,1]
            print('납품수량 합계 : %d' %orderCount)
            print('품번 : %s' % dataframe.iloc[0, 0])
            print('납기일자 : %s' % dataframe.iloc[0, 2])

        else :
            # 동일품번 다른납기
            orderCount = dataframe.iloc[0,1]
            print('납품잔량 합계 : %d' % orderCount)
            print('품번 : %s' % dataframe.iloc[0, 0])
            print('납기일자 : %s' % dataframe.iloc[0, 2])
            print('-----------------------------------------------------')
            orderCount = dataframe.iloc[1,1]
            print('납품잔량 합계 : %d' % orderCount)
            print('품번 : %s' % dataframe.iloc[1, 0])
            print('납기일자 : %s' % dataframe.iloc[1, 2])
# 데이터가 2개일 때 실행되는 로직 END
# 데이터가 3개 이상일 때 실행되는 로직 START
else :
    for i in range(len(dataframe) - 1):
        print('-----------------------------------------------------')
        if (dataframe.iloc[i, 0] == dataframe.iloc[i + 1, 0]):
            if (dataframe.iloc[i, 2] == dataframe.iloc[i + 1, 2]):
                if (i >= len(dataframe) - 2):
                    checkValue = True
                    orderCount = orderCount + dataframe.iloc[i + 1, 1]
                    print('납품수량 합계 : %d' % orderCount)
                    print('품번 : %s' % dataframe.iloc[i, 0])
                    print('납품잔량 : %d' % orderCount)
                    print('납기일자 : %s' % dataframe.iloc[i, 2])
                    print('-----------------------------------------------------')

                # 원래 동일품번 동일납기 로직 수행
                checkValue = True
                orderCount = orderCount + dataframe.iloc[i + 1, 1]
                print('납품수량 합계 : %d' % orderCount)
                print('품번 : %s' % dataframe.iloc[i+1, 0])
                print('납품잔량 : %d' % dataframe.iloc[i+1, 0])
                print('납기일자 : %s' % dataframe.iloc[i+1, 2])
                print('-----------------------------------------------------')

            else:
                # 원래 동일품번 다른납기 로직 수행
                # orderCount 기록 후 0으로 초기화
                # 1 orderCount 기록
                if (checkValue == True):
                    checkValue = False
                    orderCount = orderCount + dataframe.iloc[i, 1]
                else:
                    orderCount = orderCount + dataframe.iloc[i, 1]
                print('납품수량 합계 : %d' % orderCount)
                print('품번 : %s' % dataframe.iloc[i, 0])
                print('납품잔량 : %d' % dataframe.iloc[i, 1])
                print('납기일자 : %s' % dataframe.iloc[i, 2])
                print('-----------------------------------------------------')
                # 2 orderCount = 0 초기화
                orderCount = 0
                # 마지막 index 수행
                if (i == len(dataframe) - 2):
                    print('마지막 index 실행')
                    orderCount = orderCount + dataframe.iloc[i + 1, 1]
                    print('납품수량 합계 : %d' % orderCount)
                    print('품번 : %s' % dataframe.iloc[i+1, 0])
                    print('납품잔량 : %d' % dataframe.iloc[i+1, 1])
                    print('납기일자 : %s' % dataframe.iloc[i+1, 2])
                    orderCount = 0

        else:
            # 다른 품번
            # orderCount 기록 후 0으로 초기화
            # 1 orderCount 기록
            if (checkValue == True):
                checkValue = False
                orderCount = orderCount + dataframe.iloc[i, 1]
            else:
                orderCount = orderCount + dataframe.iloc[i, 1]
            print('납품수량 합계 : %d' % orderCount)
            print('품번 : %s' % dataframe.iloc[i, 0])
            print('납품잔량 : %d' % dataframe.iloc[i, 1])
            print('납기일자 : %s' % dataframe.iloc[i, 2])
            print('-----------------------------------------------------')
            # 2 orderCount = 0 초기화
            orderCount = 0
            # 마지막 index 수행
            if (i == len(dataframe) - 2):
                print('마지막 index 실행')
                orderCount = orderCount + dataframe.iloc[i + 1, 1]
                print('납품수량 합계 : %d' % orderCount)
                print('품번 : %s' % dataframe.iloc[i+1, 0])
                print('납품잔량 : %d' % dataframe.iloc[i+1, 1])
                print('납기일자 : %s' % dataframe.iloc[i+1, 2])
                print('-----------------------------------------------------')
                orderCount = 0
# 데이터가 3개 이상일 때 실행되는 로직 END
# 품번, 납기일, 납품잔량 전처리 END
# print(dataframe)





