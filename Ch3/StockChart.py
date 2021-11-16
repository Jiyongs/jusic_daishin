import win32com.client

'''
3-2. 과거 데이터 구하기
'''
ins = win32com.client.Dispatch("CpSysDib.StockChart")

# request/reply 방식 호출
# 1. 대신증권의 최근 10일간의 종가를 가져오기
ins.SetInputValue(0, "A003540")
ins.SetInputValue(1, ord('2'))  # ord : 유니코드 반환
ins.SetInputValue(4, 10)        # 과거데이터 기간 : 10일
ins.SetInputValue(5, 5)         # 과거데이터 종류 : 종가
ins.SetInputValue(6, ord('D'))
ins.SetInputValue(9, ord('1'))

ins.BlockRequest()

# 수신 데이터 개수 확인
numData = ins.GetHeaderValue(3)
# 수신 데이터의 종가 출력
for i in range(numData):
    print(ins.GetDataValue(0,i)) # 첫번째인자 : 수신데이터의 인덱스


# 2. 대신증권의 일자별 시가, 고가, 저가, 종가, 거래량 가져오기
ins.SetInputValue(5, (0,2,3,4,5,8))

ins.BlockRequest()

numData = ins.GetHeaderValue(3)
numField = ins.GetHeaderValue(1)

# 일자별로 6개의 데이터가 반환되므로 중첩 for문 사용하여 데이터 출력
for i in range(numData):
    for j in range(numField):
        print(ins.GetDataValue(j,i), end=" ")
    print()

# 3. 대신증권의 특정 기간 내 시가, 고가, 저가, 종가, 거래량 가져오기
ins.SetInputValue(1, ord('1'))
ins.SetInputValue(2, 20210110) # 종료날짜
ins.SetInputValue(3, 20210101) # 시작날짜

ins.BlockRequest()

numData = ins.GetHeaderValue(3)
numField = ins.GetHeaderValue(1)

for i in range(numData):
    for j in range(numField):
        print(ins.GetDataValue(j,i), end=" ")
    print()
