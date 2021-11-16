import win32com.client
import time

'''
4-1. 거래량 분석을 통한 대박 주 포착
※ 대박주 기준
1) 대량 거래(거래량이 1000%이상 급증) 종목
2) 대량 거래 시점에서 PBR이 4보다 작음
'''

ins = win32com.client.Dispatch("CpSysDib.StockChart")

# 1. 대신증권이 대량 거래 종목인지 확인
# 대신증권의 최근 60일치 거래량 저장
ins.SetInputValue(0, "A003540")
ins.SetInputValue(1, ord('2'))
ins.SetInputValue(4, 60)
ins.SetInputValue(5, 8)
ins.SetInputValue(6, ord('D'))
ins.SetInputValue(9, ord('1'))

ins.BlockRequest()

volumes = []
numData = ins.GetHeaderValue(3)
for i in range(numData):
    v = ins.GetDataValue(0, i)
    volumes.append(v)

print(volumes)

# 거래량이 1000% 급증했는가
# volumes[0] = 가장 최근 거래일의 거래량이므로, 이를 제외한 나머지 59개의 거래량 평균을 구하자
# 먼저, 60일치 거래량의 총합에 최근 거래량을 빼고, 59일치 거래량의 합을 계산하자
# 그리고, 이 값을 59로 나누면 평균 거래량!
avg = (sum(volumes) - volumes[0]) / (len(volumes) - 1)

if(volumes[0] > avg*10):
    print("대박 주")
else:
    print("일반 주", volumes[0] / avg) # 최근 거래량과 평균 거래량의 비율 출력


# 2. 한 종목에 대한 대박 주 찾기
# 함수생성
def CheckVolume(ins, code):
    # setting
    ins.SetInputValue(0, code)
    ins.SetInputValue(1, ord('2'))
    ins.SetInputValue(4, 60)
    ins.SetInputValue(5, 8)
    ins.SetInputValue(6, ord('D'))
    ins.SetInputValue(9, ord('1'))

    # request
    ins.BlockRequest()

    # getData
    volumes = []
    numData = ins.GetHeaderValue(3)
    for i in range(numData):
        v = ins.GetDataValue(0, i)
        volumes.append(v)

    # calculate average
    avg = (sum(volumes) - volumes[0]) / (len(volumes) - 1)

    if(volumes[0] > avg * 10):
        return 1
    else:
        return 0

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")

buyList = []
codeList = instCpCodeMgr.GetStockListByMarket(1)
for code in codeList[:100]:  # 100개만 요청하기
    if CheckVolume(ins, code) == 1:
        buyList.append(code)
        print(code)
    time.sleep(1) # 연속요청횟수 제한을 피하기 위해 1초 쉬기

'''
result
A000020
'''