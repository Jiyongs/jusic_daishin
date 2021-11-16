import win32com.client

'''
4-2. 업종별 PER 분석을 통한 유망 종목 찾기
'''

ins = win32com.client.Dispatch("CpUtil.CpCodeMgr")

# 1. 업종별 평균 PER 계산
# 업종별 코드 리스트 얻기
codeList = ins.GetIndustryList()

for c in codeList:
    print(c, ins.GetIndustryName(c))

# 음식료품 업종의 종목 리스트 얻기
foodCodeList = ins.GetGroupCodeList(5)

for c in foodCodeList:
    print(c, ins.CodeToName(c))

# 음식료품 업종의 평균 PER 계산하기
insM = win32com.client.Dispatch("CpSysDib.MarketEye")

# get PER
insM.SetInputValue(0, 67)
insM.SetInputValue(1, foodCodeList)

insM.BlockRequest()

numStock = insM.GetHeaderValue(2)

sumPer = 0
for i in range(numStock):
    sumPer += insM.GetDataValue(0, i)

print("Average PER : ", sumPer / numStock)
# result
# Average PER :  12.772765950953707

