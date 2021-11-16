import win32com.client

'''
3-3. PER, EPS 데이터 구하기
'''

ins = win32com.client.Dispatch("CpSysDib.MarketEye")

# 1. 대신증권의 현재가, PER, EPS, 최근분기년월 구하기
ins.SetInputValue(0, (4, 67, 70, 111))
ins.SetInputValue(1, 'A003540')

ins.BlockRequest()

print("현재가 : ", ins.GetDataValue(0,0))
print("PER : ", ins.GetDataValue(1,0))
print("EPS : ", ins.GetDataValue(2,0))
print("최근분기년월 : ", ins.GetDataValue(3,0))

