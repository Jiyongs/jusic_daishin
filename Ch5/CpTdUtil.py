import win32com.client

'''
5-2. 모의투자 매수/매도하기
*주의* 주식시장 운영 시간에 코드 실행해야 함
'''

ins = win32com.client.Dispatch("CpTrade.CpTdUtil")
ins2 = win32com.client.Dispatch("CpTrade.CpTd0311")

# 초기화 > 대화상자 떠야 함 ... 안 뜨넹
ins.TradeInit()

# 대신증권 10주를 13000원에 매수
accountNum = ins.AccountNumber[0]
ins2.SetInputValue(0, 2)            # 1:매도, 2:매수
ins2.SetInputValue(1, accountNum)   # 계좌번호
ins2.SetInputValue(3, "A003540")    # 종목코드
ins2.SetInputValue(4, 10)           # 주문 수량
ins2.SetInputValue(5, 21000)        # 주문 단가

ins2.BlockRequest()

