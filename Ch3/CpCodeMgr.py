import win32com.client

'''
3-1. 종목 코드 가져오기
'''
instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")

# 1 : 유가증권시장 종목코드
codeList = instCpCodeMgr.GetStockListByMarket(1)

# print(codeList)

kospi = {}
for code in codeList:
    name = instCpCodeMgr.CodeToName(code)
    kospi[code] = name

# print(kospi)
# csv 파일로 저장
f = open("C:\\Users\\jiyoung\\Documents\\jusic_prac\\kospi.csv", "w")
for k, v in kospi.items():
    f.write("%s, %s\n" % (k, v))
f.close()

# 인덱스, 종목 코드, 부 구분코드, 종목명 출력
for i, code in enumerate(codeList):
    secondCode = instCpCodeMgr.GetStockSectionKind(code)
    name = instCpCodeMgr.CodeToName(code)
    print(i, code, secondCode, name)


def get_master_code_name(self, code):
    instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    name = instCpCodeMgr.CodeToName(code)
    return name
