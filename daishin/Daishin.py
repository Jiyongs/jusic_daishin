import win32com.client

class Daishin():
    def __init__(self):
        super().__init__()

    def get_master_code_name(self, code):
        instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        codeList = instCpCodeMgr.GetStockListByMarket(1) # 1 : 유가증권시장 종목코드
        name = instCpCodeMgr.CodeToName(code)
        return name
