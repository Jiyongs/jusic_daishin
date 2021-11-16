import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
import win32com.client

form_class = uic.loadUiType("pyqt\\pytrader.ui")[0]
instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 주문 초기화
objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")

initCheck = objTrade.TradeInit(0)

if (initCheck != 0):
    print("주문 초기화 실패")
    exit()

def getAccount():
    return objTrade.AccountNumber[0] #계좌번호

def orderTrade(account, code, count, price, kind, order):
    # 주식 매수 주문
    accFlag = objTrade.GoodsList(account, 1)  # 주식상품 구분
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, int(order))   # 2: 매수
    objStockOrder.SetInputValue(1, account)   #  계좌번호
    objStockOrder.SetInputValue(2, accFlag[0])   # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, code)   # 종목코드 - 필요한 종목으로 변경 필요
    objStockOrder.SetInputValue(4, count)   # 매수수량 - 요청 수량으로 변경 필요
    objStockOrder.SetInputValue(5, price)   # 주문단가 - 필요한 가격으로 변경 필요
    objStockOrder.SetInputValue(7, "0")   # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
    objStockOrder.SetInputValue(8, kind)   # 주문호가 구분코드 - 01: 보통

    # 매수 주문 요청
    nRet = objStockOrder.BlockRequest()
    if (nRet != 0) :
        print("주문요청 오류", nRet)
        # 0: 정상,  그 외 오류, 4: 주문요청제한 개수 초과
        exit()

    rqStatus = objStockOrder.GetDibStatus()
    errMsg = objStockOrder.GetDibMsg1()
    if rqStatus != 0:
        print("주문 실패: ", rqStatus, errMsg)
        exit()
    else:
        print("주문 성공! ")


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        account = getAccount()
        self.lineEdit_account.setText(account)

        # line edit
        self.lineEdit_name.textChanged.connect(self.code_changed)

        # pushButton
        self.pushButton.clicked.connect(self.btn_clicked)

    def code_changed(self):
        code = self.lineEdit_name.text()
        name = instCpCodeMgr.CodeToName(code)
        self.lineEdit_code.setText(name)

    def btn_clicked(self):
        order_type_lookup = {'신규매도': "1", '신규매수': "2"}
        hoga_lookup = {'지정가': "01", '시장가': "03"}

        account = self.lineEdit_account.text()
        order = order_type_lookup[self.comboBox_order.currentText()]
        code = self.lineEdit_name.text()
        kind = hoga_lookup[self.comboBox_kind.currentText()]
        count = self.spinBox_count.value()
        price = self.spinBox_price.value()

        orderTrade(account, code, count, price, kind, order)

    def closeEvent(self, event):
        self.deleteLater()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()

    #동화약품 : A000020