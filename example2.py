import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QTimer

import signal
signal.signal(signal.SIGINT, signal.SIG_DFL)

class KiwoomAPI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self._set_signal_slots()
        self.kiwoom.dynamicCall("CommConnect()")

    def _set_signal_slots(self):
        self.kiwoom.OnEventConnect.connect(self._event_connect) 
        self.kiwoom.OnReceiveTrData.connect(self._receive_tr_data)       

    def _event_connect(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        else:
            print(f"로그인 실패: {err_code}")
        self.after_login()

    def after_login(self):
        self.get_basic_stock_info("005930")

    def get_basic_stock_info(self, stock_code):
        if stock_code is not None and len(stock_code) == 6:
            self.set_input_value("종목코드", stock_code)
            self.comm_rq_data("opt10001_req", "opt10001", 0, "5000")

    def set_input_value(self, id, input_value):
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", id, input_value)
    
    def comm_rq_data(self, rq_name, tr_code, prev_next, screen_no):
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", rq_name, tr_code, prev_next, screen_no)

    def _receive_tr_data(self, screen_no, rq_name, tr_code, record_name, prev_next, data_length, error_code, message, splm_msg):
        if rq_name == "opt10001_req":
            self.on_opt10001_req(tr_code,rq_name)

    def get_comm_data(self, tr_code, rq_name, index, item_name):
        ret = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, index, item_name) 
        return ret.strip()     

    def get_repeat_cnt(self, tr_code, rq_name):
        ret = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", tr_code, rq_name)
        return ret

    def on_opt10001_req(self, tr_code, rq_name):
        종목코드 = self.get_comm_data(tr_code, rq_name, 0, "종목코드").replace("A", "").strip()
        PER = float(self.get_comm_data(tr_code, rq_name, 0, "PER").strip())
        시가 = abs(int(self.get_comm_data(tr_code, rq_name, 0, "시가")))
        저가 = abs(int(self.get_comm_data(tr_code, rq_name, 0, "저가")))
        현재가 = abs(int(self.get_comm_data(tr_code, rq_name, 0, "현재가")))
        상한가 = abs(int(self.get_comm_data(tr_code, rq_name, 0, "상한가")))
        하한가 = abs(int(self.get_comm_data(tr_code, rq_name, 0, "하한가")))
        기준가 = abs(int(self.get_comm_data(tr_code, rq_name, 0, "기준가")))
        print(f"종목코드: {종목코드}, PER: {PER}, 시가: {시가}, 저가: {저가}, 현재가: {현재가}, 상한가: {상한가}, 하한가: {하한가}, 기준가: {기준가}")
        
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom_api =KiwoomAPI()
    sys.exit(app.exec_())
    timer = QTimer()
    timer.start(200)
    timer.timeout.connect(lambda: None)
        