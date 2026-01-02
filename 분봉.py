import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QTimer

import signal
signal.signal(signal.SIGINT, signal.SIG_DFL)
import datetime

import pandas as pd


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
        self.req_opt10080("005930_AL", "1:1분") # 삼성전자 1분봉 데이터 요청

    def set_input_value(self, id, input_value):
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", id, input_value)
    
    def get_comm_data(self, tr_code, rq_name, index, item_name):
        ret = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", tr_code, rq_name, index, item_name) 
        return ret.strip()     

    def get_repeat_cnt(self, tr_code, rq_name):
        ret = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", tr_code, rq_name)
        return ret

    def comm_rq_data(self, rq_name, tr_code, prev_next, screen_no):
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", rq_name, tr_code, prev_next, screen_no)        

    def req_opt10080(self, code,tick_range="1:1분"):
            
        self.set_input_value("종목코드", code)
        self.set_input_value("틱범위", tick_range)
        self.set_input_value("수정주가구분", 1)
        self.set_input_value("종가매매분봉",0)
        self.comm_rq_data("opt10080_req", "opt10080", 0, "5000")

    def _receive_tr_data(self, screen_no, rq_name, tr_code, record_name, prev_next, data_length, error_code, message, splm_msg):
    # 조회 요청 응답을 받거나 조회데이터를 수신했을 때 호출되는 이벤트
    # 조회 데이터는 이 이벤트에서 GetCommData 메서드를 사용하여 가져올 수 있습니다.
        if rq_name == "opt10080_req":
            self.on_opt10080_req(tr_code,rq_name)

    def on_opt10080_req(self, tr_code, rq_name):

        stock_code = self.get_comm_data(tr_code, rq_name, 0, "종목코드").replace("_AL", "")
        data_cnt = self.get_repeat_cnt(tr_code, rq_name)

        rows = []
        for i in range(data_cnt):
            _time = self.get_comm_data(tr_code, rq_name, i, "체결시간")
            open_price = self.get_comm_data(tr_code, rq_name, i, "시가")
            high_price = self.get_comm_data(tr_code, rq_name, i, "고가")
            low_price = self.get_comm_data(tr_code, rq_name, i, "저가")
            close_price = self.get_comm_data(tr_code, rq_name, i, "현재가")
            volume = self.get_comm_data(tr_code, rq_name, i, "거래량")
            rows.append([_time, int(open_price), int(high_price), int(low_price), int(close_price), int(volume)])

        daily_data_df = pd.DataFrame(rows, columns=["date", "open", "high", "low", "close", "volume"])
        daily_data_df = daily_data_df[::-1].reset_index(drop=True)  # 날짜 오름차순 정렬(역순)  일반적 df

        print(f"종목코드: {stock_code}")
        print(daily_data_df)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom_api =KiwoomAPI()
    sys.exit(app.exec_())
    timer = QTimer()
    timer.start(200)
    timer.timeout.connect(lambda: None)
        