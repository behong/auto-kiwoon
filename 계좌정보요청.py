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
        self.using_account_num=""
        self.account_info_df = pd.DataFrame(
            columns=["계좌번호","종목코드","종목명","보유수량","매매가능수량","평균단가","현재가"]
        ),
        index = pd.Index([], name="종목코드")

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
        account_nums = str(self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"]).rstrip(';'))
        print(f"계좌정보 : {account_nums}")
        account_list = account_nums.split(';')
        self.using_account_num = account_list[-7]  ## 계좌번호들 중 가장 마지막 계좌번호를 사용
        self.get_account_balance()
        
    def get_account_balance(self,next=0):
        print(f"계좌정보 조회요청 : {len(self.using_account_num)}")
        print(f"계좌정보 조회요청 : {self.using_account_num}")
        if len(self.using_account_num) > 0:
            self.set_input_value("계좌번호", self.using_account_num)
            self.set_input_value("비밀번호", "")
            self.set_input_value("비밀번호입력매체구분", "00")
            self.set_input_value("조회구분", "1")
            self.set_input_value("거래소구분", "")
            self.comm_rq_data("opw00018_req","opw00018", next, "5000")


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

    def _receive_tr_data(self, screen_no, rq_name, tr_code, record_name, prev_next, data_length, error_code, message, splm_msg):
        # 조회 요청 응답을 받거나 조회데이터를 수신했을 때 호출되는 이벤트
        # 조회 데이터는 이 이벤트에서 GetCommData 메서드를 사용하여 가져올 수 있습니다.
        print(f"리 시 브 : {rq_name}")

        if rq_name == "opw00018_req":
            self.on_opt10080_req(tr_code,rq_name)

    def on_opt10080_req(self, tr_code, rq_name):

        try:
            현재평가잔고 = self.get_comm_data(tr_code, rq_name, 0, "추정예탁자산")
        except:
            현재평가잔고 = 0

        print(f"현재평가잔고: {현재평가잔고}")

        data_cnt = self.get_repeat_cnt(tr_code, rq_name)

        for i in range(data_cnt):
            try:
                종목코드 = self.get_comm_data(tr_code, rq_name, i, "종목번호").lstrip('A')
                종목명 = self.get_comm_data(tr_code, rq_name, i, "종목명")
                매매가능수량 = int(self.get_comm_data(tr_code, rq_name, i, "매매가능수량").lstrip('0'))
                보유수량 = int(self.get_comm_data(tr_code, rq_name, i, "보유수량").lstrip('0'))
                현재가 = int(self.get_comm_data(tr_code, rq_name, i, "현재가").lstrip('0').replace('+','').replace('-',''))
                매입가 = int(self.get_comm_data(tr_code, rq_name, i, "매입가").lstrip('0'))
                self.account_info_df.loc[종목코드] = {
                    "계좌번호":self.using_account_num,
                    "종목코드":종목코드,
                    "종목명":종목명,
                    "보유수량":보유수량,
                    "매매가능수량":매매가능수량,
                    "평균단가":매입가,
                    "현재가":현재가 
                }
            except Exception as e:
                logger.exception(f"계좌정보 조회 오류: {e}")

        print(self.account_info_df)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom_api =KiwoomAPI()
    sys.exit(app.exec_())
    timer = QTimer()
    timer.start(200)
    timer.timeout.connect(lambda: None)
        