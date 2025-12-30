import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QAxContainer import QAxWidget

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.OnEventConnect.connect(self._event_connect)
        self.kiwoom.dynamicCall("CommConnect()")

    def _event_connect(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        else:
            print(f"로그인 실패: {err_code}")
        self.after_login()

    def after_login(self):
        if self.kiwoom.dynamicCall("GetConnectState()") == 0:
            print("미연결 상태")
        else:
            print("연결 중")    

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    # window.show()
    sys.exit(app.exec_())

        