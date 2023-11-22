import sys

from PyQt5.QtCore import pyqtSignal, QThread, Qt, QDate
from PyQt5.QtGui import QPainter, QPen, QColor, QFont
from PyQt5.QtWidgets import *
from datetime import datetime, timedelta
from Ui import Ui_dialog
from SeleniumThread import seleniumm
from PyQt5 import uic
from OVERSEA import Overseas
from MARKET import Market
app = QApplication(sys.argv)
# form_class = uic.loadUiType("MainUi.ui")[0]

class Stock(QMainWindow, Ui_dialog, Overseas, Market):
    progress_start = pyqtSignal(int)
    progress_finish = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.selen = None
        self.driver = 0
        self.SearhUrl = ""
        self.savename =""
        self.CheckBoxList = []
        self.CheckNum = []
        self.FinishBoxList = []
        self.URL =["https://finance.naver.com/world/sise.naver?symbol=SPI@SPX",
                   "https://finance.naver.com/world/sise.naver?symbol=NAS@IXIC",
                   "https://finance.naver.com/marketindex/worldOilDetail.naver?marketindexCd=OIL_CL&fdtc=2",
                   "https://finance.naver.com/marketindex/materialDetail.naver?marketindexCd=CMDT_NG",
                   "https://finance.naver.com/marketindex/worldGoldDetail.naver?marketindexCd=CMDT_GC&fdtc=2",
                   "https://finance.naver.com/marketindex/materialDetail.naver?marketindexCd=CMDT_CDY",
                   "https://finance.naver.com/marketindex/worldExchangeDetail.naver?marketindexCd=FX_GBPUSD"
                   ]
        self.titleList = ["미국S&P",
                          "나스닥",
                          "WTi",
                          "NaturalGas",
                          "Gold",
                          "Cupper",
                          "PoundPerdollar"
                          ]
        # self.Condition.setStyleSheet("background-color: red;")
        self.setupUi(self)
        self.StartBtn.clicked.connect(self.Start)
        self.StartURLBtn.clicked.connect(self.StartURL)
        self.AllBtn.stateChanged.connect(self.AllCheck)
        self.CHECK1.stateChanged.connect(self.CheckState1)
        self.CHECK2.stateChanged.connect(self.CheckState2)
        self.start = ""
        self.end = ""
        yesterday = datetime.now() - timedelta(days=1)
        qdate = QDate(yesterday.year, yesterday.month, yesterday.day)
        self.EndDay.setDate(qdate)
        self.EndDay.setCalendarPopup(True)
        self.StartDay.setDate(qdate)  # 현재 날짜로 초기화
        self.StartDay.setCalendarPopup(True)

    def paintEvent(self, event):
        painter = QPainter(self)
        pen = QPen(Qt.black)
        pen.setWidth(2)
        painter.setPen(pen)
        painter.setRenderHint(QPainter.Antialiasing)  # 안티앨리어싱을 사용하여 부드러운 선을 그립니다.
        painter.drawRoundedRect(20, 70, 500, 750, 20, 20)
        color = QColor(153, 204, 255)  # RGB 값으로 색상 생성 (빨간색)
        pen = QPen(color)
        pen.setWidth(5)
        painter.setPen(pen)
        painter.drawLine(570, 200, 570, 600)

        painter.drawLine(570, 200, 620, 400)
        painter.drawLine(570, 600, 620, 400)

    def show_alert(self, text):
        alert = QMessageBox()
        alert.setWindowTitle("알림")
        alert.setText(text)
        alert.setIcon(QMessageBox.Information)
        alert.setStandardButtons(QMessageBox.Ok)
        alert.exec_()

    def GetDay(self):
        self.start = self.StartDay.date().toString("yyyy-MM-dd")
        self.end = self.EndDay.date().toString("yyyy-MM-dd")
        # 2020-02-20
        if self.start > self.end:
            self.show_alert("날짜 오류")
            return [0]
        if self.start == "" or self.end == "" or len(self.start) != 10 or len(self.end) != 10:
            self.show_alert("날짜 오류")
            return [0]
        start = self.start[:4] +"."+self.start[5:7] + "." + self.start[8:]
        end = self.end[:4] + "." + self.end[5:7] + "." + self.end[8:]

        if self.start > str(datetime.today())[0:10] or self.end > str(datetime.today())[0:10] or self.start < "1900.01.01" or self.end < "1900.01.01":
            self.show_alert("날짜 오류")
            return [0]
        return [start, end]
    def SetTable(self):
        self.Checked()
        self.Result.setRowCount(0)  # 모든 행 삭제
        self.Result.setColumnCount(0)  # 모든 열 삭제
        self.Result.setRowCount(len(self.CheckNum))
        self.Result.setColumnCount(2)
        self.Result.setRowHeight(1, 87)
        font = QFont()
        font.setPointSize(12)
        self.Result.setColumnWidth(1, 230)
        for i in range(len(self.CheckNum)):
            title = QTableWidgetItem(self.titleList[self.CheckNum[i]])
            status = QTableWidgetItem("추출 준비")
            title.setFont(font)
            status.setFont(font)
            self.Result.setItem(i, 0, title)
            self.FinishBoxList.append(self.Result)
            self.Result.setItem(i, 1, status)
            self.Result.setRowHeight(i, 90)
            self.Result.setColumnWidth(i, 230)

        # self.Result.setShowGrid(False)
        self.Result.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # 수평 스크롤바 비활성화
        self.Result.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # 수직 스크롤바 비활성화

    def Start(self):
        DayList = self.GetDay()
        self.Checked()
        for check in self.CheckBoxList:
            if check:
                break
        else:
            self.show_alert("아무것도 선택되지 않았습니다.")
            return
        self.selen = seleniumm(self.URL,self.titleList, self.CheckBoxList, DayList, self.Result, self.CheckNum,self.ResultLabel2)
        self.selen.start()
        self.SetTable()
    def StartURL(self):
        self.SearhUrl = self.URLSearch.text()
        self.savename = self.SaveName.text()
        DayList= self.GetDay()
        if self.CHECK1.isChecked():
            self.selen = seleniumm(self.SearhUrl,self.savename, [True,0,0,0,0,0,0,0,0,0], DayList, self.Result, self.CheckNum, self.ResultLabel2)
            self.selen.start()

        elif self.CHECK2.isChecked():
            self.selen = seleniumm(self.SearhUrl, self.savename, [False,0,0,0,0,0,0,0,0,0], DayList, self.Result, self.CheckNum, self.ResultLabel2)
            self.selen.start()
        else:
            self.show_alert("선택 옵션이 선택되지 않았습니다.")
            return
        self.selen.wait()
        if self.selen.ErrorNum == 1:
            self.show_alert("URL이 잘못 되었습니다.")
        elif self.selen.ErrorNum == 2:
            self.show_alert("정보 추출이 완료 되었습니다.")
        elif self.selen.ErrorNum == 3:
            self.show_alert("저장 이름이 없습니다.")


    def Stop(self):
        self.selen.Stop()
    def CheckState1(self):
        if self.CHECK1.isChecked():
            self.CHECK2.setChecked(False)
        else:
            self.CHECK2.setChecked(True)
    def CheckState2(self):
        if self.CHECK2.isChecked():
            self.CHECK1.setChecked(False)
        else:
            self.CHECK1.setChecked(True)

    def Checked(self):
        self.CheckBoxList = []
        self.CheckNum= []
        for i in range(1, 8):
            CheckBox =getattr(self, f"checkBox_{i}")
            if CheckBox.isChecked():
                self.CheckNum.append(i - 1)
                self.CheckBoxList.append(True)
            else:
                self.CheckBoxList.append(False)
    def AllCheck(self):
        if self.AllBtn.isChecked():
            for i in range(1, 8):
                CheckBox = getattr(self, f"checkBox_{i}")
                CheckBox.setChecked(True)
        else:
            for i in range(1, 8):
                CheckBox = getattr(self, f"checkBox_{i}")
                CheckBox.setChecked(False)

if __name__ == "__main__":
    sp = Stock()
    sp.show()
    sys.exit(app.exec_())
    # sp.SaveFile()