from PyQt5.QtCore import pyqtSignal, QThread, Qt, QDate
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QTableWidgetItem, QMessageBox
from MARKET import Market
from OVERSEA import Overseas


class seleniumm(QThread):
    def __init__(self, URL,titleList, CheckBoxList, DayList, Result, CheckNum, ResultLabel2):
        super().__init__()
        self.ErrorNum = 0
        self.CanStart = True
        self.URL = URL
        self.Result = Result
        self.CheckNum = CheckNum
        self.titleList = titleList
        self.CheckBoxList = CheckBoxList
        self.ResultLabel2 = ResultLabel2
        self.DayList = DayList
        self.running = True
    def run(self):
        try:
            if len(self.DayList) == 1:
                self.ErrorNum = 0
                self.CanStart = False
                return
            if len(self.CheckBoxList) == 10:
                if self.titleList == "":
                    self.ErrorNum = 3
                elif self.CheckBoxList[0]:
                    oversea = Overseas(self.URL, self.titleList, self.DayList[0], self.DayList[1])
                    overseaCan = oversea.SaveFile()
                    if overseaCan:
                        self.ErrorNum = 2
                    else:
                        self.ErrorNum = 1

                else:
                    market = Market(self.URL, self.titleList, self.DayList[0], self.DayList[1])
                    marketCan = market.SaveFile()
                    if marketCan:
                        self.ErrorNum = 2
                    else:
                        self.ErrorNum = 1

                return
            index = 0
            score = 0
            self.ResultLabel2.setText(f"0/{len(self.CheckNum)}")
            font = QFont()
            font.setPointSize(12)
            for i in range(len(self.CheckBoxList)):
                Status_1 = QTableWidgetItem("추출 중")
                Status_2 = QTableWidgetItem("추출 성공")
                Status_3 = QTableWidgetItem("추출 실패")
                Status_1.setFont(font)
                Status_2.setFont(font)
                Status_3.setFont(font)
                if not self.running:
                    break

                if (i == 0 or i == 1) and self.CheckBoxList[i]:

                    self.Result.setItem(index, 1, Status_1)
                    oversea = Overseas(self.URL[i], self.titleList[i], self.DayList[0], self.DayList[1])
                    value = oversea.SaveFile()
                    if value:
                        self.Result.setItem(index, 1, Status_2)
                        score += 1
                    else:
                        self.Result.setItem(index, 1, Status_3)
                    index += 1

                elif self.CheckBoxList[i]:
                    self.Result.setItem(index, 1, Status_1)
                    market = Market(self.URL[i], self.titleList[i], self.DayList[0], self.DayList[1])
                    value = market.SaveFile()
                    if value:
                        self.Result.setItem(index, 1, Status_2)
                        score += 1
                    else:
                        self.Result.setItem(index, 1, Status_3)
                    index += 1
                self.ResultLabel2.setText(f"{score}/{len(self.CheckNum)}")
        except Exception as e:
            print(e)
    def show_alert(self, text):
        alert = QMessageBox()
        alert.setWindowTitle("알림")
        alert.setText(text)
        alert.setIcon(QMessageBox.Information)
        alert.setStandardButtons(QMessageBox.Ok)
        alert.exec_()


    def Stop(self):
        self.running = False

