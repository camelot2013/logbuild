#!/usr/bin/python
# -*- coding: utf-8 -*-


from WorkLogXls import *
from WeekLog import create_weeklog
from MonthLog import creat_month_log
from SeasonLog import create_season_log
from CorpSeasonLog import create_corp_season_log
import traceback
from mainFace import Ui_MainWindow
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PySide6 import QtCore
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import Signal, QObject
import sys
# from time import sleep
from threading import Thread
# from PySide6.QtUiTools import QUiLoader
# from PySide6.QtCore import QFile, QIODevice
from rc_app import *


class ProgressSignals(QObject):
    def __init__(self):
        super(ProgressSignals, self).__init__()

    progress_change = Signal(int, int)
    alert_error_message = Signal(str)
    alert_info_message = Signal(str)


class MainFace(QMainWindow):
    
    @property
    def wd(self):
        return self.__wd

    def __init__(self):
        super(MainFace, self).__init__()
        try:
            self.__wd = sys._MEIPASS
        except AttributeError:
            self.__wd = os.getcwd()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.work_log_xls = None

        # 界面元素均可通过self.ui这个对象来获取
        self.ui.progressBar.setValue(0)
        self.progress = ProgressSignals()

        self.progress.progress_change.connect(self.setProgressValue)
        self.progress.alert_error_message.connect(self.alert_error_message)
        self.progress.alert_info_message.connect(self.alert_info_message)
        # noinspection PyBroadException
        try:
            person_cfg = read_cfg()
            if person_cfg:
                for person_name in person_cfg['person_info']:
                    info = person_cfg['person_info'][person_name]
                    self.ui.comboBox.addItem(person_name, info)
                info = self.ui.comboBox.currentData()

                self.ui.lineEdit_company.setText(info['company'])
                self.ui.lineEdit_station.setText(info['station'])
                self.ui.lineEdit_level.setText(info['level'])
                self.ui.lineEdit_MonthlyUnitPrice.setText(str(info['MonthlyUnitPrice']))
        except Exception:
            QMessageBox.critical(self, '错误', traceback.format_exc())

        self.ui.btn_file.clicked.connect(self.btn_file_click)
        self.ui.btn_WeekLog.clicked.connect(self.btn_weeklog_thread_click)
        self.ui.btn_MonthLog.clicked.connect(self.btn_monthlog_click)
        self.ui.btn_SeasonLog.clicked.connect(self.btn_seasonlog_click)
        self.ui.btn_CorpSeasonLog.clicked.connect(self.btn_corpseasonlog_click)
        self.ui.comboBox.currentIndexChanged.connect(self.selectionchange1)

    def btn_file_click(self):
        file_dialog = QFileDialog(self)
        file_path = file_dialog.getOpenFileName(self, "日志详情导出文件", ".", "xls Files(*.xls)")
        if file_path:
            self.ui.lineEdit_FilePath.setText(file_path[0])
            self.work_log_xls = WorkLogXls()
            # noinspection PyBroadException
            try:
                self.work_log_xls.read_xls(file_path[0])
                self.ui.btn_WeekLog.setEnabled(True)
                self.ui.btn_MonthLog.setEnabled(True)
                self.ui.btn_SeasonLog.setEnabled(True)
                self.ui.btn_CorpSeasonLog.setEnabled(True)
                for work_log in self.work_log_xls.work_logs:
                    if work_log['person_name'] != self.ui.comboBox.currentText():
                        self.ui.btn_WeekLog.setEnabled(False)
                        self.ui.btn_MonthLog.setEnabled(False)
                        self.ui.btn_SeasonLog.setEnabled(False)
                        self.ui.btn_CorpSeasonLog.setEnabled(False)
                        QMessageBox.critical(self, '错误', '日志详情文件中人员名称与当前选择的不一致')
                        return
            except Exception:
                QMessageBox.critical(self, '错误', traceback.format_exc())
                self.ui.btn_WeekLog.setEnabled(False)
                self.ui.btn_MonthLog.setEnabled(False)
                self.ui.btn_SeasonLog.setEnabled(False)
                self.ui.btn_CorpSeasonLog.setEnabled(False)

    def selectionchange1(self):
        info = self.ui.comboBox.currentData()
        self.ui.lineEdit_company.setText(info['company'])
        self.ui.lineEdit_station.setText(info['station'])
        self.ui.lineEdit_level.setText(info['level'])
        self.ui.lineEdit_MonthlyUnitPrice.setText(str(info['MonthlyUnitPrice']))

    def closeBtn(self):
        self.ui.btn_WeekLog.setEnabled(False)
        self.ui.btn_MonthLog.setEnabled(False)
        self.ui.btn_SeasonLog.setEnabled(False)
        self.ui.btn_CorpSeasonLog.setEnabled(False)

    def openBtn(self):
        self.ui.btn_WeekLog.setEnabled(True)
        self.ui.btn_MonthLog.setEnabled(True)
        self.ui.btn_SeasonLog.setEnabled(True)
        self.ui.btn_CorpSeasonLog.setEnabled(True)

    def btn_weeklog_thread_click(self):
        self.ui.progressBar.setValue(0)
        thread = Thread(target=self.btn_weeklog_click)
        thread.start()

    def btn_weeklog_click(self):
        if not self.work_log_xls:
            QMessageBox.critical(self, '错误', '请先选择日志详情导出文件')
            return
        self.closeBtn()
        # noinspection PyBroadException
        try:
            if self.work_log_xls:
                self.work_log_xls.split_weeks()
                for index, week in enumerate(self.work_log_xls.weeks):
                    create_weeklog(week)
                    # self.setProgressValue(self.work_log_xls.weeks.__len__(), pro_num)
                    self.progress.progress_change.emit(self.work_log_xls.weeks.__len__(), index + 1)
                    # sleep(1)  # 为测试多线程增加的
                self.progress.alert_info_message.emit('工作周报生成完毕')
                # QMessageBox.information(self.ui, '信息', '工作周报生成完毕') #引入多线程后这里的消息提示出问题了，弹出一个白框，并且程序卡死
                self.openBtn()
        except Exception:
            QMessageBox.critical(self, '错误', traceback.format_exc())

    def setProgressValue(self, total, current):
        self.ui.progressBar.setValue(current / total * 100)

    def alert_error_message(self, error_message):
        QMessageBox.critical(self, '错误', error_message)

    def alert_info_message(self, info_message):
        QMessageBox.information(self, '信息', info_message)

    def btn_monthlog_click(self):
        if not self.work_log_xls:
            QMessageBox.critical(self, '错误', '请先选择日志详情导出文件')
            return
        self.closeBtn()
        # noinspection PyBroadException
        try:
            if self.work_log_xls:
                self.work_log_xls.split_months()
                self.setProgressValue(self.work_log_xls.months.__len__(), 0)
                for index, month_log in enumerate(self.work_log_xls.months):
                    creat_month_log(month_log)
                    self.setProgressValue(self.work_log_xls.months.__len__(), index + 1)

            self.openBtn()
        except Exception:
            QMessageBox.critical(self, '错误', traceback.format_exc())
            print(traceback.format_exc())

    def btn_seasonlog_click(self):
        if not self.work_log_xls:
            QMessageBox.critical(self, '错误', '请先选择日志详情导出文件')
            return
        self.closeBtn()
        # noinspection PyBroadException
        try:
            if self.work_log_xls:
                self.work_log_xls.split_season()
                self.setProgressValue(self.work_log_xls.seasons.__len__(), 0)
                for index, season in enumerate(self.work_log_xls.seasons):
                    create_season_log(season)
                    self.setProgressValue(self.work_log_xls.seasons.__len__(), index + 1)

            self.openBtn()
        except Exception:
            QMessageBox.critical(self, '错误', traceback.format_exc())

    def btn_corpseasonlog_click(self):
        if not self.work_log_xls:
            QMessageBox.critical(self, '错误', '请先选择日志详情导出文件')
            return
        self.closeBtn()
        # noinspection PyBroadException
        try:
            if self.work_log_xls:
                self.work_log_xls.split_season()
                self.setProgressValue(self.work_log_xls.seasons.__len__(), 0)
                for index, season in enumerate(self.work_log_xls.seasons):
                    create_corp_season_log(season)
                    self.setProgressValue(self.work_log_xls.seasons.__len__(), index + 1)

            self.openBtn()
        except Exception:
            QMessageBox.critical(self, '错误', traceback.format_exc())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainFace()

    # 禁止最大化按钮
    window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint | QtCore.Qt.WindowCloseButtonHint)
    # 禁止拉伸窗口大小
    window.setFixedSize(window.width(), window.height())

    appIcon = QIcon(QPixmap(":/images/window.ico"))
    window.setWindowIcon(appIcon)

    window.show()

    sys.exit(app.exec())
