#!/usr/bin/python
# -*- coding: utf-8 -*-


from WorkLogXls import *
from WeekLog import create_weeklog
from MonthLog import creat_month_log
from SeasonLog import create_season_log
from CorpSeasonLog import create_corp_season_log
import traceback
from mainFace import Ui_MainWindow
from PySide2.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PySide2 import QtCore
from PySide2.QtGui import QIcon
from PySide2.QtCore import Signal,QObject
import sys
from time import sleep
from threading import Thread
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, QIODevice


class progressSignals(QObject):
    progress_change = Signal(int, int)


class mainFace(QMainWindow):
    def __init__(self):
        # super(mainFace, self).__init__()
        # self.ui = Ui_MainWindow()
        # self.ui.setupUi(self)
        # -----动态加载ui文件-------#
        face_ui = QFile("mainFace.ui")  # 导入Qt designer生成的界面ui文件
        if not face_ui.open(QIODevice.ReadOnly):
            print("Cannot open {}: {}".format("mainFace.ui", face_ui.errorString()))
            sys.exit(-1)
        face_ui.close()
        self.ui = QUiLoader().load(face_ui)
        self.work_log_xls = None

        # 界面元素均可通过self.ui这个对象来获取
        self.ui.progressBar.setValue(0)
        self.progress = progressSignals()
        self.progress.progress_change.connect(self.setProgressValue)
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
        except Exception as er:
            QMessageBox.critical(self, '错误', traceback.format_exc())

        self.ui.btn_file.clicked.connect(self.btn_file_click)
        # self.ui.btn_WeekLog.clicked.connect(self.btn_weeklog_click)
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
            except Exception as er:
                QMessageBox.critical(self, '错误', traceback.format_exc())
                self.ui.btn_WeekLog.setEnabled(False)
                self.ui.btn_MonthLog.setEnabled(False)
                self.ui.btn_SeasonLog.setEnabled(False)
                self.ui.btn_CorpSeasonLog.setEnabled(False)

    def selectionchange1(self, i):
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
        QMessageBox.information(self, '信息', '工作周报生成完毕')

    def btn_weeklog_click(self):
        if not self.work_log_xls:
            QMessageBox.critical(self, '错误', '请先选择日志详情导出文件')
            return
        self.closeBtn()
        try:
            if self.work_log_xls:
                self.work_log_xls.split_weeks()
                pro_num = 0
                for week in self.work_log_xls.weeks:
                    create_weeklog(week)
                    pro_num = pro_num + 1
                    # self.setProgressValue(self.work_log_xls.weeks.__len__(), pro_num)
                    self.progress.progress_change.emit(self.work_log_xls.weeks.__len__(), pro_num)
                    # sleep(1)  # 为测试多线程增加的
                # QMessageBox.information(self, '信息', '工作周报生成完毕') #引入多线程后这里的消息提示出问题了，弹出一个白框，并且程序卡死
                self.openBtn()
        except Exception as er:
            QMessageBox.critical(self, '错误', traceback.format_exc())

    def setProgressValue(self, total, current):
        self.ui.progressBar.setValue(current / total * 100)

    def btn_monthlog_click(self):
        if not self.work_log_xls:
            QMessageBox.critical(self, '错误', '请先选择日志详情导出文件')
            return
        self.ui.btn_MonthLog.setEnabled(False)
        try:
            if self.work_log_xls:
                self.work_log_xls.split_months()
                for month_log in self.work_log_xls.months:
                    creat_month_log(month_log)

            self.ui.btn_MonthLog.setEnabled(True)
        except Exception as er:
            QMessageBox.critical(self, '错误', traceback.format_exc())

    def btn_seasonlog_click(self):
        if not self.work_log_xls:
            QMessageBox.critical(self, '错误', '请先选择日志详情导出文件')
            return
        self.ui.btn_SeasonLog.setEnabled(False)
        try:
            if self.work_log_xls:
                self.work_log_xls.split_season()
                for season in self.work_log_xls.seasons:
                    create_season_log(season)

            self.ui.btn_SeasonLog.setEnabled(True)
        except Exception as er:
            QMessageBox.critical(self, '错误', traceback.format_exc())

    def btn_corpseasonlog_click(self):
        if not self.work_log_xls:
            QMessageBox.critical(self, '错误', '请先选择日志详情导出文件')
            return
        self.ui.btn_CorpSeasonLog.setEnabled(False)
        try:
            if self.work_log_xls:
                self.work_log_xls.split_season()
                for season in self.work_log_xls.seasons:
                    create_corp_season_log(season)

            self.ui.btn_CorpSeasonLog.setEnabled(True)
        except Exception as er:
            QMessageBox.critical(self, '错误', traceback.format_exc())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # window = mainFace()
    # window.setWindowTitle('工作报告生成器')
    # # 禁止最大化按钮
    # window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint | QtCore.Qt.WindowCloseButtonHint)
    # # 禁止拉伸窗口大小
    # window.setFixedSize(window.width(), window.height());
    # # 设置图标
    # current_path = os.getcwd()
    # ico_file = os.path.join(current_path, "window.ico")
    # appIcon = QIcon(ico_file)
    # window.setWindowIcon(appIcon)
    #
    # # window.setIconModes()
    #
    # window.show()
    # sys.exit(app.exec_())
    window = mainFace()

    window.ui.setWindowTitle('工作报告生成器')
    # 禁止最大化按钮
    window.ui.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint | QtCore.Qt.WindowCloseButtonHint)
    # 禁止拉伸窗口大小
    window.ui.setFixedSize(window.ui.width(), window.ui.height())
    # 设置图标
    current_path = os.getcwd()
    ico_file = os.path.join(current_path, "window.ico")
    appIcon = QIcon(ico_file)
    window.ui.setWindowIcon(appIcon)

    window.ui.show()
    app.exec_()
