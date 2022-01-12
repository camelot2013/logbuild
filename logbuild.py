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
import sys


class mainFace(QMainWindow):
    def __init__(self):
        super(mainFace, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.work_log_xls = None

        # 界面元素均可通过self.ui这个对象来获取
        person_cfg = read_cfg()
        if person_cfg:
            for person_info in person_cfg['person_info']:
                info = person_cfg['person_info'][person_info]
                self.ui.comboBox.addItem(person_info, info)
            info = self.ui.comboBox.currentData()

            self.ui.lineEdit_company.setText(info['company'])
            self.ui.lineEdit_station.setText(info['station'])
            self.ui.lineEdit_level.setText(info['level'])
            self.ui.lineEdit_MonthlyUnitPrice.setText(str(info['MonthlyUnitPrice']))

        self.ui.btn_file.clicked.connect(self.btn_file_click)
        self.ui.btn_WeekLog.clicked.connect(self.btn_weeklog_click)
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
                print(self.work_log_xls.work_logs)
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

    def btn_weeklog_click(self):
        if not self.work_log_xls:
            QMessageBox.critical(self, '错误', '请先选择日志详情导出文件')
            return
        self.ui.btn_WeekLog.setEnabled(False)
        try:
            if self.work_log_xls:
                self.work_log_xls.split_weeks()
                for week in self.work_log_xls.weeks:
                    create_weeklog(week)
                QMessageBox.information(self, '信息', '工作周报生成完毕')
                self.ui.btn_WeekLog.setEnabled(True)
        except Exception as er:
            QMessageBox.critical(self, '错误', traceback.format_exc())

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
    window = mainFace()
    window.setWindowTitle('工作报告生成器')
    # 禁止最大化按钮
    window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint | QtCore.Qt.WindowCloseButtonHint)
    # 禁止拉伸窗口大小
    window.setFixedSize(window.width(), window.height());
    # 设置图标
    current_path = os.getcwd()
    ico_file = os.path.join(current_path, "window.ico")
    appIcon = QIcon(ico_file)
    window.setWindowIcon(appIcon)
    window.show()
    sys.exit(app.exec_())
