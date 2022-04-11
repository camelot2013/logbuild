# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'mainFace.ui'
##
## Created by: Qt User Interface Compiler version 5.15.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(800, 351)
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.comboBox = QComboBox(self.centralwidget)
        self.comboBox.setObjectName(u"comboBox")
        self.comboBox.setGeometry(QRect(80, 10, 201, 41))
        self.label = QLabel(self.centralwidget)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 10, 71, 41))
        self.btn_file = QPushButton(self.centralwidget)
        self.btn_file.setObjectName(u"btn_file")
        self.btn_file.setGeometry(QRect(650, 160, 141, 41))
        self.lineEdit_FilePath = QLineEdit(self.centralwidget)
        self.lineEdit_FilePath.setObjectName(u"lineEdit_FilePath")
        self.lineEdit_FilePath.setGeometry(QRect(10, 160, 631, 41))
        self.lineEdit_FilePath.setReadOnly(True)
        self.lineEdit_company = QLineEdit(self.centralwidget)
        self.lineEdit_company.setObjectName(u"lineEdit_company")
        self.lineEdit_company.setGeometry(QRect(370, 10, 420, 41))
        self.lineEdit_company.setReadOnly(True)
        self.label_2 = QLabel(self.centralwidget)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(290, 10, 71, 41))
        self.label_3 = QLabel(self.centralwidget)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(10, 60, 71, 41))
        self.label_4 = QLabel(self.centralwidget)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(290, 60, 71, 41))
        self.label_5 = QLabel(self.centralwidget)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(10, 110, 71, 41))
        self.lineEdit_station = QLineEdit(self.centralwidget)
        self.lineEdit_station.setObjectName(u"lineEdit_station")
        self.lineEdit_station.setGeometry(QRect(80, 60, 201, 41))
        self.lineEdit_station.setReadOnly(True)
        self.lineEdit_level = QLineEdit(self.centralwidget)
        self.lineEdit_level.setObjectName(u"lineEdit_level")
        self.lineEdit_level.setGeometry(QRect(370, 60, 201, 41))
        self.lineEdit_level.setReadOnly(True)
        self.lineEdit_MonthlyUnitPrice = QLineEdit(self.centralwidget)
        self.lineEdit_MonthlyUnitPrice.setObjectName(u"lineEdit_MonthlyUnitPrice")
        self.lineEdit_MonthlyUnitPrice.setGeometry(QRect(80, 110, 201, 41))
        self.lineEdit_MonthlyUnitPrice.setReadOnly(True)
        self.btn_WeekLog = QPushButton(self.centralwidget)
        self.btn_WeekLog.setObjectName(u"btn_WeekLog")
        self.btn_WeekLog.setGeometry(QRect(10, 210, 141, 41))
        self.btn_MonthLog = QPushButton(self.centralwidget)
        self.btn_MonthLog.setObjectName(u"btn_MonthLog")
        self.btn_MonthLog.setGeometry(QRect(170, 210, 141, 41))
        self.btn_SeasonLog = QPushButton(self.centralwidget)
        self.btn_SeasonLog.setObjectName(u"btn_SeasonLog")
        self.btn_SeasonLog.setGeometry(QRect(330, 210, 141, 41))
        self.btn_CorpSeasonLog = QPushButton(self.centralwidget)
        self.btn_CorpSeasonLog.setObjectName(u"btn_CorpSeasonLog")
        self.btn_CorpSeasonLog.setGeometry(QRect(490, 210, 141, 41))
        self.progressBar = QProgressBar(self.centralwidget)
        self.progressBar.setObjectName(u"progressBar")
        self.progressBar.setGeometry(QRect(10, 260, 781, 41))
        self.progressBar.setValue(0)
        self.label_6 = QLabel(self.centralwidget)
        self.label_6.setObjectName(u"label_6")
        self.label_6.setGeometry(QRect(20, 310, 621, 21))
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"\u5de5\u4f5c\u62a5\u544a\u751f\u6210\u5668", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"\u4eba\u5458\u540d\u79f0", None))
        self.btn_file.setText(QCoreApplication.translate("MainWindow", u"\u9009\u62e9\u65e5\u5fd7\u5bfc\u51fa\u6587\u4ef6", None))
        self.label_2.setText(QCoreApplication.translate("MainWindow", u"\u516c\u53f8\u540d\u79f0", None))
        self.label_3.setText(QCoreApplication.translate("MainWindow", u"\u5c97    \u4f4d", None))
        self.label_4.setText(QCoreApplication.translate("MainWindow", u"\u7ea7    \u522b", None))
        self.label_5.setText(QCoreApplication.translate("MainWindow", u"\u6708 \u5355 \u4ef7", None))
        self.btn_WeekLog.setText(QCoreApplication.translate("MainWindow", u"\u751f\u6210\u5468\u62a5", None))
        self.btn_MonthLog.setText(QCoreApplication.translate("MainWindow", u"\u751f\u6210\u6708\u62a5", None))
        self.btn_SeasonLog.setText(QCoreApplication.translate("MainWindow", u"\u751f\u6210\u5b63\u62a5", None))
        self.btn_CorpSeasonLog.setText(QCoreApplication.translate("MainWindow", u"\u751f\u6210\u516c\u53f8\u5b63\u62a5", None))
        self.label_6.setText(QCoreApplication.translate("MainWindow", u"\u4e3a\u7b26\u5408\u884c\u65b9\u683c\u5f0f\u8981\u6c42\uff0c\u5468\u62a5\u9700\u8bbe\u7f6e\u8868\u683c\u5bbd\u5ea6\uff1a15.18CM;   \u4e2a\u4eba\u5b63\u62a5\u9700\u8bbe\u7f6e\u8868\u683c\u5bbd\u5ea6\uff1a15.77CM", None))
    # retranslateUi

