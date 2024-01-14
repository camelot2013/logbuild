# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'mainFaceWithRcc.ui'
##
## Created by: Qt User Interface Compiler version 6.2.2
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QAction, QBrush, QColor, QConicalGradient,
    QCursor, QFont, QFontDatabase, QGradient,
    QIcon, QImage, QKeySequence, QLinearGradient,
    QPainter, QPalette, QPixmap, QRadialGradient,
    QTransform)
from PySide6.QtWidgets import (QApplication, QComboBox, QLabel, QLineEdit,
    QMainWindow, QMenu, QMenuBar, QProgressBar,
    QPushButton, QSizePolicy, QStatusBar, QWidget)
import appico_rc

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(800, 442)
        icon = QIcon()
        icon.addFile(u":/icon/images/App.ico", QSize(), QIcon.Normal, QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.actionopen = QAction(MainWindow)
        self.actionopen.setObjectName(u"actionopen")
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
        self.btn_WeekLog.setGeometry(QRect(10, 260, 141, 41))
        self.btn_MonthLog = QPushButton(self.centralwidget)
        self.btn_MonthLog.setObjectName(u"btn_MonthLog")
        self.btn_MonthLog.setGeometry(QRect(170, 260, 141, 41))
        self.btn_SeasonLog = QPushButton(self.centralwidget)
        self.btn_SeasonLog.setObjectName(u"btn_SeasonLog")
        self.btn_SeasonLog.setGeometry(QRect(330, 260, 141, 41))
        self.btn_CorpSeasonLog = QPushButton(self.centralwidget)
        self.btn_CorpSeasonLog.setObjectName(u"btn_CorpSeasonLog")
        self.btn_CorpSeasonLog.setGeometry(QRect(490, 260, 141, 41))
        self.progressBar = QProgressBar(self.centralwidget)
        self.progressBar.setObjectName(u"progressBar")
        self.progressBar.setGeometry(QRect(10, 310, 781, 41))
        self.progressBar.setValue(0)
        self.label_6 = QLabel(self.centralwidget)
        self.label_6.setObjectName(u"label_6")
        self.label_6.setGeometry(QRect(10, 360, 621, 21))
        self.lineEdit_FilePath_EpibolyTotal = QLineEdit(self.centralwidget)
        self.lineEdit_FilePath_EpibolyTotal.setObjectName(u"lineEdit_FilePath_EpibolyTotal")
        self.lineEdit_FilePath_EpibolyTotal.setGeometry(QRect(10, 210, 631, 41))
        self.lineEdit_FilePath_EpibolyTotal.setReadOnly(True)
        self.btn_file_EpibolyTotal = QPushButton(self.centralwidget)
        self.btn_file_EpibolyTotal.setObjectName(u"btn_file_EpibolyTotal")
        self.btn_file_EpibolyTotal.setGeometry(QRect(650, 210, 141, 41))
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menuBar = QMenuBar(MainWindow)
        self.menuBar.setObjectName(u"menuBar")
        self.menuBar.setGeometry(QRect(0, 0, 800, 22))
        self.menu = QMenu(self.menuBar)
        self.menu.setObjectName(u"menu")
        MainWindow.setMenuBar(self.menuBar)

        self.menuBar.addAction(self.menu.menuAction())
        self.menu.addAction(self.actionopen)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"\u5de5\u4f5c\u62a5\u544a\u751f\u6210\u5668-2.0", None))
        self.actionopen.setText(QCoreApplication.translate("MainWindow", u"\u9009\u62e9\u6587\u4ef6", None))
#if QT_CONFIG(tooltip)
        self.actionopen.setToolTip(QCoreApplication.translate("MainWindow", u"\u9009\u62e9\u6587\u4ef6", None))
#endif // QT_CONFIG(tooltip)
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
        self.btn_file_EpibolyTotal.setText(QCoreApplication.translate("MainWindow", u"\u9009\u62e9\u5916\u5305\u5de5\u4f5c\u91cf\u6587\u4ef6", None))
        self.menu.setTitle(QCoreApplication.translate("MainWindow", u"\u6587\u4ef6", None))
    # retranslateUi

