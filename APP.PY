# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'interfacejDuNLW.ui'
##
## Created by: Qt User Interface Compiler version 6.6.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QFrame,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QMainWindow, QPlainTextEdit, QPushButton, QSizePolicy,
    QSpacerItem, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget)

from Custom_Widgets.Widgets import (QCustomSlideMenu, QCustomStackedWidget)
import QSS_Resources_rc
import QSS_Resources_rc

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(953, 583)
        font = QFont()
        font.setBold(True)
        font.setItalic(True)
        MainWindow.setFont(font)
        icon = QIcon()
        icon.addFile(u":/newPrefix/logo/henkelsvg.png", QSize(), QIcon.Normal, QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setAutoFillBackground(False)
        MainWindow.setStyleSheet(u"*{\n"
"	border: none;\n"
"	background-color: transparent;\n"
"	background: transparent;\n"
"	padding: 0;\n"
"	margin: 0;\n"
"	color: #fff;\n"
"}\n"
"#centralwidget, #homeBtn, #mainBodyContent,QLineEdit,#plainTextEdit{\n"
"	background-color: #1b1b27;\n"
"}\n"
"#header,#mainBody,#footer{\n"
"	background-color: #27263c;\n"
"}\n"
"#pushButton{\n"
"	border: 2px  solid#dacd39;\n"
"	border-radius: 16px ;\n"
"	text-align: center;\n"
"}\n"
"#addUserBtn,#commencerBtn{\n"
"	background-color: #dacd39;\n"
"	border-radius:10px;\n"
"}\n"
"QPushButton{\n"
"	text-align:left;\n"
"	padding:5px 10px;\n"
"	border-top-left-radius:10px;\n"
"	border-bottom-left-radius:10px;\n"
"}\n"
"#homeBtn{\n"
"border-left: 3px  solid#dacd39;\n"
"font-weight:bold;\n"
"}\n"
"QLineEdit{\n"
"	padding:5px 10px;\n"
"	border-radius: 5px ;\n"
"}\n"
"#image{\n"
"border-radius: 5px ;\n"
"}\n"
"#btn1{\n"
"text-align:left;\n"
"	padding:5px 10px;\n"
"	border-top-left-radius:10px;\n"
"	border-bottom-left-radius:10px;\n"
"}\n"
"#commencerBtn,#pushButton_7,#push"
                        "Button_8{\n"
"	background-color: #d3dd67;\n"
"	border-radius:10px;\n"
"}\n"
"#pushButton_9{\n"
"	background-color: rgb(75, 167, 162);\n"
"	border-radius:10px;\n"
"}\n"
"#label_17,#label_18,#label_13,#label_27{\n"
"	\n"
"	border: 5px  solid#dacd39;\n"
"	border-radius: 16px ;\n"
"	text-align: center;\n"
"}\n"
"#toolBox,#toolBox_2{\n"
"	padding:5px 10px;\n"
"	border: 5px  solid#dacd39;\n"
"	border-radius: 16px ;\n"
"	text-align: center;\n"
"	background-color: transparent;\n"
"	background: transparent;\n"
"}\n"
"#comboBox{\n"
"   text-align: center;\n"
"padding:5px 10px;\n"
"   background-color: rgb(75, 167, 162);\n"
"	border-radius:10px;\n"
"	\n"
"}\n"
"")
        MainWindow.setIconSize(QSize(30, 30))
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.verticalLayout = QVBoxLayout(self.centralwidget)
        self.verticalLayout.setSpacing(10)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.header = QWidget(self.centralwidget)
        self.header.setObjectName(u"header")
        self.header.setMaximumSize(QSize(16777215, 50))
        self.horizontalLayout = QHBoxLayout(self.header)
        self.horizontalLayout.setSpacing(0)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.frame_2 = QFrame(self.header)
        self.frame_2.setObjectName(u"frame_2")
        font1 = QFont()
        font1.setBold(True)
        self.frame_2.setFont(font1)
        self.frame_2.setFrameShape(QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QFrame.Raised)
        self.horizontalLayout_4 = QHBoxLayout(self.frame_2)
        self.horizontalLayout_4.setObjectName(u"horizontalLayout_4")
        self.menuBtn = QPushButton(self.frame_2)
        self.menuBtn.setObjectName(u"menuBtn")
        self.menuBtn.setCursor(QCursor(Qt.PointingHandCursor))
        icon1 = QIcon()
        icon1.addFile(u":/icons/Icons/align-justify.png", QSize(), QIcon.Normal, QIcon.Off)
        self.menuBtn.setIcon(icon1)
        self.menuBtn.setIconSize(QSize(24, 24))

        self.horizontalLayout_4.addWidget(self.menuBtn, 0, Qt.AlignLeft)

        self.label_3 = QLabel(self.frame_2)
        self.label_3.setObjectName(u"label_3")
        font2 = QFont()
        font2.setFamilies([u"Times New Roman"])
        font2.setPointSize(19)
        font2.setBold(True)
        font2.setItalic(True)
        self.label_3.setFont(font2)

        self.horizontalLayout_4.addWidget(self.label_3)

        self.label = QLabel(self.frame_2)
        self.label.setObjectName(u"label")

        self.horizontalLayout_4.addWidget(self.label)


        self.horizontalLayout.addWidget(self.frame_2, 0, Qt.AlignLeft|Qt.AlignTop)

        self.frame = QFrame(self.header)
        self.frame.setObjectName(u"frame")
        self.frame.setFrameShape(QFrame.StyledPanel)
        self.frame.setFrameShadow(QFrame.Raised)
        self.horizontalLayout_3 = QHBoxLayout(self.frame)
        self.horizontalLayout_3.setObjectName(u"horizontalLayout_3")
        self.pushButton_4 = QPushButton(self.frame)
        self.pushButton_4.setObjectName(u"pushButton_4")
        self.pushButton_4.setCursor(QCursor(Qt.PointingHandCursor))
        icon2 = QIcon()
        icon2.addFile(u":/icons/Icons/search.png", QSize(), QIcon.Normal, QIcon.Off)
        self.pushButton_4.setIcon(icon2)

        self.horizontalLayout_3.addWidget(self.pushButton_4)

        self.pushButton_3 = QPushButton(self.frame)
        self.pushButton_3.setObjectName(u"pushButton_3")
        self.pushButton_3.setCursor(QCursor(Qt.PointingHandCursor))
        icon3 = QIcon()
        icon3.addFile(u":/icons/Icons/bell.png", QSize(), QIcon.Normal, QIcon.Off)
        self.pushButton_3.setIcon(icon3)

        self.horizontalLayout_3.addWidget(self.pushButton_3)

        self.pushButton_2 = QPushButton(self.frame)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setCursor(QCursor(Qt.PointingHandCursor))
        icon4 = QIcon()
        icon4.addFile(u":/icons/Icons/activity.png", QSize(), QIcon.Normal, QIcon.Off)
        self.pushButton_2.setIcon(icon4)
        self.pushButton_2.setIconSize(QSize(16, 16))

        self.horizontalLayout_3.addWidget(self.pushButton_2)

        self.pushButton = QPushButton(self.frame)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setMinimumSize(QSize(33, 33))
        self.pushButton.setMaximumSize(QSize(33, 33))
        self.pushButton.setCursor(QCursor(Qt.PointingHandCursor))
        icon5 = QIcon()
        icon5.addFile(u":/icons/Icons/user.png", QSize(), QIcon.Normal, QIcon.Off)
        self.pushButton.setIcon(icon5)
        self.pushButton.setIconSize(QSize(32, 32))

        self.horizontalLayout_3.addWidget(self.pushButton)


        self.horizontalLayout.addWidget(self.frame, 0, Qt.AlignRight|Qt.AlignTop)


        self.verticalLayout.addWidget(self.header)

        self.mainBody = QWidget(self.centralwidget)
        self.mainBody.setObjectName(u"mainBody")
        self.mainBody.setMinimumSize(QSize(800, 411))
        self.horizontalLayout_2 = QHBoxLayout(self.mainBody)
        self.horizontalLayout_2.setSpacing(0)
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.horizontalLayout_2.setContentsMargins(0, -1, 0, 0)
        self.leftMenu = QCustomSlideMenu(self.mainBody)
        self.leftMenu.setObjectName(u"leftMenu")
        self.leftMenu.setMinimumSize(QSize(0, 402))
        self.leftMenu.setMaximumSize(QSize(200, 16777215))
        self.verticalLayout_2 = QVBoxLayout(self.leftMenu)
        self.verticalLayout_2.setSpacing(0)
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 20)
        self.widget = QWidget(self.leftMenu)
        self.widget.setObjectName(u"widget")
        self.verticalLayout_5 = QVBoxLayout(self.widget)
        self.verticalLayout_5.setSpacing(0)
        self.verticalLayout_5.setObjectName(u"verticalLayout_5")
        self.verticalLayout_5.setContentsMargins(10, 0, 0, 0)
        self.frame_3 = QFrame(self.widget)
        self.frame_3.setObjectName(u"frame_3")
        self.frame_3.setFrameShape(QFrame.StyledPanel)
        self.frame_3.setFrameShadow(QFrame.Raised)
        self.verticalLayout_4 = QVBoxLayout(self.frame_3)
        self.verticalLayout_4.setSpacing(10)
        self.verticalLayout_4.setObjectName(u"verticalLayout_4")
        self.verticalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.homeBtn = QPushButton(self.frame_3)
        self.homeBtn.setObjectName(u"homeBtn")
        self.homeBtn.setCursor(QCursor(Qt.PointingHandCursor))
        self.homeBtn.setStyleSheet(u"background-color: #1b1b27;")
        icon6 = QIcon()
        icon6.addFile(u":/icons/Icons/home.png", QSize(), QIcon.Normal, QIcon.Off)
        self.homeBtn.setIcon(icon6)

        self.verticalLayout_4.addWidget(self.homeBtn)

        self.reportsBtn = QPushButton(self.frame_3)
        self.reportsBtn.setObjectName(u"reportsBtn")
        self.reportsBtn.setFont(font1)
        self.reportsBtn.setCursor(QCursor(Qt.PointingHandCursor))
        icon7 = QIcon()
        icon7.addFile(u":/icons/Icons/monitor.png", QSize(), QIcon.Normal, QIcon.Off)
        self.reportsBtn.setIcon(icon7)

        self.verticalLayout_4.addWidget(self.reportsBtn)

        self.accountBtn = QPushButton(self.frame_3)
        self.accountBtn.setObjectName(u"accountBtn")
        self.accountBtn.setCursor(QCursor(Qt.PointingHandCursor))
        icon8 = QIcon()
        icon8.addFile(u":/icons/Icons/users.png", QSize(), QIcon.Normal, QIcon.Off)
        self.accountBtn.setIcon(icon8)

        self.verticalLayout_4.addWidget(self.accountBtn)

        self.verticalSpacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)

        self.verticalLayout_4.addItem(self.verticalSpacer)


        self.verticalLayout_5.addWidget(self.frame_3)

        self.frame_4 = QFrame(self.widget)
        self.frame_4.setObjectName(u"frame_4")
        self.frame_4.setFrameShape(QFrame.StyledPanel)
        self.frame_4.setFrameShadow(QFrame.Raised)
        self.verticalLayout_3 = QVBoxLayout(self.frame_4)
        self.verticalLayout_3.setSpacing(10)
        self.verticalLayout_3.setObjectName(u"verticalLayout_3")
        self.verticalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.sittingsBtn = QPushButton(self.frame_4)
        self.sittingsBtn.setObjectName(u"sittingsBtn")
        self.sittingsBtn.setCursor(QCursor(Qt.PointingHandCursor))
        icon9 = QIcon()
        icon9.addFile(u":/icons/Icons/database.png", QSize(), QIcon.Normal, QIcon.Off)
        self.sittingsBtn.setIcon(icon9)

        self.verticalLayout_3.addWidget(self.sittingsBtn)

        self.helpBtn = QPushButton(self.frame_4)
        self.helpBtn.setObjectName(u"helpBtn")
        self.helpBtn.setCursor(QCursor(Qt.PointingHandCursor))
        icon10 = QIcon()
        icon10.addFile(u":/icons/Icons/clock.png", QSize(), QIcon.Normal, QIcon.Off)
        self.helpBtn.setIcon(icon10)

        self.verticalLayout_3.addWidget(self.helpBtn)

        self.aboutBtn = QPushButton(self.frame_4)
        self.aboutBtn.setObjectName(u"aboutBtn")
        self.aboutBtn.setCursor(QCursor(Qt.PointingHandCursor))
        icon11 = QIcon()
        icon11.addFile(u":/icons/Icons/key.png", QSize(), QIcon.Normal, QIcon.Off)
        self.aboutBtn.setIcon(icon11)
        self.aboutBtn.setIconSize(QSize(16, 16))

        self.verticalLayout_3.addWidget(self.aboutBtn)


        self.verticalLayout_5.addWidget(self.frame_4)


        self.verticalLayout_2.addWidget(self.widget)


        self.horizontalLayout_2.addWidget(self.leftMenu)

        self.mainBodyContent = QWidget(self.mainBody)
        self.mainBodyContent.setObjectName(u"mainBodyContent")
        self.horizontalLayout_5 = QHBoxLayout(self.mainBodyContent)
        self.horizontalLayout_5.setObjectName(u"horizontalLayout_5")
        self.mainPages = QCustomStackedWidget(self.mainBodyContent)
        self.mainPages.setObjectName(u"mainPages")
        self.homePage = QWidget()
        self.homePage.setObjectName(u"homePage")
        self.horizontalLayout_12 = QHBoxLayout(self.homePage)
        self.horizontalLayout_12.setObjectName(u"horizontalLayout_12")
        self.frame_6 = QFrame(self.homePage)
        self.frame_6.setObjectName(u"frame_6")
        self.frame_6.setFrameShape(QFrame.StyledPanel)
        self.frame_6.setFrameShadow(QFrame.Raised)
        self.horizontalLayout_7 = QHBoxLayout(self.frame_6)
        self.horizontalLayout_7.setObjectName(u"horizontalLayout_7")
        self.widget_3 = QWidget(self.frame_6)
        self.widget_3.setObjectName(u"widget_3")
        self.label_10 = QLabel(self.widget_3)
        self.label_10.setObjectName(u"label_10")
        self.label_10.setGeometry(QRect(80, 110, 411, 21))
        font3 = QFont()
        font3.setFamilies([u"Snap ITC"])
        font3.setPointSize(11)
        font3.setBold(False)
        font3.setItalic(True)
        self.label_10.setFont(font3)
        self.commencerBtn = QPushButton(self.widget_3)
        self.commencerBtn.setObjectName(u"commencerBtn")
        self.commencerBtn.setGeometry(QRect(120, 300, 151, 41))
        font4 = QFont()
        font4.setFamilies([u"Segoe UI"])
        font4.setPointSize(10)
        font4.setBold(True)
        font4.setItalic(False)
        self.commencerBtn.setFont(font4)
        self.commencerBtn.setCursor(QCursor(Qt.PointingHandCursor))
        icon12 = QIcon()
        icon12.addFile(u":/icons/Icons/airplay_disabled.png", QSize(), QIcon.Normal, QIcon.Off)
        self.commencerBtn.setIcon(icon12)
        self.commencerBtn.setIconSize(QSize(32, 32))
        self.widget_5 = QWidget(self.widget_3)
        self.widget_5.setObjectName(u"widget_5")
        self.widget_5.setGeometry(QRect(789, 82, 281, 311))
        self.verticalLayout_10 = QVBoxLayout(self.widget_5)
        self.verticalLayout_10.setObjectName(u"verticalLayout_10")
        self.label_12 = QLabel(self.widget_3)
        self.label_12.setObjectName(u"label_12")
        self.label_12.setGeometry(QRect(0, 10, 781, 103))
        font5 = QFont()
        font5.setFamilies([u"Palatino Linotype"])
        font5.setPointSize(48)
        #font5.setWeight(QFont.)
        font5.setItalic(True)
        self.label_12.setFont(font5)
        self.label_12.setStyleSheet(u"color:rgb(253, 255, 126);\n""font: 75 italic 48pt \"Palatino Linotype\";")
        self.label_16 = QLabel(self.widget_3)
        self.label_16.setObjectName(u"label_16")
        self.label_16.setGeometry(QRect(240, 240, 151, 31))
        self.label_16.setFont(font2)
        self.label_32 = QLabel(self.widget_3)
        self.label_32.setObjectName(u"label_32")
        self.label_32.setGeometry(QRect(110, 190, 111, 103))
        self.label_32.setPixmap(QPixmap(u":/icons/Icons/monitor.png"))
        self.label_32.setScaledContents(True)
        self.label_6 = QLabel(self.widget_3)
        self.label_6.setObjectName(u"label_6")
        self.label_6.setGeometry(QRect(400, 0, 151, 71))
        self.label_6.setPixmap(QPixmap(u"../../.designer/backup/QSS/logo/henkelsvg.png"))
        self.label_6.setScaledContents(True)
        self.label_5 = QLabel(self.widget_3)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(390, 110, 361, 293))
        self.label_5.setPixmap(QPixmap(u"QSS/Icons/1-removebg-preview.png"))
        self.label_5.setScaledContents(True)

        self.horizontalLayout_7.addWidget(self.widget_3)


        self.horizontalLayout_12.addWidget(self.frame_6)

        self.mainPages.addWidget(self.homePage)
        self.reportPage = QWidget()
        self.reportPage.setObjectName(u"reportPage")
        self.verticalLayout_15 = QVBoxLayout(self.reportPage)
        self.verticalLayout_15.setObjectName(u"verticalLayout_15")
        self.frame_9 = QFrame(self.reportPage)
        self.frame_9.setObjectName(u"frame_9")
        self.frame_9.setFrameShape(QFrame.StyledPanel)
        self.frame_9.setFrameShadow(QFrame.Raised)
        self.label_7 = QLabel(self.frame_9)
        self.label_7.setObjectName(u"label_7")
        self.label_7.setGeometry(QRect(30, 20, 421, 41))
        self.pushButton_7 = QPushButton(self.frame_9)
        self.pushButton_7.setObjectName(u"pushButton_7")
        self.pushButton_7.setGeometry(QRect(40, 80, 191, 41))
        self.pushButton_7.setCursor(QCursor(Qt.PointingHandCursor))
        icon13 = QIcon()
        icon13.addFile(u":/icons/Icons/folder-plus.png", QSize(), QIcon.Normal, QIcon.Off)
        self.pushButton_7.setIcon(icon13)
        self.pushButton_7.setIconSize(QSize(30, 30))
        self.pushButton_7.setCheckable(False)
        self.pushButton_8 = QPushButton(self.frame_9)
        self.pushButton_8.setObjectName(u"pushButton_8")
        self.pushButton_8.setGeometry(QRect(30, 230, 191, 41))
        self.pushButton_8.setCursor(QCursor(Qt.PointingHandCursor))
        self.pushButton_8.setIcon(icon13)
        self.pushButton_8.setIconSize(QSize(30, 30))
        self.pushButton_8.setCheckable(False)
        self.pushButton_9 = QPushButton(self.frame_9)
        self.pushButton_9.setObjectName(u"pushButton_9")
        self.pushButton_9.setGeometry(QRect(290, 270, 151, 41))
        self.pushButton_9.setCursor(QCursor(Qt.PointingHandCursor))
        icon14 = QIcon()
        icon14.addFile(u":/icons/Icons/play.png", QSize(), QIcon.Normal, QIcon.Off)
        self.pushButton_9.setIcon(icon14)
        self.pushButton_9.setIconSize(QSize(30, 30))
        self.pushButton_9.setCheckable(False)
        self.label_8 = QLabel(self.frame_9)
        self.label_8.setObjectName(u"label_8")
        self.label_8.setGeometry(QRect(40, 190, 421, 31))
        self.label_11 = QLabel(self.frame_9)
        self.label_11.setObjectName(u"label_11")
        self.label_11.setGeometry(QRect(110, 350, 101, 31))
        self.label_11.setStyleSheet(u"font: 14pt \"Showcard Gothic\";")
        self.label_13 = QLabel(self.frame_9)
        self.label_13.setObjectName(u"label_13")
        self.label_13.setGeometry(QRect(230, 340, 251, 51))
        self.label_13.setStyleSheet(u"font: 75 italic 12pt \"Comic Sans MS\";")
        self.comboBox = QComboBox(self.frame_9)
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.setObjectName(u"comboBox")
        self.comboBox.setGeometry(QRect(40, 150, 191, 22))
        self.comboBox.setStyleSheet(u"font: 75 8pt \"MS Shell Dlg 2\" center;")
        self.frame_7 = QFrame(self.frame_9)
        self.frame_7.setObjectName(u"frame_7")
        self.frame_7.setGeometry(QRect(640, -10, 141, 421))
        self.frame_7.setFrameShape(QFrame.StyledPanel)
        self.frame_7.setFrameShadow(QFrame.Raised)
        self.verticalLayout_11 = QVBoxLayout(self.frame_7)
        self.verticalLayout_11.setObjectName(u"verticalLayout_11")
        self.checkBox = QCheckBox(self.frame_7)
        self.checkBox.setObjectName(u"checkBox")

        self.verticalLayout_11.addWidget(self.checkBox)

        self.checkBox_4 = QCheckBox(self.frame_7)
        self.checkBox_4.setObjectName(u"checkBox_4")

        self.verticalLayout_11.addWidget(self.checkBox_4)

        self.checkBox_2 = QCheckBox(self.frame_7)
        self.checkBox_2.setObjectName(u"checkBox_2")

        self.verticalLayout_11.addWidget(self.checkBox_2)

        self.checkBox_3 = QCheckBox(self.frame_7)
        self.checkBox_3.setObjectName(u"checkBox_3")

        self.verticalLayout_11.addWidget(self.checkBox_3)

        self.checkBox_6 = QCheckBox(self.frame_7)
        self.checkBox_6.setObjectName(u"checkBox_6")

        self.verticalLayout_11.addWidget(self.checkBox_6)

        self.checkBox_7 = QCheckBox(self.frame_7)
        self.checkBox_7.setObjectName(u"checkBox_7")

        self.verticalLayout_11.addWidget(self.checkBox_7)

        self.checkBox_8 = QCheckBox(self.frame_7)
        self.checkBox_8.setObjectName(u"checkBox_8")

        self.verticalLayout_11.addWidget(self.checkBox_8)

        self.checkBox_9 = QCheckBox(self.frame_7)
        self.checkBox_9.setObjectName(u"checkBox_9")

        self.verticalLayout_11.addWidget(self.checkBox_9)

        self.checkBox_5 = QCheckBox(self.frame_7)
        self.checkBox_5.setObjectName(u"checkBox_5")

        self.verticalLayout_11.addWidget(self.checkBox_5)

        self.checkBox_10 = QCheckBox(self.frame_7)
        self.checkBox_10.setObjectName(u"checkBox_10")

        self.verticalLayout_11.addWidget(self.checkBox_10)

        self.checkBox_20 = QCheckBox(self.frame_7)
        self.checkBox_20.setObjectName(u"checkBox_20")

        self.verticalLayout_11.addWidget(self.checkBox_20)

        self.checkBox_11 = QCheckBox(self.frame_7)
        self.checkBox_11.setObjectName(u"checkBox_11")

        self.verticalLayout_11.addWidget(self.checkBox_11)

        self.checkBox_13 = QCheckBox(self.frame_7)
        self.checkBox_13.setObjectName(u"checkBox_13")

        self.verticalLayout_11.addWidget(self.checkBox_13)

        self.checkBox_12 = QCheckBox(self.frame_7)
        self.checkBox_12.setObjectName(u"checkBox_12")

        self.verticalLayout_11.addWidget(self.checkBox_12)

        self.checkBox_14 = QCheckBox(self.frame_7)
        self.checkBox_14.setObjectName(u"checkBox_14")

        self.verticalLayout_11.addWidget(self.checkBox_14)

        self.checkBox_15 = QCheckBox(self.frame_7)
        self.checkBox_15.setObjectName(u"checkBox_15")

        self.verticalLayout_11.addWidget(self.checkBox_15)

        self.checkBox_16 = QCheckBox(self.frame_7)
        self.checkBox_16.setObjectName(u"checkBox_16")

        self.verticalLayout_11.addWidget(self.checkBox_16)

        self.checkBox_17 = QCheckBox(self.frame_7)
        self.checkBox_17.setObjectName(u"checkBox_17")

        self.verticalLayout_11.addWidget(self.checkBox_17)

        self.checkBox_18 = QCheckBox(self.frame_7)
        self.checkBox_18.setObjectName(u"checkBox_18")

        self.verticalLayout_11.addWidget(self.checkBox_18)

        self.checkBox_19 = QCheckBox(self.frame_7)
        self.checkBox_19.setObjectName(u"checkBox_19")

        self.verticalLayout_11.addWidget(self.checkBox_19)

        self.checkBox_22 = QCheckBox(self.frame_7)
        self.checkBox_22.setObjectName(u"checkBox_22")

        self.verticalLayout_11.addWidget(self.checkBox_22)


        self.verticalLayout_15.addWidget(self.frame_9)

        self.mainPages.addWidget(self.reportPage)
        self.accountPage = QWidget()
        self.accountPage.setObjectName(u"accountPage")
        self.verticalLayout_12 = QVBoxLayout(self.accountPage)
        self.verticalLayout_12.setObjectName(u"verticalLayout_12")
        self.widget_4 = QWidget(self.accountPage)
        self.widget_4.setObjectName(u"widget_4")
        self.verticalLayout_13 = QVBoxLayout(self.widget_4)
        self.verticalLayout_13.setObjectName(u"verticalLayout_13")
        self.frame_8 = QFrame(self.widget_4)
        self.frame_8.setObjectName(u"frame_8")
        self.frame_8.setFrameShape(QFrame.StyledPanel)
        self.frame_8.setFrameShadow(QFrame.Raised)
        self.horizontalLayout_9 = QHBoxLayout(self.frame_8)
        self.horizontalLayout_9.setObjectName(u"horizontalLayout_9")
        self.label_4 = QLabel(self.frame_8)
        self.label_4.setObjectName(u"label_4")
        font6 = QFont()
        font6.setPointSize(24)
        font6.setBold(True)
        self.label_4.setFont(font6)

        self.horizontalLayout_9.addWidget(self.label_4, 0, Qt.AlignLeft)

        self.showUserFormBtn_1 = QPushButton(self.frame_8)
        self.showUserFormBtn_1.setObjectName(u"showUserFormBtn_1")
        font7 = QFont()
        font7.setPointSize(10)
        font7.setBold(True)
        self.showUserFormBtn_1.setFont(font7)
        icon15 = QIcon()
        icon15.addFile(u":/icons/Icons/plus-square.png", QSize(), QIcon.Normal, QIcon.Off)
        self.showUserFormBtn_1.setIcon(icon15)
        self.showUserFormBtn_1.setIconSize(QSize(24, 24))

        self.horizontalLayout_9.addWidget(self.showUserFormBtn_1, 0, Qt.AlignRight)


        self.verticalLayout_13.addWidget(self.frame_8)

        self.tableWidget = QTableWidget(self.widget_4)
        if (self.tableWidget.columnCount() < 4):
            self.tableWidget.setColumnCount(4)
        __qtablewidgetitem = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, __qtablewidgetitem)
        __qtablewidgetitem1 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, __qtablewidgetitem1)
        __qtablewidgetitem2 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(2, __qtablewidgetitem2)
        __qtablewidgetitem3 = QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, __qtablewidgetitem3)
        self.tableWidget.setObjectName(u"tableWidget")

        self.verticalLayout_13.addWidget(self.tableWidget)


        self.verticalLayout_12.addWidget(self.widget_4)

        self.mainPages.addWidget(self.accountPage)
        self.sittingsPage = QWidget()
        self.sittingsPage.setObjectName(u"sittingsPage")
        self.verticalLayout_14 = QVBoxLayout(self.sittingsPage)
        self.verticalLayout_14.setObjectName(u"verticalLayout_14")
        self.frame_10 = QFrame(self.sittingsPage)
        self.frame_10.setObjectName(u"frame_10")
        self.frame_10.setFrameShape(QFrame.StyledPanel)
        self.frame_10.setFrameShadow(QFrame.Raised)
        self.frame_15 = QFrame(self.frame_10)
        self.frame_15.setObjectName(u"frame_15")
        self.frame_15.setGeometry(QRect(640, 20, 571, 461))
        self.frame_15.setFrameShape(QFrame.StyledPanel)
        self.frame_15.setFrameShadow(QFrame.Raised)
        self.label_29 = QLabel(self.frame_15)
        self.label_29.setObjectName(u"label_29")
        self.label_29.setGeometry(QRect(140, 90, 471, 61))
        font8 = QFont()
        font8.setFamilies([u"Times New Roman"])
        font8.setPointSize(20)
        font8.setBold(True)
        font8.setItalic(True)
        self.label_29.setFont(font8)
        self.plainTextEdit_4 = QPlainTextEdit(self.frame_15)
        self.plainTextEdit_4.setObjectName(u"plainTextEdit_4")
        self.plainTextEdit_4.setGeometry(QRect(200, 170, 361, 101))
        font9 = QFont()
        font9.setFamilies([u"Times New Roman"])
        font9.setPointSize(16)
        font9.setBold(True)
        font9.setItalic(True)
        self.plainTextEdit_4.setFont(font9)
        self.plainTextEdit_5 = QPlainTextEdit(self.frame_15)
        self.plainTextEdit_5.setObjectName(u"plainTextEdit_5")
        self.plainTextEdit_5.setGeometry(QRect(200, 280, 361, 101))
        self.plainTextEdit_5.setFont(font9)

        self.verticalLayout_14.addWidget(self.frame_10)

        self.mainPages.addWidget(self.sittingsPage)
        self.aboutPage = QWidget()
        self.aboutPage.setObjectName(u"aboutPage")
        self.horizontalLayout_10 = QHBoxLayout(self.aboutPage)
        self.horizontalLayout_10.setObjectName(u"horizontalLayout_10")
        self.frame_12 = QFrame(self.aboutPage)
        self.frame_12.setObjectName(u"frame_12")
        self.frame_12.setFrameShape(QFrame.StyledPanel)
        self.frame_12.setFrameShadow(QFrame.Raised)
        self.verticalLayout_16 = QVBoxLayout(self.frame_12)
        self.verticalLayout_16.setObjectName(u"verticalLayout_16")
        self.frame_11 = QFrame(self.frame_12)
        self.frame_11.setObjectName(u"frame_11")
        self.frame_11.setMinimumSize(QSize(0, 80))
        self.frame_11.setMaximumSize(QSize(16777215, 80))
        self.frame_11.setFrameShape(QFrame.StyledPanel)
        self.frame_11.setFrameShadow(QFrame.Raised)

        self.verticalLayout_16.addWidget(self.frame_11)

        self.frame_13 = QFrame(self.frame_12)
        self.frame_13.setObjectName(u"frame_13")
        self.frame_13.setFrameShape(QFrame.StyledPanel)
        self.frame_13.setFrameShadow(QFrame.Raised)
        self.horizontalLayout_8 = QHBoxLayout(self.frame_13)
        self.horizontalLayout_8.setObjectName(u"horizontalLayout_8")
        self.widget_6 = QWidget(self.frame_13)
        self.widget_6.setObjectName(u"widget_6")

        self.horizontalLayout_8.addWidget(self.widget_6)


        self.verticalLayout_16.addWidget(self.frame_13)


        self.horizontalLayout_10.addWidget(self.frame_12)

        self.mainPages.addWidget(self.aboutPage)
        self.helpPage = QWidget()
        self.helpPage.setObjectName(u"helpPage")
        self.horizontalLayout_11 = QHBoxLayout(self.helpPage)
        self.horizontalLayout_11.setObjectName(u"horizontalLayout_11")
        self.frame_16 = QFrame(self.helpPage)
        self.frame_16.setObjectName(u"frame_16")
        self.frame_16.setMinimumSize(QSize(400, 0))
        self.frame_16.setMaximumSize(QSize(400, 16777215))
        self.frame_16.setFrameShape(QFrame.StyledPanel)
        self.frame_16.setFrameShadow(QFrame.Raised)

        self.horizontalLayout_11.addWidget(self.frame_16)

        self.frame_17 = QFrame(self.helpPage)
        self.frame_17.setObjectName(u"frame_17")
        self.frame_17.setFrameShape(QFrame.StyledPanel)
        self.frame_17.setFrameShadow(QFrame.Raised)

        self.horizontalLayout_11.addWidget(self.frame_17)

        self.mainPages.addWidget(self.helpPage)

        self.horizontalLayout_5.addWidget(self.mainPages)


        self.horizontalLayout_2.addWidget(self.mainBodyContent)

        self.rightMenu = QCustomSlideMenu(self.mainBody)
        self.rightMenu.setObjectName(u"rightMenu")
        self.rightMenu.setMinimumSize(QSize(0, 0))
        self.rightMenu.setMaximumSize(QSize(0, 16777215))
        self.verticalLayout_6 = QVBoxLayout(self.rightMenu)
        self.verticalLayout_6.setObjectName(u"verticalLayout_6")
        self.widget_2 = QWidget(self.rightMenu)
        self.widget_2.setObjectName(u"widget_2")
        self.widget_2.setMinimumSize(QSize(162, 242))
        self.verticalLayout_7 = QVBoxLayout(self.widget_2)
        self.verticalLayout_7.setObjectName(u"verticalLayout_7")
        self.label_2 = QLabel(self.widget_2)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setMinimumSize(QSize(70, 70))
        self.label_2.setMaximumSize(QSize(70, 70))
        self.label_2.setPixmap(QPixmap(u":/icons/Icons/edit.png"))
        self.label_2.setScaledContents(True)

        self.verticalLayout_7.addWidget(self.label_2, 0, Qt.AlignHCenter)

        self.frame_5 = QFrame(self.widget_2)
        self.frame_5.setObjectName(u"frame_5")
        self.frame_5.setFrameShape(QFrame.StyledPanel)
        self.frame_5.setFrameShadow(QFrame.Raised)
        self.verticalLayout_8 = QVBoxLayout(self.frame_5)
        self.verticalLayout_8.setObjectName(u"verticalLayout_8")
        self.userName = QLineEdit(self.frame_5)
        self.userName.setObjectName(u"userName")

        self.verticalLayout_8.addWidget(self.userName)

        self.email = QLineEdit(self.frame_5)
        self.email.setObjectName(u"email")

        self.verticalLayout_8.addWidget(self.email)

        self.phoneNo = QLineEdit(self.frame_5)
        self.phoneNo.setObjectName(u"phoneNo")

        self.verticalLayout_8.addWidget(self.phoneNo)


        self.verticalLayout_7.addWidget(self.frame_5)

        self.addUserBtn = QPushButton(self.widget_2)
        self.addUserBtn.setObjectName(u"addUserBtn")
        self.addUserBtn.setFont(font1)
        self.addUserBtn.setCursor(QCursor(Qt.PointingHandCursor))
        self.addUserBtn.setMouseTracking(False)
        icon16 = QIcon()
        icon16.addFile(u":/icons/Icons/file-plus.png", QSize(), QIcon.Normal, QIcon.Off)
        self.addUserBtn.setIcon(icon16)
        self.addUserBtn.setIconSize(QSize(24, 24))

        self.verticalLayout_7.addWidget(self.addUserBtn, 0, Qt.AlignHCenter)


        self.verticalLayout_6.addWidget(self.widget_2, 0, Qt.AlignHCenter|Qt.AlignTop)


        self.horizontalLayout_2.addWidget(self.rightMenu)


        self.verticalLayout.addWidget(self.mainBody)

        self.footer = QWidget(self.centralwidget)
        self.footer.setObjectName(u"footer")
        self.footer.setMinimumSize(QSize(0, 60))
        self.horizontalLayout_6 = QHBoxLayout(self.footer)
        self.horizontalLayout_6.setObjectName(u"horizontalLayout_6")
        self.label_9 = QLabel(self.footer)
        self.label_9.setObjectName(u"label_9")
        font10 = QFont()
        font10.setFamilies([u"Segoe UI"])
        font10.setPointSize(10)
        font10.setBold(True)
        self.label_9.setFont(font10)

        self.horizontalLayout_6.addWidget(self.label_9, 0, Qt.AlignHCenter)


        self.verticalLayout.addWidget(self.footer)

        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)

        self.mainPages.setCurrentIndex(0)
        self.pushButton_7.setDefault(False)
        self.pushButton_8.setDefault(False)
        self.pushButton_9.setDefault(False)


        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"SARL ANDROMEDE DISTRIBUTION", None))
        self.menuBtn.setText("")
        self.label_3.setText(QCoreApplication.translate("MainWindow", u"SARL ANDROMEDE DISTRIBUTION", None))
        self.label.setText("")
        self.pushButton_4.setText("")
        self.pushButton_3.setText("")
        self.pushButton_2.setText("")
        self.pushButton.setText("")
        self.homeBtn.setText(QCoreApplication.translate("MainWindow", u"Accueil", None))
        self.reportsBtn.setText(QCoreApplication.translate("MainWindow", u"BACK OFFICE", None))
        self.accountBtn.setText(QCoreApplication.translate("MainWindow", u"VENDEURS", None))
        self.sittingsBtn.setText(QCoreApplication.translate("MainWindow", u"Base De Donn\u00e9es", None))
        self.helpBtn.setText(QCoreApplication.translate("MainWindow", u"TIME ORDER ", None))
        self.aboutBtn.setText(QCoreApplication.translate("MainWindow", u"KPI / (SKU/ FACT)", None))
        self.label_10.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p>Alger cit\u00e9 chaaibia</p></body></html>", None))
        self.commencerBtn.setText(QCoreApplication.translate("MainWindow", u"Commencer", None))
        self.label_12.setText(QCoreApplication.translate("MainWindow", u"Sarl Androm\u00e9de Distribution", None))
        self.label_16.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p>Back office</p></body></html>", None))
        self.label_32.setText("")
        self.label_6.setText("")
        self.label_5.setText("")
        self.label_7.setText(QCoreApplication.translate("MainWindow", u"JOINDRE FICHIER", None))
        self.pushButton_7.setText(QCoreApplication.translate("MainWindow", u"BON CHARGEMENT", None))
        self.pushButton_8.setText(QCoreApplication.translate("MainWindow", u"BON CHARGEMENT", None))
        self.pushButton_9.setText(QCoreApplication.translate("MainWindow", u"Executer", None))
        self.label_8.setText(QCoreApplication.translate("MainWindow", u"JOINDRE FICHIER", None))
        self.label_11.setText(QCoreApplication.translate("MainWindow", u"TOTAL:", None))
        self.label_13.setText("")
        self.comboBox.setItemText(0, QCoreApplication.translate("MainWindow", u"PS16F01", None))
        self.comboBox.setItemText(1, QCoreApplication.translate("MainWindow", u"PS16F02", None))
        self.comboBox.setItemText(2, QCoreApplication.translate("MainWindow", u"PS16F03", None))
        self.comboBox.setItemText(3, QCoreApplication.translate("MainWindow", u"PS16F04", None))
        self.comboBox.setItemText(4, QCoreApplication.translate("MainWindow", u"PS16F05", None))
        self.comboBox.setItemText(5, QCoreApplication.translate("MainWindow", u"PS16F06", None))
        self.comboBox.setItemText(6, QCoreApplication.translate("MainWindow", u"PS16F07", None))
        self.comboBox.setItemText(7, QCoreApplication.translate("MainWindow", u"PS16F08", None))
        self.comboBox.setItemText(8, QCoreApplication.translate("MainWindow", u"PS16F09", None))
        self.comboBox.setItemText(9, QCoreApplication.translate("MainWindow", u"PS16F10", None))
        self.comboBox.setItemText(10, QCoreApplication.translate("MainWindow", u"PS16F11", None))
        self.comboBox.setItemText(11, QCoreApplication.translate("MainWindow", u"PS16F12", None))
        self.comboBox.setItemText(12, QCoreApplication.translate("MainWindow", u"PS16F13", None))
        self.comboBox.setItemText(13, QCoreApplication.translate("MainWindow", u"PS16F14", None))
        self.comboBox.setItemText(14, QCoreApplication.translate("MainWindow", u"PS16F15", None))
        self.comboBox.setItemText(15, QCoreApplication.translate("MainWindow", u"PS16F16", None))
        self.comboBox.setItemText(16, QCoreApplication.translate("MainWindow", u"PS16F17", None))
        self.comboBox.setItemText(17, QCoreApplication.translate("MainWindow", u"PS16F18", None))
        self.comboBox.setItemText(18, QCoreApplication.translate("MainWindow", u"PS16F19", None))
        self.comboBox.setItemText(19, QCoreApplication.translate("MainWindow", u"PS16F20", None))
        self.comboBox.setItemText(20, QCoreApplication.translate("MainWindow", u"PS16F21", None))

        self.checkBox.setText(QCoreApplication.translate("MainWindow", u"PS16F01", None))
        self.checkBox_4.setText(QCoreApplication.translate("MainWindow", u"PS16F02", None))
        self.checkBox_2.setText(QCoreApplication.translate("MainWindow", u"PS16F03", None))
        self.checkBox_3.setText(QCoreApplication.translate("MainWindow", u"PS16F04", None))
        self.checkBox_6.setText(QCoreApplication.translate("MainWindow", u"PS16F05", None))
        self.checkBox_7.setText(QCoreApplication.translate("MainWindow", u"PS16F06", None))
        self.checkBox_8.setText(QCoreApplication.translate("MainWindow", u"PS16F07", None))
        self.checkBox_9.setText(QCoreApplication.translate("MainWindow", u"PS16F08", None))
        self.checkBox_5.setText(QCoreApplication.translate("MainWindow", u"PS16F09", None))
        self.checkBox_10.setText(QCoreApplication.translate("MainWindow", u"PS16F10", None))
        self.checkBox_20.setText(QCoreApplication.translate("MainWindow", u"PS16F11", None))
        self.checkBox_11.setText(QCoreApplication.translate("MainWindow", u"PS16F12", None))
        self.checkBox_13.setText(QCoreApplication.translate("MainWindow", u"PS16F13", None))
        self.checkBox_12.setText(QCoreApplication.translate("MainWindow", u"PS16F14", None))
        self.checkBox_14.setText(QCoreApplication.translate("MainWindow", u"PS16F15", None))
        self.checkBox_15.setText(QCoreApplication.translate("MainWindow", u"PS16F16", None))
        self.checkBox_16.setText(QCoreApplication.translate("MainWindow", u"PS16F17", None))
        self.checkBox_17.setText(QCoreApplication.translate("MainWindow", u"PS16F18", None))
        self.checkBox_18.setText(QCoreApplication.translate("MainWindow", u"PS16F19", None))
        self.checkBox_19.setText(QCoreApplication.translate("MainWindow", u"PS16F20", None))
        self.checkBox_22.setText(QCoreApplication.translate("MainWindow", u"PS16F21", None))
        self.label_4.setText(QCoreApplication.translate("MainWindow", u"VENDEURS", None))
        self.showUserFormBtn_1.setText(QCoreApplication.translate("MainWindow", u"Ajouter VENDEURS", None))
        ___qtablewidgetitem = self.tableWidget.horizontalHeaderItem(0)
        ___qtablewidgetitem.setText(QCoreApplication.translate("MainWindow", u"Nouvelle colonne", None));
        ___qtablewidgetitem1 = self.tableWidget.horizontalHeaderItem(1)
        ___qtablewidgetitem1.setText(QCoreApplication.translate("MainWindow", u"Utilisateur", None));
        ___qtablewidgetitem2 = self.tableWidget.horizontalHeaderItem(2)
        ___qtablewidgetitem2.setText(QCoreApplication.translate("MainWindow", u"Email", None));
        ___qtablewidgetitem3 = self.tableWidget.horizontalHeaderItem(3)
        ___qtablewidgetitem3.setText(QCoreApplication.translate("MainWindow", u"Num\u00e9ro", None));
        self.label_29.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p><br/></p></body></html>", None))
        self.label_2.setText("")
        self.userName.setPlaceholderText(QCoreApplication.translate("MainWindow", u"Nom d'utilisateur", None))
        self.email.setPlaceholderText(QCoreApplication.translate("MainWindow", u"E-mail", None))
        self.phoneNo.setPlaceholderText(QCoreApplication.translate("MainWindow", u"Num\u00e9ro de T\u00e9l\u00e9phone ", None))
        self.addUserBtn.setText(QCoreApplication.translate("MainWindow", u"Ajouter User", None))
        self.label_9.setText(QCoreApplication.translate("MainWindow", u"Copyright-sarl - andromede-distribution 2023", None))
    # retranslateUi

