from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QIcon


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(979, 787)
        MainWindow.setWindowIcon(QIcon('web.png'))


        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(77)
        sizePolicy.setVerticalStretch(75)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())


        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setMinimumSize(QtCore.QSize(75, 75))
        MainWindow.setAutoFillBackground(False)
        MainWindow.setStyleSheet(
            # "background-color: #fff2bd"
            "background-color: #fffeb5;"
            "background-color: #f7f7d0;"
            )
        MainWindow.setDocumentMode(False)
        MainWindow.setTabShape(QtWidgets.QTabWidget.Rounded)
        MainWindow.setUnifiedTitleAndToolBarOnMac(False)


        self.centralwidget = QtWidgets.QWidget(MainWindow)


        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.centralwidget.sizePolicy().hasHeightForWidth())


        self.centralwidget.setSizePolicy(sizePolicy)
        self.centralwidget.setMinimumSize(QtCore.QSize(409, 518))
        self.centralwidget.setObjectName("centralwidget")


        self.gridLayout_2 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_2.setSpacing(0)
        self.gridLayout_2.setObjectName("gridLayout_2")



        self.mainLayout = QtWidgets.QFrame(self.centralwidget)
        self.mainLayout.setEnabled(True)


        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(100)
        sizePolicy.setVerticalStretch(1)
        sizePolicy.setHeightForWidth(self.mainLayout.sizePolicy().hasHeightForWidth())


        self.mainLayout.setSizePolicy(sizePolicy)
        self.mainLayout.setObjectName("mainLayout")


        self.gridLayout = QtWidgets.QGridLayout(self.mainLayout)
        self.gridLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.gridLayout.setContentsMargins(0, 0, 9, 0)
        self.gridLayout.setSpacing(5)
        self.gridLayout.setObjectName("gridLayout")


        self.Showing = QtWidgets.QFrame(self.mainLayout)
        self.Showing.setStyleSheet(
            # "background-color: #ffedcb"
            # "background-color: #fffeb5;"
            "background-color: #f7f7d0;"
            
            )
        self.Showing.setLineWidth(0)
        self.Showing.setObjectName("Showing")


        self.gridLayout_3 = QtWidgets.QGridLayout(self.Showing)
        self.gridLayout_3.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.gridLayout_3.setContentsMargins(11, -1, 11, -1)
        self.gridLayout_3.setSpacing(2)
        self.gridLayout_3.setObjectName("gridLayout_3")


        self.scrollArea = QtWidgets.QScrollArea(self.Showing)
        self.scrollArea.setStyleSheet("")
        self.scrollArea.setLineWidth(0)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")

        self.scrollArea.setStyleSheet(
            "background-color:#fefff7"
            )


        # self.scrollAreaWidgetContents_2 = QtWidgets.QWidget()
        # self.scrollAreaWidgetContents_2.setGeometry(QtCore.QRect(0, 0, 458, 645))
        # self.scrollAreaWidgetContents_2.setObjectName("scrollAreaWidgetContents_2")


        # self.scrollArea.setWidget(self.scrollAreaWidgetContents_2)


        self.gridLayout_3.addWidget(self.scrollArea, 0, 0, 1, 1)


        self.gridLayout.addWidget(self.Showing, 0, 1, 1, 1)


        # self.pushButton = QtWidgets.QPushButton(self.mainLayout)
        # font = QtGui.QFont()
        # font.setFamily("lato")
        # font.setPointSize(20)
        # font.setBold(False)
        # font.setItalic(False)
        # font.setUnderline(False)
        # font.setWeight(50)
        # font.setKerning(True)
        # self.pushButton.setFont(font)

        # self.pushButton.setIconSize(QtCore.QSize(30, 30))
        # self.pushButton.setObjectName("pushButton")


        # self.gridLayout.addWidget(self.pushButton, 1, 1, 1, 1)


        self.Reading = QtWidgets.QFrame(self.mainLayout)
        self.Reading.setStyleSheet(
            # "background-color:#fff4c7;\n"
            # "background-color: #fffeb5;"
            "background-color: #f7f7d0;"
"")
        self.Reading.setLineWidth(0)
        self.Reading.setObjectName("Reading")


        self.gridLayout_5 = QtWidgets.QGridLayout(self.Reading)
        self.gridLayout_5.setHorizontalSpacing(7)
        self.gridLayout_5.setObjectName("gridLayout_5")


       
        self.scrollArea_2 = QtWidgets.QScrollArea(self.Reading)
        self.scrollArea_2.setLineWidth(0)
        self.scrollArea_2.setWidgetResizable(True)
        self.scrollArea_2.setObjectName("scrollArea_2")
        self.scrollArea_2.setStyleSheet(
            "background-color:#fefff7;\n"
# "font-size:30px;\n"
"font-family:lato;\n"
"color:black;\n")




        self.gridLayout_5.addWidget(self.scrollArea_2, 0, 0, 1, 1)


        self.gridLayout.addWidget(self.Reading, 0, 0, 2, 1)


        self.gridLayout_2.addWidget(self.mainLayout, 0, 0, 1, 2)




#         self.pushButton = QtWidgets.QPushButton("открыть/закрыть правое окно")
#         font = QtGui.QFont()
#         font.setFamily("lato")
#         font.setPointSize(20)
#         font.setBold(False)
#         font.setItalic(False)
#         font.setUnderline(False)
#         font.setWeight(50)
#         font.setKerning(True)
#         self.pushButton.setFont(font)
#         # self.pushButton.setStyleSheet("background-color:#fcdc9d;\n"
# # "font-family:lato;")
#         self.pushButton.setIconSize(QtCore.QSize(30, 30))
#         self.pushButton.setObjectName("pushButton")


#         self.gridLayout_5.addWidget(self.pushButton, 1, 0, 1, 1)
#         self.pushButton.clicked.connect(self.show_close)



        MainWindow.setCentralWidget(self.centralwidget)


        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 979, 41))
        self.menubar.setStyleSheet(
            # "background-color:#fff5e2;\n"
            "background-color: #fffeb5;"
"font-size:30px;\n"
"font-family:lato;\n"
"color:blue;")
        self.menubar.setObjectName("menubar")


        self.menuSettings = QtWidgets.QMenu(self.menubar)
        self.menuSettings.setObjectName("menuSettings")


        self.menuAdd = QtWidgets.QMenu(self.menubar)
        self.menuAdd.setObjectName("menuAdd")


        self.menuHeader = QtWidgets.QMenu(self.menuAdd)
        self.menuHeader.setObjectName("menuHeader")


        self.menuSettings_2 = QtWidgets.QMenu(self.menubar)
        self.menuSettings_2.setObjectName("menuSettings_2")


        MainWindow.setMenuBar(self.menubar)


        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setStyleSheet(
# "background-color:#fff5e2"

# "background-color: #fffeb5;"
"background-color: #f7f7d0;"
)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)


        self.actionSave = QtWidgets.QAction(MainWindow)
        self.actionSave.setObjectName("actionSave")


        self.actionOpen = QtWidgets.QAction(MainWindow)
        self.actionOpen.setObjectName("actionOpen")


        self.actionHeader_1 = QtWidgets.QAction(MainWindow)


        self.actionHeader_1.setObjectName("actionHeader_1")


        self.actionHeader_2 = QtWidgets.QAction(MainWindow)


        self.actionHeader_2.setObjectName("actionHeader_2")


        self.actionHeader_3 = QtWidgets.QAction(MainWindow)
        self.actionHeader_3.setObjectName("actionHeader_3")


        self.actionText = QtWidgets.QAction(MainWindow)
        self.actionText.setObjectName("actionText")


        self.actionListing = QtWidgets.QAction(MainWindow)
        self.actionListing.setObjectName("actionListing")


        self.actionList = QtWidgets.QAction(MainWindow)
        self.actionList.setObjectName("actionList")


        self.actionTable = QtWidgets.QAction(MainWindow)
        self.actionTable.setObjectName("actionTable")


        self.actionUsed_Info = QtWidgets.QAction(MainWindow)
        self.actionUsed_Info.setObjectName("actionUsed_Info")

        self.actionPicture = QtWidgets.QAction(MainWindow)
        self.actionPicture.setObjectName("actionPicture")


        # self.actionTitle_page = QtWidgets.QAction(MainWindow)
        # self.actionTitle_page.setObjectName("actionTitle_page")


        # self.actionNew_file = QtWidgets.QAction(MainWindow)
        # self.actionNew_file.setObjectName("actionNew_file")
    

        # self.actionCompress_photos = QtWidgets.QAction(MainWindow)
        # self.actionCompress_photos.setObjectName("actionCompress_photos")


        self.actionNumeration = QtWidgets.QAction(MainWindow)
        self.actionNumeration.setObjectName("actionNumeration")


        # self.actionFooter = QtWidgets.QAction(MainWindow)
        # self.actionFooter.setObjectName("actionFooter")


        self.menuSettings.addAction(self.actionSave)
        self.menuSettings.addAction(self.actionOpen)
        # self.menuSettings.addAction(self.actionNew_file)


        self.menuHeader.addAction(self.actionHeader_1)
        self.menuHeader.addAction(self.actionHeader_2)
        self.menuHeader.addAction(self.actionHeader_3)


        self.menuAdd.addAction(self.menuHeader.menuAction())
        self.menuAdd.addAction(self.actionText)
        self.menuAdd.addAction(self.actionListing)
        self.menuAdd.addAction(self.actionList)
        self.menuAdd.addAction(self.actionTable)
        self.menuAdd.addAction(self.actionUsed_Info)
        self.menuAdd.addAction(self.actionPicture)
        


        # self.menuSettings_2.addAction(self.actionTitle_page)
        # self.menuSettings_2.addAction(self.actionCompress_photos)
        # self.menuSettings_2.addAction(self.actionNumeration)
        # self.menuSettings_2.addAction(self.actionFooter)


        self.menubar.addAction(self.menuSettings.menuAction())
        self.menubar.addAction(self.menuAdd.menuAction())
        self.menubar.addAction(self.menuSettings_2.menuAction())


        self.retranslateUi(MainWindow)


        QtCore.QMetaObject.connectSlotsByName(MainWindow)



        self.count = 0

    def retranslateUi(self, MainWindow):


        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        # self.pushButton.setText(_translate("MainWindow", "Сохранить"))
        self.menuSettings.setTitle(_translate("MainWindow", "Настройки"))
        self.menuAdd.setTitle(_translate("MainWindow", "Добавить"))
        self.menuHeader.setTitle(_translate("MainWindow", "Заголовок"))
        # self.menuSettings_2.setTitle(_translate("MainWindow", "Титульный лист"))
        self.actionSave.setText(_translate("MainWindow", "Сохранить файл"))
        self.actionOpen.setText(_translate("MainWindow", "Открыть файл"))
        self.actionHeader_1.setText(_translate("MainWindow", "Заголовок 1 уровня"))
        self.actionHeader_2.setText(_translate("MainWindow", "Заголовок 2 уровня"))
        self.actionHeader_3.setText(_translate("MainWindow", "Заголовок 3 уровня"))
        self.actionText.setText(_translate("MainWindow", "Обычный текст"))
        self.actionListing.setText(_translate("MainWindow", "Листинг"))
        self.actionList.setText(_translate("MainWindow", "Список"))
        self.actionTable.setText(_translate("MainWindow", "Таблица"))
        self.actionUsed_Info.setText(_translate("MainWindow", "Список использованных источников"))
        self.actionPicture.setText(_translate("MainWindow", "Рисунки"))
        # self.actionTitle_page.setText(_translate("MainWindow", "Титульный лист"))
        # self.actionNew_file.setText(_translate("MainWindow", "Новый файл"))
        # self.actionCompress_photos.setText(_translate("MainWindow", "Сompress photos"))
        # self.actionNumeration.setText(_translate("MainWindow", "Numeration"))
        # self.actionFooter.setText(_translate("MainWindow", "Footer"))





if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
