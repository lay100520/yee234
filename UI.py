# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'UI.ui'
#
# Created by: PyQt5 UI code generator 5.15.2
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(564, 322)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.groupBox_4 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_4.setGeometry(QtCore.QRect(9, 9, 546, 180))
        self.groupBox_4.setObjectName("groupBox_4")
        self.groupBox = QtWidgets.QGroupBox(self.groupBox_4)
        self.groupBox.setGeometry(QtCore.QRect(10, 70, 521, 51))
        self.groupBox.setObjectName("groupBox")
        self.lineEdit_import_file_2 = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit_import_file_2.setGeometry(QtCore.QRect(10, 20, 411, 20))
        self.lineEdit_import_file_2.setObjectName("lineEdit_import_file_2")
        self.toolButton_import_file_2 = QtWidgets.QToolButton(self.groupBox)
        self.toolButton_import_file_2.setGeometry(QtCore.QRect(440, 20, 61, 23))
        self.toolButton_import_file_2.setObjectName("toolButton_import_file_2")
        self.groupBox_3 = QtWidgets.QGroupBox(self.groupBox_4)
        self.groupBox_3.setGeometry(QtCore.QRect(10, 20, 521, 51))
        self.groupBox_3.setObjectName("groupBox_3")
        self.lineEdit_import_file_1 = QtWidgets.QLineEdit(self.groupBox_3)
        self.lineEdit_import_file_1.setGeometry(QtCore.QRect(10, 20, 411, 20))
        self.lineEdit_import_file_1.setObjectName("lineEdit_import_file_1")
        self.toolButton_import_file_1 = QtWidgets.QToolButton(self.groupBox_3)
        self.toolButton_import_file_1.setGeometry(QtCore.QRect(440, 20, 61, 23))
        self.toolButton_import_file_1.setObjectName("toolButton_import_file_1")
        self.groupBox_2 = QtWidgets.QGroupBox(self.groupBox_4)
        self.groupBox_2.setGeometry(QtCore.QRect(10, 120, 521, 51))
        self.groupBox_2.setObjectName("groupBox_2")
        self.comboBox = QtWidgets.QComboBox(self.groupBox_2)
        self.comboBox.setGeometry(QtCore.QRect(8, 20, 491, 22))
        self.comboBox.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItem("")
        self.groupBox_5 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_5.setGeometry(QtCore.QRect(10, 190, 546, 111))
        self.groupBox_5.setTitle("")
        self.groupBox_5.setObjectName("groupBox_5")
        self.open_catia = QtWidgets.QPushButton(self.groupBox_5)
        self.open_catia.setGeometry(QtCore.QRect(20, 10, 81, 91))
        self.open_catia.setObjectName("open_catia")
        self.mod_design = QtWidgets.QPushButton(self.groupBox_5)
        self.mod_design.setGeometry(QtCore.QRect(160, 10, 81, 91))
        self.mod_design.setObjectName("mod_design")
        self.drawing = QtWidgets.QPushButton(self.groupBox_5)
        self.drawing.setGeometry(QtCore.QRect(300, 10, 81, 91))
        self.drawing.setObjectName("drawing")
        self.closs_system = QtWidgets.QPushButton(self.groupBox_5)
        self.closs_system.setGeometry(QtCore.QRect(430, 10, 81, 91))
        self.closs_system.setObjectName("closs_system")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.lineEdit_import_file_1, self.toolButton_import_file_1)
        MainWindow.setTabOrder(self.toolButton_import_file_1, self.lineEdit_import_file_2)
        MainWindow.setTabOrder(self.lineEdit_import_file_2, self.toolButton_import_file_2)
        MainWindow.setTabOrder(self.toolButton_import_file_2, self.open_catia)
        MainWindow.setTabOrder(self.open_catia, self.mod_design)
        MainWindow.setTabOrder(self.mod_design, self.drawing)
        MainWindow.setTabOrder(self.drawing, self.closs_system)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.groupBox_4.setTitle(_translate("MainWindow", "匯入檔案"))
        self.groupBox.setTitle(_translate("MainWindow", "2.匯入逗點檔"))
        self.toolButton_import_file_2.setText(_translate("MainWindow", "選擇檔案"))
        self.groupBox_3.setTitle(_translate("MainWindow", "1.匯入線段檔"))
        self.toolButton_import_file_1.setText(_translate("MainWindow", "選擇檔案"))
        self.groupBox_2.setTitle(_translate("MainWindow", "3.選擇模具類型"))
        self.comboBox.setItemText(0, _translate("MainWindow", "剪切模"))
        self.open_catia.setText(_translate("MainWindow", "開啟Catai"))
        self.mod_design.setText(_translate("MainWindow", "模具生成"))
        self.drawing.setText(_translate("MainWindow", "出圖"))
        self.closs_system.setText(_translate("MainWindow", "關閉系統"))
