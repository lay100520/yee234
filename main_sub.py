from PyQt5 import QtWidgets, QtGui, QtCore
from UI import Ui_MainWindow
import sys, psutil, os, re
import subprocess
from subprocess import CREATE_NEW_CONSOLE
from PyQt5.QtWidgets import *
import global_var as gvar
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import csv
import ModuleMain


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.open_catia.clicked.connect(self.start_CATIA)
        self.ui.toolButton_import_file_1.clicked.connect(self.import_file_root_1)
        self.ui.toolButton_import_file_2.clicked.connect(self.strip_parameters_file_root)
        self.ui.closs_system.clicked.connect(self.close_you)
        self.ui.mod_design.clicked.connect(self.mod_design)

    def start_CATIA(self):  # 開啟CATIA
        # try to CATIA enviorment file
        env_dir = 'C:\ProgramData\DassaultSystemes\CATEnv'
        list_dir = os.listdir(env_dir)

        print(list_dir)
        if any('V5-6R' in file for file in list_dir):
            for file in list_dir:
                if 'V5-6R' in file:
                    env_file = open(env_dir + '\\' + file, 'rt')
                    env_line = env_file.read().splitlines()
                    for line in env_line:
                        if 'CATInstallPath' in line:
                            CATIA_dir = re.sub('CATInstallPath=', '', line)
                            env_name = re.sub('.txt', '', file)
                            print('get CATIA dir and env is %s , %s' % (CATIA_dir, env_name))
        else:
            pass

        chk = [p.info for p in psutil.process_iter(attrs=['pid', 'name']) if 'CNEXT' in p.info['name']]
        print(chk)
        if chk == []:
            args = [r"%s\code\bin\CATSTART.exe" % CATIA_dir, "-run", "CNEXT.exe", "-env %s -direnv" % env_name,
                    "C:\ProgramData\DassaultSystemes\CATEnv", "-nowindow"]
            print(args)
            request = subprocess.Popen(args, shell=False, creationflags=CREATE_NEW_CONSOLE)
            print(str(request))
            print(os.getpid())
            return False
        else:
            pass

    def import_file_root_1(self):
        self.route = QtWidgets.QFileDialog.getOpenFileName(None, "選取檔案")
        # print(str(self.route[0]))
        self.ui.lineEdit_import_file_1.setText(str(self.route[0]))
        gvar.import_file_root = (str(self.route[0]))

    def strip_parameters_file_root(self):
        self.route = QtWidgets.QFileDialog.getOpenFileName(None, "選取檔案")
        # print(str(self.route[0]))
        self.ui.lineEdit_import_file_2.setText(str(self.route[0]))
        gvar.strip_parameters_file_root = (str(self.route[0]))

    def close_you(self):
        self.reply = QMessageBox.question(self, "警示", "您確定離開本系統?\nAre you sure you want to close?", QMessageBox.Yes,
                                          QMessageBox.No)
        if self.reply == QMessageBox.Yes:
            self.close()
        elif self.reply == QMessageBox.No:
            pass

    def mod_design(self):
        # 料帶資料設定
        with open(gvar.strip_parameters_file_root) as csvFile:
            rows = csv.reader(csvFile)
            strip_parameter_list = list(tuple(rows)[0])
            gvar.strip_parameter_list = strip_parameter_list
        ModuleMain.ModuleMain()



if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
