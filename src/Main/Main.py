'''
Created on 18.12.2015

@author: rleber
'''

from PyQt4 import QtCore, QtGui
from Database import Database
from PyQt4.Qt import QColor

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    _fromUtf8 = lambda s: s

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName(_fromUtf8("Dialog"))
        Dialog.resize(231, 520)
        self.lineEdit_host = QtGui.QLineEdit(Dialog)
        self.lineEdit_host.setGeometry(QtCore.QRect(125, 20, 91, 20))
        self.lineEdit_host.setText(_fromUtf8(""))
        self.lineEdit_host.setObjectName(_fromUtf8("lineEdit_host"))
        self.lineEdit_user = QtGui.QLineEdit(Dialog)
        self.lineEdit_user.setGeometry(QtCore.QRect(125, 60, 91, 20))
        self.lineEdit_user.setText(_fromUtf8(""))
        self.lineEdit_user.setObjectName(_fromUtf8("lineEdit_user"))
        self.lineEdit_passwd = QtGui.QLineEdit(Dialog)
        self.lineEdit_passwd.setGeometry(QtCore.QRect(125, 100, 91, 20))
        self.lineEdit_passwd.setText(_fromUtf8(""))
        self.lineEdit_passwd.setObjectName(_fromUtf8("lineEdit_passwd"))
        self.lineEdit_db = QtGui.QLineEdit(Dialog)
        self.lineEdit_db.setGeometry(QtCore.QRect(125, 140, 91, 20))
        self.lineEdit_db.setText(_fromUtf8(""))
        self.lineEdit_db.setObjectName(_fromUtf8("lineEdit_db"))
        self.lineEdit_path = QtGui.QLineEdit(Dialog)
        self.lineEdit_path.setGeometry(QtCore.QRect(125, 180, 91, 20))
        self.lineEdit_path.setText(_fromUtf8(""))
        self.lineEdit_path.setObjectName(_fromUtf8("lineEdit_path"))
        self.lineEdit_tableKey = QtGui.QLineEdit(Dialog)
        self.lineEdit_tableKey.setGeometry(QtCore.QRect(125, 270, 91, 20))
        self.lineEdit_tableKey.setObjectName(_fromUtf8("lineEdit_tableKey"))
        self.lineEdit_attNameKey = QtGui.QLineEdit(Dialog)
        self.lineEdit_attNameKey.setGeometry(QtCore.QRect(125, 310, 91, 20))
        self.lineEdit_attNameKey.setObjectName(_fromUtf8("lineEdit_attNameKey"))
        self.lineEdit_attTypeKey = QtGui.QLineEdit(Dialog)
        self.lineEdit_attTypeKey.setGeometry(QtCore.QRect(125, 350, 91, 20))
        self.lineEdit_attTypeKey.setObjectName(_fromUtf8("lineEdit_attTypeKey"))
        self.lineEdit_dataKey = QtGui.QLineEdit(Dialog)
        self.lineEdit_dataKey.setGeometry(QtCore.QRect(125, 390, 91, 20))
        self.lineEdit_dataKey.setObjectName(_fromUtf8("lineEdit_dataKey"))
        self.textBrowser = QtGui.QTextBrowser(Dialog)
        self.textBrowser.setGeometry(QtCore.QRect(125, 430, 91, 20))
        self.textBrowser.setText(_fromUtf8(""))
        self.textBrowser.setObjectName(_fromUtf8("textBrowser"))
        #self.textBrowser.setStyleSheet("QTextBrowser{background: black;}")
        self.label = QtGui.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(15, 270, 81, 20))
        self.label.setObjectName(_fromUtf8("label"))
        self.label_2 = QtGui.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(125, 240, 61, 16))
        self.label_2.setObjectName(_fromUtf8("label_2"))
        self.label_4 = QtGui.QLabel(Dialog)
        self.label_4.setGeometry(QtCore.QRect(15, 310, 101, 20))
        self.label_4.setObjectName(_fromUtf8("label_4"))
        self.label_5 = QtGui.QLabel(Dialog)
        self.label_5.setGeometry(QtCore.QRect(15, 350, 101, 20))
        self.label_5.setObjectName(_fromUtf8("label_5"))
        self.label_6 = QtGui.QLabel(Dialog)
        self.label_6.setGeometry(QtCore.QRect(15, 390, 101, 20))
        self.label_6.setObjectName(_fromUtf8("label_6"))
        self.label_8 = QtGui.QLabel(Dialog)
        self.label_8.setGeometry(QtCore.QRect(15, 20, 71, 20))
        self.label_8.setObjectName(_fromUtf8("label_8"))
        self.label_9 = QtGui.QLabel(Dialog)
        self.label_9.setGeometry(QtCore.QRect(15, 140, 101, 20))
        self.label_9.setObjectName(_fromUtf8("label_9"))
        self.label_10 = QtGui.QLabel(Dialog)
        self.label_10.setGeometry(QtCore.QRect(15, 60, 101, 20))
        self.label_10.setObjectName(_fromUtf8("label_10"))
        self.label_11 = QtGui.QLabel(Dialog)
        self.label_11.setGeometry(QtCore.QRect(15, 180, 91, 20))
        self.label_11.setObjectName(_fromUtf8("label_11"))
        self.label_12 = QtGui.QLabel(Dialog)
        self.label_12.setGeometry(QtCore.QRect(15, 100, 101, 20))
        self.label_12.setObjectName(_fromUtf8("label_12"))
        self.label_13 = QtGui.QLabel(Dialog)
        self.label_13.setGeometry(QtCore.QRect(15, 430, 101, 20))
        self.label_13.setObjectName(_fromUtf8("label_13"))
        self.line = QtGui.QFrame(Dialog)
        self.line.setGeometry(QtCore.QRect(10, 220, 211, 20))
        self.line.setFrameShape(QtGui.QFrame.HLine)
        self.line.setFrameShadow(QtGui.QFrame.Sunken)
        self.line.setObjectName(_fromUtf8("line"))
        self.pushButton_Parse = QtGui.QPushButton(Dialog)
        self.pushButton_Parse.setGeometry(QtCore.QRect(125, 470, 91, 23))
        self.pushButton_Parse.setObjectName(_fromUtf8("pushButton_Parse"))
        self.pushButton_Parse.clicked.connect(self.handleButtonParsedClicked)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QtGui.QApplication.translate("Dialog", "Dialog", None, QtGui.QApplication.UnicodeUTF8))
        self.label.setText(QtGui.QApplication.translate("Dialog", "Table name key:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_2.setText(QtGui.QApplication.translate("Dialog", "Keywords:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_4.setText(QtGui.QApplication.translate("Dialog", "Attribute names key:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_5.setText(QtGui.QApplication.translate("Dialog", "Attribute types key:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_6.setText(QtGui.QApplication.translate("Dialog", "Data key:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_8.setText(QtGui.QApplication.translate("Dialog", "Host:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_9.setText(QtGui.QApplication.translate("Dialog", "Database:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_10.setText(QtGui.QApplication.translate("Dialog", "User:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_11.setText(QtGui.QApplication.translate("Dialog", "Path:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_12.setText(QtGui.QApplication.translate("Dialog", "Password:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_13.setText(QtGui.QApplication.translate("Dialog", "Output:", None, QtGui.QApplication.UnicodeUTF8))
        self.pushButton_Parse.setText(QtGui.QApplication.translate("Dialog", "Parse", None, QtGui.QApplication.UnicodeUTF8))

    def handleButtonParsedClicked(self):
        path = 'C:/Users/rleber/Desktop/150908/'
        #str(self.lineEdit_path.text())
        hostname = str(self.lineEdit_host.text())
        username = str(self.lineEdit_user.text())
        password = str(self.lineEdit_passwd.text())
        database = str(self.lineEdit_db.text()) 
        tableKey= str(self.lineEdit_tableKey.text())
        attNameKey= str(self.lineEdit_attNameKey.text())
        attTypeKey= str(self.lineEdit_attTypeKey.text())
        dataKey= str(self.lineEdit_dataKey.text())
    
        # Open database connection
        db = pymysql.connect(host = hostname, # your host, usually localhost 
                             user = username, # your username 
                             passwd = password, # your password 
                             db = database) # name of the data base
        
        for filename in os.listdir(path):
            csv_data = csv.reader(open(path+filename), delimiter=";")
            database = Database(db, csv_data)
            database.createTable(tableKey, attNameKey, attTypeKey, dataKey)
        self.textBrowser.setText("Finished")

if __name__ == "__main__":
    import sys
    import os
    import pymysql
    import csv
    app = QtGui.QApplication(sys.argv)
    Dialog = QtGui.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())

