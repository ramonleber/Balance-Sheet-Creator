'''
Created on 18.12.2016

@author: rleber
'''

from PyQt4 import QtCore, QtGui
import threading
import sys
import os
import csv
import ctypes
import webbrowser
import Tkinter
import tkFileDialog
from multiprocessing import Event
from symbol import except_clause

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    _fromUtf8 = lambda s: s

#Overwrite of showPop() to create handle to dropdown menue click
class ComboBox(QtGui.QComboBox):
    popupAboutToBeShown = QtCore.pyqtSignal()

    def showPopup(self):
        self.popupAboutToBeShown.emit()
        super(ComboBox, self).showPopup()

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName(_fromUtf8("Dialog"))
        Dialog.resize(331, 581)
        self.label_AddIssuer = QtGui.QLabel(Dialog)
        self.label_AddIssuer.setGeometry(QtCore.QRect(20, 450, 121, 20))
        self.label_AddIssuer.setObjectName(_fromUtf8("label_AddIssuer"))
        self.label_PathToBalanceSheet = QtGui.QLabel(Dialog)
        self.label_PathToBalanceSheet.setGeometry(QtCore.QRect(20, 50, 121, 20))
        self.label_PathToBalanceSheet.setObjectName(_fromUtf8("label_PathTBalanceSheet"))
        self.pushButton_OrderBalance = QtGui.QPushButton(Dialog)
        self.pushButton_OrderBalance.setEnabled(False)
        self.pushButton_OrderBalance.setGeometry(QtCore.QRect(220, 530, 91, 23))
        self.pushButton_OrderBalance.setObjectName(_fromUtf8("pushButton_OrderBalance"))
        self.pushButton_OrderBalance.clicked.connect(self.handleButtonOrderBalance)
        
        super(Ui_Dialog, self).__init__()
        self.comboBox_AddType = ComboBox(Dialog)
        self.comboBox_AddType.setEnabled(False)
        self.comboBox_AddType.setGeometry(QtCore.QRect(140, 450, 171, 22))
        self.comboBox_AddType.setObjectName(_fromUtf8("comboBox_AddType"))
        self.comboBox_AddType.activated.connect(self.handleAddType)
        self.comboBox_AddType.popupAboutToBeShown.connect(self.handleComboBox)
        
        self.line = QtGui.QFrame(Dialog)
        self.line.setGeometry(QtCore.QRect(20, 230, 291, 20))
        self.line.setFrameShape(QtGui.QFrame.HLine)
        self.line.setFrameShadow(QtGui.QFrame.Sunken)
        self.line.setObjectName(_fromUtf8("line"))
        self.label_NewIssuer = QtGui.QLabel(Dialog)
        self.label_NewIssuer.setGeometry(QtCore.QRect(20, 250, 121, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_NewIssuer.setFont(font)
        self.label_NewIssuer.setObjectName(_fromUtf8("label_NewIssuer"))
        self.label_IssuerName = QtGui.QLabel(Dialog)
        self.label_IssuerName.setGeometry(QtCore.QRect(20, 290, 121, 20))
        self.label_IssuerName.setObjectName(_fromUtf8("label_IssuerName"))
        self.label_DateOfIssue = QtGui.QLabel(Dialog)
        self.label_DateOfIssue.setGeometry(QtCore.QRect(20, 330, 121, 20))
        self.label_DateOfIssue.setObjectName(_fromUtf8("label_DateOfIssue"))
        self.label_IBAN = QtGui.QLabel(Dialog)
        self.label_IBAN.setGeometry(QtCore.QRect(20, 370, 91, 20))
        self.label_IBAN.setObjectName(_fromUtf8("label_IBAN"))
        self.label_Amount = QtGui.QLabel(Dialog)
        self.label_Amount.setGeometry(QtCore.QRect(20, 410, 121, 20))
        self.label_Amount.setObjectName(_fromUtf8("label_Amount"))
        self.line_2 = QtGui.QFrame(Dialog)
        self.line_2.setGeometry(QtCore.QRect(20, 490, 291, 16))
        self.line_2.setFrameShape(QtGui.QFrame.HLine)
        self.line_2.setFrameShadow(QtGui.QFrame.Sunken)
        self.line_2.setObjectName(_fromUtf8("line_2"))
        self.lineEdit_IssuerName = QtGui.QLineEdit(Dialog)
        self.lineEdit_IssuerName.setEnabled(False)
        self.lineEdit_IssuerName.setGeometry(QtCore.QRect(140, 290, 171, 20))
        self.lineEdit_IssuerName.setText(_fromUtf8(""))
        self.lineEdit_IssuerName.setReadOnly(True)
        self.lineEdit_IssuerName.setObjectName(_fromUtf8("lineEdit_IssuerName"))
        self.lineEdit_DateOfIssue = QtGui.QLineEdit(Dialog)
        self.lineEdit_DateOfIssue.setEnabled(False)
        self.lineEdit_DateOfIssue.setGeometry(QtCore.QRect(140, 330, 171, 20))
        self.lineEdit_DateOfIssue.setText(_fromUtf8(""))
        self.lineEdit_DateOfIssue.setReadOnly(True)
        self.lineEdit_DateOfIssue.setObjectName(_fromUtf8("lineEdit_DateOfIssue"))
        self.lineEdit_IBAN = QtGui.QLineEdit(Dialog)
        self.lineEdit_IBAN.setEnabled(False)
        self.lineEdit_IBAN.setGeometry(QtCore.QRect(140, 370, 171, 20))
        self.lineEdit_IBAN.setText(_fromUtf8(""))
        self.lineEdit_IBAN.setReadOnly(True)
        self.lineEdit_IBAN.setObjectName(_fromUtf8("lineEdit_IBAN"))
        self.lineEdit_Amount = QtGui.QLineEdit(Dialog)
        self.lineEdit_Amount.setEnabled(False)
        self.lineEdit_Amount.setGeometry(QtCore.QRect(140, 410, 171, 20))
        self.lineEdit_Amount.setText(_fromUtf8(""))
        self.lineEdit_Amount.setReadOnly(True)
        self.lineEdit_Amount.setObjectName(_fromUtf8("lineEdit_Amount"))
        self.pushButton_BrowseBalanceSheet = QtGui.QPushButton(Dialog)
        self.pushButton_BrowseBalanceSheet.setGeometry(QtCore.QRect(220, 50, 91, 23))
        self.pushButton_BrowseBalanceSheet.setObjectName(_fromUtf8("pushButton_BrowseBalanceSheet"))
        self.pushButton_BrowseBalanceSheet.clicked.connect(self.handleButtonBrowse)
        self.line_3 = QtGui.QFrame(Dialog)
        self.line_3.setGeometry(QtCore.QRect(20, 130, 291, 20))
        self.line_3.setFrameShape(QtGui.QFrame.HLine)
        self.line_3.setFrameShadow(QtGui.QFrame.Sunken)
        self.line_3.setObjectName(_fromUtf8("line_3"))
        self.label_OpenBalanceSheet = QtGui.QLabel(Dialog)
        self.label_OpenBalanceSheet.setGeometry(QtCore.QRect(20, 10, 121, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_OpenBalanceSheet.setFont(font)
        self.label_OpenBalanceSheet.setObjectName(_fromUtf8("label_OpenBalanceSheet"))
        self.label_TypesOfIssuer = QtGui.QLabel(Dialog)
        self.label_TypesOfIssuer.setGeometry(QtCore.QRect(20, 150, 121, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_TypesOfIssuer.setFont(font)
        self.label_TypesOfIssuer.setObjectName(_fromUtf8("label_TypesOfIssuer"))
        self.label_ListTypes = QtGui.QLabel(Dialog)
        self.label_ListTypes.setGeometry(QtCore.QRect(20, 190, 121, 20))
        self.label_ListTypes.setObjectName(_fromUtf8("label_ListTypes"))
        self.pushButton_ShowTypes = QtGui.QPushButton(Dialog)
        self.pushButton_ShowTypes.setEnabled(False)
        self.pushButton_ShowTypes.setGeometry(QtCore.QRect(220, 190, 91, 23))
        self.pushButton_ShowTypes.setObjectName(_fromUtf8("pushButton_ShowTypes"))
        self.pushButton_ShowTypes.clicked.connect(self.handleButtonShowTypes)
        self.label_PathToBalanceSheet_2 = QtGui.QLabel(Dialog)
        self.label_PathToBalanceSheet_2.setGeometry(QtCore.QRect(20, 90, 111, 20))
        self.label_PathToBalanceSheet_2.setObjectName(_fromUtf8("label_PathTBalanceSheet_2"))
        self.lineEdit_ShowSelectedFile = QtGui.QLineEdit(Dialog)
        self.lineEdit_ShowSelectedFile.setGeometry(QtCore.QRect(140, 90, 171, 20))
        self.lineEdit_ShowSelectedFile.setAcceptDrops(True)
        self.lineEdit_ShowSelectedFile.setText(_fromUtf8(""))
        self.lineEdit_ShowSelectedFile.setReadOnly(True)
        self.lineEdit_ShowSelectedFile.setObjectName(_fromUtf8("lineEdit_ShowSelectedFile"))

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QtGui.QApplication.translate("Dialog", "Dialog", None, QtGui.QApplication.UnicodeUTF8))
        self.label_AddIssuer.setText(QtGui.QApplication.translate("Dialog", "Add Issuer to a Type:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_PathToBalanceSheet.setText(QtGui.QApplication.translate("Dialog", "Path to Balance Sheet:", None, QtGui.QApplication.UnicodeUTF8))
        self.pushButton_OrderBalance.setText(QtGui.QApplication.translate("Dialog", "Order Balance", None, QtGui.QApplication.UnicodeUTF8))
        self.label_NewIssuer.setText(QtGui.QApplication.translate("Dialog", "New Issuer Found:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_IssuerName.setText(QtGui.QApplication.translate("Dialog", "Issuer Name:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_DateOfIssue.setText(QtGui.QApplication.translate("Dialog", "Date of Issue:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_IBAN.setText(QtGui.QApplication.translate("Dialog", "IBAN:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_Amount.setText(QtGui.QApplication.translate("Dialog", "Amount:", None, QtGui.QApplication.UnicodeUTF8))
        self.pushButton_BrowseBalanceSheet.setText(QtGui.QApplication.translate("Dialog", "Browse", None, QtGui.QApplication.UnicodeUTF8))
        self.label_OpenBalanceSheet.setText(QtGui.QApplication.translate("Dialog", "Open Balance Sheet:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_TypesOfIssuer.setText(QtGui.QApplication.translate("Dialog", "Types of Issuer:", None, QtGui.QApplication.UnicodeUTF8))
        self.label_ListTypes.setText(QtGui.QApplication.translate("Dialog", "Show/Edit List of Types:", None, QtGui.QApplication.UnicodeUTF8))
        self.pushButton_ShowTypes.setText(QtGui.QApplication.translate("Dialog", "List of Types", None, QtGui.QApplication.UnicodeUTF8))
        self.label_PathToBalanceSheet_2.setText(QtGui.QApplication.translate("Dialog", "Selected File:", None, QtGui.QApplication.UnicodeUTF8))
   
    def getIssuerTypesFilename(self):
        listOfTypesPath = os.getcwd()
        issuerTypesFileName =  listOfTypesPath  + "\issuer_types.csv"
        return issuerTypesFileName
    
    def getBalanceSheet(self):
        balanceSheet = str(self.lineEdit_ShowSelectedFile.text())
        return balanceSheet
        
    def getTypesList(self):
        typesList = list()
        issuerTypes = ui.getIssuerTypesFilename()
        typesSheetData = csv.reader(open(issuerTypes), delimiter=";")
        typesSheetList = list(typesSheetData)
        for typesRow in typesSheetList:
            typesList.append(typesRow[0])
        
        return typesList
    
    def setComboBoxTypes(self):
        if os.path.isfile(ui.getIssuerTypesFilename()):
            self.comboBox_AddType.clear()
            self.comboBox_AddType.addItems(ui.getTypesList())
            
    def setNewType(self, newType):
        self.lineEdit_IssuerName.setText(newType[3])
        self.lineEdit_DateOfIssue.setText(newType[0])
        self.lineEdit_IBAN.setText(newType[5])
        self.lineEdit_Amount.setText(newType[11]+' '+newType[10]+' '+newType[12])
   
    def writeFile(self, fileParam, mode, param):
        with open(fileParam, mode) as myfile:
                wr = csv.writer(myfile, dialect='excel', delimiter=';')
                wr.writerow(param)
    
    def isOrdered(self):
        balanceSheetData = csv.reader(open(ui.getBalanceSheet()), delimiter=";")
        balanceSheetList = list(balanceSheetData)
        for row in balanceSheetList:
            if row[0] == 'Show Ordered Table:':
                return True
        
        return False
    
    def createDefaultTypesList(self, openBrowser):
        listOfTypesPath = os.getcwd()
        
        if not os.path.isfile(ui.getIssuerTypesFilename()):
            MessageBox = ctypes.windll.user32.MessageBoxA
            MessageBox(None, 'No list of Issuer Types was found. The list must be in the home directory: "'+listOfTypesPath+'" and must be called "issuer_types.csv". A new default list will now be created. If you have already created a list of Issuer Types before, search for "issuer_types.csv" in your file explorer and replace the new created default list in the home directory.', 'No List of Types found!', 0)
            with open(ui.getIssuerTypesFilename(), 'wb') as myfile:
                defaultList1 = ['Housing:','','','','','','','','','','','','']
                defaultList2 = ['Insurance:','','','','','','','','','','','','']
                defaultList3 = ['Shopping:','','','','','','','','','','','','']
                defaultList4 = ['Car:','','','','','','','','','','','','']
                defaultList5 = ['Banking:','','','','','','','','','','','','']
                defaultList6 = ['Miscellaneous:','','','','','','','','','','','','']
                wr = csv.writer(myfile, dialect='excel', delimiter=';',quoting=csv.QUOTE_NONE)
                wr.writerow(defaultList1)
                wr.writerow(defaultList2)
                wr.writerow(defaultList3)
                wr.writerow(defaultList4)
                wr.writerow(defaultList5)
                wr.writerow(defaultList6)
        
        elif openBrowser:
            webbrowser.open(ui.getIssuerTypesFilename())
           
    def enableElements(self):
        self.pushButton_BrowseBalanceSheet.setEnabled(True)
        self.lineEdit_ShowSelectedFile.setEnabled(True)
           
    def closeDialog(self):
        webbrowser.open(ui.getBalanceSheet())
        Dialog.close()
        
    def handleComboBox(self):
        ui.setComboBoxTypes()
   
    def handleButtonBrowse(self):
        FILEOPENOPTIONS = dict(defaultextension='.csv', filetypes=[('All files','*.*'), ('cvs file','*.csv')])
        root = Tkinter.Tk()
        root.withdraw()
        filename = tkFileDialog.askopenfilename(parent=root,**FILEOPENOPTIONS)
        self.lineEdit_ShowSelectedFile.setText(filename)
        self.pushButton_OrderBalance.setEnabled(True)
        self.pushButton_ShowTypes.setEnabled(True)
        self.comboBox_AddType.setEnabled(True)
        self.lineEdit_IssuerName.setEnabled(True)
        self.lineEdit_DateOfIssue.setEnabled(True)
        self.lineEdit_IBAN.setEnabled(True)
        self.lineEdit_Amount.setEnabled(True)
            
    def handleButtonShowTypes(self):
        ui.createDefaultTypesList(True)
        
    def handleButtonOrderBalance(self):
        if ui.isOrdered():
            MessageBox = ctypes.windll.user32.MessageBoxA
            MessageBox(None, 'This Balance Sheet has already been processed', 'Balance already processed!', 0)
        
        else:    
            ui.writeFile(ui.getBalanceSheet(), 'ab', ['','','','','','','','','','','','',''])
            ui.writeFile(ui.getBalanceSheet(), 'ab', ['Show Ordered Table:','','','','','','','','','','','',''])
            self.pushButton_BrowseBalanceSheet.setEnabled(False)
            self.lineEdit_ShowSelectedFile.setEnabled(False)
            self.pushButton_OrderBalance.setEnabled(False)
            ui.createDefaultTypesList(False)    
            t1.start()
        
    def handleAddType(self):
        listOfTypeLists = []
        listFinished = False
        listExtended = False
        selectedType = self.comboBox_AddType.currentText()
        issuerTypesFile = ui.getIssuerTypesFilename()
        typesSheetData = csv.reader(open(issuerTypesFile), delimiter=";")
        typesSheetList = list(typesSheetData)
        
        #Checks which type got selected in the dropdown and adds it to the typeslist
        for typesRow in typesSheetList:
            if typesRow[0] == selectedType:
                columnIndex = 0
                for column in typesRow:
                    if column == '':
                        typesRow[columnIndex] = self.lineEdit_IssuerName.text()
                        listFinished = True
                        break
                    
                    columnIndex += 1
                
                if not listFinished:
                    typesRow.append(self.lineEdit_IssuerName.text())
                    listExtended = True
                    listLength = len(typesRow)
                
                listOfTypeLists.append(typesRow)
            
            elif not typesRow[0] == selectedType:
                listOfTypeLists.append(typesRow)
        
        #Fills the types which not got selected with blanks, to keep the csv-file persistant
        if listExtended:    
            for typesList in listOfTypeLists:
                if not listLength == len(typesList):
                    typesList.append('')
        
        print listOfTypeLists
        
        first = True
        
        for listOf in listOfTypeLists:
            if first :
                ui.writeFile(issuerTypesFile, 'wb', listOf)
                first = False
            
            else:
                ui.writeFile(issuerTypesFile, 'ab', listOf)
        
        e.set()
        
def startOrder(e):
        habenSollIndex = 12     #Column of debit or credit sign in balance sheet
        amountIndex = 11        #Column of amount in balance sheet
        startList = 14          #Line to start searching for values in balance sheet
        debit = 'S'             #Character to search for if its debit
        credit = 'H'            #Character to search for if its credit
        
        blanceSheet = ui.getBalanceSheet()
        balanceSheetData = csv.reader(open(blanceSheet), delimiter=";")
        balanceSheetList = list(balanceSheetData)
        issuerTypes = ui.getIssuerTypesFilename()
        typesSheetData = csv.reader(open(issuerTypes), delimiter=";")
        typesSheetList = list(typesSheetData)
        
        #Checks if a type is already in the types list. 
        #If a type is not in the types list, it will be added.
        typesListLong = list()
        for typesList in typesSheetList:
            typesListLong = typesListLong + typesList
        
        rowIndex = 0    
        for balanceRow in balanceSheetList:
            rowIndex += 1
            if rowIndex >= startList and not balanceRow[3] == '' and balanceRow[3] not in typesListLong:
                print balanceRow[3]
                eventComboBoxSelected = e.wait(0)
                e.clear()
                while not e.isSet():
                    if eventComboBoxSelected:
                        a = 0
                    
                    else:
                        ui.setNewType(balanceRow)
                            
            issuerTypes = ui.getIssuerTypesFilename()
            typesSheetData = csv.reader(open(issuerTypes), delimiter=";")
            typesSheetList = list(typesSheetData)
            typesListLong = list()
            for typesList in typesSheetList:
                typesListLong = typesListLong + typesList    
                
        #Sort of types and calculate the total amount of each type
        for typesRow in typesSheetList:
            ui.writeFile(blanceSheet, 'ab', ['','','','','','','','','','','','',''])
            ui.writeFile(blanceSheet, 'ab', [typesRow[0],'','','','','','','','','','','',''])
            total= 0
            
            for typesColumn in typesRow:
                rowIndex = 0
                for balanceRow in balanceSheetList:
                    rowIndex += 1
                    balanceRowTmp = balanceRow
                    if balanceRowTmp[3] == typesColumn and not balanceRowTmp[3] == '':    
                        ui.writeFile(blanceSheet, 'ab', balanceRowTmp)
                        if balanceRowTmp[habenSollIndex] == debit:
                            total = total - float(balanceRowTmp[amountIndex].replace('.','').replace(',','.'))
                            
                        elif balanceRowTmp[habenSollIndex] == credit:
                            total = total + float(balanceRowTmp[amountIndex].replace('.','').replace(',','.'))
            
            ui.writeFile(blanceSheet, 'ab', ['','','','','','','','','','','Total:',total,''])
            ui.enableElements()                      
        
        ui.closeDialog()
        
if __name__ == "__main__":    
    app = QtGui.QApplication(sys.argv)
    Dialog = QtGui.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    e = threading.Event()
    t1 = threading.Thread(name='order', target=startOrder, args=(e,))
    sys.exit(app.exec_())
    