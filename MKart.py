import Cust_Functions as F
import xml_v_drevo as XML
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWinExtras import QtWin
import os
import time
import subprocess
from mydesign import Ui_MainWindow  # импорт нашего сгенерированного файла
#from mydesign2 import Ui_Dialog  # импорт нашего сгенерированного файла
import sys

def showDialog(self, msg):
    msgBox = QtWidgets.QMessageBox()
    msgBox.setIcon(QtWidgets.QMessageBox.Information)
    msgBox.setText(msg)
    msgBox.setWindowTitle("Внимание!")
    msgBox.setStandardButtons(QtWidgets.QMessageBox.Ok)  # | QtWidgets.QMessageBox.Cancel)
    returnValue = msgBox.exec()

#class mywindow2(QtWidgets.QDialog):  # диалоговое окно
#    def __init__(self,parent=None,item_o="",p1=0,p2=0):
#        self.item_o = item_o
#        self.p1 = p1
#        self.p2 = p2
#        self.myparent = parent
#        super(mywindow2, self).__init__()
#        self.ui2 = Ui_Dialog()
#        self.ui2.setupUi(self)
#        self.setWindowModality(QtCore.Qt.ApplicationModal)

class mywindow(QtWidgets.QMainWindow):
    resized = QtCore.pyqtSignal()
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)

        #self.resized.connect(self.widths)

        tree = self.ui.treeWidget
        tree.setColumnCount(8)
        tree.headerItem().setText(0, QtCore.QCoreApplication.translate("MainW", "Наименование"))
        tree.headerItem().setText(1, QtCore.QCoreApplication.translate("MainW", "Обозначение"))
        tree.headerItem().setText(2, QtCore.QCoreApplication.translate("MainW", "Количество"))
        tree.headerItem().setText(3, QtCore.QCoreApplication.translate("MainW", "Материал"))
        tree.headerItem().setText(4, QtCore.QCoreApplication.translate("MainW", "Материал2"))
        tree.headerItem().setText(5, QtCore.QCoreApplication.translate("MainW", "Материал3"))
        tree.headerItem().setText(6, QtCore.QCoreApplication.translate("MainW", "ID"))
        tree.headerItem().setText(7, QtCore.QCoreApplication.translate("MainW", "Количество на изделие"))


        #tree.setColumnWidth(1, int(tree.width() * 0.1))
        #tree.setColumnWidth(0, int(tree.width() - tree.columnWidth(1) - 81) - 5)
        tree.setStyleSheet(
            "QTreeView {background-color: rgb(212, 212, 212);} QTreeView::item:hover {background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop:"
            " 0 #e7effd, stop: 1 #cbdaf1);border: 1px solid #bfcde4;} ")
        tree.setFocusPolicy(15)

        actionXML = self.ui.action_XML
        actionXML.triggered.connect(self.viborXML)
        #self.action = self.findChild(QtWidgets.QAction, "action")
        #self.action.triggered.connect(self.Smena_Parol)


    def viborXML(self):
        putt = F.f_dialog_name(self,'Выбрать XML','',"Файлы *.xml")
        s=XML.spisok_iz_xml(putt)
        self.zapoln_tree_spiskom(s)

    def zapoln_tree_spiskom(self,spisok):
        tree = self.ui.treeWidget
        tree.clear()
        n = 0
        for i in range(0, len(spisok)):
            if spisok[i][20] == '0' or spisok[i][20] == 0:
                root = QtWidgets.QTreeWidgetItem(tree)
                tmp = root
            if spisok[i][20] == '1' or spisok[i][20] == 1:
                root = QtWidgets.QTreeWidgetItem(tmp)
                tmp2 = root
            if spisok[i][20] == '2' or spisok[i][20] == 2:
                root = QtWidgets.QTreeWidgetItem(tmp2)
                tmp3 = root
            if spisok[i][20] == '3' or spisok[i][20] == 3:
                root = QtWidgets.QTreeWidgetItem(tmp3)
                tmp4 = root
            if spisok[i][20] == '4' or spisok[i][20] == 4:
                root = QtWidgets.QTreeWidgetItem(tmp4)
                tmp5 = root
            if spisok[i][20] == '5' or spisok[i][20] == 5:
                root = QtWidgets.QTreeWidgetItem(tmp5)
                tmp6 = root
            if spisok[i][20] == '6' or spisok[i][20] == 6:
                root = QtWidgets.QTreeWidgetItem(tmp6)
                tmp7 = root
            if spisok[i][20] == '7' or spisok[i][20] == 7:
                root = QtWidgets.QTreeWidgetItem(tmp7)
                tmp8 = root
            if spisok[i][20] == '8' or spisok[i][20] == 8:
                root = QtWidgets.QTreeWidgetItem(tmp8)
                tmp9 = root
            if spisok[i][20] == '9' or spisok[i][20] == 9:
                root = QtWidgets.QTreeWidgetItem(tmp9)
                tmp10 = root
            if spisok[i][20] == '10' or spisok[i][20] == 10:
                root = QtWidgets.QTreeWidgetItem(tmp10)

            for j in range(0,len(spisok[i])):
                root.setText(j, str(spisok[i][j]))
            tree.addTopLevelItem(root)
            tree.expandItem(root)
            tree.setCurrentItem(root)

            n+=1
app = QtWidgets.QApplication([])

myappid = 'Powerz.BAG.SustControlWork.0.0.0'  # !!!
QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
app.setWindowIcon(QtGui.QIcon(os.path.join("icons", "icon.png")))

S = F.scfg('Stile').split(",")
app.setStyle(S[1])

application = mywindow()
application.show()

sys.exit(app.exec())