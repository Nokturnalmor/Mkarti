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

        but_add_bd = self.ui.pushButton_add_v_bd
        but_add_bd.clicked.connect(self.dob_izd_k_bd)

        but_cr_mk = self.ui.pushButton_create_MK
        but_cr_mk.clicked.connect(self.create_mk)
        tabl = self.ui.table_zayavk
        shapka = ['Файл', 'Изделие', 'Кол-во']
        tabl.setColumnCount(3)
        tabl.setHorizontalHeaderLabels(shapka)

        #tree.setColumnWidth(1, int(tree.width() * 0.1))
        #tree.setColumnWidth(0, int(tree.width() - tree.columnWidth(1) - 81) - 5)
        tree.setStyleSheet(
            "QTreeView {background-color: rgb(212, 212, 212);} QTreeView::item:hover {background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop:"
            " 0 #e7effd, stop: 1 #cbdaf1);border: 1px solid #bfcde4;} ")
        tree.setFocusPolicy(15)

        actionXML = self.ui.action_XML
        actionXML.triggered.connect(self.viborXML)

    def kol_po_zayav(self,sp_xml_tmp,kol):
        for i in sp_xml_tmp:
            i[2] = int(i[2]) * int(kol)
            i[7] = int(i[7]) * int(kol)
        return sp_xml_tmp

    def create_mk(self):
        tabl = self.ui.table_zayavk
        s_vert = []

        sp_izd = F.spisok_iz_wtabl(tabl)
        for i in sp_izd:
            putt = i[0]
            sp_xml_tmp = XML.spisok_iz_xml(putt)
            sp_xml_tmp = self.kol_po_zayav(sp_xml_tmp,i[2])
            for j in sp_xml_tmp:
                if len(s_vert) == 0:
                    s_vert.append(['' for x in list(range(0, len(j)))])
                    nach_sod = len(j)
                s_vert.append(j)

        bd = F.otkr_f(F.tcfg('BD_dse'), False, "|")
        for i in range(1, len(s_vert)):
            ima = s_vert[i][0]
            nn = s_vert[i][1]
            kol_det_vseg = s_vert[i][7]
            flag = 0
            for j in bd:
                if ima == j[1] and nn == j[0]:
                    flag = 1
                    break
            if flag == 0:
                showDialog(self,'Не найден в БД ' + ima + ' ' + nn)
                return
            if j[2] == '':
                showDialog(self, 'Не найдена техкарта ' + ima + ' ' + nn)
                return
            tk = F.otkr_f(F.scfg('add_docs') + os.sep + j[2] + "_" + nn + '.txt', False, "|")
            tk = self.grup_tk_po_rabc(tk,kol_det_vseg)
            self.ogran = nach_sod-1
            for k in tk:
                s_vert = self.dob_etap(s_vert,k[0],k[1],k[2],i,self.ogran)

        F.zapoln_wtabl(self,s_vert,tabl,0,0,"","",200,True,'')

    def summa_rc(self,rc):
        s = ''
        for i in rc:
            if F.is_numeric(i):
                s+=str(i)
        s = int(s)
        return s

    def dob_kol(self,spis,nomer,ima):
        for i in range(0, len(spis)):
            if i == 0:
                spis[i].insert(nomer,ima)
            else:
                spis[i].insert(nomer, '')
        return spis

    def dob_etap(self,spis,rc,vrem,oper,stroka,ogran):

        flag = 0
        for i in range(ogran+1,len(spis[0])):
            if flag == 1:
                break
            if spis[0][i] == rc:
                flag = 1
                spis[stroka][i] = str(vrem) + "$" + oper
                self.ogran = i-1
                break
            if self.summa_rc(rc) < self.summa_rc(spis[0][i]):
                j = i-1
                while j>=self.ogran:
                    if j <= ogran:
                        spis = self.dob_kol(spis, j + 1, rc)
                        spis[stroka][j + 1] = str(vrem) + "$" + oper
                        self.ogran = j
                        flag = 1
                        break
                    if self.summa_rc(rc) == self.summa_rc(spis[0][j]):
                        spis[stroka][j] = str(vrem) + "$" + oper
                        self.ogran = j-1
                        flag = 1
                        break
                    if self.summa_rc(rc) > self.summa_rc(spis[0][j]):
                        spis = self.dob_kol(spis, j+1, rc)
                        spis[stroka][j+1] = str(vrem) + "$" + oper
                        self.ogran = j-1
                        flag = 1
                        break
                    if j == self.ogran:
                        spis = self.dob_kol(spis, self.ogran , rc)
                        spis[stroka][self.ogran] = str(vrem) + "$" + oper
                        flag = 1
                        break
                    j-=1
            else:
                j = i + 1
                while j <= len(spis[0])-1:
                    if self.summa_rc(rc) == self.summa_rc(spis[0][j]):
                        spis[stroka][j] = str(vrem) + "$" + oper
                        self.ogran = j
                        flag = 1
                        break
                    if self.summa_rc(rc) < self.summa_rc(spis[0][j]):
                        spis = self.dob_kol(spis, j -1, rc)
                        spis[stroka][j - 1] = str(vrem) + "$" + oper
                        self.ogran = j - 1
                        flag = 1
                        break
                    j += 1
                if flag == 0:
                    spis = self.dob_kol(spis, len(spis[0]), rc)
                    spis[stroka][len(spis[0])-1] = str(vrem) + "$" + oper
                    self.ogran = len(spis[0])-1
                    flag = 1
                    break

        if flag == 0:
            spis = self.dob_kol(spis,len(spis[0]),rc)
            spis[stroka][len(spis[0])-1]= str(vrem) + "$" + oper
            self.ogran = len(spis[0]) - 1
        return spis



    def grup_tk_po_rabc(self,tk,kol_det_vseg):
        spis = []
        flag = 0
        for itk in tk:
            if len(itk) == 21:
                if itk[20] == '0' and flag == 1:
                    return
                if itk[20] == '0' and flag == 0:
                    flag = 1
                if itk[20] == '1':
                    rc = itk[4]
                    vrem = F.valm(itk[6]) + kol_det_vseg * F.valm(itk[7])
                    vrem = round(vrem,1)
                    n_op = itk[2]
                    if len(spis) > 0:
                        if spis[-1][0] == rc:
                            spis[-1][1]= round(spis[-1][1]+ vrem)
                            spis[-1][2]+=';' + n_op
                        else:
                            spis.append([rc, vrem, n_op])
                    else:
                        spis.append([rc,vrem,n_op])
        return spis


    def dob_izd_k_bd(self):
            tree = self.ui.treeWidget
            s = F.spisok_dreva(tree)
            bd = F.otkr_f(F.tcfg('BD_dse'),False,"|")
            n = 0
            for i in s:
                ima = i[0]
                nn = i[1]
                flag = 0
                for j in bd:
                    if ima==j[1] and nn == j[0]:
                        flag=1
                        break
                if flag == 0:
                    bd.append([nn,ima,'',''])
                    n+=1
            F.zap_f(F.tcfg('BD_dse'),bd,'|')
            if n == 0:
                showDialog(self,'Новых ДСЕ не добавлено')
            else:
                showDialog(self, 'Добавлено ' + str(n) + 'ед. ДСЕ')

    def viborXML(self):
        vklad = self.ui.tabWidget
        putt = F.f_dialog_name(self,'Выбрать XML','',"Файлы *.xml")
        if putt == '':
            return
        s=XML.spisok_iz_xml(putt)
        if vklad.currentIndex() == 0:
            self.zapoln_tree_spiskom(s)
        if vklad.currentIndex() == 1:
            self.dob_izd(s,putt)

    def nalich_dannih_v_tk(self,n_dse,n_tk,nomer_st):
        tk = F.otkr_f(F.scfg('add_docs')+ os.sep + n_tk +"_"+n_dse+'.txt', False, "|")
        flag = 0
        for i in tk:
            if len(i) == 21:
                if i[20] == '0' and flag == 1:
                    return
                if i[20] == '0' and flag == 0:
                    flag = 1
                if i[20] == '1':
                    if i[nomer_st] == "":
                        return False
        if flag == 1:
            return True
        else:
            return False



    def nalich_tk(self,spisok):
        bd = F.otkr_f(F.tcfg('BD_dse'), False, "|")
        s_bd = []
        for i in spisok:
            ima = i[0]
            nn = i[1]
            flag_bd = 0
            flag_tk = 0
            flag_marsh = 0
            flag_vrema = 0
            for j in bd:
                if ima == j[1] and nn == j[0]:
                    flag_bd = 1
                    if j[2] != '':
                        flag_tk = 1
                        if self.nalich_dannih_v_tk(nn,j[2],4) == True:
                            flag_marsh = 1
                        if self.nalich_dannih_v_tk(nn,j[2],6) == True and self.nalich_dannih_v_tk(nn,j[2],7) == True:
                            flag_vrema = 1
                    break
            if flag_bd == 0:
                s_bd.append('нет в базе ' + " " + nn + ' ' + ima)
            if flag_tk == 0:
                s_bd.append('нет техкарты '+ " " + nn+ " " + ima)
            if flag_marsh == 0:
                s_bd.append('нет маршрутов в тк '+ " " + nn+ " " + ima)
            if flag_vrema == 0:
                s_bd.append('нет времени в тк '+ " " + nn+ " " + ima)

        return s_bd

    def dob_izd(self,spisok,putt):
        sp_tk = self.nalich_tk(spisok)
        if len(sp_tk) > 0:
            viv= ''
            for i in sp_tk:
                viv += i + '\n'
            F.copy_bufer(viv)
            showDialog(self,"Скопировано в буфер:" + '\n' + viv)
            return
        tabl = self.ui.table_zayavk
        s = F.spisok_iz_wtabl(tabl,'',True)
        s.append([putt,spisok[0][0],''])
        edit = {2}
        F.zapoln_wtabl(self,s,tabl,0,edit,(),(),200,True,"")

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