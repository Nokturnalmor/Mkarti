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

def showDialog_Y_C(self, msg):
    msgBox = QtWidgets.QMessageBox()
    msgBox.setIcon(QtWidgets.QMessageBox.Information)
    msgBox.setText(msg)
    msgBox.setWindowTitle("Внимание!")
    msgBox.setStandardButtons(QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel)
    returnValue = msgBox.exec()
    return returnValue

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
        self.setWindowTitle("Создание маршрутных карт")
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
        for _ in range(0,8):
            tree.resizeColumnToContents(_)
        tree.setSelectionBehavior(1)
        F.ust_cvet_videl_tab(tree)
        tree.setSelectionMode(1)

        tab = self.ui.tabWidget
        tab.currentChanged.connect(self.tab_click)

        self.tabl_nomenk = self.ui.table_nomenkl
        F.ust_cvet_videl_tab(self.tabl_nomenk)
        self.tabl_nomenk.setSelectionBehavior(1)
        self.tabl_nomenk.setSelectionMode(1)

        self.tabl_mk = self.ui.table_spis_MK
        F.ust_cvet_videl_tab(self.tabl_mk)
        self.tabl_mk.setSelectionBehavior(1)
        self.tabl_mk.setSelectionMode(1)
        self.tabl_mk.cellChanged[int, int].connect(self.corr_mk)

        self.tabl_brak = self.ui.table_brak
        F.ust_cvet_videl_tab(self.tabl_brak)
        self.tabl_brak.setSelectionBehavior(1)
        self.tabl_brak.setSelectionMode(1)
        self.tabl_brak.clicked.connect(self.click_brak)
        self.tabl_brak.doubleClicked.connect(self.tabl_brak_dbl_clk)

        butt_vib_nomen = self.ui.pushButton_ass_nomen_MK
        butt_vib_nomen.clicked.connect(self.ass_dse_to_mk)
        F.ust_cvet_videl_tab(butt_vib_nomen)

        combo_PY = self.ui.comboBox_PY
        combo_PY.activated[int].connect(self.vibor_PY)
        combo_proekt = self.ui.comboBox_np
        combo_proekt.activated[int].connect(self.vibor_pr)

        spis = F.otkr_f(F.tcfg('BD_Proect'), False, '|', False, False)

        for i in range(10,len(spis)):
            if spis[i][3] == 'к производству':
                combo_PY.addItem(spis[i][1])
                combo_proekt.addItem(spis[i][0])

        tabl_cr_stukt = self.ui.table_razr_MK
        self.clear_mk2()
        self.edit_cr_mk = {2, 3, 4, 5, 8, 9, 19}

        but_clear_mk = self.ui.pushButton_create_mk_clear
        but_clear_mk.clicked.connect(self.clear_mk2)


        but_add_gl_uzel = self.ui.pushButton_create_koren
        but_add_gl_uzel.clicked.connect(self.add_gl_uzel)

        but_add_vhod = self.ui.pushButton_create_vxodyash
        but_add_vhod.clicked.connect(self.add_vhod)

        but_udal_uzel = self.ui.pushButton_create_udalituzel
        but_udal_uzel.clicked.connect(self.del_uzel)

        but_add_bd = self.ui.pushButton_add_v_bd
        but_add_bd.clicked.connect(self.dob_izd_k_bd)

        but_cr_mk = self.ui.pushButton_create_MK
        but_cr_mk.clicked.connect(self.create_mk)

        but_save_mk = self.ui.pushButton_save_MK
        but_save_mk.clicked.connect(self.save_mk)

        but_add_v_mk = self.ui.pushButton_add_v_MK
        but_add_v_mk.clicked.connect(self.add_v_mk)

        self.but_ass_brak_to_mk = self.ui.pushButton_ass_brak_to_mk
        self.but_ass_brak_to_mk.clicked.connect(self.ass_brak_to_mk)
        self.ui.pushButton_ass_brak_to_mk.setEnabled(False)

        self.but_open_mk = self.ui.pushButton_open_mk
        self.but_open_mk.clicked.connect(self.open_mk)

        self.but_close_mk = self.ui.pushButton_close_mk
        self.but_close_mk.clicked.connect(self.close_mk)

        self.but_del_mk = self.ui.pushButton_del_mk
        self.but_del_mk.clicked.connect(self.del_mk)

        self.ui.pushButton_clear_label.clicked.connect(self.del_ass)

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

    def tabl_brak_dbl_clk(self):
        label = self.ui.label_opis_braka
        strok = self.tabl_brak.currentRow()
        label_brak = self.ui.label_opis_braka
        n_k = 4
        if self.tabl_brak.currentColumn() == n_k:
            if self.tabl_brak.item(self.tabl_brak.currentRow(), n_k).text().replace('Фото:','') != "":
                sp_foto = self.tabl_brak.item(self.tabl_brak.currentRow(), n_k).text().split(')(')
                sp_pap = F.spis_files(F.scfg('foto_brak'))[0][1]
                for j in range(len(sp_foto)):
                    sp_foto[j] = sp_foto[j].replace(')', '')
                    sp_foto[j] = sp_foto[j].replace('(', '')
                    for i in range(len(sp_pap)):
                        if F.nalich_file(F.scfg('foto_brak') + os.sep + sp_pap[i] + os.sep + sp_foto[j]) == True:
                            F.zapyst_file(F.scfg('foto_brak') + os.sep + sp_pap[i] + os.sep + sp_foto[j])
                return
            return
        if label == "":
            return
        nom_id = self.naiti_parametr_v_stroke(label.text(), 'ID:')

        if nom_id == "":
            return
        kol_det = self.tabl_brak.item(strok,8).text().replace('Количество:','')


        nom_mk_isprav = self.naiti_parametr_v_stroke(label_brak.text(), 'Изгот.вновь по МК:')
        if nom_mk_isprav != '':
            F.msgbox(f'ДСЕ уже изготавливается по МК №{nom_mk_isprav}')
            return
        tree = self.ui.treeWidget
        spis_tree = F.spisok_dreva(tree)
        if spis_tree == []:
            F.msgbox("Не открыто древо")
            self.viborXML()
            return
        nom_kol_id = F.nom_kol_po_imen(tree,'ID')
        for i in range(len(spis_tree)):
            if spis_tree[i][nom_kol_id] == nom_id:
                uroven = spis_tree[i][20]
                rez = F.videlit_tree_nom(tree, i + 1)
                if rez == False:
                    F.msgbox(f'Деталь {spis_tree[i][20]} не найдена')
                    return
                self.add_v_mk()

                table = self.ui.table_razr_MK
                table.item(table.rowCount() - 1, F.nom_kol_po_imen(table, 'Кол. по заявке')).setText(kol_det)
                table.setCurrentCell(table.rowCount() - 1, 1)
                for j in range(i+1, len(spis_tree)):
                    if spis_tree[j][20] > uroven:
                        rez = F.videlit_tree_nom(tree,i+1)
                        if rez == False:
                            F.msgbox(f'Деталь {spis_tree[j][20]} не найдена')
                            self.ui.tabWidget.setCurrentIndex(0)
                            return
                        self.add_v_mk()
                    else:
                        break

                return





    def click_brak(self):
        label = self.ui.label_opis_braka
        strok = self.tabl_brak.currentRow()
        label.setText('')
        nom_mk = ""
        nom_ID = ""
        if strok == -1:
            return
        if self.tabl_brak.currentColumn() == 4:
            return
        nom_acta = self.tabl_brak.item(strok, 0).text().replace('Номер акта:', '')
        nom_nar = self.tabl_brak.item(strok,3).text().replace('Номер наряда:','')
        spis_nar = F.otkr_f(F.tcfg('Naryad'), False, '|')
        kol_nom_nar = F.nom_kol_po_im_v_shap(spis_nar,'№')
        kol_nom_mk = F.nom_kol_po_im_v_shap(spis_nar,'МК')
        kol_nom_ID =F.nom_kol_po_im_v_shap(spis_nar,'ID')
        kol_nom_oper = F.nom_kol_po_im_v_shap(spis_nar,'N операции')
        for i in range(len(spis_nar)-1,0,-1):
            if spis_nar[i][kol_nom_nar] == nom_nar:
                nom_mk = spis_nar[i][kol_nom_mk]
                nom_ID = spis_nar[i][kol_nom_ID]
                nom_oper = spis_nar[i][kol_nom_oper]
                break
        if nom_mk == "":
            F.msgbox(f'Наряд {nom_nar} не найден.')
            return
        if nom_ID == '':
            return

        sp_tabl_mk = F.otkr_f(F.scfg('mk_data') + os.sep + nom_mk + '.txt', False, '|')
        if sp_tabl_mk == ['']:
            F.msgbox(f'Маршрутная карта {nom_mk} не найдена.')
            return
        kol_nom_ID = F.nom_kol_po_im_v_shap(sp_tabl_mk,'ID')
        kol_naim_det = F.nom_kol_po_im_v_shap(sp_tabl_mk,'Наименование')
        kol_nom_det =F.nom_kol_po_im_v_shap(sp_tabl_mk,'Обозначение')
        for i in range(len(sp_tabl_mk)):
            if sp_tabl_mk[i][kol_nom_ID] == nom_ID:
                naim_det = sp_tabl_mk[i][kol_naim_det]
                nom_det = sp_tabl_mk[i][kol_nom_det]
                break

        spis_mk = F.otkr_f(F.tcfg('bd_mk'), False, '|', False, False)
        nom_izg_vn_mk = F.naiti_v_spis_1_1(spis_mk,F.nom_kol_po_im_v_shap(spis_mk,'Основание'),nom_acta,0,True)
        label.setText(f'ID:{nom_ID}       ДСЕ:{naim_det.strip()} {nom_det.strip()}      МК:{nom_mk}      '
                      f'№ операции:{nom_oper}      Изгот.вновь по МК:{nom_izg_vn_mk}')
        return

    def del_ass(self):
        self.ui.label_ass.clear()

    def close_mk(self):
        if self.tabl_mk.currentRow() == -1:
            return
        row = self.tabl_mk.currentRow()
        kol = F.nom_kol_po_imen(self.tabl_mk, 'Статус')
        spis_mk = F.otkr_f(F.tcfg('bd_mk'), False, '|', False, False)
        if spis_mk[row + 1][kol] == "Открыта":
            spis_mk[row + 1][kol] = "Закрыта"
            nom_tek_mk = spis_mk[row + 1][F.nom_kol_po_imen(self.tabl_mk, 'Номер')]
            F.zap_f(F.tcfg('bd_mk'), spis_mk, '|')
            kol_prog = F.nom_kol_po_imen(self.tabl_mk, 'Прогресс')
            if spis_mk[row + 1][kol_prog] == "Завершено":
                ass = spis_mk[row + 1][F.nom_kol_po_imen(self.tabl_mk, 'Основание')]
                if ass != "":
                    spis_brak = F.otkr_f(F.tcfg('BDact'), False, '|', False, False)
                    spis_nar = F.otkr_f(F.tcfg('Naryad'), False, '|')
                    arr_tmp_ass = ass.split(';')
                    for nom_acta in arr_tmp_ass:

                        nom_naryada = F.naiti_v_spis_1_1(spis_brak,0,'Номер акта:'+nom_acta,3).replace('Номер наряда:','')
                        nom_mk = F.naiti_v_spis_1_1(spis_nar,0,nom_naryada,1)
                        stroka_acta = F.naiti_v_spis_1_1(spis_brak,0,'Номер акта:'+nom_acta)
                        if '(Изгот.вновь по МК №' not in spis_brak[stroka_acta][7]:
                            spis_brak[stroka_acta][7] = spis_brak[stroka_acta][7] + '(Изгот.вновь по МК №' + nom_tek_mk + ')'

                        #spis_det_mk = F.otkr_f(F.scfg('mk_data') + os.sep + nom_mk + '.txt', False, '|')
                    F.zap_f(F.tcfg('BDact'), spis_brak, '|')
                    for nom_acta in arr_tmp_ass:

                        nom_naryada = F.naiti_v_spis_1_1(spis_brak,0,'Номер акта:'+nom_acta,3).replace('Номер наряда:','')
                        self.zapis_v_mk(spis_nar,nom_naryada,spis_brak)
            self.tab_click()

    def zapis_v_mk(self,sp_nar,nom_naryada,spis_brak):
        nom_mk = F.naiti_v_spis_1_1(sp_nar, 0, nom_naryada, 1)
        n_op = F.naiti_v_spis_1_1(sp_nar, 0, nom_naryada, F.nom_kol_po_im_v_shap(sp_nar, 'N операции'))
        id_det = F.naiti_v_spis_1_1(sp_nar, 0, nom_naryada,
                                    F.nom_kol_po_im_v_shap(sp_nar, 'ID'))
        sp_tabl_mk  = F.otkr_f(F.scfg('mk_data') + os.sep + nom_mk + '.txt',False,'|')
        if sp_tabl_mk  == []:
            self.showDialog('Не корректное содержимое МК')
            return
        spis_op = self.spis_op_po_mk_id_op(sp_tabl_mk,id_det,n_op)
        if spis_op == None:
            self.showDialog('Не корректное содержимое списка операций')
            return
        self.otmetka_v_mk(nom_mk,spis_op, id_det,sp_tabl_mk,sp_nar,spis_brak)

    def spis_op_po_mk_id_op(self,sp_tabl_mk,id,op):
        for j in range(1, len(sp_tabl_mk)):
            if sp_tabl_mk[j][6] == id:
                for i in range(11, len(sp_tabl_mk[0]), 4):
                    if sp_tabl_mk[j][i].strip() != '':
                        obr = sp_tabl_mk[j][i].strip().split('$')
                        obr2 = obr[-1].split(";")
                        if op in obr2:
                            return obr2
                return None

    def otmetka_v_mk(self,nom,spis_op,id,sp_tabl_mk,spis_nar,spis_brak):
        for j in range(1,len(sp_tabl_mk)):
            if sp_tabl_mk[j][6]==id:
                for i in range(11, len(sp_tabl_mk[0]), 4):
                    if sp_tabl_mk[j][i].strip() != '':
                        obr = sp_tabl_mk[j][i].strip().split('$')
                        obr2 = obr[-1].split(";")
                        if spis_op == obr2:
                            text_act = self.spis_act_po_mk_id_op(nom,id,spis_op,spis_nar,spis_brak)
                            sp_tabl_mk[j][i + 3] = '$'.join(text_act)
                            F.zap_f(F.scfg('mk_data') + os.sep + nom + '.txt',sp_tabl_mk,'|')
                            return

    def spis_act_po_mk_id_op(self, mk, id, spis_op,nar,sp_act):
        sp = []

        for i in range(1, len(nar)):
            if nar[i][1].strip() == str(mk) and nar[i][25].strip() == str(id) and nar[i][24].strip() in spis_op:
                for j in range(len(sp_act)):
                    if sp_act[j][3] == 'Номер наряда:' + nar[i][0].strip():
                        sost = ''
                        if '(Исправлен по наряду №' in sp_act[j][7]:
                            sost = 'Исправлен'
                        elif sp_act[j][6] == 'Категория брака:Неисправимый':
                            if '(Изгот.вновь по МК №' in sp_act[j][7]:
                                sost = 'Изгот.вновь'
                            else:
                                sost = 'Неисп-мый'
                        nom_acta = sp_act[j][0].replace('Номер акта:', '')
                        sp.append(nom_acta + ' ' + sost)
        return sp

    def del_mk(self):
        if self.tabl_mk.currentRow() == -1:
            return
        nom_mk = self.tabl_mk.item(self.tabl_mk.currentRow(),0).text()
        progress = self.tabl_mk.item(self.tabl_mk.currentRow(),F.nom_kol_po_imen(self.tabl_mk,'Прогресс')).text()
        if progress != '':
            F.msgbox('Нельзя удалить начатую МК')
            return
        otv = showDialog_Y_C(self,'Точно удалить полность маршрутную карту №'
                             + nom_mk + '?')
        if otv == 1024:
            row = self.tabl_mk.currentRow()

            spis_mk = F.otkr_f(F.tcfg('bd_mk'), False, '|', False, False)
            spis_mk.pop(row + 1)
            F.zap_f(F.tcfg('bd_mk'), spis_mk, '|')
            self.tab_click()
            if F.nalich_file(F.scfg('mk_data') + os.sep + nom_mk + '.txt') == True:
                F.udal_file(F.scfg('mk_data') + os.sep + nom_mk + '.txt')
            F.msgbox(nom_mk + ' удалена успешно')

    def open_mk(self):
        if self.tabl_mk.currentRow() == -1:
            return
        row = self.tabl_mk.currentRow()
        kol = F.nom_kol_po_imen(self.tabl_mk,'Статус')
        spis_mk = F.otkr_f(F.tcfg('bd_mk'), False, '|', False, False)
        if spis_mk[row+1][kol] == "Закрыта":
            if spis_mk[row+1][F.nom_kol_po_imen(self.tabl_mk,'Прогресс')] == "Завершено":
                F.msgbox('Нельзя открыть завершенную закрытую МК')
                return
            spis_mk[row+1][kol] = "Открыта"
            F.zap_f(F.tcfg('bd_mk'),spis_mk,'|')
            self.tab_click()


    def corr_mk(self,row,kol):
        if self.tabl_mk.hasFocus() == True:
            if self.tabl_mk.currentRow() == -1:
                return
            spis_mk = F.otkr_f(F.tcfg('bd_mk'), False, '|', False, False)
            spis_mk[row+1][kol] = self.tabl_mk.item(row,kol).text()
            F.zap_f(F.tcfg('bd_mk'),spis_mk,'|')

    def naiti_parametr_v_stroke(self,stroka,parametr):
        arr = stroka.split('  ')
        for i in arr:
            if i != '':
                if i.strip().startswith(parametr) == True:
                    return i.replace(parametr,'').strip()


    def ass_brak_to_mk(self):
        if self.tabl_brak.currentIndex() == -1:
            return
        tabl_razr_mk = self.ui.table_razr_MK
        label = self.ui.label_ass
        label_brak = self.ui.label_opis_braka
        nom_oper = self.naiti_parametr_v_stroke(label_brak.text(),'№ операции:')
        nom_id = self.naiti_parametr_v_stroke(label_brak.text(),'ID:')
        nom_kol_nach_tabl = F.nom_kol_po_imen(tabl_razr_mk,'Сумм.кол-во')
        if nom_kol_nach_tabl == None:
            F.msgbox('Не создана МК')
            return
        kol_id = F.nom_kol_po_imen(tabl_razr_mk,'ID')
        flag_naid = False
        for i in range(tabl_razr_mk.rowCount()):
            if flag_naid == True:
                break
            if tabl_razr_mk.item(i,kol_id).text() == nom_id:
                for j in range(tabl_razr_mk.columnCount()-1,nom_kol_nach_tabl,-1):
                    if tabl_razr_mk.item(i,j).text() != '':
                        if nom_oper not in tabl_razr_mk.item(i,j).text():
                            tabl_razr_mk.item(i, j).setText('')
                            flag_pust=True
                            for k in range(tabl_razr_mk.rowCount()):
                                if tabl_razr_mk.item(k,j).text() != '':
                                    flag_pust = False
                                    break
                            if flag_pust == True:
                                for k in range(5):
                                    tabl_razr_mk.removeColumn(j)
                        else:
                            flag_naid = True
                            break
        if flag_naid == False:
            F.msgbox('Не найдена деталь для ассоциации ' + nom_id)
            return


        nom = self.tabl_brak.item(self.tabl_brak.currentRow(),0).text().replace('Номер акта:','')

        spis_ass = self.ui.label_ass.text().split(';')
        if spis_ass[0]=='':
            spis_ass.pop(0)
        spis_ass.append(nom)
        self.ui.label_ass.setText(';'.join(spis_ass))

    def tab_click(self):
        tab = self.ui.tabWidget
        if tab.currentIndex() == 2: # номенклатура
            tabl_nomenk = self.ui.table_nomenkl
            spis = F.otkr_f(F.tcfg('BD_dse'), False, '|', False, False)
            F.zapoln_wtabl(self, spis, tabl_nomenk, 0, 0, (), (), 200, True, '')
        if tab.currentIndex() == 3: # брак
            spis_brak = F.otkr_f(F.tcfg('BDact'), False, '|', False, False)
            spis_itog = []
            for i in range(len(spis_brak)):
                if spis_brak[i][6] == "Категория брака:Неисправимый":
                    if '(Изгот.вновь по МК №' not in spis_brak[i][7]:
                        if F.is_numeric(spis_brak[i][3].replace('Номер наряда:','')) == True:
                            spis_itog.append(spis_brak[i])
            F.zapoln_wtabl(self, spis_itog, self.tabl_brak, 0, 0, (), (), 200, False, '')
        if tab.currentIndex() == 4: # мк
            tabl_mk = self.ui.table_spis_MK
            if tabl_mk.currentIndex() != -1:
                tmp_poz = tabl_mk.currentIndex()
            spis = F.otkr_f(F.tcfg('bd_mk'), False, '|', False, False)
            spis_korr = {6}
            F.zapoln_wtabl(self, spis, tabl_mk,0, spis_korr, (), (), 200, True, '')

            F.cvet_cell_wtabl(tabl_mk,'Прогресс','','Завершено')
            F.cvet_cell_wtabl(tabl_mk, 'Статус', '', 'Закрыта')
            F.cvet_cell_wtabl(tabl_mk, 'Статус', '', 'Открыта',203,176,0,False)

            tabl_mk.setCurrentIndex(tmp_poz)



    def clear_mk2(self):
        tabl_cr_stukt = self.ui.table_razr_MK
        but_add_gl_uzel = self.ui.pushButton_create_koren
        but_add_vhod = self.ui.pushButton_create_vxodyash
        but_udal_uzel = self.ui.pushButton_create_udalituzel
        tabl_cr_stukt.clearContents()
        tabl_cr_stukt.setRowCount(0)
        tabl_cr_stukt.setColumnCount(21)
        shapka = ['Наименование', 'Обозначение', 'Количество', 'Материал', 'Материал2', 'Материал3',
                  'ID', 'Количество на изделие', 'Вес', 'ПКИ', '', '', '', '', '', '', '', ''
            , '', 'Кол. по заявке', 'Уровень']
        tabl_cr_stukt.setHorizontalHeaderLabels(shapka)
        tabl_cr_stukt.resizeColumnsToContents()
        for i in range(11, 19):
            tabl_cr_stukt.setColumnHidden(i, True)
        F.ust_cvet_videl_tab(tabl_cr_stukt)
        tabl_cr_stukt.setSelectionMode(1)

        but_add_gl_uzel.setEnabled(True)
        but_add_vhod.setEnabled(True)
        but_udal_uzel.setEnabled(True)
        self.ui.pushButton_ass_brak_to_mk.setEnabled(True)

    def vibor_PY(self,param):
        combo_PY = self.ui.comboBox_PY
        combo_proekt = self.ui.comboBox_np
        combo_proekt.setCurrentIndex(param)
        return

    def vibor_pr(self,param):
        combo_PY = self.ui.comboBox_PY
        combo_proekt = self.ui.comboBox_np
        combo_PY.setCurrentIndex(param)
        return

    def cr_mk2(self):
        butt_add_gl_uzel = self.ui.pushButton_create_koren
        tabl_cr_stukt = self.ui.table_razr_MK


        nk = F.nom_kol_po_imen(tabl_cr_stukt,'Уровень')
        nk_kol_p_z = F.nom_kol_po_imen(tabl_cr_stukt, 'Кол. по заявке')
        nk_kol = F.nom_kol_po_imen(tabl_cr_stukt, 'Количество')

        if nk == None:
            return
        if nk_kol_p_z == None:
            return

        min =1000
        for i in range(tabl_cr_stukt.rowCount()):
            if int(tabl_cr_stukt.item(i,nk).text()) < min:
                min = int(tabl_cr_stukt.item(i,nk).text())
        if min > 0:
            for i in range(tabl_cr_stukt.rowCount()):
                tabl_cr_stukt.item(i, nk).setText(str(int(tabl_cr_stukt.item(i, nk).text())-min))
        for i in range(tabl_cr_stukt.rowCount()):
            if int(tabl_cr_stukt.item(i,nk).text()) == 0:
                if tabl_cr_stukt.item(i,nk_kol_p_z).text() == "":
                    F.msgbox('не указано кол-во комплектов на ' + tabl_cr_stukt.item(i,1).text())
                    tabl_cr_stukt.setCurrentCell(i,nk_kol_p_z)
                    return False
                else:
                    kol = int(tabl_cr_stukt.item(i,nk_kol_p_z).text())
            tabl_cr_stukt.item(i,nk_kol).setText(str(int(tabl_cr_stukt.item(i,nk_kol).text())* int(kol)))
        return


    def add_v_mk(self):
        tab = self.ui.tabWidget
        tab2 = self.ui.tabWidget_2
        tabl_cr_stukt = self.ui.table_razr_MK

        tree = self.ui.treeWidget
        if tree.currentIndex().row() == -1:
            F.msgbox('Не выбран узел')
            return
        item = tree.currentItem()
        if item == None:
            return
        nk =  F.nom_kol_po_imen(tree,'ID')
        current_ID = item.text(nk)
        sp_tree = F.spisok_dreva(tree)
        flag_naid = -1
        for i in range(len(sp_tree)):
            if sp_tree[i][nk] == current_ID:
                flag_naid = i
                break
        if flag_naid == -1:
            F.msgbox("Не найден выбранный узел")
            return

        q_strok = tabl_cr_stukt.currentRow()
        q_column = tabl_cr_stukt.currentColumn()
        spisok = F.spisok_iz_wtabl(tabl_cr_stukt, "", True)

        ur = int(sp_tree[flag_naid][20])
        k = 0
        spisok.insert(q_strok + 2+k, sp_tree[flag_naid])

        for i in range(flag_naid+1,len(sp_tree)):
            if int(sp_tree[i][20]) > ur:
                k=k+1
                spisok.insert(q_strok + 2 + k, sp_tree[i])
            else:
                break


        F.zapoln_wtabl(self, spisok, tabl_cr_stukt, 0, self.edit_cr_mk, (), (), 200, True, '', 30)
        tabl_cr_stukt.clearSelection()
        tabl_cr_stukt.setCurrentCell(q_strok, q_column)
        tab.setCurrentIndex(1)
        tab2.setCurrentIndex(1)

    def ass_dse_to_mk(self):
        tabl_cr_stukt = self.ui.table_razr_MK
        tabl_nomenk = self.ui.table_nomenkl
        if tabl_cr_stukt.currentRow() == -1:
            F.msgbox('Не выбрана позиция в МК')
            return
        if tabl_nomenk.currentRow() == -1:
            F.msgbox('Не выбрана ДСЕ')
            return

        naim = F.znach_vib_strok_po_kol(tabl_nomenk,'Наименование')
        nn = F.znach_vib_strok_po_kol(tabl_nomenk,'Номенклатурный номер')
        F.zapis_znach_vib_strok_po_kol(tabl_cr_stukt,'Наименование',naim)
        F.zapis_znach_vib_strok_po_kol(tabl_cr_stukt, 'Обозначение', nn)

    def del_uzel(self):
        butt_add_gl_uzel = self.ui.pushButton_create_koren
        tabl_cr_stukt = self.ui.table_razr_MK
        if tabl_cr_stukt.currentRow() == -1:
            F.msgbox('Не выбрана позиция в МК')
            return
        q_strok = tabl_cr_stukt.currentRow()
        q_column = tabl_cr_stukt.currentColumn()

        spisok = F.spisok_iz_wtabl(tabl_cr_stukt, "", True)
        spisok_tmp = spisok.copy()
        k=0
        spisok.pop(q_strok + 1)
        k+=1
        ur = int(tabl_cr_stukt.item(q_strok,20).text())
        for i in range(q_strok+2,len(spisok_tmp)):
            if int(spisok_tmp[i][20]) > ur:
                spisok.pop(i-k)
                k += 1
            else:
                break


        F.zapoln_wtabl(self, spisok, tabl_cr_stukt, 0, self.edit_cr_mk, (), (), 200, True, '', 30)
        tabl_cr_stukt.setCurrentCell(q_strok, q_column)

    def add_vhod(self):
        butt_add_gl_uzel = self.ui.pushButton_create_koren
        tabl_cr_stukt = self.ui.table_razr_MK
        if tabl_cr_stukt.currentRow() == -1:
            F.msgbox('Не выбрана позиция в МК')
            return
        q_strok = tabl_cr_stukt.currentRow()
        q_column = tabl_cr_stukt.currentColumn()

        spisok = F.spisok_iz_wtabl(tabl_cr_stukt, "", True)
        strok = []
        for i in range(20):
            strok.append('')
        strok.append(str(int(tabl_cr_stukt.item(q_strok,20).text())+1))
        strok[6] = F.time_metka()

        spisok.insert(q_strok+2,strok)

        F.zapoln_wtabl(self, spisok, tabl_cr_stukt, 0, self.edit_cr_mk, (), (), 200, True, '', 30)
        tabl_cr_stukt.clearSelection()

        tabl_cr_stukt.setCurrentCell(q_strok,q_column)


    def add_gl_uzel(self):
        butt_add_gl_uzel = self.ui.pushButton_create_koren
        tabl_cr_stukt = self.ui.table_razr_MK
        spisok = F.spisok_iz_wtabl(tabl_cr_stukt,"",True)
        strok = []
        for i in range(20):
            strok.append('')
        strok.append('0')
        strok[6] = F.time_metka()

        spisok.append(strok)

        F.zapoln_wtabl(self,spisok,tabl_cr_stukt,0,self.edit_cr_mk,(),(),200,True,'',30)



    def save_mk(self):
        nom_pu = self.ui.comboBox_PY
        nom_pr = self.ui.comboBox_np
        prim = self.ui.lineEdit_prim
        tab2 = self.ui.tabWidget_2
        if nom_pu.currentText() == "":
            showDialog(self,"Не введен номер ПУ")
            return
        if 'ПУ00-' not in nom_pu.currentText():
            if len(nom_pu.currentText()) > 6:
                showDialog(self, "Не верно введен номер ПУ")
                return
            if F.is_numeric(nom_pu.currentText()) == False:
                showDialog(self, "Не верно введен номер ПУ")
                return
            nom_pu.setText('ПУ00-' + '0'*(6-len(nom_pu.currentText())) + nom_pu.currentText())

        sp_proektov = F.otkr_f(F.tcfg('BD_Proect'),separ='|')
        flag = 0
        for i in range(len(sp_proektov)-1):
            if len(sp_proektov[i])>1 and sp_proektov[i][1] == nom_pu.currentText() and sp_proektov[i][0] == nom_pr.currentText():
                flag = 1
                break
        if flag == 0:
            showDialog(self, 'Не верно выбран номер проекта')
            return
        tablrazr_MK = self.ui.table_razr_MK
        tabl = self.ui.table_zayavk
        if tab2.currentIndex() == 1:

            if tablrazr_MK.rowCount() == 0:
                return
        if tab2.currentIndex() == 0:

            if tabl.rowCount() == 0:
                return

        if F.nalich_file(F.tcfg('bd_mk')) == False:
            showDialog(self, 'Не найдена база МК')
            return
        bd = F.otkr_f(F.tcfg('bd_mk'),separ="|")
        nom =  int(bd[len(bd)-1][0]) + 1
        nom = '0'*(6-len(str(nom)))+str(nom)
        sp_projects = ''
        spisok = F.spisok_iz_wtabl(tabl, '', True)

        if self.ui.tabWidget_2.currentIndex() == 0:
            for i in range(1, len(spisok)):
                if spisok[i][1].startswith(' ') == False:
                    sp_projects = sp_projects + spisok[i][1] + ';'
            sp_projects = sp_projects[0:-1]

        osnovanie = self.ui.label_ass.text()
        bd.append([nom,F.date(2),'Закрыта',sp_projects,nom_pu.currentText().strip(),nom_pr.currentText().strip(),prim.text().replace('\n',' '),osnovanie,''])
        F.zap_f(F.tcfg('bd_mk'),bd,"|")

        if self.ui.tabWidget_2.currentIndex() == 0:
            spisok = F.spisok_iz_wtabl(tabl, '', True)
        else:
            spisok = F.spisok_iz_wtabl(tablrazr_MK , '', True)
        for i in range(0,len(spisok)):
            for j in range(9,len(spisok[0])):
                if '\n' in spisok[i][j]:
                    spisok[i][j] = spisok[i][j].replace('\n','$')
        for i in range(0, len(spisok)):
            spisok[i] = "|".join(spisok[i])
        F.zap_f(F.scfg('mk_data') + os.sep + str(nom) + '.txt',spisok,"")
        showDialog(self,'маршрутная карта ' + str(nom) + ' успешно сохранена' )
        if self.ui.tabWidget_2.currentIndex() == 1:
            self.clear_mk2()



    def kol_po_zayav(self,sp_xml_tmp,kol):
        for i in sp_xml_tmp:
            i[2] = int(i[2]) * int(kol)

        return sp_xml_tmp

    def create_mk(self):
        tab2 = self.ui.tabWidget_2
        tabl_cr_stukt = self.ui.table_razr_MK
        if tab2.currentIndex() == 1:

            if tabl_cr_stukt.rowCount() == 0:
                return
            but_add_gl_uzel = self.ui.pushButton_create_koren
            but_add_vhod = self.ui.pushButton_create_vxodyash
            but_udal_uzel = self.ui.pushButton_create_udalituzel


            rez = self.cr_mk2()
            if rez == False:
                return
            s_vert = F.spisok_iz_wtabl(tabl_cr_stukt,"",True)
            nach_sod = len(s_vert[0])
            but_add_gl_uzel.setEnabled(False)
            but_add_vhod.setEnabled(False)
            but_udal_uzel.setEnabled(False)

        if tab2.currentIndex() == 0:
            tabl = self.ui.table_zayavk

            if tabl.rowCount() == 0:
                return
            if tabl.columnCount() > 5:
                return
            s_vert = []

            sp_izd = F.spisok_iz_wtabl(tabl)
            for i in sp_izd:
                putt = i[0]
                sp_xml_tmp = XML.spisok_iz_xml(putt)
                for itm in range(len(sp_xml_tmp)):
                    for j in range(itm+1,len(sp_xml_tmp)):
                        if itm < len(sp_xml_tmp)-1:
                            if sp_xml_tmp[itm][0] == sp_xml_tmp[j][0] and sp_xml_tmp[itm][1] == sp_xml_tmp[j][1] \
                                    and sp_xml_tmp[itm][20] == sp_xml_tmp[j][20]:
                                sp_xml_tmp[itm][2] = str(int(sp_xml_tmp[itm][2]) + int(sp_xml_tmp[j][2]))
                                sp_xml_tmp[j][0] = "deletes" + str(F.time_metka())
                            else:
                                break
                sp_xml_tmp_ = []
                for itm in range(len(sp_xml_tmp)):
                    if sp_xml_tmp[itm][0].startswith('deletes') == False:
                        sp_xml_tmp_.append(sp_xml_tmp[itm])
                sp_xml_tmp = sp_xml_tmp_


                if i[2] == '':
                    showDialog(self,"Не указано кол-во по заявке")
                    return
                sp_xml_tmp = self.kol_po_zayav(sp_xml_tmp,i[2])
                for j in sp_xml_tmp:
                    if len(s_vert) == 0:
                        s_vert.append(['' for x in list(range(0, len(j)))])
                        nach_sod = len(j)
                        s_vert[0][0] = "Наименование"
                        s_vert[0][1] = "Обозначение"
                        s_vert[0][2] = "Кол-во"
                        s_vert[0][3] = "Материал"
                        s_vert[0][4] = "Материал2"
                        s_vert[0][5] = "Материал3"
                        s_vert[0][6] = "ID"
                        s_vert[0][7] = "Кол-во на изд."
                        s_vert[0][8] = 'Масса'
                        s_vert[0][9] = 'Покуп. изд.'
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
            tk = F.otkr_f(F.scfg('add_docs') + os.sep + j[2] + "_" + nn + '.txt', False, "|",True,True)
            tk = self.grup_tk_po_rabc(tk,kol_det_vseg)
            self.ogran = nach_sod-1
            for k in tk:
                s_vert = self.dob_etap(s_vert,k[0],k[1],k[2],i,self.ogran)
        s_vert = self.oformlenie_sp_pod_mk(s_vert)
        if tab2.currentIndex() == 1:
            for i in range(tabl_cr_stukt.columnCount()):
                tabl_cr_stukt.setColumnHidden(i,False)
            F.zapoln_wtabl(self, s_vert, tabl_cr_stukt, 0, 0, "", "", 200, True, '', 90)
            tabl_cr_stukt.setSelectionBehavior(1)
            self.oformlenie_formi_mk(tabl_cr_stukt, s_vert)
            self.ui.pushButton_ass_brak_to_mk.setEnabled(True)
        if tab2.currentIndex() == 0:
            for i in range(tabl.columnCount()):
                tabl.setColumnHidden(i,False)
            F.zapoln_wtabl(self,s_vert,tabl,0,0,"","",200,True,'',90)
            tabl.setSelectionBehavior(1)
            self.oformlenie_formi_mk(tabl,s_vert)

    def oformlenie_formi_mk(self, tabl,s):
        for i in range(11,len(s[0])-1,4):
            for j in range(0, len(s)-1):
                #if tabl.item(j,i) == None:
                #    cellinfo = QtWidgets.QTableWidgetItem('')
                #    tabl.setItem(j,i, cellinfo)
                F.ust_color_wtab(tabl, j, i, 227, 227, 227)


                #tabl.item(j,i).setBackground(QtGui.QColor(227,227,227))
        for i in range(0,11):
            for j in range(0, len(s)-1):
                #if tabl.item(j,i) == None:
                #    cellinfo = QtWidgets.QTableWidgetItem('')
                #    tabl.setItem(j,i, cellinfo)
                #tabl.item(j,i).setBackground(QtGui.QColor(227,227,227))
                F.ust_color_wtab(tabl, j, i, 227, 227, 227)

    def summ_kol(self,s,i):
        naim = s[i][0].strip()
        nn = s[i][1].strip()
        summ = 0
        for j in range(1,len(s)):
            if s[j][0].strip() == naim and s[j][1].strip() == nn:
                summ+= int(s[j][2])
        return summ

    def udal_kol(self,s,nom):
        for i in range(0, len(s)):
            s[i].pop(nom)
        return s

    def oformlenie_sp_pod_mk(self,s):
        for i in range(1,len(s)):
            s[i][0] = '    ' * int(s[i][20]) + s[i][0].strip()
            s[i][1] = '    ' * int(s[i][20]) + s[i][1].strip()
            s[i][9] = s[i][9].replace('0','')
            s[i][9] = s[i][9].replace('1', '+')

        for j in range(11, 21):
            s = self.udal_kol(s,11)

        s[0][10] = 'Сумм.кол-во'
        for i in range(1, len(s)):
            s[i][10] = self.summ_kol(s,i)
        for j in s:
            for i in range(11, len(s[0])):
                if '$' in str(j[i]):
                    vrem, oper = [x for x in j[i].split("$")]
                    j[i] = 'Время: ' + vrem + ' мин.' + '\n' + 'Операции:' + '\n' + oper
        i = 12
        sp_ins = ['комплектация','изготовление','контроль']
        while i:
            if i > len(s[0]):
                break
            for j in sp_ins:
                s = self.dob_kol(s,i,j)
                i+=1
            i+=1
        return s

    def summa_rc(self,rc):
        s = ''
        if F.is_numeric(rc):
            return int(rc)
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
    def poporyadku(self,i,spis,stroka,oper):
        for j in range(i, len(spis[0])):
            if spis[stroka][j] != "":
                arr_tmp_nom = spis[stroka][j].split(';')
                arr_tmp_nom2 = arr_tmp_nom[-1].split('$')
                tmp_nom = arr_tmp_nom2[-1]
                if int(tmp_nom) < int(oper):
                    return False

        return True

    def dob_etap(self,spis,rc,vrem,oper,stroka,ogran):

        flag = 0
        for i in range(ogran+1,len(spis[0])):
            if flag == 1:
                break
            if spis[0][i] == rc and self.poporyadku(i,spis,stroka,oper) == True:
                flag = 1
                spis[stroka][i] = str(vrem) + "$" + oper
                self.ogran = i-1
                break
            if self.summa_rc(rc) < self.summa_rc(spis[0][i]):
                j = i-1
                while j>=self.ogran:
                    if self.poporyadku(j+1,spis,stroka,oper) == False:
                        if spis[0][j + 2] != rc:
                            spis = self.dob_kol(spis, j + 2, rc)
                        spis[stroka][j + 2] = str(vrem) + "$" + oper
                        self.ogran = j + 2
                        flag = 1
                        break
                    if j <= ogran:
                        if spis[0][j + 1] != rc:
                            spis = self.dob_kol(spis, j + 1, rc)
                        spis[stroka][j + 1] = str(vrem) + "$" + oper
                        self.ogran = j +1
                        flag = 1
                        break
                    if self.summa_rc(rc) >= self.summa_rc(spis[0][j]):
                        if spis[0][j + 1] != rc:
                            spis = self.dob_kol(spis, j+1, rc)
                        spis[stroka][j+1] = str(vrem) + "$" + oper
                        self.ogran = j+1
                        flag = 1
                        break
                    if j == self.ogran:
                        if spis[0][self.ogran] != rc:
                            spis = self.dob_kol(spis, self.ogran , rc)
                        spis[stroka][self.ogran] = str(vrem) + "$" + oper
                        flag = 1
                        break
                    j-=1
            else:
                j = i + 1
                while j <= len(spis[0])-1:
                    if self.summa_rc(rc) == self.summa_rc(spis[0][j]) and self.poporyadku(j+1,spis,stroka,oper) == True:
                        spis[stroka][j] = str(vrem) + "$" + oper
                        self.ogran = j
                        flag = 1
                        break
                    if self.summa_rc(rc) < self.summa_rc(spis[0][j]) and self.poporyadku(j+1,spis,stroka,oper) == True:
                        spis = self.dob_kol(spis, j - 1, rc)
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
                    vrem = F.valm(itk[6]) + int(kol_det_vseg) * F.valm(itk[7])
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
        tabl = self.ui.table_zayavk
        tree = self.ui.treeWidget
        tab = self.ui.tabWidget
        if tab.currentIndex() > 2:
            self.ui.tabWidget.setCurrentIndex(0)
        putt = F.f_dialog_name(self, 'Выбрать XML', '', "Файлы *.xml")
        if putt == '':
            return
        s = XML.spisok_iz_xml(putt)
        if vklad.currentIndex() == 0:
            self.zapoln_tree_spiskom(s)
            for _ in range(0, 8):
                tree.resizeColumnToContents(_)
        if vklad.currentIndex() == 1:
            tabl.setSelectionBehavior(0)
            self.dob_izd(s, putt)


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
        if tabl.columnCount() > 5:
            tabl.clear()
            tabl.clearContents()
            shapka = ['Файл', 'Изделие', 'Кол-во']
            tabl.setColumnCount(3)
            tabl.setRowCount(0)
            tabl.setHorizontalHeaderLabels(shapka)
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