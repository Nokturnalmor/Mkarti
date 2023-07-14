# -*- coding: utf-8 -*-
import pprint

import project_cust_38.Cust_Functions as F
import project_cust_38.xml_v_drevo as XML
from PyQt5 import QtWidgets, QtCore, QtGui,QtDesigner
from PyQt5.QtWinExtras import QtWin
import os
import project_cust_38.Cust_Qt as CQT
CQT.conver_ui_v_py()
from mk_gui import Ui_MainWindow  # импорт нашего сгенерированного файла

import sys
import project_cust_38.Cust_SQLite as CSQ
import project_cust_38.Cust_mes as CMS
import project_cust_38.Cust_Excel as CEX
import project_cust_38.Zamechaniya as ZMCH
import obespechenie as OBSP
import industrial_capacity as IND
import export_docs_mkarts as EXPD
import Selector_conversation as SLCT
import calculate_vo as CVO
import kal_plan as KPL
import gui_kal_plan as GKPL
import gui_vol_plan as GVKPL
import ERP_conn as ERP
import project_cust_38.Cust_b24 as CB24
from dataclasses import dataclass

@dataclass
class Data_plan:
    db_kplan = F.bdcfg('DB_kplan')
    # ======= KAL PLAN======================
    NAPR_DEYAT = CSQ.zapros(db_kplan, f"""SELECT * FROM napravl_deyat""", rez_dict=True)
    DICT_NAPR_DEYAT = F.raskrit_dict(NAPR_DEYAT, 'Пномер')
    DICT_NAPR_DEYAT_NAME = F.raskrit_dict(NAPR_DEYAT, 'Имя')
    VID_PO_NAPR = CSQ.zapros(db_kplan, f"""SELECT * FROM Виды_по_напр""", rez_dict=True)
    DICT_VID_PO_NAPR = F.raskrit_dict(VID_PO_NAPR, 'Пномер')
    DICT_VID_PO_NAPR_NAME = F.raskrit_dict(VID_PO_NAPR, 'Имя')
    STATUS_POZ =        CSQ.zapros(db_kplan, f"""SELECT * FROM status_poz""", rez_dict=True)
    DICT_STATUS_POZ = F.raskrit_dict(STATUS_POZ, 'Пномер')
    DICT_STATUS_POZ_NAME = F.raskrit_dict(STATUS_POZ, 'Имя')
    STATUS_ETAPI_ERP =        CSQ.zapros(db_kplan, f"""SELECT * FROM status_etapi_erp""", rez_dict=True)
    DICT_STATUS_ETAPI_ERP = F.raskrit_dict(STATUS_ETAPI_ERP, 'Пномер')
    DICT_STATUS_ETAPI_ERP_NAME = F.raskrit_dict(STATUS_ETAPI_ERP, 'Имя')

    DICT_NAPRAVLENIE = F.raskrit_dict(
        CSQ.zapros(db_kplan, f"""SELECT * FROM napravlenie""", rez_dict=True), 'Пномер')

    DICT_CLD = CMS.DICT_CLD_KPLAN(db_kplan)
    DICT_PODR = F.raskrit_dict(CSQ.zapros(db_kplan, """SELECT * FROM podrazdel""", rez_dict=True), 'Имя')

    STATUS_NORM = CSQ.zapros(db_kplan, f"""SELECT * FROM status_norm""", rez_dict=True)
    DICT_STATUS_NORM = F.raskrit_dict(STATUS_NORM, 'Код')
    DICT_STATUS_NORM_NAME = F.raskrit_dict(STATUS_NORM, 'Имя')

class mywindow(QtWidgets.QMainWindow):
    resized = QtCore.pyqtSignal()
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.versia = '1.7.5'
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        self.setWindowTitle("Создание маршрутных карт")
        #self.showMaximized()
        #F.test_path()
        # self.resized.connect(self.widths)
        self.ip_srv = ''
        self.mk_file_founding = ''
        CMS.load_ip_srv(self)
        self.bd_naryad_TEST = F.scfg('Naryad') + F.sep() + 'old' + F.sep() + 'Naryad.db'
        self.bd_naryad = F.bdcfg('Naryad')
        self.bd_act = F.bdcfg('BDact')
        self.db_users = F.bdcfg('BD_users')
        self.bd_files = F.bdcfg('files')
        self.db_kplan = F.bdcfg('DB_kplan')
        self.db_mater = self.bd_nomen = F.bdcfg('nomenklatura_erp')
        self.db_selector = F.bdcfg('BD_selector')
        self.db_resxml = F.bdcfg('db_resxml')
        self.db_dse = F.bdcfg('BD_dse')
        self.res = ''
        self.BD = F.bdcfg('BD_projects')
        self.glob_nom_mk_obesp = ''
        self.regim = ''
        self.list_vars_vo = []
        self.SPIS_OP = CSQ.zapros(self.bd_naryad,"""SELECT * FROM operacii""")
        #self.DICT_FILTR = F.raskrit_dict(CSQ.zapros(self.db_mater, f"""SELECT * FROM complex_filtr""", rez_dict=True), 'kod')
        self.DICT_MAT = F.raskrit_dict(CSQ.zapros(self.db_mater,f"""SELECT * FROM nomen""",rez_dict=True),'Код')
        if self.SPIS_OP == False:
            CQT.msgbox(f'БД занята, пробуй позже')
            quit()
        self.DICT_OP = F.list_to_dict(self.SPIS_OP,'kod')
        conn_users, cur_users = CSQ.connect_bd(self.db_users)
        self.SPIS_RC = CSQ.zapros(self.db_users,"""SELECT * FROM rab_c""",conn=conn_users,cur=cur_users)
        self.DICT_RC = F.list_to_dict(self.SPIS_RC,'Код')
        if self.SPIS_RC == False:
            CSQ.close_bd(conn_users,cur_users)
            CQT.msgbox(f'БД занята, пробуй позже')
            quit()
        self.SPIS_OB = CSQ.zapros(self.db_users,"""SELECT Инв_номер,Наименование,Примечание FROM equipment""",conn=conn_users,cur=cur_users)
        if self.SPIS_OB == False:
            CSQ.close_bd(conn_users,cur_users)
            CQT.msgbox(f'БД занята, пробуй позже')
            quit()
        self.SPIS_PROF = CSQ.zapros(self.db_users,"""SELECT * FROM professions""",conn=conn_users,cur=cur_users)
        if self.SPIS_PROF == False:
            CSQ.close_bd(conn_users,cur_users)
            CQT.msgbox(f'БД занята, пробуй позже')
            quit()
        CSQ.close_bd(conn_users,cur_users)


        # ================CALENDAR===================================
        self.ui.cld_obespechenie.clicked.connect(lambda _, x=self: OBSP.data_obespech(x))
        self.ui.calendarWidget.clicked.connect(lambda: KPL.clck_cld(self))
        # ==============TREE=========================================

        # ==================================================================
        # ===============TAB================================================
        tab = self.ui.tabWidget
        tab.currentChanged[int].connect(self.tab_click)
        self.ui.tabWidget_2.currentChanged[int].connect(self.tab_click2)
        self.ui.tabWidget_3.currentChanged[int].connect(self.tab_mk_click)
        self.ui.tabWidget_4.currentChanged[int].connect(self.tab_zagruzka_rc)
        # ============================================================
        # ==================TABLE=====================================
        self.ui.tbl_selector_proj_view.clicked.connect(self.SLCT_click)
        self.ui.tbl_selector_proj_view.cellChanged[int, int].connect(self.SLCT_edit_primech)
        self.ui.tbl_selector_proj_view_zamech.cellDoubleClicked[int, int].connect(self.SLCT_edit_zamech_from_view)
        self.ui.tbl_selector_proj_view.itemSelectionChanged.connect(self.SLCT_selector_proj_view_itemSelection)
        self.ui.tbl_selector_proj_view.cellDoubleClicked[int, int].connect(self.SLCT_add_new_zamech)
        self.ui.tbl_selector.cellDoubleClicked[int, int].connect(self.SLCT_select_zamech)
        self.tabl_nomenk = self.ui.table_nomenkl
        self.tabl_nomenk.cellDoubleClicked[int, int].connect(self.zapusk_docs)
        self.ui.tbl_obespechenie.cellDoubleClicked[int, int].connect(self.OBSP_select_obesp_po_mk_from_table)
        self.ui.tbl_rc.cellChanged[int, int].connect(self.IND_cellChanged)
        CQT.ust_cvet_videl_tab(self.tabl_nomenk)
        self.tabl_nomenk.setSelectionBehavior(1)
        self.tabl_nomenk.setSelectionMode(1)
        self.tabl_mk = self.ui.table_spis_MK
        CQT.ust_cvet_videl_tab(self.tabl_mk)
        self.tabl_mk.setSelectionBehavior(1)
        self.tabl_mk.setSelectionMode(1)
        self.tabl_mk.cellChanged[int, int].connect(self.corr_mk)
        # self.tabl_mk.cellActivated[int, int].connect(self.corr_mk)
        self.tabl_mk.clicked.connect(self.spis_MK_clck)
        self.tabl_brak = self.ui.table_brak
        CQT.ust_cvet_videl_tab(self.tabl_brak)
        self.tabl_brak.setSelectionBehavior(1)
        self.tabl_brak.setSelectionMode(1)
        self.tabl_brak.clicked.connect(self.click_brak)
        self.tabl_brak.doubleClicked.connect(self.tabl_brak_dbl_clk)

        self.ui.table_spis_MK.setSelectionBehavior(1)
        self.ui.table_spis_MK.setSelectionMode(1)
        CQT.ust_cvet_videl_tab(self.ui.table_spis_MK)

        tabl = self.ui.table_zayavk
        shapka = ['Файл', 'Изделие', 'Количество']
        tabl.setColumnCount(3)
        tabl.setHorizontalHeaderLabels(shapka)
        self.ui.tbl_kal_pl.itemSelectionChanged.connect(lambda  x=self, y=self.ui.tbl_kal_pl: KPL.clck_tbl_kal_pl_tbl(x,y))
        self.ui.tbl_preview.itemSelectionChanged.connect(lambda x=self, y=self.ui.tbl_preview: KPL.clck_tbl_preview(x, y))
        self.ui.tbl_pl_gaf.itemSelectionChanged.connect(
            lambda x=self, y=self.ui.tbl_pl_gaf: KPL.clck_tbl_preview(x, y))
        #self.ui.tbl_kal_pl.clicked.connect(lambda  x=self: KPL.clck_tbl(x))
        self.ui.tbl_rc.itemSelectionChanged.connect(self.clck_tbl_rc)
        self.ui.tbl_rc.clicked.connect(self.clck_tbl_rc)
        self.ui.tbl_rc.doubleClicked.connect(lambda _, x=self: IND.select_schema_dbl_clk(x))
        self.ui.tbl_pl_add_poz.doubleClicked.connect(lambda: KPL.dbl_clk_tbl_add_poz(self))
        self.ui.tbl_tabeli.doubleClicked.connect(lambda: IND.set_old_val(self))
        self.ui.tbl_tabeli.clicked.connect(lambda: IND.set_tooltip_val(self))
        self.ui.tbl_preview.doubleClicked.connect(lambda: KPL.select_field_from_kgui(self))
        self.ui.tbl_pl_gaf.horizontalScrollBar().valueChanged.connect(
            self.ui.tbl_pl_gaf_filtr.horizontalScrollBar().setValue)
        self.ui.tbl_pl_gaf.horizontalScrollBar().valueChanged.connect(
            self.ui.tbl_pl_gaf_svod.horizontalScrollBar().setValue)
        self.ui.tbl_kal_pl.horizontalScrollBar().valueChanged.connect(
            self.ui.tbl_filtr_kal_pl.horizontalScrollBar().setValue)
        #self.ui.tbl_pl_gaf_svod.clicked.connect(lambda: GVKPL.set_tooltip_val(self))
        self.ui.tbl_pl_gaf_svod.setMouseTracking(True)
        self.ui.tbl_pl_gaf_svod.mouseMoveEvent = self.tbl_pl_gaf_svod_mouseMoveEvent
        self.ui.tbl_pl_gaf_svod.doubleClicked.connect(lambda: GVKPL.dbl_clk_svod_select_etap(self))
        self.ui.tbl_pl_gaf.doubleClicked.connect(lambda: GVKPL.dbl_clk_select_etap(self))
        self.ui.tbl_preview.setMouseTracking(True)
        self.ui.tbl_preview.mouseMoveEvent = self.tbl_preview_mouseMoveEvent
        self.ui.tbl_pl_gaf.setMouseTracking(True)
        self.ui.tbl_pl_gaf.mouseMoveEvent = self.tbl_pl_gaf_mouseMoveEvent
        self.ui.btn_save_pl.clicked.connect(lambda: GVKPL.save_kpl_plan(self))
        # =================================================================
        # ==============BUTTON==========================================


        self.ui.btn_add_rm.clicked.connect(lambda _, x=self: IND.add_rm(x))
        self.ui.btn_tabeli_ok.clicked.connect(lambda _, x=self: IND.load_tabel_in_db(x))
        self.ui.btn_obespechenie_spis_po_mk.clicked.connect(lambda _, x=self: OBSP.spis_obesp_po_mk(x))
        self.ui.btn_obespechenie_zapisat.clicked.connect(lambda _, x=self: OBSP.zapisat(x))
        self.ui.btn_add_zamech.clicked.connect(lambda _, x=self: ZMCH.add_zamech(x))
        self.ui.btn_edit_zamech.clicked.connect(lambda _, x=self: ZMCH.load_zamech_to_edit(x))
        self.ui.btn_add_v_planetapi.clicked.connect(self.add_v_planetapi)
        self.ui.btn_zaversh.clicked.connect(self.zaversh_mkards)

        self.ui.btn_zapoln_osv_zav_po_nar.clicked.connect(self.zapoln_osv_zav_po_nar)
        self.ui.btn_obnov_po_strukt.clicked.connect(self.obnov_po_strukt)
        self.ui.btn_obnovit_naruadi_po_mk.clicked.connect(self.obnovit_naruadi_po_mk)

        self.ui.btn_edit_res_xml.clicked.connect(self.edit_res_xml)

        btn_korr_nom = self.ui.btn_korr_nom
        btn_korr_nom.clicked.connect(self.btn_korr_nom)

        btn_del_nom = self.ui.btn_del_poz_nom
        btn_del_nom.clicked.connect(self.del_nom)

        butt_vib_nomen = self.ui.pushButton_ass_nomen_MK
        butt_vib_nomen.clicked.connect(self.ass_dse_to_mk)
        CQT.ust_cvet_videl_tab(butt_vib_nomen)

        self.ui.btn_obnov_pr.clicked.connect(self.obn_spis_pr)

        but_clear_mk = self.ui.pushButton_create_mk_clear
        but_clear_mk.clicked.connect(self.clear_mk2)

        but_add_gl_uzel = self.ui.pushButton_create_koren
        but_add_gl_uzel.clicked.connect(self.add_gl_uzel)

        but_add_vhod = self.ui.pushButton_create_vxodyash
        but_add_vhod.clicked.connect(self.add_vhod)

        but_add_paral = self.ui.pushButton_create_paralel
        but_add_paral.clicked.connect(self.add_paral)

        but_udal_uzel = self.ui.pushButton_create_udalituzel
        but_udal_uzel.clicked.connect(self.del_uzel)

        btn_save_cust_drevo = self.ui.btn_save_cust_drevo
        btn_save_cust_drevo.clicked.connect(self.save_cust_drevo)

        btn_load_cust_drevo = self.ui.btn_load_cust_drevo
        btn_load_cust_drevo.clicked.connect(self.load_cust_drevo)

        but_add_bd = self.ui.pushButton_add_v_bd
        but_add_bd.clicked.connect(self.dob_izd_k_bd)

        but_cr_mk = self.ui.pushButton_create_MK
        but_cr_mk.clicked.connect(self.create_mk)

        but_save_mk = self.ui.pushButton_save_MK
        but_save_mk.clicked.connect(self.save_mk)

        but_add_v_mk = self.ui.pushButton_add_v_MK
        but_add_v_mk.clicked.connect(self.add_v_mk)

        but_add_v_nomenk = self.ui.pushButton_add_v_bd_2
        but_add_v_nomenk.clicked.connect(self.add_v_nomenkl)

        btn_normi = self.ui.btn_vigruzka_norm
        btn_normi.clicked.connect(self.vigruzka_norm)

        btn_normi = self.ui.btn_vigruzka_norm_mat
        btn_normi.clicked.connect(self.vigruzka_norm_mat)

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
        self.ui.pushButton_clear_label.setToolTip('Удалить ассоциации с актами о браке')

        self.ui.btn_update_norm.clicked.connect(lambda _, x='vrem': self.update_norm(x))
        self.ui.btn_update_norm_prof.clicked.connect(lambda _, x='prof': self.update_norm(x))
        self.ui.btn_update_norm_rc.clicked.connect(lambda _, x='rc': self.update_norm(x))
        self.ui.btn_update_norm_mat.clicked.connect(lambda _, x='mat': self.update_norm(x))
        self.ui.btn_ochistit.clicked.connect(self.clear_res)
        self.ui.btn_selector_add.clicked.connect(lambda _, x=self: SLCT.add_zamech(x))
        self.ui.btn_selector_edit.clicked.connect(lambda _, x=self: SLCT.edit_zamech(x))
        self.ui.btn_pl_add_poz.clicked.connect(lambda _, x=self: KPL.btn_pl_add_poz_click(x))
        self.ui.btn_pl_ok_add_poz.clicked.connect(lambda _, x=self: KPL.btn_pl_ok_add_poz_click(x))
        self.ui.btn_pl_edit_poz.clicked.connect(lambda _, x=self: KPL.btn_pl_edit_poz_click(x))
        self.ui.btn_settings.clicked.connect(lambda _, x=self: KPL.btn_pl_settings(x))
        self.ui.btn_pl_mode.clicked.connect(lambda _, x=self: KPL.btn_pl_mode(x))
        self.ui.btn_kal_pl_left.clicked.connect(lambda _, x=self: KPL.kal_pl_left(x))
        self.ui.btn_kal_pl_right.clicked.connect(lambda _, x=self: KPL.kal_pl_right(x))
        self.ui.btn_fdate_res_erp.clicked.connect(self.clk_fdate_res_erp)
        self.ui.btn_edit_local_gant_left.clicked.connect(lambda: GKPL.move_left(self))
        self.ui.btn_edit_local_gant_right.clicked.connect(lambda: GKPL.move_right(self))
        self.ui.btn_show_svod.clicked.connect(lambda: GVKPL.show_svod(self))
        self.ui.btn_pl_tabel.clicked.connect(lambda: KPL.show_tabel(self))
        self.ui.btn_load_file_mk_founfing.clicked.connect(self.load_file_mk_founfing)
        self.ui.btn_show_file_founding_mk.clicked.connect(self.show_file_founding_mk)
        self.ui.btn_pl_open_dir.clicked.connect(lambda: KPL.btn_pl_open_dir(self))
        self.ui.btn_pl_add_trbl.clicked.connect(lambda: KPL.btn_pl_add_trbl(self))
        self.ui.btn_pl_load_norm.clicked.connect(lambda: KPL.btn_pl_load_norm(self))

        # =================================================================
        # ===========COMBOBOX===========================================

        combo_proekt = self.ui.comboBox_np
        combo_proekt.activated[int].connect(self.vibor_pr)
        combo_nap = self.ui.comboBox_napravlenia
        spis_napr = F.otkr_f(F.scfg('mk_data') + os.sep + 'Направления.txt')
        for i in range(len(spis_napr)):
            combo_nap.addItem(spis_napr[i])

        combo_vid = self.ui.comboBox_vid
        spis_vid = F.otkr_f(F.scfg('mk_data') + os.sep + 'Виды.txt')
        for i in range(len(spis_vid)):
            combo_vid.addItem(spis_vid[i])

        dict_tip = CSQ.zapros(self.bd_naryad,"""SELECT * FROM Тип_мк""",rez_dict=True)
        self.DICT_TIP_MK = F.raskrit_dict(dict_tip,'Имя')
        self.ui.cmb_tip_mk.addItem('')
        self.ui.cmb_tip_mk.addItems([*self.DICT_TIP_MK.keys()])


        dict_tip = CSQ.zapros(self.bd_naryad,"""SELECT * FROM тип_дорезок""",rez_dict=True)
        self.DICT_TIP_DOREZ = F.raskrit_dict(dict_tip,'Имя')


        self.ui.cmb_schems.activated[int].connect(lambda _, x=self: IND.select_schema(x))
        self.ui.cmb_tip_mk.activated[int].connect(self.cmb_tip_click)
        self.ui.cmb_etap.activated[int].connect(lambda: KPL.select_etap_edit(self))

        # =================================================================
        # ==========================LINEEDIT=============================

        self.ui.lineEdit_naim.textEdited.connect(self.poisk_nn)
        self.ui.lineEdit_nom_n.textEdited.connect(self.poisk_nn)
        self.ui.lineEdit_primech.textEdited.connect(self.poisk_nn)

        # =================================================================
        # =================SLIDER==========================================
        #self.ui.sl_mash_local.valueChanged[int].connect(self.sl_mash_change)
        # =================DATE_EDIT==========================================
        self.ui.de_vol_pl.dateChanged.connect(lambda: GVKPL.save_diapazon_month(self))
        self.ui.de_vol_pl_end.dateChanged.connect(lambda: GVKPL.save_diapazon_month(self))
        # =================================================================
        # ===================Check_box=================================
        self.ui.chk_kpl_zaversch.clicked.connect(lambda: KPL.set_params_kpl(self))

        # ========================ACTIONS=================================

        actionXML = self.ui.action_XML
        actionXML.triggered.connect(self.viborXML)

        action_mat = self.ui.action_mat
        action_mat.triggered.connect(self.export_mat_spec)

        action_mat_tk = self.ui.action_mat_tk
        action_mat_tk.triggered.connect(self.export_mat_spec_tk)

        action_json = self.ui.action_JSON
        action_json.triggered.connect(self.export_json)

        action_json_Excel = self.ui.action_JSON_Excel
        action_json_Excel.triggered.connect(self.export_json_Excel)

        self.ui.actionexcel.triggered.connect(self.export_table)
        self.ui.action_txt.triggered.connect(self.export_table_txt)
        self.ui.action_opn_dir_mk.triggered.connect(self.open_zayavk)
        self.ui.action_res_ERP.triggered.connect(lambda _, x=self: EXPD.export_res_erp(x))
        self.ui.action_plan_date.triggered.connect(lambda _, x=self: ERP.export_date_plan(x))
        self.ui.action_load_plan.triggered.connect(lambda _, x=self: KPL.import_exel_plan(x))
        self.ui.action4_px.triggered.connect(lambda: self.sl_mash_change(4))
        self.ui.action6_px.triggered.connect(lambda: self.sl_mash_change(6))
        self.ui.action8_px.triggered.connect(lambda: self.sl_mash_change(8))
        self.ui.action10_px.triggered.connect(lambda: self.sl_mash_change(10))
        self.ui.action12_px.triggered.connect(lambda: self.sl_mash_change(12))
        self.ui.action14_px.triggered.connect(lambda: self.sl_mash_change(14))
        self.ui.action16_px.triggered.connect(lambda: self.sl_mash_change(16))
        self.ui.action18_px.triggered.connect(lambda: self.sl_mash_change(18))
        # =================================================================
        # =============LOADS========================================
        #KPL.load_gui(self)
        ZMCH.init_zamech_const(self)
        SLCT.load_dicts_for_selector(self)
        #===================
        conn_naryad, cur_naryad = CSQ.connect_bd(self.bd_naryad)
        rez = CMS.dict_etapi(self, self.bd_naryad,conn_naryad)
        if rez == False:
            CSQ.close_bd(conn_naryad, cur_naryad)
            CQT.msgbox(f'база нарядов занята')
            quit()
        rez = self.DICT_KOD_OPER = F.raskrit_dict(
            CSQ.zapros(self.bd_naryad, f"""SELECT kod, name FROM operacii""", rez_dict=True, conn=conn_naryad, cur = cur_naryad), 'name')
        if rez == False:
            CSQ.close_bd(conn_naryad,cur_naryad)
            CQT.msgbox(f'база нарядов занята')
            quit()
        CSQ.close_bd(conn_naryad,cur_naryad)
        #======================== nomen
        conn_nomen, cur_nomen = CSQ.connect_bd(self.bd_nomen)
        self.DICT_NOMEN = F.raskrit_dict(CSQ.zapros(self.bd_nomen, f'''SELECT * FROM nomen''',
                                            one=False, shapka=True, rez_dict=True,conn= conn_nomen, cur = cur_nomen),'Код')
        if self.DICT_NOMEN == False:
            CSQ.close_bd(conn_nomen,cur_nomen)
            CQT.msgbox(f'база номенклатуры занята')
            quit()
        self.DICT_FILTR_NOMEN = CSQ.zapros(self.bd_nomen, f'''SELECT * FROM complex_filtr''',
                                     one=False, shapka=False, rez_dict=True, conn=conn_nomen, cur = cur_nomen)
        if self.DICT_FILTR_NOMEN == False:
            CSQ.close_bd(conn_nomen, cur_nomen)
            CQT.msgbox(f'база номенклатуры занята')
            quit()
        CSQ.close_bd(conn_nomen, cur_nomen)
        #====================== txt
        CMS.dict_projects(self,F.tcfg('BD_Proect'))
        CVO.DICT_VAR_OPER(self)
        CMS.DICT_PLACES(self,self.db_users)

        #======================= users
        conn_users, cur_users = CSQ.connect_bd(self.db_users)
        rez = CMS.dict_professions(self, self.db_users,conn_users)
        if rez == False:
            CSQ.close_bd(conn_users, cur_users)
            CQT.msgbox(f'база users занята')
            quit()
        rez = CMS.dict_rc(self, self.db_users,conn_users)

        if rez == False:
            CSQ.close_bd(conn_users, cur_users)
            CQT.msgbox(f'база users занята')
            quit()
        self.DICT_EMPLOEE = CMS.dict_emploee(self.db_users,conn_users)
        self.DICT_EMPLOEE_FULL = CMS.dict_emploee_full(self.db_users,conn_users)

        if self.DICT_EMPLOEE == False:
            CSQ.close_bd(conn_users, cur_users)
            CQT.msgbox(f'база users занята')
            quit()
        rez = CMS.dict_emploee_rc(self,conn_users)
        if rez == False:
            CSQ.close_bd(conn_users, cur_users)
            CQT.msgbox(f'база users занята')
            quit()
        CMS.dict_rab_mesta(self, self.db_users,conn_users)
        CSQ.close_bd(conn_users, cur_users)

        #=================================
        self.load_lbl_schema()

        self.obn_spis_pr()
        self.clear_mk2()
        self.edit_cr_mk = {2, 3, 4, 5, 8, 9, 19}
        self.edit_cr_mk_ruch = {0, 1, 2, 3, 4, 5, 8, 9, 19, 20}
        self.TIP_NEGRUZ_DSE = ('Сборочный чертёж', 'Изделие проекта','Монтажный чертёж','Материал')
        #CMS.add_menu(self)

        # self.ui.menu.addAction(self.ui.menu_22.menuAction())

        # self.ui.tabWidget.setCurrentIndex(3)

        self.sp_ins = ['комплектация', 'изготовление', 'контроль']
        self.nom_mk_dlya_korr = None
        self.spis_nom_tk_kor_mk = []
        self.spis_nom_tk_del_kor_mk = []
        try:
            spis_tablic_db = CSQ.spis_tablic(self.bd_naryad)
        except:
            CQT.msgbox(f'База занята')
            quit()

        places = CSQ.zapros(self.bd_naryad,"""SELECT Имя FROM places""",shapka=False)
        for place in places:
            self.ui.action_place = QtWidgets.QAction(place[0], self)
            self.ui.action_place.triggered.connect(lambda checked, item=place[0]: self.select_place(place[0]))
            self.ui.menu_3.addAction(self.ui.action_place)
        if F.nalich_file(CMS.tmp_dir() + F.sep() + 'place.txt'):
            self.place = int(F.load_file(CMS.tmp_dir() + F.sep() + 'place.txt'))
        else:
            self.place = None


        path = F.scfg('mk_data') + F.sep() + 'schems' + F.sep()
        if F.nalich_file(path):
            list_files = F.spis_files(path)
            for file in list_files[0][2]:
                if F.ostavit_rasshir(file) == '.jpg':
                    self.ui.cmb_schems.addItem(F.ubrat_rasshir(file))

        if not CMS.user_access(self.bd_naryad, 'mkart_mk_korrect_res_xml', F.user_name(), msg=False):
            self.ui.btn_update_norm.setEnabled(False)
            self.ui.btn_update_norm_rc.setEnabled(False)
            self.ui.btn_update_norm_prof.setEnabled(False)
            self.ui.btn_obnov_po_strukt.setEnabled(False)
            self.ui.btn_obnovit_naruadi_po_mk.setEnabled(False)
            self.ui.btn_ochistit.setEnabled(False)
            self.ui.btn_open_korr_mk.setEnabled(False)
            self.ui.btn_add_v_planetapi.setEnabled(False)
            self.ui.btn_zapoln_osv_zav_po_nar.setEnabled(False)

        # ==============VREMENNO========================================
        # self.ui.btn_zaversh.setDisabled(1)
        #self.VREMENNO_pereschet_vesa_mk()
        self.ui.btn_open_korr_mk.setDisabled(1)
        #VREMENNO   self.miration_data_sql()
        # self.ui.btn_vigruzka_norm.setDisabled(1)
        # ==============================================================

        self.ui.lbl_shema.mousePressEvent = self.getPos
        CQT.load_css(self)
        CQT.load_icons(self,24)

    def tbl_pl_gaf_svod_mouseMoveEvent(self,e):
        GVKPL.hover_tbl_pl_gaf_svod(self,e)
    def tbl_preview_mouseMoveEvent(self,e):
        GKPL.hover_tbl_preview(self,e)
    @CQT.onerror
    def tbl_pl_gaf_mouseMoveEvent(self,e):
        GKPL.hover_tbl_pl_gaf(self,e)


    def getPos(self, event):
        x = event.pos().x()
        y = event.pos().y()
        CQT.statusbar_text(self,'Mouse coords: ( %d : %d )' % (x, y))
        F.copy_bufer(str(x)+ ";" + str(y))



    def keyReleaseEvent(self, e):
        if e.key() == QtCore.Qt.Key_F11:
            if self.isFullScreen():
                self.showNormal()
            else:
                self.showFullScreen()

        if self.ui.tbl_tabeli.hasFocus():
            if e.key() == 86 and e.modifiers() == QtCore.Qt.ControlModifier:
                IND.load_tabel_from_bufer(self)
        if self.ui.tbl_pl_gaf_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_pl_gaf_filtr, self.ui.tbl_pl_gaf)
                self.ui.tbl_pl_gaf.setRowHidden(0, True)
                self.ui.tbl_pl_gaf.setRowHidden(1, True)
                GVKPL.load_svod(self)

        if self.ui.tbl_filtr_kal_pl.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_filtr_kal_pl, self.ui.tbl_kal_pl)
        if self.ui.tbl_vacant_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_vacant_filtr, self.ui.tbl_vacant)
        if self.ui.tbl_emploee_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_emploee_filtr, self.ui.tbl_emploee)
        if self.ui.tbl_tabeli_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_tabeli_filtr, self.ui.tbl_tabeli)
        if self.ui.tbl_rc_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_rc_filtr, self.ui.tbl_rc)
        if self.ui.tbl_rc.hasFocus():
            if e.key() == 16777220:
                row = self.ui.tbl_rc.currentRow()
                column = self.ui.tbl_rc.currentColumn()

                IND.cellChanged(self, row, column)
        if e.key() == 67 and e.modifiers() == (QtCore.Qt.ControlModifier | QtCore.Qt.ShiftModifier):
            if CQT.focus_is_QTableWidget():
                CQT.copy_bufer_table(QtWidgets.QApplication.focusWidget())


        if self.ui.tbl_selector_proj_view_zamech_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_selector_proj_view_zamech_filtr, self.ui.tbl_selector_proj_view_zamech)
        if self.ui.tbl_selector_proj_view.hasFocus():
            if e.key() == 16777220:
                if self.ui.tbl_selector_proj_view.currentColumn() == CQT.nom_kol_po_imen(self.ui.tbl_selector_proj_view,'Примечание'):
                    self.SLCT_edit_primech(self.ui.tbl_selector_proj_view.currentRow(),self.ui.tbl_selector_proj_view.currentColumn())
        if self.ui.tbl_selector_proj_view_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_selector_proj_view_filtr, self.ui.tbl_selector_proj_view)
        if self.ui.tbl_selector_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_selector_filtr, self.ui.tbl_selector)
        if self.ui.tbl_obespechenie_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_obespechenie_filtr, self.ui.tbl_obespechenie)
        if self.ui.tbl_zamech_filtr.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_zamech_filtr, self.ui.tbl_zamech)
        if self.ui.tbl_filtr_nomenkl.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_filtr_nomenkl, self.ui.table_nomenkl)
        if self.ui.tbl_filtr_mk.hasFocus():
            if e.key() == 16777220:
                CMS.primenit_filtr(self, self.ui.tbl_filtr_mk, self.ui.table_spis_MK)
        if self.ui.table_spis_MK.hasFocus() == True:
            if e.key() == 16777223 and e.modifiers() == QtCore.Qt.ShiftModifier:
                self.clear_res()
            if e.key() == 16777220:
                self.corr_mk(self.ui.table_spis_MK.currentRow(), self.ui.table_spis_MK.currentColumn())
            if e.key() == 16777222 and e.modifiers() == QtCore.Qt.ShiftModifier:
                self.create_and_add_res_to_mk()
        if self.ui.table_zayavk.hasFocus() == True:
            if e.key() == QtCore.Qt.Key_Delete:
                self.ui.table_zayavk.removeRow(self.ui.table_zayavk.currentRow())
    def cmb_tip_click(self,nom):
        if self.ui.cmb_tip_mk.currentText() == '':
            self.fill_cmb_dorez(True)
            return
        if self.DICT_TIP_MK[self.ui.cmb_tip_mk.currentText()]['Пномер'] == 2:
            self.fill_cmb_dorez()
        else:
            self.fill_cmb_dorez(True)


    def fill_cmb_dorez(self,clear=False):
        if clear:
            self.ui.cmb_tip_dorez.clear()
        else:
            self.ui.cmb_tip_dorez.addItem('')
            self.ui.cmb_tip_dorez.addItems([*self.DICT_TIP_DOREZ.keys()])
            self.ui.cmb_tip_dorez.adjustSize()

    def load_lbl_schema(self):
        self.ui.tabWidget_6.setCurrentIndex(CQT.nom_tab_po_imen(self.ui.tabWidget_6, 'Схема'))
        self.SIZE_SCHEMA_LBL = self.ui.lbl_shema.size()
        self.ui.lbl_shema.setScaledContents(False)

    def load_tree(self, spis_xml: list):

        tree = self.ui.treeWidget
        # tree.setColumnCount(10)
        list_user = ["Наименование"
            , "Обозначение полное"
            , "Количество"
            , "Ед.изм."
            , "Масса/М1,М2,М3"
            , "Ссылка на объект DOCs"
            , "ID"
            , "Количество на изделие"
            , "Примечание"
            , 'Покупное изделие'
            , "Классификатор изделия"
            , "Код ERP"
            , 'Раздел'
                     ]
        set_shapka = set()
        for item in spis_xml:
            for name in item['data'].keys():
                set_shapka.add(name)
        list_shapka = sorted(list(set_shapka))

        for item in list_shapka:
            if item not in list_user:
                list_user.append(item)
        list_user.append('Уровень')
        iter = 0
        for item in list_user:
            tree.headerItem().setText(iter, QtCore.QCoreApplication.translate("MainW", item))
            iter += 1

        for _ in range(0, len(list_user)):
            tree.resizeColumnToContents(_)

        tree.setSelectionBehavior(1)
        CQT.ust_cvet_videl_tab(tree)
        tree.setSelectionMode(1)

        # tree.setColumnWidth(1, int(tree.width() * 0.1))
        # tree.setColumnWidth(0, int(tree.width() - tree.columnWidth(1) - 81) - 5)
        tree.setStyleSheet(
            "QTreeView {background-color: rgb(212, 212, 212);} QTreeView::item:hover {background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1, stop:"
            " 0 #e7effd, stop: 1 #cbdaf1);border: 1px solid #bfcde4;} ")
        tree.setFocusPolicy(15)
        return list_user


    def tab_click(self, ind):
        def get_params_kpl(self:mywindow):
            kpl_bool_load_zav = 0
            try:
                kpl_bool_load_zav = F.valm(CMS.load_tmp_path('kpl_bool_load_zav'))
            except:
                pass
            self.ui.chk_kpl_zaversch.setChecked(kpl_bool_load_zav)

        if CMS.kontrol_ver(self.versia, 'МКарты') == False:
            quit()
        tab = self.ui.tabWidget
        if self.place == None:
            CQT.msgbox('Не выбрано подразделение (Файл->Подразделение)')
            tab.setCurrentIndex(CQT.nom_tab_po_imen(tab,'Просмотр структуры'))
        if tab.currentIndex() == 2:  # номенклатура
            self.zapoln_tabl_nomenkl()
        if tab.currentIndex() == 3:  # брак
            if F.nalich_file(self.bd_act):
                usl = "Категория_брака == 'Неисправимый' AND Ном_мк_повт_изг == ''"
                spis_itog = CSQ.naiti_v_bd(self.bd_act, 'act', {}, siroe_usl=usl, shapka=True)
            else:
                spis_itog = [["Не найдена база данных BDact"]]
            CQT.zapoln_wtabl(self, spis_itog, self.tabl_brak, 0, 0, (), (), 200, True, '')
        if tab.currentIndex() == CQT.nom_tab_po_imen(tab, 'Маршрутные карты'):  # мк
            self.load_table_mk()
        if tab.tabText(ind) == 'Загрузка рабочих центров':
            if self.ui.tabWidget_4.currentIndex() == CQT.nom_tab_po_imen(self.ui.tabWidget_4, 'Рабочие места'):
                IND.zagruzka_rc(self)
                self.load_lbl_schema()
        if tab.tabText(ind) == 'Замечания по МК':
            ZMCH.load_table_add(self)
            ZMCH.load_table_zamech(self)
        if tab.tabText(ind) == 'Селектор':
            SLCT.load_table_db(self)
            SLCT.load_table_add(self)
        if tab.tabText(ind) == 'Объемно-календарное планирование':
            self.kpl_mode = 0
            self.data_kpl= Data_plan()
            #self.LIST_ETAPS = [ _ for _ in CSQ.spis_tablic(self.db_kplan) if 'пл_' in _ ]
            self.list_tbl_info = []
            self.edit_tabel_mode = False
            if "val_masht" not in dir(self):
                self.val_masht = 12
                try:
                    self.val_masht = int(CMS.load_tmp_path('mk_val_masht'))
                except:
                    pass
            GVKPL.load_diapazon_month(self)
            KPL.load_gui(self)
            self.ui.splitter_pl.setSizes([400, 180])
            get_params_kpl(self)
            GVKPL.get_max_mosh_from_db(self)
            self.glob_kpl_summ_selct_tbl = ''
            self.dict_form_kpl = ''


    def tab_zagruzka_rc(self,nom):
        if self.ui.tabWidget_4.tabText(nom) == 'Рабочие места':
            IND.zagruzka_rc(self)
        if self.ui.tabWidget_4.tabText(nom) == 'Список сотрудников':
            IND.load_emploee(self)
        if self.ui.tabWidget_4.tabText(nom) == 'Вакантные места':
            IND.load_deficit_emploee(self)
        if self.ui.tabWidget_4.tabText(nom) == 'Табели':
            IND.add_list_of_months_to_cmb(self)

    def tab_click2(self, nom):
        if self.ui.tabWidget_2.tabText(nom) == 'Разработка МК':
            self.ui.btn_vigruzka_norm.setEnabled(False)
            self.ui.btn_vigruzka_norm_mat.setEnabled(True)
            self.list_vars_vo = []
            self.spis_poziciy_rez_ruchnoi = []
            self.res = ''
        if self.ui.tabWidget_2.tabText(nom) == 'Создание МК из *.XML':
            self.ui.btn_vigruzka_norm.setEnabled(True)
            self.ui.btn_vigruzka_norm_mat.setEnabled(False)
        if self.ui.tabWidget_2.tabText(nom) == 'Корректировка':
            self.load_mk_to_edit()

    def tab_mk_click(self, nom):
        self.glob_nom_mk_obesp = ''
        if self.ui.tabWidget_3.tabText(nom) == 'Обеспечение':
            OBSP.load_obesp_mk(self)

    def clck_tbl_rc(self):
        IND.clck_tbl_rc(self)

    def load_mk_to_edit(self):
        if not CMS.user_access(self.bd_naryad,'mkart_mk_korrect_res_xml',F.user_name()):
            return
        tbl = self.ui.table_spis_MK
        if tbl.currentRow() == -1:
            CQT.msgbox(f'Не выбрана МК')
            return
        nk_nom_mk = CQT.nom_kol_po_imen(tbl,'Пномер')
        nom_mk = int(tbl.item(tbl.currentRow(),nk_nom_mk).text())
        res = CSQ.zapros(self.db_resxml,f'''SELECT data FROM res WHERE Номер_мк == {nom_mk};''',shapka=False,one=True)
        if res == False:
            CQT.msgbox(f'ОШибка')
            return
        try:
            res = F.from_binary_pickle(res[-1][0])
            self.ui.txt_res.setPlainText(pprint.pformat(res))
        except:
            CQT.msgbox(f'Некорректные данные')
        return


    def edit_res_xml(self):
        tbl = self.ui.table_spis_MK
        if tbl.currentRow() == -1:
            CQT.msgbox(f'Не выбрана МК')
            return
        nk_nom_mk = CQT.nom_kol_po_imen(tbl, 'Пномер')
        nom_mk = int(tbl.item(tbl.currentRow(), nk_nom_mk).text())
        if self.ui.tabWidget_6.currentIndex() == CQT.nom_tab_po_imen(self.ui.tabWidget_6,'RES'):
            res = self.ui.txt_res.toPlainText()
            res = eval(res)
            blob = F.to_binary_pickle(res)
            if not CQT.msgboxgYN(f'ТОчно внести правку?'):
                return
            CSQ.zapros(self.db_resxml,f'''UPDATE res SET data = {blob} WHERE Номер_мк == {nom_mk};''')
        

    def select_place(self,item):
        kod = CSQ.zapros(self.bd_naryad,f'''SELECT Код FROM places WHERE Имя == "{item}" ''',shapka=False)
        spis = kod
        self.place = kod[-1][0]
        F.save_file(CMS.tmp_dir() + F.sep() + 'place.txt',spis)



    def SLCT_click(self):
        CMS.on_section_resized(self)

    def SLCT_edit_primech(self,row,column):
        if self.ui.tbl_selector_proj_view.hasFocus():
            SLCT.edit_primech(self,row,column)

    def SLCT_edit_zamech_from_view(self,row,column):
        SLCT.edit_zamech_from_view(self, row, column)

    def SLCT_add_new_zamech(self,row,column):
        SLCT.add_new_zamech(self, row, column)

    def SLCT_selector_proj_view_itemSelection(self):
        SLCT.selector_proj_view_itemSelection(self)

    def SLCT_select_zamech(self,row,column):
        SLCT.select_zamech(self,row,column)

    def OBSP_select_obesp_po_mk_from_table(self,row,column):
        OBSP.select_obesp_po_mk_from_table(self,row,column)

    def IND_cellChanged(self,row,column):
        if self.ui.tbl_rc.hasFocus():
            IND.cellChanged(self,row,column)

    def load_table_mk(self, fitr=''):
        self.load_tab_mk()
        tbl = self.ui.table_spis_MK
        CMS.zapolnit_filtr(self, self.ui.tbl_filtr_mk, tbl, fitr)
        nk_ststus = CQT.nom_kol_po_imen(self.ui.tbl_filtr_mk, 'Статус')
        self.ui.tbl_filtr_mk.item(0, nk_ststus).setText('!Закрыта')
        CMS.primenit_filtr(self, self.ui.tbl_filtr_mk, tbl)
        # === процент выполнения====
        nk_nom_mk = CQT.nom_kol_po_imen(tbl, 'Пномер')
        nk_progress = CQT.nom_kol_po_imen(tbl, 'Прогресс_01')
        if nk_progress != None:
            conn, cur = CSQ.connect_bd(self.db_resxml)
            for i in range(tbl.rowCount()):
                if not tbl.isRowHidden(i):
                    query = f'''SELECT data FROM res WHERE Номер_мк == {int(tbl.item(i, nk_nom_mk).text())}
                                        '''
                    res = F.from_binary_pickle(CSQ.zapros(self.db_resxml, query, conn)[-1][0])
                    if res == False:
                        CSQ.close_bd(conn,cur)
                        CQT.msgbox(f'БД занята пробуй позже')
                        return
                    if res == None:
                        tbl.item(i, nk_progress).setText('0|0')
                    else:
                        tbl.item(i, nk_progress).setText(CMS.procent_vip(res, '01'))
                else:
                    tbl.item(i, nk_progress).setText('0|0')

            CQT.zapolnit_progress(self, tbl, nk_progress)
            CSQ.close_bd(conn,cur)
        # =======
        tbl.selectRow(tbl.rowCount() - 1)
        tbl.scrollToBottom()

    def zapoln_osv_zav_po_nar(self):
        tbl = self.ui.table_spis_MK
        row = tbl.currentRow()
        if row == -1:
            return
        nk_pnom = CQT.nom_kol_po_imen(tbl, 'Пномер')
        nom_mk = int(tbl.item(row, nk_pnom).text())
        if not CQT.msgboxgYN(f'Заполнить по нарядам МК №{nom_mk}?'):
            return

        zapros = f'''SELECT Пномер, ДСЕ_ID, Операции, Опер_колво,ФИО,Фвремя,ФИО2,Фвремя2 FROM naryad WHERE Номер_мк == {nom_mk}'''
        rez = CSQ.zapros(self.bd_naryad, zapros, rez_dict=True)
        res = CMS.load_res(nom_mk)
        if res == False:
            CQT.msgbox(f'ОШибка загрузки ресурсной из бд')
            return
        for i in range(len(res)):
            for j in range(len(res[i]['Операции'])):
                res[i]['Операции'][j]['Освоено,шт.'] = 0
                res[i]['Операции'][j]['Закрыто,шт.'] = 0
        for naryad in rez:
            spis_id = naryad['ДСЕ_ID'].split('|')
            spis_op = naryad['Операции'].split('|')
            spis_kol = naryad['Опер_колво'].split('|')
            flag_zav = self.zaversh_naryad(naryad['Пномер'],'')
            for i in range(len(spis_id)):
                id = int(spis_id[i])
                op_nom = spis_op[i].split('$')[0]
                op_ima = spis_op[i].split('$')[1]
                kol = int(spis_kol[i])
                for j in range(len(res)):
                    if res[j]['Номерпп'] == id:
                        kol_det = res[j]['Количество']
                        for k in range(len(res[j]['Операции'])):
                            if res[j]['Операции'][k]['Опер_наименовние'] == op_ima and res[j]['Операции'][k][
                                'Опер_номер'] == op_nom:
                                res[j]['Операции'][k]['Освоено,шт.'] += kol
                                if res[j]['Операции'][k]['Освоено,шт.'] > kol_det:
                                    res[j]['Операции'][k]['Освоено,шт.'] = kol_det
                                if flag_zav:
                                    res[j]['Операции'][k]['Закрыто,шт.'] += kol
                                    if res[j]['Операции'][k]['Закрыто,шт.'] > kol_det:
                                        res[j]['Операции'][k]['Закрыто,шт.'] = kol_det

        CMS.save_res(self.db_resxml, nom_mk, res)
        CQT.msgbox(f'маршрутка {nom_mk} обновлена по бд')


    def zaversh_naryad(self, nom_nar, conn):
        zapros = f'''SELECT ФИО, Фвремя, ФИО2, Фвремя2 FROM naryad WHERE Пномер == {nom_nar}'''
        naryad = CSQ.zapros(self.bd_naryad, zapros, rez_dict=True, conn=conn)[0]
        flag_zav = True
        if naryad['ФИО'] == "" and naryad['ФИО2'] == "":
            flag_zav = False
        if flag_zav:
            if naryad['ФИО'] != "":
                if naryad['Фвремя'] == "":
                    flag_zav = False
        if flag_zav:
            if naryad['ФИО2'] != "":
                if naryad['Фвремя2'] == "":
                    flag_zav = False
        return flag_zav

    def check_zaversheni_naruady(self,mkarti:list) -> bool:
        '''True if all naruads closed'''
        zapros = f'''SELECT mk.Пномер, naryad.Пномер as Номер_Наряда, naryad.ФИО , naryad.ФИО2  FROM mk INNER JOIN naryad ON mk.Пномер = naryad.Номер_мк WHERE
                (naryad.ФИО != "" and naryad.Фвремя == "") or (naryad.ФИО2 != "" and naryad.Фвремя2 == "")'''
        list_nezversh_nar = CSQ.zapros(self.bd_naryad, zapros)
        list_nezav = []
        for mk in mkarti:
            for i in range(len(list_nezversh_nar)):
                if mk == list_nezversh_nar[i][0]:
                    list_nezav.append(f'МК № {mk}, Наряд № {list_nezversh_nar[i][1]} ({list_nezversh_nar[i][2]},{list_nezversh_nar[i][3]})')
        if list_nezav != []:
            CQT.msgbox(f'По маршрутке {mkarti} не завершены наряды, закрытие невозможно.{list_nezav}')
            return False
        return True

    def zaversh_mkards(self):
        if not CMS.user_access(self.bd_naryad,'мкарт_маршрутные_завершить',F.user_name()):
            return
        modifiers = CQT.get_key_modifiers(self)
        tbl = self.ui.table_spis_MK
        nk_nommk = CQT.nom_kol_po_imen(tbl, 'Пномер')
        if modifiers == ['shift']:
            spis_mk = []
            list_naruads = []
            for i in range(tbl.rowCount()):
                if tbl.isRowHidden(i) == False:
                    nom_mk = int(tbl.item(i, nk_nommk).text())
                    list_naruads.append(nom_mk)
            if not self.check_zaversheni_naruady(list_naruads):
                return
            for i in range(tbl.rowCount()):
                if tbl.isRowHidden(i) == False:
                    spis_mk.append(tbl.item(i,nk_nommk).text())
            spis_mk_text = ", ".join(spis_mk)
            if CQT.msgboxgYN(f'Будут принудительно закрыты маршрутные карты №№ {spis_mk_text}'):
                conn, cur = CSQ.connect_bd(self.bd_naryad)
                for i in range(tbl.rowCount()):
                    if tbl.isRowHidden(i) == False:
                        rez  = self.zaversh_mk(i ,conn=conn,cur=cur)
                        if rez == False:
                            CSQ.close_bd(conn, cur)
                            return
                CSQ.close_bd(conn,cur)
                CQT.msgbox(f'Успешно завершено')
                filtr = CMS.znach_filtr(self, self.ui.tbl_filtr_mk)
                self.load_table_mk(filtr)
                try:
                    msg = f"{F.user_full_namre()} ЗАВЕРШИЛ МК №№ {spis_mk_text}"
                    self.send_info_mk_b24(msg, 'chat41228')
                except:
                    print('Ошибка отправки в Б24')
        else:
            nom_mk = int(tbl.item(tbl.currentRow(), nk_nommk).text())

            project = tbl.item(tbl.currentRow(),
                                        CQT.nom_kol_po_imen(tbl, 'Номенклатура')).text()
            nom_pu_r = tbl.item(tbl.currentRow(),
                                         CQT.nom_kol_po_imen(tbl, 'Номер_заказа')).text()
            nom_pr_r = tbl.item(tbl.currentRow(),
                                         CQT.nom_kol_po_imen(tbl, 'Номер_проекта')).text()
            kolvo = tbl.item(tbl.currentRow(), CQT.nom_kol_po_imen(tbl, 'Количество')).text()
            prim = tbl.item(tbl.currentRow(), CQT.nom_kol_po_imen(tbl, 'Примечание')).text()
            osnovanie = tbl.item(tbl.currentRow(),
                                          CQT.nom_kol_po_imen(tbl, 'Основание')).text()
            if not self.check_zaversheni_naruady([nom_mk]):
                return
            conn, cur = CSQ.connect_bd(self.bd_naryad)
            self.zaversh_mk(conn=conn, cur =cur)
            CSQ.close_bd(conn,cur)
            try:
                msg = f"{F.user_full_namre()} ЗАВРЕШИЛ МК № {str(nom_mk)}:\n{project} - {str(kolvo)} шт.\n{nom_pu_r.strip()} Проект: {nom_pr_r.strip()}\n" \
                      f"Прим.: {prim} {osnovanie}"
                self.send_info_mk_b24(msg, 'chat41228')
            except:
                print('Ошибка отправки в Б24')


    def zaversh_mk(self, row:int = '', conn='',cur = ''):
        tbl = self.ui.table_spis_MK
        mode_one = False
        if row == "":
            mode_one = True
            row = tbl.currentRow()
        if row == -1:
            return
        nk_pnom = CQT.nom_kol_po_imen(tbl, 'Пномер')
        nom_mk = int(tbl.item(row, nk_pnom).text())
        if mode_one:
            if not CQT.msgboxgYN(f'Завершить МК №{nom_mk}?'):
                return
        zapros = f'''SELECT Номер_мк FROM res WHERE Номер_мк == {nom_mk}'''
        rez = CSQ.zapros(self.db_resxml, zapros, rez_dict=True, one=True)
        if rez == False:
            CQT.msgbox(f'Не найдена ресурсная. Нужно переоткрыть')
            return False
        zapros = f'''SELECT mk.Дата_завершения, mk.Статус, mk.Основание  FROM mk
            WHERE Пномер == {nom_mk}'''
        rez = CSQ.zapros(self.bd_naryad, zapros, rez_dict=True, one=True,conn=conn, cur = cur)
        if rez['Дата_завершения'] != '':
            CQT.msgbox(f'Нельзя завершить ранее завершенную МК №{nom_mk}')
            return False
        if rez['Статус'] == "Закрыта":
            CQT.msgbox(f'Нельзя завершить закрытую МК №{nom_mk}')
            return False
        if rez['Статус'] == "Открыта":
            if mode_one:
                conn_res, cur_res = CSQ.connect_bd(self.db_resxml)
                res = CMS.load_res(nom_mk, conn=conn_res, cur=cur_res)
                CSQ.close_bd(conn_res, cur_res)
                neosvoeno = self.check_gotovnost_mk(res)
                flag_zav = False
                if neosvoeno == None:
                    flag_zav = True
                else:
                    F.copy_bufer(neosvoeno)
                    CQT.msgbox(f'(Скопировано в буфер)* По мк {nom_mk} еще не завершено {neosvoeno}')
                    if CQT.msgboxgYN(F'Завершить принудительно?'):
                        flag_zav = True
            else:
                flag_zav = True
            if flag_zav:
                zapros = f'''UPDATE mk SET Статус = "Закрыта", Дата_завершения = "{F.now()}" WHERE Пномер =={nom_mk}'''
                CSQ.zapros(self.bd_naryad, zapros,conn=conn, cur = cur)
                if rez['Основание'] != "":
                    arr_tmp_ass = rez['Основание'].split(';')
                    for nom_acta in arr_tmp_ass:
                        CSQ.update_bd_sql(self.bd_act, 'act', {'Ном_мк_повт_изг': int(nk_pnom)},
                                          {'Пномер': int(nom_acta)}, conn=conn,cur=cur)
                if mode_one:
                    CQT.msgbox(f'Успешно завершено')
                    filtr = CMS.znach_filtr(self, self.ui.tbl_filtr_mk)
                    self.load_table_mk(filtr)

    def check_gotovnost_mk(self, res):
        rez = []
        for i in range(len(res)):
            kolich = res[i]['Количество']
            nn = res[i]['Номенклатурный_номер']
            naim = res[i]['Наименование']
            for oper in res[i]['Операции']:
                osv = 0
                zav = 0
                oper_name = f"{oper['Опер_номер']} {oper['Опер_наименовние']}"
                if 'Освоено,шт.' in oper:
                    osv = oper['Освоено,шт.']
                if 'Закрыто,шт.' in oper:
                    zav = oper['Закрыто,шт.']
                if osv < kolich:
                    rez.append(f'{nn} {naim} {oper_name} не освоено {kolich - osv} шт.')
                if zav < kolich:
                    rez.append(f'{nn} {naim} {oper_name} не закрыто {kolich - zav} шт.')
        if rez == []:
            return
        return '\n'.join(rez)


    def update_norm(self, vid = ''):

        tbl = self.ui.table_spis_MK
        if tbl.currentRow() == -1:
            return
        nk_nommk = CQT.nom_kol_po_imen(tbl, 'Пномер')
        nom_mk = tbl.item(tbl.currentRow(), nk_nommk).text()
        if not CQT.msgboxgYN(f'Обновить нормы для МК {nom_mk}?'):
            return

        conn_res, cur_res = CSQ.connect_bd(self.db_resxml)
        res = CMS.load_res(int(nom_mk), conn=conn_res, cur=cur_res)
        CSQ.close_bd(conn_res, cur_res)

        if res:
            conn1, cur1 = CSQ.connect_bd(self.db_dse)
            nk_urov = 20
            schet_izm = 0
            for i in range(0, len(res)):
                nn = res[i]['Номенклатурный_номер'].strip()
                naim = res[i]['Наименование'].strip()
                kolvo = int(res[i]['Количество'])

                #nom_tk = CSQ.naiti_v_bd(self.db_dse, 'dse', {'Номенклатурный_номер': nn,
                #                                                'Наименование': naim},
                #                        ['Номер_техкарты'], all=False, conn=conn1, cur=cur1)

                nom_tk = CSQ.zapros(self.db_dse,f"""SELECT Номер_техкарты FROM dse WHERE Номенклатурный_номер == '{nn}' 
                and  Наименование == '{naim}'""",conn=conn1,cur=cur1,one=True)
                if len(nom_tk) == 1:
                    CQT.msgbox(f'Не найден номер техкарты/не сделана на {nn} {naim}')
                    continue
                nom_tk = nom_tk[1][0]
                putf = F.scfg('mk_data') + os.sep + nom_mk + os.sep + nom_tk + '_' + nn + '.pickle'
                if F.nalich_file(putf) == False:
                    F.skopir_file(F.scfg('add_docs') + os.sep + nom_tk + "_" + nn + '.pickle',
                                  F.scfg('mk_data') + os.sep + nom_mk + os.sep + nom_tk + '_' + nn + '.pickle')

                    #putf = F.scfg('add_docs') + os.sep + nom_tk + "_" + nn + '.pickle'
                    #uroven_dse = int(res[i][nk_urov])
                uroven_dse = CMS.uroven(nn)
                if F.nalich_file(putf):
                    rez_spis_op = []
                    nk_rc_tk = 4
                    nk_ur_tk = 20
                    nk_op_tk = 2
                    nk_op_oborud = 5
                    nk_mat_tk = 10
                    nk_doc_tk = 15
                    nk_textper = 0
                    nk_op_doc = 13
                    nk_op_tpz = 6
                    nk_op_tst = 7
                    nk_op_prof = 8
                    nk_op_KR = 9
                    nk_op_KOID = 11
                    nk_per_ima = 0
                    nk_per_instr = 12
                    nk_per_osn = 11
                    nk_per_doc = 13
                    sp_tk = F.otkr_f(putf, False, "|", pickl=True)
                    tk_docs = sp_tk[10][nk_doc_tk].split('$')
                    for j in range(11, len(sp_tk)):
                        rez_spis_mat = []
                        rez_spis_instr = []
                        rez_spis_osn = []
                        rez_spis_doc = []
                        if sp_tk[j][nk_ur_tk] == '0':
                            break
                        if sp_tk[j][nk_ur_tk] == '1':
                            mat_str = sp_tk[j][nk_mat_tk].split('{')
                            for k in range(len(mat_str)):
                                mat_str_str = mat_str[k].split('$')
                                if len(mat_str_str) == 4:
                                    rez_spis_mat.append({'Мат_код': mat_str_str[0], "Мат_наименование": mat_str_str[1],
                                                         "Мат_ед_изм": mat_str_str[2],
                                                         "Мат_норма": round(F.valm(mat_str_str[3]) * kolvo, 6),
                                                         "Мат_норма_ед": round(F.valm(mat_str_str[3]), 6),
                                                         'Мат_параметрика': dict()})
                                else:
                                    pass
                            rez_spis_instr = []
                            rez_spis_osn = []
                            rez_spis_doc = sp_tk[j][nk_op_doc].split('$')
                            rez_spis_doc = F.clear_free_items(rez_spis_doc)
                            spis_per = []
                            for k in range(j + 1, len(sp_tk)):
                                if sp_tk[k][nk_ur_tk] == '1' or sp_tk[k][nk_ur_tk] == '0':
                                    break
                                spis_per.append(sp_tk[k][nk_per_ima])
                                spis_instr = sp_tk[k][nk_per_instr].split('$')
                                for item in spis_instr:
                                    rez_spis_instr.append(item)
                                spis_osn = sp_tk[k][nk_per_osn].split('$')
                                for item in spis_osn:
                                    rez_spis_osn.append(item)
                                spis_doc = sp_tk[k][nk_per_doc].split('$')
                                for item in spis_doc:
                                    rez_spis_doc.append(item)

                            rez_spis_op.append({"Опер_наименовние": sp_tk[j][nk_textper],
                                                "Опер_код": CMS.kod_oper_po_ima(self.SPIS_OP ,sp_tk[j][nk_textper]),
                                                "Опер_номер": sp_tk[j][nk_op_tk],
                                                "Опер_РЦ_наименовние": CMS.ima_rc_po_nom(self.SPIS_RC, sp_tk[j][nk_rc_tk]),
                                                "Опер_РЦ_код": sp_tk[j][nk_rc_tk],
                                                "Опер_оборудование_наименовние": sp_tk[j][nk_op_oborud],
                                                "Опер_оборудование_код": CMS.kod_oborud_po_ima(self.SPIS_OB, sp_tk[j][nk_op_oborud]),
                                                "Опер_Тпз": round(F.valm(sp_tk[j][nk_op_tpz]), 6),
                                                "Опер_Тшт": round(F.valm(sp_tk[j][nk_op_tst]) * kolvo, 6),
                                                "Опер_профессия_наименование": CMS.ima_prof_po_kod(self.SPIS_PROF,
                                                    sp_tk[j][nk_op_prof]),
                                                "Опер_профессия_код": sp_tk[j][nk_op_prof],
                                                "Опер_КР": F.valm(sp_tk[j][nk_op_KR]),
                                                "Опер_КОИД": F.valm(sp_tk[j][nk_op_KOID]),
                                                "Опер_документы": rez_spis_doc,
                                                "Опер_инстумент": rez_spis_instr, "Опер_оснастка": rez_spis_osn,
                                                "Материалы": rez_spis_mat, "Переходы": spis_per})

                    for oper_j in range(len(res[i]['Операции'])):
                        for oper_j_tmp in range(len(rez_spis_op)):
                            if res[i]['Операции'][oper_j]['Опер_номер'] == rez_spis_op[oper_j_tmp][
                                'Опер_номер']:
                                if vid == '':
                                    if res[i]['Операции'][oper_j]['Опер_код'] != rez_spis_op[oper_j_tmp]['Опер_код']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_код'] = rez_spis_op[oper_j_tmp]['Опер_код']

                                    if res[i]['Операции'][oper_j]['Опер_наименовние'] != rez_spis_op[oper_j_tmp]['Опер_наименовние']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_наименовние'] = rez_spis_op[oper_j_tmp]['Опер_наименовние']
                                        if rez_spis_op[oper_j_tmp]['Опер_наименовние'] not in self.DICT_ETAPI:
                                            CQT.msgbox(f'{nn} операция {rez_spis_op[oper_j_tmp]["Опер_наименовние"]} не найден "Этап"')
                                            etap = 'Неопознан'
                                        else:
                                            etap = self.DICT_ETAPI[rez_spis_op[oper_j_tmp]['Опер_наименовние']]
                                        res[i]['Операции'][oper_j]['Этап'] = rez_spis_op[oper_j_tmp][
                                            etap]

                                    if res[i]['Операции'][oper_j]['Опер_КР'] != rez_spis_op[oper_j_tmp]['Опер_КР']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_КР'] = rez_spis_op[oper_j_tmp]['Опер_КР']

                                    if res[i]['Операции'][oper_j]['Опер_КОИД'] != rez_spis_op[oper_j_tmp]['Опер_КОИД']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_КОИД'] = rez_spis_op[oper_j_tmp]['Опер_КОИД']

                                    if res[i]['Операции'][oper_j]['Опер_документы'] != rez_spis_op[oper_j_tmp][
                                        'Опер_документы']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_документы'] = rez_spis_op[oper_j_tmp][
                                            'Опер_документы']

                                    if res[i]['Операции'][oper_j]['Опер_инстумент'] != rez_spis_op[oper_j_tmp][
                                        'Опер_инстумент']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_инстумент'] = rez_spis_op[oper_j_tmp][
                                            'Опер_инстумент']

                                    if res[i]['Операции'][oper_j]['Опер_оснастка'] != rez_spis_op[oper_j_tmp][
                                        'Опер_оснастка']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_оснастка'] = rez_spis_op[oper_j_tmp][
                                            'Опер_оснастка']



                                    if res[i]['Операции'][oper_j]['Переходы'] != rez_spis_op[oper_j_tmp]['Переходы']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Переходы'] = rez_spis_op[oper_j_tmp]['Переходы']
                                if vid == 'mat':

                                    if res[i]['Операции'][oper_j]['Материалы'] != rez_spis_op[oper_j_tmp]['Материалы']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Материалы'] = rez_spis_op[oper_j_tmp]['Материалы']

                                if vid == 'vrem':
                                    if res[i]['Операции'][oper_j]['Опер_Тпз'] != rez_spis_op[oper_j_tmp]['Опер_Тпз']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_Тпз'] = rez_spis_op[oper_j_tmp]['Опер_Тпз']
                                    if res[i]['Операции'][oper_j]['Опер_Тшт'] != rez_spis_op[oper_j_tmp]['Опер_Тшт']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_Тшт'] = rez_spis_op[oper_j_tmp]['Опер_Тшт']
                                        res[i]['Операции'][oper_j]['Опер_Тшт_ед'] = rez_spis_op[oper_j_tmp][
                                                                                        'Опер_Тшт'] / res[i][
                                                                                        'Количество']
                                if vid == 'rc':
                                    if res[i]['Операции'][oper_j]['Опер_РЦ_код'] != rez_spis_op[oper_j_tmp]['Опер_РЦ_код']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_РЦ_код'] = rez_spis_op[oper_j_tmp]['Опер_РЦ_код']

                                if vid == 'prof':
                                    if res[i]['Операции'][oper_j]['Опер_профессия_наименование'] != \
                                            rez_spis_op[oper_j_tmp]['Опер_профессия_наименование']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_профессия_наименование'] = \
                                            rez_spis_op[oper_j_tmp]['Опер_профессия_наименование']
                                    if res[i]['Операции'][oper_j]['Опер_профессия_код'] != rez_spis_op[oper_j_tmp][
                                        'Опер_профессия_код']:
                                        schet_izm += 1
                                        res[i]['Операции'][oper_j]['Опер_профессия_код'] = rez_spis_op[oper_j_tmp][
                                            'Опер_профессия_код']


                                break
            CSQ.close_bd(conn1, cur1)
            CMS.save_res(self.db_resxml, nom_mk, res)
            CQT.msgbox(f'маршрутка {nom_mk} обновлена, {schet_izm} изменений')
        else:
            CQT.msgbox(f'маршрутка {nom_mk} ОТСУТСТВУЕТ')

    def clk_fdate_res_erp(self):
        tbl = self.ui.table_spis_MK
        nk_date_etap = CQT.nom_kol_po_imen(tbl,'Ресурсная_дата')
        r = tbl.currentRow()
        if r == None or r == -1:
            return
        val_date = tbl.item(r,nk_date_etap).text()
        nk_nom_mk = CQT.nom_kol_po_imen(tbl,'Пномер')
        nom_mk = int(tbl.item(r,nk_nom_mk).text())
        now = F.now("%Y-%m-%d")
        buf = F.paste_bufer()
        if buf.strip() != '':
            if F.is_date(buf,"%Y-%m-%d"):
                if CQT.msgboxgYN(f'Использовать дату {buf} вместо текущей {now} для МК {nom_mk}? '):
                    now = buf

        if val_date != now:
            if not CQT.msgboxgYN(f'Заменить дату {val_date} на {now} для МК {nom_mk}? '):
                return
        
            request = f"""UPDATE mk SET Ресурсная_дата = '{now}' where Пномер == {nom_mk}"""
            CSQ.zapros(self.bd_naryad,request)
            tbl.item(r,nk_date_etap).setText(now)
            CQT.msgbox(f'Успешно')
        pass

    def raschet_etapa(self, fio):
        for key in self.DICT_EMPLOEE_RC:
            if fio in key:
                return self.DICT_EMPLOEE_RC[key]
        return

    def clear_res(self):
        if CQT.msgboxgYN('Очистить мк?'):
            tbl = self.ui.table_spis_MK
            n_k = CQT.nom_kol_po_imen(tbl, 'Пномер')
            nom_mk = int(tbl.item(tbl.currentRow(), n_k).text())
            zapros = f"""UPDATE res SET data = '' WHERE Номер_мк == {nom_mk}"""
            CSQ.zapros(self.bd_naryad, zapros)
            CQT.msgbox(f'маршрутка {nom_mk} очищена')

    def obnovit_naruadi_po_mk(self):
        tbl = self.ui.table_spis_MK
        row = tbl.currentRow()
        if row == -1:
            return
        nk_pnom = CQT.nom_kol_po_imen(tbl, 'Пномер')
        nom_mk = int(tbl.item(row, nk_pnom).text())
        if not CQT.msgboxgYN(f'Заполнить наряды по новым нормам и количеству из МК №{nom_mk}?'):
            return
        tmp_zapros = []
        zapros = f'''SELECT Пномер, Внеплан, Задание, ФИО, ФИО2, Твремя, ДСЕ_ID, Операции, Опер_время, Опер_колво FROM naryad WHERE Номер_мк == {nom_mk} AND Внеплан == 0'''
        rez = CSQ.zapros(self.bd_naryad, zapros, rez_dict=True)
        conn_res, cur_res = CSQ.connect_bd(self.db_resxml)
        res = CMS.load_res(int(nom_mk), conn=conn_res, cur=cur_res)
        CSQ.close_bd(conn_res, cur_res)
        for dse in res:
            id = dse['Номерпп']
            kolvo_dse = dse['Количество']
            nn = dse['Номенклатурный_номер']
            naim = dse['Наименование']
            for oper in dse['Операции']:
                nom = oper['Опер_номер']
                nazv = oper['Опер_наименовние']
                kolvo_oper = kolvo_dse
                norma_ed = oper['Опер_Тшт'] / dse['Количество']
                for n, naryad in enumerate(rez):
                    spis_dse_nar = naryad['ДСЕ_ID'].split('|')
                    spis_dse_kol = naryad['Опер_колво'].split('|')
                    spis_dse_vrem = naryad['Опер_время'].split('|')
                    spis_dse_oper = naryad['Операции'].split('|')
                    koef_rab = 1
                    zadanie = ''
                    if naryad['ФИО'] != '' and naryad['ФИО2'] != '':
                        koef_rab = 2
                    summ_tvr = 0
                    for i in range(len(spis_dse_nar)):
                        if int(spis_dse_nar[i]) == id and spis_dse_oper[i] == '$'.join([nom, nazv]):
                            kol_nar = int(spis_dse_kol[i])
                            if kol_nar <= kolvo_oper:
                                kolvo_oper -= kol_nar
                                norma_rez = norma_ed * kol_nar
                                spis_dse_vrem[i] = str(round(norma_rez))
                            else:
                                if kolvo_oper == 0:
                                    norma_rez = 0
                                    kol_rez = 0
                                else:
                                    kol_rez = kolvo_oper
                                    norma_def = norma_ed * kol_rez
                                    norma_rez = norma_def
                                    kolvo_oper = 0
                                    if round(norma_rez) == 0:
                                        norma_rez += 1
                                spis_dse_vrem[i] = str(round(norma_rez))
                                spis_dse_kol[i] = str(kol_rez)
                            break

                    for item in spis_dse_vrem:
                        summ_tvr += F.valm(item)
                    rez[n]['Опер_время'] = '|'.join(spis_dse_vrem)
                    rez[n]['Опер_колво'] = '|'.join(spis_dse_kol)
                    rez[n]['Твремя'] = round(summ_tvr / koef_rab)

        conn, cur = CSQ.connect_bd(self.bd_naryad)
        for n, naryad in enumerate(rez):
            spis_dse_nar = naryad['ДСЕ_ID'].split('|')
            spis_dse_kol = naryad['Опер_колво'].split('|')
            spis_dse_vrem = naryad['Опер_время'].split('|')
            spis_dse_oper = naryad['Операции'].split('|')
            for i in range(len(spis_dse_nar)):
                id = spis_dse_nar[i]
                kol = spis_dse_kol[i]
                vrem = spis_dse_vrem[i]
                oper_ima = spis_dse_oper[i]
                docs = ''
                perehod = ''
                naim = ''
                nn = ''
                for dse in res:
                    if dse['Номерпп'] == int(id):
                        naim = dse['Наименование']
                        nn = dse['Номенклатурный_номер']
                        for oper in dse['Операции']:
                            if oper_ima == '$'.join([oper['Опер_номер'], oper['Опер_наименовние']]):
                                docs = "; ".join(oper['Опер_документы'])
                                perehod = "; ".join(oper['Переходы'])
                                break
                        break
                head = f'{naim} {nn} ' \
                       f'({kol} шт.) - {vrem} мин.'
                body = f' {docs} ' + '\n' + \
                       f'    {oper["Опер_номер"]} {oper["Опер_наименовние"]}' + '\n' + \
                       f'   {perehod}'
                zadanie += head + body + '\n' + '\n'

            rez[n]['Задание'] = zadanie
            zapros = f'''UPDATE naryad SET Задание = "{rez[n]['Задание']}", Твремя = {rez[n]['Твремя']}, 
                        Опер_время = "{rez[n]['Опер_время']}", Опер_колво = "{rez[n]['Опер_колво']}" WHERE Пномер == {rez[n]['Пномер']} '''
            CSQ.zapros(self.bd_naryad, zapros, rez_dict=True, conn=conn, cur =cur)
        CSQ.close_bd(conn, cur)
        CQT.msgbox(f'Наряды по мк {nom_mk} обновлены')

    def obnov_po_strukt(self):
        # ============================Обновление количества по структуре==========================================================================
        tbl = self.ui.table_spis_MK
        n_k = CQT.nom_kol_po_imen(tbl, 'Пномер')
        nom_mk = int(tbl.item(tbl.currentRow(), n_k).text())
        query = f'''SELECT mk.Количество FROM mk 
            WHERE Пномер == {int(nom_mk)}
                        '''
        rez = CSQ.zapros(self.bd_naryad, query)

        conn_res, cur_res = CSQ.connect_bd(self.db_resxml)
        res = CMS.load_res(int(nom_mk), conn=conn_res, cur=cur_res)
        if res == False:
            CQT.msgbox('Нет ресурсной')
            return

        query = f'''SELECT data, Head FROM xml 
            WHERE Номер_мк == {int(nom_mk)}
                        '''
        rez_xml = CSQ.zapros(self.db_resxml, query)
        xml = rez_xml[-1][0]
        xml_head = rez_xml[-1][1]
        if xml == '':
            CQT.msgbox('Нет хмл файла')
            return

        CSQ.close_bd(conn_res, cur_res)

        res_new = CMS.resursnaya_from_xml(self, self.podgotovka_xml(XML.spisok_iz_xml(str_f=xml),xml_head), kol_vo_izdeliy=rez[-1][1])
        if len(res) != len(res_new):
            CQT.msgbox(f'Несовпадение числа деталей')
            return
        if res:
            for i in range(len(res)):
                if res[i]['Номенклатурный_номер'] != res_new[i]['Номенклатурный_номер']:
                    CQT.msgbox(f'Несовпадение номеров старой и новой мк')
                    return
                if res[i]['Количество'] != res_new[i]['Количество']:
                    print(
                        f"мК {nom_mk}, {res[i]['Номенклатурный_номер']} было {res[i]['Количество']}, стало {res_new[i]['Количество']}")
                    res[i]['Количество'] = res_new[i]['Количество']
                    for j in range(len(res[i]['Операции'])):
                        if 'Освоено,шт.' in res[i]['Операции'][j]:
                            if res[i]['Операции'][j]['Освоено,шт.'] >= res_new[i]['Количество']:
                                res[i]['Операции'][j]['Освоено,шт.'] = res_new[i]['Количество']
                        if 'Закрыто,шт.' in res[i]['Операции'][j]:
                            if res[i]['Операции'][j]['Закрыто,шт.'] >= res_new[i]['Количество']:
                                res[i]['Операции'][j]['Закрыто,шт.'] = res_new[i]['Количество']
                        res[i]['Операции'][j]['Опер_Тшт'] = res_new[i]['Операции'][j]['Опер_Тшт']
                        res[i]['Операции'][j]['Материалы'] = res_new[i]['Операции'][j]['Материалы']
            CMS.save_res(self.db_resxml, nom_mk, res)
            CQT.msgbox(f'маршрутка {nom_mk} обновлена')

    # ====================================================================================================================

    def load_cust_drevo(self):
        tmp_path = CMS.load_tmp_path('razr_mk')
        put_ima = CQT.f_dialog_name(self, 'выбор файла', tmp_path, '*.pickle', True)
        if put_ima == '.':
            return
        CMS.save_tmp_path('razr_mk',put_ima)
        spis = F.otkr_f(put_ima, False, separ='', pickl=True)
        modifiers = CQT.get_key_modifiers(self)
        if modifiers == ['shift']:
            row = self.ui.table_razr_MK.currentRow()
            if row == -1:
                uroven = 0
                row = 0
            else:
                uroven =  int(CQT.znach_vib_strok_po_kol(self.ui.table_razr_MK,'Уровень'))+1
            main_list = CQT.spisok_iz_wtabl(self.ui.table_razr_MK,shapka=True)
            nk_uroven = F.nom_kol_po_im_v_shap(spis,'Уровень')
            koef = 1
            for i in range(1,len(spis)):
                spis[i][nk_uroven] = str(int(spis[i][nk_uroven]) + uroven)
                main_list.insert(row+koef+1,spis[i])
                koef+=1
            spis = main_list
        CQT.zapoln_wtabl(self, spis, self.ui.table_razr_MK, 0, self.edit_cr_mk_ruch, (), (), 200, True, '', 30)
        self.spis_poziciy_rez_ruchnoi = []

    def save_cust_drevo(self):
        tbl = self.ui.table_razr_MK
        spis = CQT.spisok_iz_wtabl(tbl, '', True)
        tmp_path = CMS.load_tmp_path('razr_mk_save')

        put_ima = CQT.f_dialog_save(self, 'Выбрать куда сохранить', F.put_po_umolch(), '*.pickle')
        if put_ima == '.':
            return
        CMS.save_tmp_path('razr_mk_save', put_ima)
        F.zap_f(put_ima, spis, separ='', pickl=True)
        CQT.msgbox('Успешно')

    def export_mat_spec_tk(self):
        list_mk = ''
        """list_mk = (281, 472, 529, 294, 618, 537, 277, 321, 470, 471, 485, 493, 571, 327, 339, 397, 457, 285, 518,
                   551, 279, 320, 329, 614, 284, 465, 351, 497, 483, 432, 369, 283, 218, 292, 530, 543, 434, 492,
                   521, 237, 289, 398, 373, 393, 613, 466, 487, 382, 441, 386, 476, 216, 309, 568, 367, 219, 442,
                   473, 484, 616, 536, 547, 268, 435, 236, 532, 368, 307, 190, 385, 634, 310, 336, 467, 464, 364,
                   167, 217, 344, 540, 273, 308, 528, 356, 491, 520, 562, 221, 544, 437, 468, 524, 567, 459, 220,
                   522, 301, 443, 474, 527, 566, 486, 525, 235, 278, 288, 298, 365, 546, 456, 282, 157, 505, 286,
                   343, 295, 345, 637, 342, 381, 378, 296, 313, 158, 366, 270, 316, 488, 478, 379, 326, 436, 392,
                   384, 489, 275, 330, 318, 306, 463, 297, 325, 691, 475, 615, 274, 440, 458, 214, 390, 534, 350,
                   215, 291, 287, 523, 337, 612, 328, 572, 433, 280, 315, 340, 290, 317, 469, 516, 573, 490, 374,
                   276, 619, 399, 560, 461, 271, 538, 391, 383, 168, 380, 526, 293, 311, 455, 617, 462, 375, 477,
                   300, 496, 267, 269, 633, 389, 519, 531, 312,
                    675, 648, 362, 635, 644, 679, 620, 676, 650, 494, 685, 561, 495, 569, 570, 649, 564, 674, 363, 360)"""
        self.export_mat_spec(po_tk = True,list_mk=list_mk)

    def export_mat_spec(self, po_tk = False, list_mk = ''):
        if self.ui.tabWidget.tabText(self.ui.tabWidget.currentIndex()) != 'Маршрутные карты':
            CQT.msgbox('Не выбрана вкладка Маршрутные карты')
            return
        if self.ui.table_spis_MK.currentRow() == -1:
            CQT.msgbox('Не выбрана маршрутная карта')
            return
        put_tmp = CMS.load_tmp_path('ved_mat')
        put = CQT.getDirectory(self, put_tmp)
        if put == '':
            return
        CMS.save_tmp_path('ved_mat', put)

        nk_pnomer = CQT.nom_kol_po_imen(self.ui.table_spis_MK, 'Пномер')
        nk_nomenk = CQT.nom_kol_po_imen(self.ui.table_spis_MK, 'Номенклатура')

        nom_mk = self.ui.table_spis_MK.item(self.ui.table_spis_MK.currentRow(), nk_pnomer).text()
        nomenk = self.ui.table_spis_MK.item(self.ui.table_spis_MK.currentRow(), nk_nomenk).text()
        spis_filtr_mat = CSQ.zapros(self.bd_nomen, f"""SELECT kod, filtr FROM complex_filtr""", shapka=False)
        if list_mk != '':
            query = f'''SELECT xml.data as xml, xml.Head as xml_head, mk.Количество, mk.Пномер, mk.Номенклатура, mk.Направление FROM mk 
                     INNER JOIN xml ON mk.Пномер = xml.Номер_мк
                     WHERE Пномер in {list_mk}
                                 '''
            rez_mk = CSQ.zapros(self.bd_naryad, query,shapka=False,rez_dict=True)
            for mk in rez_mk:
                xml = mk['xml']
                xml_head = mk['xml_head']
                kol_vo = mk['Количество']
                nom_mk = mk['Пномер']
                nomenk = mk['Номенклатура']
                napr = mk['Направление']
                if xml == '':
                    CQT.msgbox('Нет хмл файла')
                    continue
                spis_dse = CMS.resursnaya_from_xml(self, self.podgotovka_xml(XML.spisok_iz_xml(str_f=xml),xml_head),
                                                    kol_vo_izdeliy=kol_vo)
                spis = CMS.spis_mat_erp(self, nom_mk, spis_filtr_mat, po_tk = po_tk, spis_dse= spis_dse)
                CEX.zap_spis(spis, put, f'Нормы материалов на МК{nom_mk}_{napr}_{nomenk}_{kol_vo}шт.xlsx', '1', 1, 1)
        else:
            nk_kol_vo = CQT.nom_kol_po_imen(self.ui.table_spis_MK, 'Количество')
            kol_vo = int(self.ui.table_spis_MK.item(self.ui.table_spis_MK.currentRow(), nk_kol_vo).text())
            spis = CMS.spis_mat_erp(self, nom_mk, spis_filtr_mat, po_tk = po_tk)
            CEX.zap_spis(spis, put, f'Нормы материалов на МК{nom_mk}_{nomenk}_{kol_vo}шт.xlsx', '1', 1, 1)
        F.otkr_papky(put)

    def export_json_Excel(self):
        self.export_json(True)


    def export_table_txt(self):
        tab = self.ui.tabWidget
        if tab.currentIndex() == CQT.nom_tab_po_imen(tab, 'Отчеты'):
            pass

    def export_table(self):
        tab = self.ui.tabWidget
        if tab.currentIndex() == CQT.nom_tab_po_imen(tab, 'Отчеты'):
            pass

    def export_spis_etapov_s_vesom_for_plan_exel(self,putt, spis_poziciy_rez):
        dict_etapov = {}
        ves = 0
        for poz in spis_poziciy_rez:
            kolvo = poz['kol_zayavk']
            for dse in poz['data']:
                mat = dse['Мат_кд'].split('/')
                if mat[1] != '' and mat[2] != '':
                    ves += F.valm(mat[0])*dse['Количество']
                for oper in dse['Операции']:
                    if oper['Опер_наименовние'] not  in self.DICT_ETAPI:
                        CQT.msgbox(f"Операция {oper['Опер_наименовние']} отсутствует в списке операций БД. не выгружено")
                        return
                    etap = self.DICT_ETAPI[oper['Опер_наименовние']]
                    vrema = oper['Опер_Тпз'] + oper['Опер_Тшт']
                    if etap in dict_etapov:
                        dict_etapov[etap]+= vrema
                    else:
                        dict_etapov[etap] = vrema
        spis_etapov = [[k, round(v)] for k, v in dict_etapov.items()]
        spis_etapov.insert(0,['Вес',round(ves)])
        F.save_file(putt,spis_etapov,utf=False)
        CQT.msgbox(f'Многоуважаемый {F.user_name()}! Нормы для плана успешно выгружены, хорошего дня.')




    def export_json(self, exel=False):
        
        def generate(self,exel,rez,nom_mk,py):
            put = F.put_po_umolch()
            if F.nalich_file(CMS.tmp_dir() + F.sep() + 'json_dir_cache.txt'):
                sod_f = F.load_file(CMS.tmp_dir() + F.sep() + 'json_dir_cache.txt')
                if sod_f != []:
                    put = sod_f
            dir = CQT.getDirectory(self, put)
            if dir == ['.'] or dir == '.':
                return
            F.save_file(CMS.tmp_dir() + F.sep() + 'json_dir_cache.txt', dir)
            if exel == False:
                path = dir + F.sep() + f'{nom_mk}_{py}.json'
                F.zapis_json(rez, path, False)
                CQT.msgbox(f'Готово')
                F.otkr_papky(dir)
            else:
                spis_shab_mk = [
                    ['Номерпп', "Наименование", 'Номенклатурный_номер', 'Количество', 'Уровень', 'Опер_наименовние',
                     'Опер_код',
                     'Опер_номер', 'Опер_РЦ_наименовние', 'Опер_РЦ_код', 'Опер_оборудование_наименовние',
                     'Опер_оборудование_код',
                     'Опер_Тпз', 'Опер_Тшт*Кол', 'Опер_профессия_наименование', 'Опер_профессия_код', 'Опер_КР',
                     'Опер_КОИД', 'Опер_документы', 'Опер_инстумент', 'Опер_оснастка',
                     'Материалы*Кол', 'Переходы', 'Параметрика', 'Документы', 'ПКИ', 'Мат_кд', 'Ссылка', 'Прим']]
                for i, dse in enumerate(rez):
                    nomerpp = dse['Номерпп']
                    naim = dse['Наименование']
                    nn = dse['Номенклатурный_номер']
                    kolich = dse['Количество']
                    uroven = dse['Уровень']
                    mat = dse['Мат_кд']
                    parametrika = '@'.join(list(dse['Параметрика']))
                    dokumenty = ';'.join(dse['Документы'])
                    mat_kd = dse['Мат_кд']
                    ssil = dse['Ссылка']
                    prim = dse['Прим']
                    pki = dse['ПКИ']
    
                    for j, oper in enumerate(dse['Операции']):
                        oper_naim = oper['Опер_наименовние']
                        oper_nom = oper['Опер_номер']
                        oper_kod = oper['Опер_код']
                        oper_rc_naimenovnie = oper['Опер_РЦ_наименовние']
                        oper_rc_kod = oper['Опер_РЦ_код']
                        oper_oborud = oper['Опер_оборудование_наименовние']
                        oper_oborudovanie_kod = oper['Опер_оборудование_код']
                        oper_tpz = oper['Опер_Тпз']
                        oper_tst = oper['Опер_Тшт']
                        oper_prof = oper['Опер_профессия_наименование']
                        oper_professiya_kod = oper['Опер_профессия_код']
                        oper_kr = oper['Опер_КР']
                        oper_koid = oper['Опер_КОИД']
                        docs = '; '.join(dse['Документы']) + "; " + '; '.join(oper['Опер_документы'])
                        perehod = '; '.join(oper['Переходы'])
                        oper_instument = '; '.join(oper['Опер_инстумент'])
                        oper_osnastka = '; '.join(oper['Опер_оснастка'])
    
                        materialy = ':'
    
                        spis_shab_mk.append([nomerpp, naim, nn, kolich, uroven, oper_naim, oper_kod,
                                             oper_nom, oper_rc_naimenovnie, oper_rc_kod, oper_oborud,
                                             oper_oborudovanie_kod,
                                             oper_tpz, oper_tst, oper_prof, oper_professiya_kod, oper_kr, oper_koid,
                                             docs, oper_instument, oper_osnastka,
                                             materialy, perehod, parametrika, dokumenty, pki, mat, ssil, prim])
                        for mat_i in range(len(oper['Материалы'])):
                            spis_shab_mk.append(['', '', '', '', '', '', '',
                                                 '', '', '', '',
                                                 '',
                                                 '', '', '', '', '',
                                                 '', '', '', '',
                                                 oper['Материалы'][mat_i]['Мат_код'],
                                                 oper['Материалы'][mat_i]['Мат_наименование'],
                                                 oper['Материалы'][mat_i]['Мат_ед_изм'],
                                                 str(oper['Материалы'][mat_i]['Мат_норма']),
                                                 '@'.join(list(oper['Материалы'][mat_i]['Мат_параметрика'])), '', '',
                                                 ''])
    
                CEX.zap_spis(spis_shab_mk, dir, f'{nom_mk}_{py}.xlsx', 'Рес', 1, 1)
                CQT.msgbox('Готово')
                F.otkr_papky(dir)
        def from_abstract_mk(self,exel):
            if self.res == '':
                CQT.msgbox('Не создана ресурсная')
                return
            rez = self.res
            try:
                py = self.ui.comboBox_np.currentText().split("$")[1]
            except:
                CQT.migat_obj(self,1,self.ui.comboBox_np,f'Ошибка чтения ПУ')
                return
            generate(self, exel, rez, "ВО", py)
        def from_real_mk(self,exel):
            if self.ui.table_spis_MK.currentRow() == -1:
                CQT.msgbox(f'Не выбрана МК')
                return
            nom_mk = self.ui.table_spis_MK.item(self.ui.table_spis_MK.currentRow(),
                                                CQT.nom_kol_po_imen(self.ui.table_spis_MK, 'Пномер')).text()
            py = self.ui.table_spis_MK.item(self.ui.table_spis_MK.currentRow(),
                                            CQT.nom_kol_po_imen(self.ui.table_spis_MK, 'Номер_заказа')).text()
            rez = self.resursnaya_from_mk(nom_mk)
            generate(self,exel,rez,nom_mk,py)

                
        if self.ui.tabWidget.tabText(self.ui.tabWidget.currentIndex()) == 'Маршрутные карты':
            from_real_mk(self,exel)
        if self.ui.tabWidget.tabText(self.ui.tabWidget.currentIndex()) == 'Создание МК':
            if self.ui.tabWidget_2.tabText(self.ui.tabWidget_2.currentIndex()) == 'Разработка МК':
                from_abstract_mk(self,exel)
        return


    def resursnaya_from_cust_struktura(self, spis_dse, kol_vo_izdeliy=None, ruchnoi=False):
        rez_spis = []
        spis_dse = CMS.rasstanovka_dreva_kod_ruchnoi(self, spis_dse)
        nk_naim = F.nom_kol_po_im_v_shap(spis_dse, 'Наименование')
        nk_nn = F.nom_kol_po_im_v_shap(spis_dse, 'Обозначение')
        nk_kol = F.nom_kol_po_im_v_shap(spis_dse, 'Количество')
        nk_sumkol = F.nom_kol_po_im_v_shap(spis_dse, 'Сумм.Количество')
        nk_pki = F.nom_kol_po_im_v_shap(spis_dse, 'ПКИ')
        nk_mat = F.nom_kol_po_im_v_shap(spis_dse, 'Масса/М1,М2,М3')
        nk_ssil = F.nom_kol_po_im_v_shap(spis_dse, 'Ссылка')
        nk_prim = F.nom_kol_po_im_v_shap(spis_dse, 'Примечание')
        nk_urov = F.nom_kol_po_im_v_shap(spis_dse, 'Уровень')
        nk_dreva_kod = F.nom_kol_po_im_v_shap(spis_dse, 'dreva_kod')
        if nk_dreva_kod == None:
            nk_dreva_kod = 18
        nach = 1
        npp = 0

        conn1, cur1 = CSQ.connect_bd(self.db_dse)
        dict_dse = CSQ.zapros(self.db_dse,"""SELECT Номенклатурный_номер, Номер_техкарты FROM dse""",rez_dict=True,conn=conn1,cur=cur1)
        CSQ.close_bd(conn1,cur1)

        if dict_dse == False:
            CQT.msgbox(f'База занята, пробуй позже')
        dict_dse = F.raskrit_dict(dict_dse,'Номенклатурный_номер')
        for i in range(nach, len(spis_dse)):
            npp += 1
            nn = spis_dse[i][nk_nn].strip()
            naim = spis_dse[i][nk_naim].strip()
            pki = spis_dse[i][nk_pki].strip()
            mat = spis_dse[i][nk_mat].strip()
            ssil = spis_dse[i][nk_ssil].strip()
            prim = spis_dse[i][nk_prim].strip()
            dreva_kod = spis_dse[i][nk_dreva_kod]
            Способы_получения_материала = 'Произвести по основной спецификации'
            if pki == '1':
                Способы_получения_материала = 'Обеспечивать'

            if nn not in dict_dse:
                CQT.msgbox(f'Номер техкарты для {nn} {naim} не найден')
                return
            nom_tk = dict_dse[nn]
            if nom_tk == None or nom_tk == '' or nom_tk[0] == '':
                CQT.msgbox(f'Номер техкарты для {nn} {naim} не найден')
                return
            putf = F.scfg('add_docs') + os.sep + nom_tk + "_" + nn + '.pickle'
            uroven_dse = int(spis_dse[i][nk_urov])
            kolvo_koef = self.kol_v_uzel(spis_dse, i, nk_naim, nk_kol, nk_urov)
            kolvo_summ = kolvo_koef * int(spis_dse[i][nk_kol]) * kol_vo_izdeliy
            if F.nalich_file(putf):
                rez_tmp = self.dse_for_res(putf,kolvo_summ,npp, naim,nn,uroven_dse,pki,mat,ssil,prim,dreva_kod,Способы_получения_материала,int(spis_dse[i][nk_kol]))
                if rez_tmp == None:
                    return
            else:
                CQT.msgbox(f'Не найден файл {putf}')

                return
            rez_spis.append(rez_tmp)

        return rez_spis


    def dse_for_res(self,putf,kolvo_summ,npp, naim,nn,uroven_dse,pki,mat,ssil,prim, dreva_kod,Способы_получения_материала,kol_ed):

        rez_spis_op = []
        nk_rc_tk = 4
        nk_ur_tk = 20
        nk_op_tk = 2
        nk_op_oborud = 5
        nk_mat_tk = 10
        nk_doc_tk = 15
        nk_textper = 0
        nk_op_doc = 13
        nk_op_tpz = 6
        nk_op_tst = 7
        nk_op_prof = 8
        nk_op_KR = 9
        nk_op_KOID = 11

        nk_per_ima = 0
        nk_per_instr = 12
        nk_per_osn = 11
        nk_per_doc = 13
        sp_tk = F.otkr_f(putf, False, "|", pickl=True)
        try:
            tk_docs = sp_tk[10][nk_doc_tk].split('$')
        except:
            CQT.msgbox(f'Что-то не так в тк {putf.split(F.sep())[-1]} проверь ее!')
            return
        for j in range(11, len(sp_tk)):
            rez_spis_mat = []
            if sp_tk[j][nk_ur_tk] == '0':
                break
            if sp_tk[j][nk_ur_tk] == '1':
                if self.list_vars_vo != []:
                    sp_tk[j] = CVO.update_parametrs(self, sp_tk,j,nn)
                mat_str = sp_tk[j][nk_mat_tk].split('{')
                for k in range(len(mat_str)):
                    mat_str_str = mat_str[k].split('$')
                    if len(mat_str_str) == 4:
                        mat_cod = mat_str_str[0]
                        fl_add =False
                        if mat_cod in self.DICT_NOMEN:
                            filtr = ''
                            for item in self.DICT_FILTR_NOMEN:
                                if item['kod_oper'] == CMS.kod_oper_po_ima(self.SPIS_OP, sp_tk[j][nk_textper]) \
                                        and mat_cod == item['kod']:
                                    filtr = item['filtr']
                                    break
                            if filtr == '' or filtr == 0:
                                fl_add = True
                        if fl_add:
                            list_Материалы = self.DICT_NOMEN[mat_cod]
                            Материалы_Статья_калькуляции = 'Сырье'
                            if list_Материалы['Вид'] == 'Упаковочные материалы для складского хоз-ва 10.09':
                                Материалы_Статья_калькуляции = 'Упаковка'

                            rez_spis_mat.append({'Мат_код': mat_cod, "Мат_наименование": mat_str_str[1],
                                                 "Мат_ед_изм": mat_str_str[2],
                                                 "Мат_норма": round(F.valm(mat_str_str[3]) * kolvo_summ, 6),
                                                 "Мат_норма_ед": round(F.valm(mat_str_str[3]), 6),
                                                 'Мат_параметрика': dict(),
                                                'Материалы_Статья_калькуляции' : Материалы_Статья_калькуляции,
                                                 "Способы_получения_материала": 'Обеспечивать'
                                                 })
                    else:
                        pass
                rez_spis_instr = []
                rez_spis_osn = []
                rez_spis_doc = sp_tk[j][nk_op_doc].split('$')
                rez_spis_doc = F.clear_free_items(rez_spis_doc)
                spis_per = []
                for k in range(j + 1, len(sp_tk)):
                    if sp_tk[k][nk_ur_tk] == '1' or sp_tk[k][nk_ur_tk] == '0':
                        break
                    spis_per.append(sp_tk[k][nk_per_ima])
                    spis_instr = sp_tk[k][nk_per_instr].split('$')
                    for item in spis_instr:
                        rez_spis_instr.append(item)
                    spis_osn = sp_tk[k][nk_per_osn].split('$')
                    for item in spis_osn:
                        rez_spis_osn.append(item)
                    spis_doc = sp_tk[k][nk_per_doc].split('$')
                    for item in spis_doc:
                        rez_spis_doc.append(item)
                if CMS.ima_rc_po_nom(self.SPIS_RC, sp_tk[j][nk_rc_tk]) == None:
                    CQT.msgbox(f'{nn} операция {sp_tk[j][nk_op_tk]} не найден РЦ')
                if sp_tk[j][nk_textper] not in self.DICT_ETAPI:
                    CQT.msgbox(f'{nn} операция {sp_tk[j][nk_op_tk]} не найден "Этап"')
                    etap = 'Неопознан'
                else:
                    etap = self.DICT_ETAPI[sp_tk[j][nk_textper]]
                kod_oper = CMS.kod_oper_po_ima(self.SPIS_OP, sp_tk[j][nk_textper])
                rez_spis_op.append({"Этап": etap,
                                    "Опер_наименовние": sp_tk[j][nk_textper],
                                    "Опер_код": kod_oper,
                                    "Опер_вспомогательная": self.DICT_OP[kod_oper]['Вспомогат'],
                                    "Опер_номер": sp_tk[j][nk_op_tk],
                                    "Опер_РЦ_наименовние": CMS.ima_rc_po_nom(self.SPIS_RC, sp_tk[j][nk_rc_tk]),
                                    "Опер_РЦ_код": sp_tk[j][nk_rc_tk],
                                    "Опер_оборудование_наименовние": sp_tk[j][nk_op_oborud],
                                    "Опер_оборудование_код": CMS.kod_oborud_po_ima(self.SPIS_OB,
                                                                                   sp_tk[j][nk_op_oborud]),
                                    "Опер_Тпз": round(F.valm(sp_tk[j][nk_op_tpz]), 6),
                                    "Опер_Тшт": round(F.valm(sp_tk[j][nk_op_tst]) * kolvo_summ, 6),
                                    "Опер_Тшт_ед": round(F.valm(sp_tk[j][nk_op_tst]), 6),
                                    "Опер_профессия_наименование": CMS.ima_prof_po_kod(self.SPIS_PROF,
                                                                                       sp_tk[j][nk_op_prof]),
                                    "Опер_профессия_код": sp_tk[j][nk_op_prof],
                                    "Опер_КР": F.valm(sp_tk[j][nk_op_KR]),
                                    "Опер_КОИД": F.valm(sp_tk[j][nk_op_KOID]), "Опер_документы": rez_spis_doc,
                                    "Опер_инстумент": rez_spis_instr, "Опер_оснастка": rez_spis_osn,
                                    "Материалы": rez_spis_mat, "Переходы": spis_per})

        rez_tmp = {'Номерпп': npp, 'Наименование': naim, 'Номенклатурный_номер': nn, 'Количество': kolvo_summ, 'Количество_ед': kol_ed,
                   'Уровень': uroven_dse, "Операции": rez_spis_op, 'Параметрика': dict(), 'Документы': tk_docs,
                   'ПКИ': pki, 'Мат_кд': mat, 'Ссылка': ssil, 'Прим': prim,
                                    "dreva_kod" : dreva_kod, "Способы_получения_материала" : Способы_получения_материала}
        return rez_tmp

    def resursnaya_from_mk(self, nom_mk):
        if nom_mk == None:
           CQT.msgbox(f'Номер мк не указан')
        query = f'''SELECT data, Head FROM xml 
            WHERE Номер_мк == {int(nom_mk)}
                        '''
        rez_xml = CSQ.zapros(self.db_resxml, query)
        xml = rez_xml[-1][0]
        xml_head = rez_xml[-1][1]
        if xml == '':
            CQT.msgbox('Нет хмл файла')
            return

        query = f'''SELECT Количество FROM mk
                    WHERE Пномер == {int(nom_mk)}
                                '''
        kol_vo_izdeliy = CSQ.zapros(self.bd_naryad, query)[-1][0]

        rez_spis = CMS.resursnaya_from_xml(self,self.podgotovka_xml(XML.spisok_iz_xml(str_f=xml),xml_head),kol_vo_izdeliy,nom_mk)
        return rez_spis


    def spisok_tek_tehkarta(self, spis_sod_tk):
        nach = 0
        kon = len(spis_sod_tk) - 1
        for i in range(len(spis_sod_tk)):
            if len(spis_sod_tk[i]) == 21 and spis_sod_tk[i][-1] == '0':
                if nach == 0:
                    nach = i
                else:
                    kon = i - 1
        return spis_sod_tk[nach:kon + 1]

    def open_zayavk(self):
        nom_pr = ''
        nom_pu = ''
        if self.ui.tabWidget.currentIndex() == CQT.nom_tab_po_imen(self.ui.tabWidget,'Создание МК'):
            if self.ui.comboBox_np.currentText() == "" or '$' not in self.ui.comboBox_np.currentText():
                CQT.msgbox('Не выбран номер проекта')
                return
            nom_pr = self.ui.comboBox_np.currentText().split('$')[0]
            nom_pu = self.ui.comboBox_np.currentText().split('$')[1]
        if self.ui.tabWidget.currentIndex() == CQT.nom_tab_po_imen(self.ui.tabWidget,'Маршрутные карты'):
            if self.ui.table_spis_MK.currentRow() == None or self.ui.table_spis_MK.currentRow() == -1:
                CQT.msgbox('Не выбран номер проекта')
                return
            nk_pr = CQT.nom_kol_po_imen(self.ui.table_spis_MK,'Номер_проекта')
            nk_pu = CQT.nom_kol_po_imen(self.ui.table_spis_MK, 'Номер_заказа')
            nom_pr = self.ui.table_spis_MK.item(self.ui.table_spis_MK.currentRow(),nk_pr).text()
            nom_pu = self.ui.table_spis_MK.item(self.ui.table_spis_MK.currentRow(),nk_pu).text()
        if self.ui.tabWidget.currentIndex() == CQT.nom_tab_po_imen(self.ui.tabWidget, 'Объемно-календарное планирование'):
            nk_nom_py = CQT.nom_kol_po_imen(self.ui.tbl_kal_pl,'пл_оуп.№ERP')
            if nk_nom_py == None:
                CQT.msgbox(f'Не найдено поле пл_оуп.№ERP')
                return
            nk_nom_pr = CQT.nom_kol_po_imen(self.ui.tbl_kal_pl, 'пл_оуп.№проекта')
            if nk_nom_pr == None:
                CQT.msgbox(f'Не найдено поле пл_оуп.№№проекта')
                return
            nom_pr = self.ui.tbl_kal_pl.item(self.ui.tbl_kal_pl.currentRow(), nk_nom_pr).text()
            nom_pu = self.ui.tbl_kal_pl.item(self.ui.tbl_kal_pl.currentRow(), nk_nom_py).text()
        if nom_pu == '':
            return
        putf = CMS.Put_k_papke_s_proektom_NPPU(nom_pr, nom_pu,True)
        modifiers = CQT.get_key_modifiers(self)
        if modifiers == ['shift']:
            F.otkr_papky(putf)
            return
        list_files = F.spis_files(putf)
        if list_files == []:
            return
        for file in list_files[0][2]:
            if 'Заказ на производство №' in file and '.pdf' in file:
                F.zapyst_file(list_files[0][0] + file)
                return
        F.otkr_papky(list_files[0][0])


    def vigruzka_norm_vr_po_spis_xml(self, spis_xml, razrabotka=0):
        n_k_nn = 1
        n_k_naim = 0
        n_k_kol = 7
        if razrabotka == True:
            n_k_kol = 10
        n_k_kol_bez_summ = 2
        n_k_ves = 4
        ves = 0
        rez = []
        ima_sbor = spis_xml[0][0] + '$' + spis_xml[0][1]
        dict_dse = CSQ.zapros(self.db_dse, """SELECT Номенклатурный_номер, Номер_техкарты FROM dse""",rez_dict=True)
        dict_dse = F.raskrit_dict(dict_dse,'Номенклатурный_номер')
        for i in range(len(spis_xml)):
            ves_det = 0
            chislo_det = 0
            if spis_xml[i][n_k_ves].split('/')[1] != '' and spis_xml[i][n_k_ves].split('/')[2] != '':
                ves_det = F.valm(spis_xml[i][n_k_ves].split('/')[0])
                chislo_det = self.kol_v_uzel(spis_xml, i, 1, 2)  # F.valm(spis_xml[i][n_k_kol])
                ves += ves_det * chislo_det
                print(spis_xml[i][n_k_nn], ves_det, chislo_det)
            if spis_xml[i][n_k_nn] not in dict_dse:
                CQT.msgbox(f'Не найден номер техкарты на {spis_xml[i][n_k_nn]} {spis_xml[i][n_k_naim]}')
                return
            nom_t_k = dict_dse[spis_xml[i][n_k_nn]]
            if F.nalich_file(F.scfg('add_docs') + os.sep + nom_t_k + "_" + spis_xml[i][n_k_nn] + '.pickle') == False:
                CQT.msgbox(f'Не найдена техкарта {nom_t_k}')
                return
            spis_sod_tk = F.otkr_f(F.scfg('add_docs') + os.sep + nom_t_k + "_" + spis_xml[i][n_k_nn] + '.txt',
                                   separ='|', pickl=True)
            tek_karta = self.spisok_tek_tehkarta(spis_sod_tk)
            for j in range(len(tek_karta)):
                if tek_karta[j][-1] == '1':
                    if F.is_numeric(tek_karta[j][7]) == False or tek_karta[j][7] == '':
                        CQT.msgbox(
                            f'{spis_xml[i][n_k_naim] + "$" + spis_xml[i][n_k_nn]} операция {tek_karta[j][2]} не отнормирована')
                        return
                    rez.append([ima_sbor, spis_xml[i][n_k_naim] + "$" + spis_xml[i][n_k_nn], spis_xml[i][n_k_kol],
                                tek_karta[j][0],
                                tek_karta[j][2],
                                tek_karta[j][4],
                                tek_karta[j][6],
                                tek_karta[j][7],
                                tek_karta[j][8],
                                ves_det, chislo_det])
                    ves_det = 0
                    chislo_det = 0
        return [rez, ves]

    def vigruzka_norm_mat_po_spis_xml(self, xml, rez, kol_izd, ruchnoiy=0):#отключена
        nk_naim = 0
        nk_nn = 1
        nk_sumkol = 7
        if ruchnoiy == True:
            nk_sumkol = 10
        err_arr = []
        conn1, cur1 = CSQ.connect_bd(self.db_dse)
        dict_dse = CSQ.zapros(self.bd_naryad, """SELECT Номенклатурный_номер, Номер_техкарты FROM dse""", rez_dict=True, conn=conn1,cur=cur1)
        CSQ.close_bd(conn1,cur1)
        dict_dse = F.raskrit_dict(dict_dse, 'Номенклатурный_номер')
        for i in range(1, len(xml)):
            nn = xml[i]['data']['Обозначение'].strip()
            naim = xml[i]['data']['Наименование'].strip()
            kolvo = int(xml[i][nk_sumkol])  # * kol_izd
            if nn not in dict_dse:
                CQT.msgbox(f'{nn} в дсе не найдена')
            nom_tk = dict_dse[nn]
            putf = F.scfg('add_docs') + os.sep + nom_tk + '_' + nn + '.pickle'
            if F.nalich_file(putf):
                nk_rc_tk = 4
                nk_ur_tk = 20
                nk_op_tk = 2
                nk_mat_tk = 10
                nk_doc_tk = 15
                nk_textper = 0
                sp_tk = F.otkr_f(putf, False, "|", pickl=True)
                for j in range(11, len(sp_tk)):
                    if sp_tk[j][nk_ur_tk] == '0':
                        break
                    if sp_tk[j][nk_ur_tk] == '1':
                        mat_str = sp_tk[j][nk_mat_tk].split('{')
                        for k in range(len(mat_str)):
                            if sp_tk[j][nk_rc_tk] == '':
                                CQT.msgbox(
                                    f'В техкарте {sp_tk[0][0]} {sp_tk[1][0]} не корректно занесен РЦ на {sp_tk[j][nk_op_tk]} операцию')
                                return
                            if mat_str[k] != '':
                                rez = self.add_mat_v_sp(rez, mat_str[k] + '$' + sp_tk[j][nk_rc_tk], kolvo)
                                if rez == False:
                                    CQT.msgbox(
                                        f'В техкарте {sp_tk[0][0]} {sp_tk[1][0]} не корректно занесен материал на {sp_tk[j][nk_op_tk]} операцию')
                                    return
            else:
                CQT.msgbox(f'Не найдена техкарта {putf}')

                return

        return rez

    def add_mat_v_sp(self, spis: list, list_mat, kolvo):
        if list_mat == '':
            return spis
        list_mat = list_mat.split('$')
        if len(list_mat) < 5:
            return False
        list_mat[3] = F.valm(list_mat[3]) * int(kolvo)
        flag = False
        for i in range(len(spis)):
            if spis[i][0] == list_mat[0] and spis[i][4] == list_mat[4]:
                spis[i][3] += list_mat[3]
                flag = True
                break
        if flag == False:
            spis.append(list_mat)
        return spis

    def vigruzka_norm_mat(self):
        self.w2 = CVO.mywindow2(self)


    def vigruzka_norm(self):
        if self.ui.tabWidget_2.tabText(self.ui.tabWidget_2.currentIndex()) == 'Разработка МК':
            if self.spis_poziciy_rez_ruchnoi == []:
                CQT.msgbox(f'Не сформирована МК')
                return

        nom_pr = self.ui.comboBox_np.currentText().split('$')[0]
        tmp_putt = CMS.load_tmp_path('table_normi')


        putt = CQT.f_dialog_save(self, 'Сохранить нормы', tmp_putt[0] + F.sep() + nom_pr + '_Нормы_поэтапно',
                                 "Text files (*.txt)")
        if putt == '' or putt == '.':
            return
        CMS.save_tmp_path('table_normi',putt)
        #puttres = F.sep().join(putt.split(F.sep())[:-1]) + F.sep() + nom_pr + '_Ресурсная.pickle'
        spis_poziciy_rez = self.vigruzka_norm_resusrnaya()
        if spis_poziciy_rez == None:
            return
        self.vigruzka_norm_exel(putt.replace('.txt','.xlsx'), spis_poziciy_rez)
        self.export_spis_etapov_s_vesom_for_plan_exel(putt, spis_poziciy_rez)
        return



    def vigruzka_norm_resusrnaya(self):
        if self.ui.tabWidget_2.tabText(self.ui.tabWidget_2.currentIndex()) == 'Создание МК из *.XML':
            tbl = self.ui.table_zayavk
            if tbl.rowCount() == 0:
                CQT.msgbox('Не заполнены позиции')
                return
            spis_poziciy = CQT.spisok_iz_wtabl(tbl, '', True)
            if len(spis_poziciy) == 1:
                CQT.msgbox('Не заполнены позиции')
                return
            n_k_file = F.nom_kol_po_im_v_shap(spis_poziciy, 'Файл')
            n_k_count = F.nom_kol_po_im_v_shap(spis_poziciy, 'Количество')
            if n_k_file == None:
                CQT.msgbox('Не найден путь')
                return
            spis_poziciy_rez = []
            for i in range(1, len(spis_poziciy)):
                dict_poziciy_rez = dict()
                if spis_poziciy[i][n_k_count] == '':
                    CQT.msgbox('Не указано количество')
                    return
                if F.is_numeric(spis_poziciy[i][n_k_count]) == False:
                    CQT.msgbox('Количество не число')
                    return
                if F.nalich_file(spis_poziciy[i][n_k_file]) == False:
                    CQT.msgbox(f'Не найден XML {spis_poziciy[i][n_k_file]}')
                    return
                else:
                    spis_det_xml = self.podgotovka_xml(XML.spisok_iz_xml(spis_poziciy[i][n_k_file]))
                    kol_vo_izdeliy = int(spis_poziciy[i][n_k_count])
                    rez = CMS.resursnaya_from_xml(self, spis_det_xml, kol_vo_izdeliy)
                    if rez == None:
                        return
                    dict_poziciy_rez['name'] = f"{rez[0]['Наименование']} {rez[0]['Номенклатурный_номер']}"
                    dict_poziciy_rez['data'] = rez
                    dict_poziciy_rez['kol_zayavk'] = kol_vo_izdeliy
                    spis_poziciy_rez.append(dict_poziciy_rez)
            # F.save_file_pickle(putt, dict_poziciy_rez)
            return spis_poziciy_rez
        else:
            return self.spis_poziciy_rez_ruchnoi

    def vigruzka_norm_exel(self, putt, spis_poziciy_rez):
        if self.ui.comboBox_np.currentText() == "" or '$' not in self.ui.comboBox_np.currentText():
            CQT.msgbox('Не выбран номер проекта')
            return

        putp = os.sep.join(putt.split(os.sep)[:-1])
        file = putt.split(os.sep)[-1]
        if '.xlsx' not in file:
            CQT.msgbox('Файл должен быть *.xlsx')
            return

        summ_ves = 0
        rez = [['Позиция', 'ДСЕ', 'Количество', 'Операция', 'Номер', 'РЦ', 'Тпз', 'Тшт', 'Профессия', 'По заявке',
                'Тшт*Кол*Заяв', 'Вес',
                'Кол-во для веса']]
        for item in spis_poziciy_rez:
            resyrsnaya = item['data']
            poz = item['name']
            kol_vo_izdeliy = item['kol_zayavk']
            for dse in resyrsnaya:
                dse_name = f"{dse['Наименование']} {dse['Номенклатурный_номер']}"
                kol_vo_dse = dse['Количество'] / kol_vo_izdeliy
                ves_spis = F.valm(dse['Мат_кд'].split('/'))
                ves = 0
                if ves_spis[1] != '' and ves_spis[2] != '':
                    ves = F.valm(ves_spis[0])
                ves_kol_vo = dse['Количество']
                for oper in dse['Операции']:
                    oper_name = oper['Опер_наименовние']
                    oper_nom = oper['Опер_номер']
                    oper_rc = oper['Опер_РЦ_код']
                    oper_tpz = oper['Опер_Тпз']

                    oper_prof = oper['Опер_профессия_наименование']

                    oper_tsht_ed = oper['Опер_Тшт_ед']
                    oper_tsht  = oper['Опер_Тшт']
                    tsht_kol_zayv = oper_tsht + oper_tpz
                    rez.append([poz, dse_name, kol_vo_dse, oper_name, oper_nom, oper_rc, oper_tpz, oper_tsht_ed, oper_prof,
                                kol_vo_izdeliy,
                                tsht_kol_zayv, ves,
                                ves_kol_vo])
                    summ_ves += ves * ves_kol_vo
                    ves = 0
                    ves_kol_vo = 0
        for i in range(4):
            rez.append(['' for x in range(len(rez[0]))])
        rez[-1][3] = 'масса'
        rez[-1][4] = round(summ_ves, 1)

        file = F.ochist_strok_pod_ima_fila(file)
        CEX.zap_spis(rez, putp, file, 'Нормы времени', 0, 0)
        if F.nalich_file(putp + F.sep() + file):
            rez = CQT.msgboxgYN('эксель успешно сохранен. Запустить?')
            if rez == True:
                F.zapyst_file(putp + F.sep() + file)
        else:
            CQT.msgbox('Файл эксель не сохранен, что то пошло не так')

    def obn_spis_pr(self):
        combo_proekt = self.ui.comboBox_np
        combo_proekt.clear()
        spis = F.otkr_f(F.tcfg('BD_Proect'), False, '|', False, False)
        rez = []
        for i in range(10, len(spis)):
            if len(spis[i]) >= 21:
                if spis[i][3] == 'к производству' or spis[i][3] == 'подготовка':
                    rez.append(spis[i])
        rez.sort()
        if self.ui.tabWidget_2.tabText(self.ui.tabWidget_2.currentIndex()) == 'Разработка МК':
            combo_proekt.addItems(["$".join(['ВО', '-', "-", "КТ", "1"]),
                   "$".join(['ВО', '-', "-", "КЛ", "1"]),
                   "$".join(['ВО', '-', "-", "ШГ", "1"]),
                   "$".join(['ВО', '-', "-", "ПР", "1"])])
        for i in range(len(rez)):
            combo_proekt.addItem("$".join([rez[i][0], rez[i][1], rez[i][2], rez[i][22], rez[i][8]]))
        self.vibor_pr()

    def zapusk_docs(self, strok, kol):
        tbl = self.ui.table_nomenkl
        kol_naim = CQT.nom_kol_po_imen(tbl, 'Наименование')
        kol_nn = CQT.nom_kol_po_imen(tbl, 'Номенклатурный_номер')
        kol_pn = CQT.nom_kol_po_imen(tbl, 'Пномер')
        if kol == kol_pn:
            nn_det = tbl.item(strok, kol_nn).text()
            naim = tbl.item(strok, kol_naim).text()
            CMS.zapustit_ssicy_docs(nn_det, naim)
            # F.zapyst_file(adres,False)

    def poisk_nn(self):
        nn = self.ui.lineEdit_nom_n.text()
        naim = self.ui.lineEdit_naim.text()
        prim = self.ui.lineEdit_primech.text()
        tab_dse = self.ui.table_nomenkl
        for i in range(0, tab_dse.model().rowCount()):
            if nn in CQT.cells(i, 0, tab_dse) and naim in CQT.cells(i, 1, tab_dse) and prim in CQT.cells(i, 3, tab_dse):
                tab_dse.selectRow(i)
                return
        tab_dse.clearSelection()

    def otmizm_nom(self):
        self.zapoln_tabl_nomenkl()

    def del_nom(self):
        if self.tabl_nomenk.rowCount() < 1:
            return
        if self.tabl_nomenk.currentRow() == -1:
            CQT.msgbox('Не выбрана номенклатура')
            return
        rez = CQT.msgboxgYN(f'Удалить строку для {self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 1).text()}?')
        if rez == False:
            return
        if self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 3).text() == "":
            CSQ.zapros(self.db_dse, f"""DELETE FROM dse where Пномер = {int(self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 0).text())} """)
        else:
            CQT.msgbox(
                f'На {self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 1).text()} уже создана техкарта, удаление не возможно.')
            return
        self.zapoln_tabl_nomenkl()
        CQT.msgbox(f'Строка успешно удалена')

    def cvet_izm_nom(self):
        CQT.ust_cvet_obj(self.ui.btn_korr_nom, 155, 253, 155)

    def btn_korr_nom(self):
        if self.tabl_nomenk.rowCount() < 1:
            return
        if self.tabl_nomenk.currentRow() == -1:
            CQT.msgbox('Не выбрана номенклатура')
            return
        rez = CQT.msgboxgYN(
            f'Откорректировать строку для {self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 1).text()}?')
        if rez == False:
            return

        if self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 3).text() == "":


            CSQ.zapros(self.db_dse,f"""UPDATE dse SET Номенклатурный_номер = ?, Наименование = ?, 
            Примечание, = ? WHERE Пномер == {int(self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 0).text())};""",
                       (self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 1).text(),
                       self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 2).text(),
                       self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 4).text().replace("|", "-"))
                        )

        else:
            rez = CQT.msgboxgYN(
                f'Откорректировать <Номенклатурный_номер> невозможно, т.к. техкарта уже создана.'
                f' Внести корректировку в <Наименование> и <Примечание> ?')
            if rez == False:
                return

            CSQ.zapros(self.db_dse, f"""UPDATE dse SET Наименование = ?, 
                    Примечание, = ? WHERE Пномер == {int(self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 0).text())};""",
                   (self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 2).text(),
                   self.tabl_nomenk.item(self.tabl_nomenk.currentRow(), 4).text().replace("|", "-")))
        self.zapoln_tabl_nomenkl()
        # CQT.ust_cvet_obj(self.ui.btn_korr_nom)
        CQT.msgbox(f'Изменения успешно записаны')

    def add_v_nomenkl(self):
        le_nn = self.ui.lineEdit_nom_n
        le_naim = self.ui.lineEdit_naim
        le_prim = self.ui.lineEdit_primech
        if le_naim.text() == "":
            le_naim.setFocus()
            CQT.msgbox("Не указано наименование")
            return
        nn = F.ochist_strok_pod_ima_fila(le_nn.text())
        naim = F.ochist_strok_pod_ima_fila(le_naim.text())
        #rez = CSQ.naiti_v_bd(self.db_dse, 'dse', {'Номенклатурный_номер': nn, 'Наименование': naim})
        rez = CSQ.zapros(self.db_dse,f"""SELECT Номенклатурный_номер FROM dse WHERE Номенклатурный_номер == '{nn}' and Наименование == '{naim}'; """)
        if len(rez) > 1:
            CQT.msgbox(f'ДСЕ {nn} {naim} уже существует')
            return
        putf = CMS.Put_k_papke_s_proektom_NPPU('.'.join(nn.split('.')[:-1]), naim)
        list_add = [nn, naim, '', le_prim.text(), putf, '', '', '', '', '', '', '', '', ""]
        #CSQ.dob_strok_v_bd_sql(self.db_dse, 'dse',
        #                       [list_add])

        CSQ.zapros(self.db_dse, f"""INSERT INTO dse (Номенклатурный_номер, Наименование, 
        Номер_техкарты, Примечание, Путь_docs, Доступ, Процесс, Нормы, Материалы, Тех_заметки, Теги, Мат_кд,
         Код_ЕРП, Классификатор) VALUES ({', '.join(['?']*len(list_add))});""", spisok_spiskov=[list_add] )

        self.zapoln_tabl_nomenkl()
        self.ui.lineEdit_primech.setText(' шт.')
        CQT.msgbox(f'ДСЕ {nn} {naim} успешно занесена')

    def spis_MK_clck(self):
        param = self.tabl_mk.currentRow()
        if self.tabl_mk.item(param, CQT.nom_kol_po_imen(self.tabl_mk, "Статус")).text() == "Открыта":
            self.ui.pushButton_close_mk.setEnabled(True)
            self.ui.pushButton_open_mk.setEnabled(False)
        else:
            self.ui.pushButton_close_mk.setEnabled(False)
            self.ui.pushButton_open_mk.setEnabled(True)

    def tabl_brak_dbl_clk(self):
        return
        label = self.ui.label_opis_braka
        strok = self.tabl_brak.currentRow()
        label_brak = self.ui.label_opis_braka
        n_k = 4
        if F.nalich_file(F.scfg('foto_brak')) == False:
            CQT.msgbox(f'Недоступна папка {F.scfg("foto_brak")}')
            return
        if self.tabl_brak.currentColumn() == n_k:
            if self.tabl_brak.item(self.tabl_brak.currentRow(), n_k).text().replace('Фото:', '') != "":
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
        kol_det = self.tabl_brak.item(strok, 8).text().replace('Количество:', '')

        nom_mk_isprav = self.naiti_parametr_v_stroke(label_brak.text(), 'Изгот.вновь по МК:')
        if nom_mk_isprav != '' and nom_mk_isprav != 'None':
            CQT.msgbox(f'ДСЕ уже изготавливается по МК №{nom_mk_isprav}')
            return

        if nom_id.startswith("0x"):
            pass
        else:
            self.ui.tabWidget.setCurrentIndex(1)
            self.ui.tabWidget_2.setCurrentIndex(1)
            self.add_gl_uzel(nom_id)
            return

        tree = self.ui.treeWidget
        spis_tree = CQT.spisok_dreva(tree)
        if spis_tree == []:
            CQT.msgbox("Не открыто древо")
            self.viborXML()
            self.ui.tabWidget.setCurrentIndex(3)
            self.tabl_brak.selectRow(strok)
            self.tabl_brak_dbl_clk()
            return
        nom_kol_id = CQT.nom_kol_po_imen(tree, 'ID')
        for i in range(len(spis_tree)):
            if spis_tree[i][nom_kol_id] == nom_id:
                uroven = spis_tree[i][20]
                rez = CQT.videlit_tree_nom(tree, i + 1)
                if rez == False:
                    CQT.msgbox(f'Деталь {spis_tree[i][20]} не найдена')
                    return
                self.add_v_mk()
                table = self.ui.table_razr_MK
                for j in range(table.rowCount()):
                    if table.item(j, 6).text() == nom_id:
                        table.item(j, CQT.nom_kol_po_imen(table, 'Кол. по заявке')).setText(kol_det)
                        break
                # table.setCurrentCell(table.rowCount() - 1, 1)
                # for j in range(i + 1, len(spis_tree)):
                #    if spis_tree[j][20] > uroven:
                #        rez = CQT.videlit_tree_nom(tree, i + 1)
                #        if rez == False:
                #            CQT.msgbox(f'Деталь {spis_tree[j][20]} не найдена')
                #            self.ui.tabWidget.setCurrentIndex(0)
                #            return
                #        self.add_v_mk()
                #    else:
                #        break
                CQT.msgbox(
                    'В случе если основание создания маршрутной карты это БРАК, то необходимо, посел создания ассоциировать брак с МК. во вкладке брак, а после сохранить.')
                self.ui.pushButton_save_MK.setEnabled(False)
                return

    def click_brak(self):
        return

    def del_ass(self):
        self.ui.label_ass.clear()

    def close_mk(self):
        if self.tabl_mk.currentRow() == -1:
            return
        tbl = self.tabl_mk
        if not CMS.user_access(self.bd_naryad,'мкарт_маршрутные_закрытьоткрыть',F.user_name()):
            return
        row = self.tabl_mk.currentRow()
        nom_tek_mk = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Пномер')).text()
        nom_mk = nom_tek_mk
        project = tbl.item(tbl.currentRow(),
                           CQT.nom_kol_po_imen(tbl, 'Номенклатура')).text()
        nom_pu_r = tbl.item(tbl.currentRow(),
                            CQT.nom_kol_po_imen(tbl, 'Номер_заказа')).text()
        nom_pr_r = tbl.item(tbl.currentRow(),
                            CQT.nom_kol_po_imen(tbl, 'Номер_проекта')).text()
        kolvo = tbl.item(tbl.currentRow(), CQT.nom_kol_po_imen(tbl, 'Количество')).text()
        prim = tbl.item(tbl.currentRow(), CQT.nom_kol_po_imen(tbl, 'Примечание')).text()
        osnovanie = tbl.item(tbl.currentRow(),
                             CQT.nom_kol_po_imen(tbl, 'Основание')).text()
        if not self.check_zaversheni_naruady([int(nom_tek_mk)]):
            return
        #qery = CSQ.naiti_v_bd(self.bd_naryad, 'mk', {'Пномер': int(nom_tek_mk)},
        #                      ['Статус', 'Дата_завершения', 'Основание'])
        qery = CSQ.zapros(self.bd_naryad, f"""SELECT Статус, Дата_завершения, Основание FROM mk 
        WHERE Пномер == {int(nom_tek_mk)};""", shapka = False)

        if qery[0][0] == "Открыта":
            #rez = CSQ.update_bd_sql(self.bd_naryad, 'mk', {'Статус': 'Закрыта'}, {'Пномер': int(nom_tek_mk)})
            rez = CSQ.zapros(self.bd_naryad, f"""UPDATE mk SET Статус == 'Закрыта' WHERE Пномер = {int(nom_tek_mk)};""")

            if rez == False:
                CQT.msgbox('Не удалось записать')
                return
            self.tab_click(row)

            try:
                msg = f"{F.user_full_namre()} ЗАКРЫЛ мк № {str(nom_mk)}:\n{project} - {str(kolvo)} шт.\n{nom_pu_r.strip()} Проект: {nom_pr_r.strip()}\n" \
                      f"Прим.: {prim} {osnovanie}"
                self.send_info_mk_b24(msg, 'chat41228')
            except:
                print('Ошибка отправки в Б24')
        self.ui.pushButton_open_mk.setEnabled(True)
        self.ui.pushButton_close_mk.setEnabled(False)


    def spis_op_po_mk_id_op(self, sp_tabl_mk, id, op):
        for j in range(1, len(sp_tabl_mk)):
            if sp_tabl_mk[j][6] == id:
                for i in range(11, len(sp_tabl_mk[0]), 4):
                    if sp_tabl_mk[j][i].strip() != '':
                        obr = sp_tabl_mk[j][i].strip().split('$')
                        obr2 = obr[-1].split(";")
                        if op in obr2:
                            return obr2
                return None

    def del_mk(self):
        if self.tabl_mk.currentRow() == -1:
            return
        if CMS.user_access(self.bd_naryad,'созданиемаршрутныхкарт_удалить', F.user_name()) == False:
            return
        nom_mk = self.tabl_mk.item(self.tabl_mk.currentRow(), 0).text()
        project = self.tabl_mk.item(self.tabl_mk.currentRow(), CQT.nom_kol_po_imen(self.tabl_mk,'Номенклатура')).text()
        nom_pu_r = self.tabl_mk.item(self.tabl_mk.currentRow(), CQT.nom_kol_po_imen(self.tabl_mk,'Номер_заказа')).text()
        nom_pr_r = self.tabl_mk.item(self.tabl_mk.currentRow(), CQT.nom_kol_po_imen(self.tabl_mk,'Номер_проекта')).text()
        kolvo = self.tabl_mk.item(self.tabl_mk.currentRow(), CQT.nom_kol_po_imen(self.tabl_mk,'Количество')).text()
        prim =self.tabl_mk.item(self.tabl_mk.currentRow(), CQT.nom_kol_po_imen(self.tabl_mk,'Примечание')).text()
        osnovanie =self.tabl_mk.item(self.tabl_mk.currentRow(), CQT.nom_kol_po_imen(self.tabl_mk,'Основание')).text()
        progress = self.tabl_mk.item(self.tabl_mk.currentRow(), CQT.nom_kol_po_imen(self.tabl_mk, 'Прогресс')).text()
        if progress != '':
            CQT.msgbox('Нельзя удалить начатую МК')
            return
        otv = CQT.msgboxgYN('Точно удалить полность маршрутную карту №'
                             + nom_mk + '?')
        if otv:
            row = self.tabl_mk.currentRow()
            nom_tek_mk = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Пномер')).text()
            #rez = CSQ.delete(self.bd_naryad, 'mk', {'Пномер': int(nom_tek_mk)})
            rez = CSQ.zapros(self.bd_naryad, f"""DELETE FROM mk where Пномер = {int(nom_tek_mk)}""")
            if rez == False:
                CQT.msgbox('Запрос не выполнен')
                return
            rez = CSQ.zapros(self.bd_naryad, f"""DELETE FROM zagot where Ном_МК = {int(nom_tek_mk)}""")
            rez = CSQ.zapros(self.db_resxml, f"""DELETE FROM res where Номер_мк = {int(nom_tek_mk)}""")
            rez = CSQ.zapros(self.db_resxml, f"""DELETE FROM xml where Номер_мк = {int(nom_tek_mk)}""")
            rez = CSQ.zapros(self.bd_files, f"""DELETE FROM MK_founding where Num_mk = {int(nom_tek_mk)}""")
            try:
                rez = CSQ.zapros(self.bd_naryad, f"""DELETE FROM дорезки_мк where Номер_мк = {int(nom_tek_mk)}""")
            except:
                pass
            self.tab_click(row)
            if F.nalich_file(F.scfg('mk_data') + os.sep + nom_mk + '.txt') == True:
                F.udal_file(F.scfg('mk_data') + os.sep + nom_mk + '.txt')
            if F.nalich_file(F.scfg('mk_data') + os.sep + nom_mk):
                F.udal_papky(F.scfg('mk_data') + os.sep + nom_mk)
            CQT.msgbox(f"Маршрутная карта номер {nom_mk} удалена успешно")
            try:
                msg = f"{F.user_full_namre()} !УДАЛИЛ мк № {str(nom_mk)}:\n{project} - {str(kolvo)} шт.\n{nom_pu_r.strip()} Проект: {nom_pr_r.strip()}\n" \
                      f"Прим.: {prim} {osnovanie}"
                self.send_info_mk_b24(msg, 'chat41228')
            except:
                print('Ошибка отправки в Б24')

    def add_res_to_mk(self, xml, kol_vo_izdeliy, nom_tek_mk,xml_head, conn = '', cur = ''):
        self.res = CMS.resursnaya_from_xml(self, self.podgotovka_xml(XML.spisok_iz_xml(str_f=xml),xml_head), kol_vo_izdeliy, conn=conn,cur=cur)
        #rez = CSQ.update_bd_sql(self.bd_naryad, 'res',
        #                        {'data': F.to_binary_pickle(self.res)},
        #                        {'Номер_мк': int(nom_tek_mk)},conn=conn,cur=cur)
        rez = CSQ.zapros(self.db_resxml, f"""UPDATE res SET data = {F.to_binary_pickle(self.res)} WHERE Номер_мк = {int(nom_tek_mk)};""",conn=conn, cur = cur)
        if rez == False:
            CSQ.close_bd(conn, cur)
            CQT.msgbox('Ресурсная не обновлена')
            return False

    def create_and_add_res_to_mk(self):
        if self.tabl_mk.currentRow() == -1:
            return
        row = self.tabl_mk.currentRow()
        nom_tek_mk = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Пномер')).text()
        query = f'''SELECT xml.data as xml, mk.Количество, res.data as Ресурсная FROM mk 
                    INNER JOIN xml ON mk.Пномер = xml.Номер_мк
                    INNER JOIN res ON mk.Пномер = res.Номер_мк
                    WHERE Пномер == {int(nom_tek_mk)}'''
        rez = CSQ.zapros(self.bd_naryad, query)
        xml = rez[-1][0]
        self.add_res_to_mk(xml, rez[-1][1], nom_tek_mk)
        CQT.msgbox('Успешно добавлена ресурсная')

    def open_mk(self):
        if self.tabl_mk.currentRow() == -1:
            return
        if not CMS.user_access(self.bd_naryad,'мкарт_маршрутные_закрытьоткрыть',F.user_name()):
            return
        row = self.tabl_mk.currentRow()
        tbl = self.tabl_mk
        nom_tek_mk = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Пномер')).text()
        nom_mk = nom_tek_mk
        project = tbl.item(tbl.currentRow(),
                           CQT.nom_kol_po_imen(tbl, 'Номенклатура')).text()
        nom_pu_r = tbl.item(tbl.currentRow(),
                            CQT.nom_kol_po_imen(tbl, 'Номер_заказа')).text()
        nom_pr_r = tbl.item(tbl.currentRow(),
                            CQT.nom_kol_po_imen(tbl, 'Номер_проекта')).text()
        kolvo = tbl.item(tbl.currentRow(), CQT.nom_kol_po_imen(tbl, 'Количество')).text()
        prim = tbl.item(tbl.currentRow(), CQT.nom_kol_po_imen(tbl, 'Примечание')).text()
        osnovanie = tbl.item(tbl.currentRow(),
                             CQT.nom_kol_po_imen(tbl, 'Основание')).text()
        conn, cur = CSQ.connect_bd(self.bd_naryad)
        #qery = CSQ.naiti_v_bd(self.bd_naryad, 'mk', {'Пномер': int(nom_tek_mk)}, ['Статус', 'Прогресс', 'Основание', 'Количество'], conn=conn,cur=cur)
        qery = CSQ.zapros(self.bd_naryad, f"""SELECT Статус, Прогресс, Основание, Количество, Дата_завершения
            FROM mk WHERE Пномер == {int(nom_tek_mk)};""", conn=conn,cur=cur, rez_dict=True, one=True)

        if qery['Статус'] == "Закрыта":
            if qery['Дата_завершения'] != "":
                CSQ.close_bd(conn, cur)
                CQT.msgbox('Нельзя открыть завершенную закрытую МК')
                return
            kolvo = qery['Количество']
            if kolvo == 0:
                CSQ.close_bd(conn, cur)
                CQT.msgbox('Необходимо исправить количество, не может быть 0')
                return

            query = f'''SELECT xml.data as xml, res.data as Ресурсная, xml.Head as xml_head FROM xml 
                                            INNER JOIN res ON res.Номер_мк = xml.Номер_мк
                                            WHERE xml.Номер_мк == {int(nom_tek_mk)}'''
            rez = CSQ.zapros(self.db_resxml, query)

            if len(rez) == 1:
                CQT.msgbox(f'Ошибка выгрузки ресурсной или хмл')
                return


            if rez[-1][1] == '':
                xml = rez[-1][0]
                xml_head = rez[-1][2]
                rez = self.add_res_to_mk(xml, kolvo, nom_tek_mk,xml_head, conn = conn, cur= cur)
                if rez == False:
                    return
            #rez = CSQ.update_bd_sql(self.bd_naryad, 'mk', {'Статус': 'Открыта'}, {'Пномер': int(nom_tek_mk)}, conn = conn, cur = cur)
            rez = CSQ.zapros(self.bd_naryad,f"""UPDATE mk SET Статус = 'Открыта' WHERE Пномер = {int(nom_tek_mk)};""",conn=conn,cur=cur)
            if rez == False:
                CSQ.close_bd(conn, cur)
                CQT.msgbox('Запрос не выполнен')
                return
            self.ui.pushButton_open_mk.setEnabled(False)
            self.ui.pushButton_close_mk.setEnabled(True)
            self.tab_click(row)

            try:
                msg = f"{F.user_full_namre()} ОТКРЫЛ мк № {str(nom_mk)}:\n{project} - {str(kolvo)} шт.\n{nom_pu_r.strip()} Проект: {nom_pr_r.strip()}\n" \
                      f"Прим.: {prim} {osnovanie}"
                self.send_info_mk_b24(msg, 'chat41228')
            except:
                print('Ошибка отправки в Б24')
        else:
            CSQ.close_bd(conn,cur)



    def corr_mk(self, row, kol):
        if self.tabl_mk.hasFocus() == True:
            if self.tabl_mk.currentRow() == -1:
                return
            nom_tek_mk = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Пномер')).text()
            prim = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Примечание')).text()
            prior = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Приоритет')).text()
            paral = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Коэф_парал')).text()
            vid = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Вид')).text()
            iscl = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Искл_план_рм')).text()
            if F.is_numeric(prior) == False:
                CQT.msgbox('Не число Приоритет')
                self.load_tab_mk()
                return
            if F.is_numeric(paral) == False:
                CQT.msgbox('Не число Коэф_парал')
                self.load_tab_mk()
                return

            prior = int(prior)
            paral = int(paral)
            dict_zamen = {';':',','-':',','/':',','\\':',',' ':','}
            for key in dict_zamen.keys():
                iscl = iscl.replace(key,dict_zamen[key])
            #rez = CSQ.update_bd_sql(self.bd_naryad, 'mk',
            #                        {'Примечание': prim, 'Приоритет': prior, 'Коэф_парал': paral, 'Вид': vid, 'Искл_план_рм': iscl},
            #                        {'Пномер': int(nom_tek_mk)})
            rez = CSQ.zapros(self.bd_naryad,f"""UPDATE mk SET Примечание = '{prim}', Приоритет = {prior}, 
                    Коэф_парал = {paral}, Вид = '{vid}', Искл_план_рм = '{iscl}' WHERE Пномер = '{int(nom_tek_mk)}';""")
            if rez == False:
                CQT.msgbox('Запрос не выполнен')
                return
            self.cvet_prioritetov()

    def naiti_parametr_v_stroke(self, stroka, parametr):
        arr = stroka.split('  ')
        for i in arr:
            if i != '':
                if i.strip().startswith(parametr) == True:
                    return i.replace(parametr, '').strip()

    def ass_brak_to_mk(self):
        if self.tabl_brak.currentIndex() == -1:
            return
        tabl_razr_mk = self.ui.table_razr_MK
        label = self.ui.label_ass
        label_brak = self.ui.label_opis_braka
        nom_oper = self.naiti_parametr_v_stroke(label_brak.text(), '№ операции:')
        nom_id = self.naiti_parametr_v_stroke(label_brak.text(), 'ID:')
        dse = self.naiti_parametr_v_stroke(label_brak.text(), 'ДСЕ:')
        nom_kol_nach_tabl = CQT.nom_kol_po_imen(tabl_razr_mk, 'Сумм.Количество')
        if nom_id == None:
            return
        if self.tabl_brak.currentRow() == None or self.tabl_brak.currentRow() == -1:
            CQT.msgbox("Не выбран акт о браке")
            return
        nom = self.tabl_brak.item(self.tabl_brak.currentRow(), 0).text().replace('Номер акта:', '')

        if nom_kol_nach_tabl == None:
            CQT.msgbox('Не создана МК')
            return
        kol_id = CQT.nom_kol_po_imen(tabl_razr_mk, 'ID')
        kol_naim = CQT.nom_kol_po_imen(tabl_razr_mk, 'Наименование')
        kol_nn = CQT.nom_kol_po_imen(tabl_razr_mk, 'Обозначение')
        flag_naid = False
        for i in range(tabl_razr_mk.rowCount()):
            if flag_naid == True:
                break
            if tabl_razr_mk.item(i, kol_id).text() == nom_id or tabl_razr_mk.item(i,
                                                                                  kol_naim).text() + ' ' + tabl_razr_mk.item(
                i, kol_nn).text() == dse:
                for j in range(tabl_razr_mk.columnCount() - 1, nom_kol_nach_tabl, -1):
                    if tabl_razr_mk.item(i, j).text() != '':
                        if nom_oper not in tabl_razr_mk.item(i, j).text():
                            tabl_razr_mk.item(i, j).setText('')
                            flag_pust = True
                            for k in range(tabl_razr_mk.rowCount()):
                                if tabl_razr_mk.item(k, j).text() != '':
                                    flag_pust = False
                                    break
                            if flag_pust == True:
                                for k in range(4):
                                    t = CQT.spisok_iz_wtabl(tabl_razr_mk, "", True)
                                    tabl_razr_mk.removeColumn(j)
                        else:
                            flag_naid = True
                            break
        if flag_naid == False:
            CQT.msgbox('Не найдена деталь для ассоциации ' + nom_id)
            return

        spis_ass = self.ui.label_ass.text().split(';')
        if spis_ass[0] == '':
            spis_ass.pop(0)
        spis_ass.append(nom)
        self.ui.label_ass.setText(';'.join(spis_ass))
        CQT.msgbox(
            'Нужно проверить, что концовка новой маршрутной карты должна совпдать с местом возникновения брака по основной маршрутной карте.')
        self.ui.tabWidget.setCurrentIndex(1)
        self.ui.pushButton_save_MK.setEnabled(True)

    def zapoln_tabl_nomenkl(self):
        tabl_nomenk = self.ui.table_nomenkl
        conn, cur = CSQ.connect_bd(self.db_dse)
        #spis = CSQ.spis_iz_bd_sql(self.db_dse, 'dse', False, True,conn=conn, cur=cur)
        spis = CSQ.zapros(self.db_dse, """SELECT * FROM dse;""",shapka=True )
        CSQ.close_bd(conn,cur)
        red_nom = {1, 2, 4}
        CQT.zapoln_wtabl(self, spis, tabl_nomenk, {0, 1, 2, 3, 4, 7, 8, 9, 10}, red_nom, (), (), 200, True, '',
                         max_vis_row=20)
        CMS.zapolnit_filtr(self, self.ui.tbl_filtr_nomenkl, tabl_nomenk)
        # tabl_nomenk.setMouseTracking(True)
        # tabl_nomenk.selectRow(tabl_nomenk.rowCount()-1)
        tabl_nomenk.scrollToBottom()

        # CQT.ust_cvet_obj(self.ui.btn_korr_nom)

    def load_tab_mk(self):
        tabl_mk = self.ui.table_spis_MK
        if tabl_mk.currentIndex() != -1:
            tmp_poz = tabl_mk.currentIndex()
        zapros = f'''SELECT mk.Пномер, Тип_мк.Имя as Тип, Дата, Статус,Номенклатура,Номер_заказа,Номер_проекта,Вид,Ресурсная_дата,Примечание,Основание,
Прогресс,Приоритет,Направление,Вес,Количество, Дата_завершения, Коэф_парал, Искл_план_рм FROM mk 
        INNER JOIN Тип_мк ON Тип_мк.Пномер = mk.Тип WHERE Место == {self.place}'''
        # spis = CSQ.spis_iz_bd_sql(self.bd_naryad, 'mk', False, True)
        spis = CSQ.zapros(self.bd_naryad, zapros, '', True)

        spis_korr = {F.nom_kol_po_im_v_shap(spis, 'Примечание')
            , F.nom_kol_po_im_v_shap(spis, 'Приоритет')
            , F.nom_kol_po_im_v_shap(spis, 'Коэф_парал')
            , F.nom_kol_po_im_v_shap(spis, 'Вид')
            , F.nom_kol_po_im_v_shap(spis, 'Искл_план_рм')}

        if self.ui.chk_progress.isChecked():
            # === процент выполнения====
            spis[0].append('Прогресс_01')
            for i in range(1, len(spis)):
                spis[i].append('Прогресс_01')
        CQT.zapoln_wtabl(self, spis, tabl_mk, 0, spis_korr, (), (), 200, True, '')
        CQT.cvet_cell_wtabl(tabl_mk, 'Прогресс', '', 'Завершено')
        CQT.cvet_cell_wtabl(tabl_mk, 'Статус', '', 'Закрыта')
        CQT.cvet_cell_wtabl(tabl_mk, 'Статус', '', 'Открыта', 203, 176, 0, False)
        for key in self.DICT_TIP_MK.keys():
            r, g, b = self.DICT_TIP_MK[key]['rgb'].split(',')
            CQT.cvet_cell_wtabl(tabl_mk, 'Тип', '', key, r, g, b, False)
        self.cvet_prioritetov()
        # CQT.zapolnit_progress(self, tabl_mk, CQT.nom_kol_po_imen(tabl_mk, 'Уровень_вып'))
        tabl_mk.setCurrentIndex(tmp_poz)

    def cvet_prioritetov(self):
        tabl_mk = self.ui.table_spis_MK
        CQT.cvet_cell_wtabl(tabl_mk, 'Приоритет', '', "0", 254, 254, 254, False)
        CQT.cvet_cell_wtabl(tabl_mk, 'Приоритет', '', "1", 254, 200, 200, False)
        CQT.cvet_cell_wtabl(tabl_mk, 'Приоритет', '', "2", 254, 150, 150, False)
        CQT.cvet_cell_wtabl(tabl_mk, 'Приоритет', '', "3", 254, 100, 100, False)
        CQT.cvet_cell_wtabl(tabl_mk, 'Приоритет', '', "4", 254, 50, 50, False)
        CQT.cvet_cell_wtabl(tabl_mk, 'Приоритет', '', "5", 254, 0, 0, False)

    def clear_mk2(self):
        tabl_cr_stukt = self.ui.table_razr_MK
        but_add_gl_uzel = self.ui.pushButton_create_koren
        but_add_vhod = self.ui.pushButton_create_vxodyash
        but_udal_uzel = self.ui.pushButton_create_udalituzel
        tabl_cr_stukt.clearContents()
        tabl_cr_stukt.setRowCount(0)
        tabl_cr_stukt.setColumnCount(21)
        shapka = ['Наименование', 'Обозначение', 'Количество', 'Ед.изм.', 'Масса/М1,М2,М3', 'Ссылка',
                  'ID', 'Количество на изделие', 'Примечание', 'ПКИ', 'Сумм.Количество', '', '', '', '', '', '', ''
            , 'dreva_kod', 'Кол. по заявке', 'Уровень']
        tabl_cr_stukt.setHorizontalHeaderLabels(shapka)
        tabl_cr_stukt.resizeColumnsToContents()
        for i in range(10, 19):
            tabl_cr_stukt.setColumnHidden(i, True)
        CQT.ust_cvet_videl_tab(tabl_cr_stukt)
        tabl_cr_stukt.setSelectionMode(1)
        self.list_vars_vo = []
        but_add_gl_uzel.setEnabled(True)
        but_add_vhod.setEnabled(True)
        but_udal_uzel.setEnabled(True)
        self.ui.label_ass.clear()
        self.ui.lineEdit_ves.clear()
        self.ui.comboBox_napravlenia.setCurrentText('')
        self.ui.comboBox_vid.setCurrentText('')
        self.ui.comboBox_np.setCurrentText('')
        self.ui.pushButton_ass_brak_to_mk.setEnabled(True)
        self.ui.btn_save_cust_drevo.setEnabled(True)
        self.ui.btn_load_cust_drevo.setEnabled(True)
        self.ui.pushButton_create_paralel.setEnabled(True)
        self.ui.btn_vigruzka_norm.setEnabled(True)
        if self.ui.tabWidget_2.tabText(self.ui.tabWidget_2.currentIndex()) == 'Разработка МК':
            self.ui.btn_vigruzka_norm_mat.setEnabled(True)
        else:
            self.ui.btn_vigruzka_norm_mat.setEnabled(False)

    # def vibor_PY(self, param):
    #    combo_PY = self.ui.comboBox_PY
    #    combo_proekt = self.ui.comboBox_np
    #    combo_proekt.setCurrentIndex(param)
    #    return

    def vibor_pr(self):
        # combo_PY = self.ui.comboBox_PY
        combo_proekt = self.ui.comboBox_np.currentText().split('$')
        self.ui.comboBox_vid.setCurrentText(combo_proekt[2])
        self.ui.comboBox_napravlenia.setCurrentText(combo_proekt[3])
        """nk_kol = CQT.nom_kol_po_imen(self.ui.table_zayavk,'Количество')
        if nk_kol != None:
            try:
                self.ui.table_zayavk.item(,nk_kol).setText(combo_proekt[4])
            except:
                pass"""
        # combo_PY.setCurrentIndex(param)
        return

    def cr_mk2(self):
        butt_add_gl_uzel = self.ui.pushButton_create_koren
        tabl_cr_stukt = self.ui.table_razr_MK
        nk = CQT.nom_kol_po_imen(tabl_cr_stukt, 'Уровень')
        nk_kol_p_z = CQT.nom_kol_po_imen(tabl_cr_stukt, 'Кол. по заявке')
        nk_kol = CQT.nom_kol_po_imen(tabl_cr_stukt, 'Количество')
        nk_kol_izd = CQT.nom_kol_po_imen(tabl_cr_stukt, 'Количество на изделие')
        naim = CQT.nom_kol_po_imen(tabl_cr_stukt, 'Наименование')
        nn = CQT.nom_kol_po_imen(tabl_cr_stukt, 'Обозначение')
        nom_kol_mass = CQT.nom_kol_po_imen(tabl_cr_stukt, 'Масса/М1,М2,М3')
        if nk == False or nk_kol_p_z == False:
            CQT.msgbox('Ошибка выбора колонок')
            return
        if nk == None:
            return
        if nk_kol_p_z == None:
            return

        min = 1000
        for i in range(tabl_cr_stukt.rowCount()):
            if int(tabl_cr_stukt.item(i, nk).text()) < min:
                min = int(tabl_cr_stukt.item(i, nk).text())
        if min > 0:
            for i in range(tabl_cr_stukt.rowCount()):
                tabl_cr_stukt.item(i, nk).setText(str(int(tabl_cr_stukt.item(i, nk).text()) - min))
        for i in range(tabl_cr_stukt.rowCount()):
            flag_err = False
            if int(tabl_cr_stukt.item(i, nk).text()) == 0:
                tabl_cr_stukt.item(i, nk_kol).setText(str(1))
            if tabl_cr_stukt.item(i, nom_kol_mass).text().count('/') != 3 or F.is_numeric(
                    tabl_cr_stukt.item(i, nom_kol_mass).text().split('/')[0]) == False:
                CQT.msgbox(
                    f'Не записана масса в {i + 1} строке, {tabl_cr_stukt.horizontalHeaderItem(nom_kol_mass).text()} колонке. Значение:({tabl_cr_stukt.item(i, nom_kol_mass).text()})')
                flag_err = True
            if F.is_numeric(tabl_cr_stukt.item(i, 19).text()) == False and int(tabl_cr_stukt.item(i, nk).text()) == 0:
                CQT.msgbox(
                    f'Не число в {i + 1} строке, {tabl_cr_stukt.horizontalHeaderItem(19).text()} колонке. Значение:({tabl_cr_stukt.item(i, 19).text()})')
                flag_err = True
            if F.is_numeric(tabl_cr_stukt.item(i, 9).text()) == False and tabl_cr_stukt.item(i, 9).text() != "":
                CQT.msgbox(
                    f'Не число в {i + 1} строке, {tabl_cr_stukt.horizontalHeaderItem(9).text()} колонке. Значение:({tabl_cr_stukt.item(i, 9).text()})')
                flag_err = True
            if flag_err == True:
                return False
            if F.is_numeric(tabl_cr_stukt.item(i, nk_kol).text()) == False:
                CQT.msgbox(
                    f'в строке {i + 1}, колонке {nk_kol} не число. Значение:({tabl_cr_stukt.item(i, nk_kol).text()})')
                return False

            naim_ = tabl_cr_stukt.item(i, naim).text()
            nn_ = tabl_cr_stukt.item(i, nn).text()
            summ = 0
            for j in range(tabl_cr_stukt.rowCount()):
                if naim_ == tabl_cr_stukt.item(j, naim).text() and nn_ == tabl_cr_stukt.item(j, nn).text():
                    nach_ur = int(tabl_cr_stukt.item(j, nk).text())
                    if tabl_cr_stukt.item(j, nk_kol).text() == "":
                        CQT.msgbox(f'В колонке {nk_kol} не число')
                        return
                    kol_ = int(F.valm(tabl_cr_stukt.item(j, nk_kol).text()))
                    for k in range(j - 1, -1, -1):
                        if int(tabl_cr_stukt.item(k, nk).text()) > nach_ur:
                            break
                        if int(tabl_cr_stukt.item(k, nk).text()) < nach_ur:
                            kol_ *= int(F.valm(tabl_cr_stukt.item(k, nk_kol).text()))
                            nach_ur = int(tabl_cr_stukt.item(k, nk).text())
                        if int(tabl_cr_stukt.item(k, nk).text()) == 0:
                            break
                    summ += kol_
            tabl_cr_stukt.item(i, nk_kol_izd).setText(str(summ))

        for i in range(tabl_cr_stukt.rowCount()):
            if int(tabl_cr_stukt.item(i, nk).text()) == 0:
                if tabl_cr_stukt.item(i, nk_kol_p_z).text() == "":
                    CQT.msgbox('не указано Количество комплектов на ' + tabl_cr_stukt.item(i, 1).text())
                    tabl_cr_stukt.setCurrentCell(i, nk_kol_p_z)
                    return False

            # tabl_cr_stukt.item(i,nk_kol).setText(str(int(tabl_cr_stukt.item(i,nk_kol_izd).text())* int(kol)))
        self.ui.btn_vigruzka_norm.setEnabled(True)
        self.ui.btn_vigruzka_norm_mat.setEnabled(False)
        return

    def add_v_mk(self):
        tab = self.ui.tabWidget
        tab2 = self.ui.tabWidget_2
        tabl_cr_stukt = self.ui.table_razr_MK

        shapka = [
                'Наименование'
                ,'Обозначение'
                ,'Количество'
                ,'Ед.изм.'
                ,'Масса/М1,М2,М3'
                ,'Ссылка'
                ,'ID'
                ,'Количество на изделие'
                ,'Примечание'
                ,'ПКИ'
                ,'Сумм.Количество'
                ,''
                ,''
                ,''
                ,''
                ,''
                ,''
                ,''
                ,'dreva_kod'
                ,'Кол. по заявке'
                ,'Уровень']

        tree = self.ui.treeWidget
        if tree.currentIndex().row() == -1:
            CQT.msgbox('Не выбран узел')
            return
        item = tree.currentItem()
        if item == None:
            return
        nk = CQT.nom_kol_po_imen(tree, 'ID')
        current_ID = item.text(nk)
        sp_tree = CQT.spisok_dreva(tree, shapka=True)
        flag_naid = -1
        for i in range(len(sp_tree)):
            if sp_tree[i][nk] == current_ID:
                flag_naid = i
                break
        if flag_naid == -1:
            CQT.msgbox("Не найден выбранный узел")
            return

        q_strok = tabl_cr_stukt.currentRow() + 1
        q_column = tabl_cr_stukt.currentColumn()
        spisok = CQT.spisok_iz_wtabl(tabl_cr_stukt, "", True)
        shapka = spisok[0]

        list_add = ['Сумм.Количество','', 'dreva_kod','Кол. по заявке']
        for item in list_add:
            sp_tree[0].append(item)
            for i in range(1,len(sp_tree)):
                sp_tree[i].append('')

        dict_tree = F.list_to_dict(sp_tree,'ID')
        nk_uroven = F.nom_kol_po_im_v_shap(sp_tree,'Уровень')
        nk_uroven_out = F.nom_kol_po_im_v_shap(spisok, 'Уровень')

        dict_zamen = {'Ссылка на объект DOCs':'Ссылка', 'Покупное изделие':'ПКИ', 'Обозначение':'Обозначение2', 'Обозначение полное':'Обозначение'}


        for i in range(len(sp_tree[0])):
            for key in dict_zamen.keys():
                if sp_tree[0][i] == key:
                    sp_tree[0][i] = dict_zamen[key]

        for item in shapka:
            if item not in sp_tree[0]:
                CQT.msgbox(f'В xml не найдено поле {item}')
                return


        list_to_add = [sp_tree[0]]
        ur = int(sp_tree[flag_naid][nk_uroven])
        list_to_add.append(sp_tree[flag_naid])
        for i in range(flag_naid+1, len(sp_tree)):
            if int(sp_tree[i][nk_uroven]) <= ur:
                break
            list_to_add.append(sp_tree[i])
        dict_to_add = F.list_to_dict(list_to_add,'ID')

        for id in dict_to_add.keys():
            dict_to_add[id]['Количество на изделие'] = ''
            tmp = []
            for item in shapka:
                tmp.append(dict_to_add[id][item])
            spisok.append(tmp)



        CQT.zapoln_wtabl(self, spisok, tabl_cr_stukt, 0, self.edit_cr_mk, (), (), 200, True, '', 30)

        tabl_cr_stukt.clearSelection()
        tabl_cr_stukt.setCurrentCell(q_strok, q_column)
        tab.setCurrentIndex(1)
        tab2.setCurrentIndex(1)

    def ass_dse_to_mk(self):
        tabl_cr_stukt = self.ui.table_razr_MK
        tabl_nomenk = self.ui.table_nomenkl
        if tabl_cr_stukt.currentRow() == -1:
            CQT.msgbox('Не выбрана позиция в МК')
            return
        if tabl_nomenk.currentRow() == -1:
            CQT.msgbox('Не выбрана ДСЕ')
            return

        naim = CQT.znach_vib_strok_po_kol(tabl_nomenk, 'Наименование')
        nn = CQT.znach_vib_strok_po_kol(tabl_nomenk, 'Номенклатурный_номер')
        if nn == False or naim == False:
            return
        CQT.zapis_znach_vib_strok_po_kol(tabl_cr_stukt, 'Наименование', naim)
        CQT.zapis_znach_vib_strok_po_kol(tabl_cr_stukt, 'Обозначение', nn)
        self.ui.tabWidget.setCurrentIndex(1)
        self.ui.tabWidget_2.setCurrentIndex(1)

    def del_uzel(self):
        butt_add_gl_uzel = self.ui.pushButton_create_koren
        tabl_cr_stukt = self.ui.table_razr_MK
        if tabl_cr_stukt.currentRow() == -1:
            CQT.msgbox('Не выбрана позиция в МК')
            return
        q_strok = tabl_cr_stukt.currentRow()
        q_column = tabl_cr_stukt.currentColumn()

        spisok = CQT.spisok_iz_wtabl(tabl_cr_stukt, "", True)
        spisok_tmp = spisok.copy()
        k = 0
        spisok.pop(q_strok + 1)
        k += 1
        ur = int(tabl_cr_stukt.item(q_strok, 20).text())
        for i in range(q_strok + 2, len(spisok_tmp)):
            if int(spisok_tmp[i][20]) > ur:
                spisok.pop(i - k)
                k += 1
            else:
                break

        CQT.zapoln_wtabl(self, spisok, tabl_cr_stukt, 0, self.edit_cr_mk_ruch, (), (), 200, True, '', 30)
        tabl_cr_stukt.setCurrentCell(q_strok, q_column)

    def add_paral(self):
        self.add_vhod("0")

    def add_vhod(self, ur="1"):
        if ur == False:
            ur = "1"
        tabl_cr_stukt = self.ui.table_razr_MK
        if tabl_cr_stukt.currentRow() == -1:
            CQT.msgbox('Не выбрана позиция в МК')
            return
        q_strok = tabl_cr_stukt.currentRow()
        q_column = tabl_cr_stukt.currentColumn()

        spisok = CQT.spisok_iz_wtabl(tabl_cr_stukt, "", True)
        strok = []
        for i in range(20):
            strok.append('')
        strok.append(str(int(tabl_cr_stukt.item(q_strok, 20).text()) + int(ur)))
        strok[6] = F.time_metka()
        strok[4] = '/М1/М2/М3'
        spisok.insert(q_strok + 2, strok)

        CQT.zapoln_wtabl(self, spisok, tabl_cr_stukt, 0, self.edit_cr_mk_ruch, (), (), 200, True, '', 30)
        tabl_cr_stukt.clearSelection()

        tabl_cr_stukt.setCurrentCell(q_strok, q_column)

    def add_gl_uzel(self, id=''):
        butt_add_gl_uzel = self.ui.pushButton_create_koren
        tabl_cr_stukt = self.ui.table_razr_MK
        spisok = CQT.spisok_iz_wtabl(tabl_cr_stukt, "", True)
        strok = []
        for i in range(20):
            strok.append('')
        strok.append('0')
        strok[2] = '1'
        if id == '' or id == False:
            strok[6] = F.time_metka()
        else:
            strok[6] = id
        strok[4] = '/М1/М2/М3'
        spisok.append(strok)

        CQT.zapoln_wtabl(self, spisok, tabl_cr_stukt, 0, self.edit_cr_mk_ruch, (), (), 200, True, '', 30)
        self.ui.btn_vigruzka_norm.setEnabled(False)
        self.ui.btn_vigruzka_norm_mat.setEnabled(True)

    def kol_po_zayav(self, sp_xml_tmp, kol):
        sp_xml_tmp[10] = int(sp_xml_tmp[2]) * int(kol)
        return sp_xml_tmp

    def kol_v_uzel(self, s, j, nk_naim, nk_kol, nk_ur=''):
        koef = 1
        if nk_ur != "":
            koef_ur = int(s[j][nk_ur])
        else:
            koef_ur = CMS.uroven(s[j][nk_naim])
        for k in range(j - 1, 0, -1):
            if nk_ur != "":
                ur_tmp = int(s[k][nk_ur])
            else:
                ur_tmp = CMS.uroven(s[k][nk_naim])
            if ur_tmp < koef_ur:
                koef *= int(s[k][nk_kol])
                koef_ur = ur_tmp
            if koef_ur == 0:
                break
        return int(koef)



    def kol_na_izd(self, s, kol_po_zayav: int):
        nk_kol = F.nom_kol_po_im_v_shap(s, 'Количество')
        nk_kol_izd = F.nom_kol_po_im_v_shap(s, 'Количество на изделие')
        nk_kol_summ = F.nom_kol_po_im_v_shap(s, 'Сумм.Количество')
        nk_naim = F.nom_kol_po_im_v_shap(s, 'Наименование')
        nk_nn = F.nom_kol_po_im_v_shap(s, 'Обозначение')
        nk_ur = F.nom_kol_po_im_v_shap(s, 'Уровень')
        if s[0][nk_ur] == 'Уровень':
            for i in range(1, len(s)):
                s[i][nk_naim] = '    ' * int(s[i][nk_ur]) + s[i][nk_naim].strip()
                s[i][nk_nn] = '    ' * int(s[i][nk_ur]) + s[i][nk_nn].strip()
        for i in range(1, len(s)):
            naim = s[i][nk_naim].strip()
            nn = s[i][nk_nn].strip()
            summ = 0
            if type(F.valm(s[i][nk_kol])) == float:
                if F.valm(s[i][nk_kol]).is_integer():
                    s[i][nk_kol] = int(F.valm(s[i][nk_kol]))
            koef = self.kol_v_uzel(s, i, nk_naim, nk_kol)
            if type(F.valm(s[i][nk_kol])) == float:
                if not F.valm(s[i][nk_kol]).is_integer():
                    CQT.msgbox(f'{s[i][nk_naim]} {s[i][nk_nn]} по количеству занесен как расходник, обратиться к технологу')
                    CQT.statusbar_text(self)
                    return
            s[i][nk_kol_summ] = koef * kol_po_zayav * int(F.valm(s[i][nk_kol]))
            for j in range(1, len(s)):
                if s[j][nk_naim].strip() == naim and s[j][nk_nn].strip() == nn:
                    koef = self.kol_v_uzel(s, j, nk_naim, nk_kol)
                    summ +=int(F.valm(s[j][nk_kol])) * koef
            s[i][nk_kol_izd] = str(summ) + ' (' + str(summ * kol_po_zayav) + ')'
        return s

    def kol_na_izd_zayav_1(self, s):
        # ===========================================================доделать
        nk_kol = F.nom_kol_po_im_v_shap(s, 'Количество')
        nk_kol_p_z = F.nom_kol_po_im_v_shap(s, 'Кол. по заявке')
        nk_kol_izd = F.nom_kol_po_im_v_shap(s, 'Количество на изделие')
        s[0][10] = 'Сумм.Количество'
        nk_kol_summ = F.nom_kol_po_im_v_shap(s, 'Сумм.Количество')
        kol = 1
        if nk_kol_p_z == None:
            CQT.msgbox('Не подходящий набор данных для формирования МК')
            return
        for i in range(1, len(s)):
            koef = 1
            if int(s[i][20]) == 0:

                kol = int(s[i][nk_kol_p_z])
            else:

                tek_ur = int(s[i][20])
                for j in range(i - 1, -1, -1):
                    if int(s[j][20]) < tek_ur:
                        koef *= int(s[j][nk_kol])
                        tek_ur = int(s[j][20])
                    if int(s[j][20]) == 0:
                        break

            if i:
                s[i][nk_kol_izd] = str(s[i][nk_kol_izd]) + ' (' + str(int(s[i][nk_kol_izd]) * kol) + ')'
                s[i][nk_kol_summ] = int(s[i][nk_kol]) * kol * koef

        return s

    # ===========================================================

    def raschet_vesa_xml(self, s_vert):
        nom_kol_mat = F.nom_kol_po_im_v_shap(s_vert, 'Масса/М1,М2,М3')
        nom_kol_kol = F.nom_kol_po_im_v_shap(s_vert, 'Количество')
        ves = 0
        for i in range(1, len(s_vert)):
            if s_vert[i][nom_kol_mat].split('/')[1] != '' and s_vert[i][nom_kol_mat].split('/')[2] != '':
                if F.is_numeric(s_vert[i][nom_kol_mat].split('/')[0]) == False:
                    CQT.msgbox(f'В строке {i} вес не число')
                    return 0

                ves += (F.valm(s_vert[i][nom_kol_mat].split('/')[0]) * F.valm(s_vert[i][nom_kol_kol]))
        return ves

    def raschet_vesa_dse(self):
        self.LIST_ED_IZM_MAT = ['Килограмм', 'кг']
        #nom_kol_mat = F.nom_kol_po_im_v_shap(s_vert, 'Масса/М1,М2,М3')
        #nom_kol_kol = F.nom_kol_po_im_v_shap(s_vert, 'Количество')
        #nom_kol_naim = F.nom_kol_po_im_v_shap(s_vert, 'Наименование')
        #nom_kol_tip = F.nom_kol_po_im_v_shap(s_vert, 'Тип')
        #ves = 0
        #if ruchnoi == False:
        #    for i in range(1, len(s_vert)):
        #        if s_vert[i][nom_kol_tip] != 'Сборочная единица':
        #            if F.is_numeric(s_vert[i][nom_kol_mat].split('/')[0]) == False:
        #                CQT.msgbox(f'В строке {i} вес не число')
        #                return 0
        #            ves += (F.valm(s_vert[i][nom_kol_mat].split('/')[0]) * F.valm(s_vert[i][nom_kol_kol]) *  self.kol_v_uzel(s_vert, i,nom_kol_naim, nom_kol_kol))
        #
        #else:
        #    for i in range(1, len(s_vert)):
        #        if s_vert[i][nom_kol_mat].split('/')[1] != '' and s_vert[i][nom_kol_mat].split('/')[2] != '':
        #            if F.is_numeric(s_vert[i][nom_kol_mat].split('/')[0]) == False:
        #                CQT.msgbox(f'В строке {i} вес не число')
        #                return 0
        #            ves += (F.valm(s_vert[i][nom_kol_mat].split('/')[0]) * F.valm(s_vert[i][nom_kol_kol]) *  self.kol_v_uzel(s_vert, i,nom_kol_naim, nom_kol_kol))
        if self.res == '':
            CQT.msgbox(f'ОШибка')
            return
        ves_res = 0
        ves_res_list= 0
        list_hz_mat= []
        res = self.res
        for dse in res:
            for oper in dse['Операции']:
                for mat in oper['Материалы']:
                    if mat['Мат_ед_изм'] in self.LIST_ED_IZM_MAT:
                        ves_res += F.valm(mat['Мат_норма'])
                        # print(f"{F.valm(mat['Мат_норма'])} опер {oper['Опер_наименовние']} дет {dse['Наименование']}")
                        if mat['Мат_код'] in self.DICT_MAT and self.DICT_MAT[mat['Мат_код']]['П5'] == '1':
                            if self.DICT_MAT[mat['Мат_код']]['П6'] != '':
                                ves_res_list += F.valm(mat['Мат_норма'])
        return ves_res, round(ves_res_list,2)

    def create_sp_dreva_ruchnoi(self):
        tabl_cr_stukt = self.ui.table_razr_MK
        if tabl_cr_stukt.rowCount() == 0:
            return
        but_add_gl_uzel = self.ui.pushButton_create_koren
        but_add_vhod = self.ui.pushButton_create_vxodyash
        but_udal_uzel = self.ui.pushButton_create_udalituzel
        rez = self.cr_mk2()
        if rez == False:
            return
        s_vert = CQT.spisok_iz_wtabl(tabl_cr_stukt, "", True)
        self.kol_izdeliy = int(s_vert[1][F.nom_kol_po_im_v_shap(s_vert, 'Кол. по заявке')])
        s_vert = self.kol_na_izd(s_vert, self.kol_izdeliy)
        if s_vert == None:
            CQT.msgbox('Не подходящий набор данных для формирования МК')
            return
        but_add_gl_uzel.setEnabled(False)
        but_add_vhod.setEnabled(False)
        but_udal_uzel.setEnabled(False)
        self.xml_file = ''

        # =========================================РЕСУРСНАЯ
        self.res = self.resursnaya_from_cust_struktura(s_vert, kol_vo_izdeliy=self.kol_izdeliy, ruchnoi=True)
        if self.res == None:
            return
        dict_poziciy_rez = dict()
        dict_poziciy_rez['name'] = f"{s_vert[1][F.nom_kol_po_im_v_shap(s_vert, 'Наименование')]} {s_vert[1][F.nom_kol_po_im_v_shap(s_vert, 'Обозначение')]}"
        dict_poziciy_rez['data'] = self.res
        dict_poziciy_rez['kol_zayavk'] = self.kol_izdeliy
        self.spis_poziciy_rez_ruchnoi = [dict_poziciy_rez]
        return s_vert

    def create_sp_dreva_po_xml(self):
        tabl = self.ui.table_zayavk
        if tabl.rowCount() == 0:
            CQT.msgbox('Не добавлены заявки')
            return
        if tabl.columnCount() > 5:
            return
        sp_izd = CQT.spisok_iz_wtabl(tabl)
        if len(sp_izd) > 1:
            CQT.msgbox('Создать МК можно только на одну позицию, Del - удалить строку')
            return
        putt_xml = sp_izd[0][0]
        sp_xml_tmp = self.podgotovka_xml(XML.spisok_iz_xml(putt_xml))
        if sp_xml_tmp == None:
            CQT.msgbox('Файл не корректный')
            return
        self.xml_file = F.load_file_convert_to_binary(putt_xml)
        if sp_izd[0][2] == '':
            CQT.msgbox("Не указано Количество по заявке")
            return
        self.kol_izdeliy = int(sp_izd[0][2])
        # ==========================================РЕСУРСНАЯ
        self.res = CMS.resursnaya_from_xml(self, sp_xml_tmp, self.kol_izdeliy)
        # ==============================================какое то складывние
        """        for i in range(len(sp_xml_tmp)):
            for j in range(i + 1, len(sp_xml_tmp)):
                if i < len(sp_xml_tmp) - 1:
                    if sp_xml_tmp[i]['data']['Наименование'] == sp_xml_tmp[j]['data']['Наименование'] \
                            and sp_xml_tmp[i]['data']['Обозначение полное'] == sp_xml_tmp[j]['data']['Обозначение полное'] \
                            and sp_xml_tmp[i]['uroven'] == sp_xml_tmp[j]['uroven']:#НЕДОДАЕЛАНО!!!!
                        sp_xml_tmp[i][2] = str(int(sp_xml_tmp[i][2]) + int(sp_xml_tmp[j][2]))
                        sp_xml_tmp[j][0] = "deletes" + str(F.time_metka())
                    else:
                        break
        sp_xml_tmp_ = []
        for i in range(len(sp_xml_tmp)):
            if sp_xml_tmp[i][0].startswith('deletes') == False:
                sp_xml_tmp_.append(sp_xml_tmp[i])
        sp_xml_tmp = sp_xml_tmp_"""
        # =============================================================
        s_vert = [["Наименование"
        ,"Обозначение"
        ,"Количество"
        ,"Ед.изм."
        ,"Масса/М1,М2,М3"
        ,"Ссылка"
        ,"ID"
        ,"Количество на изделие"
        ,'Примечание'
        ,'ПКИ'
        , 'Сумм.Количество'
        , 'Уровень'
        , 'Тип']]

        for j in range(len(sp_xml_tmp)):
            for item  in ['Наименование','Обозначение полное','Количество','Единица измерения','Масса/М1,М2,М3',
                          'Ссылка на объект DOCs', 'ID', 'Количество на изделие','Примечание',  'Покупное изделие'
                          , 'Тип']:
                if item not in sp_xml_tmp[j]['data']:
                    CQT.msgbox(f'В файле отсутствует поле {item}')
                    return
        try:
            for j in range(len(sp_xml_tmp)):
                s_vert.append([sp_xml_tmp[j]['data']['Наименование'],sp_xml_tmp[j]['data']['Обозначение полное'],
                              sp_xml_tmp[j]['data']['Количество'],sp_xml_tmp[j]['data']['Единица измерения'],
                              sp_xml_tmp[j]['data']['Масса/М1,М2,М3'],sp_xml_tmp[j]['data']['Ссылка на объект DOCs'],
                              sp_xml_tmp[j]['data']['ID'],sp_xml_tmp[j]['data']['Количество на изделие'],
                              sp_xml_tmp[j]['data']['Примечание'],sp_xml_tmp[j]['data']['Покупное изделие'],
                              '',sp_xml_tmp[j]['uroven'],sp_xml_tmp[j]['data']['Тип']])
        except:
            CQT.msgbox(f'Ошибка обработки ХМЛ')
            return
        s_vert = self.kol_na_izd(s_vert, self.kol_izdeliy)
        if s_vert == None:
            return
        return s_vert

    def create_mk(self):
        self.ves_res_list = 0
        if self.ui.comboBox_napravlenia.currentText() == '':
            CQT.msgbox('Не указано направление')
            return
        if self.ui.tabWidget_2.currentIndex() == 1:  # вручную# вручную# вручную# вручную# вручную# вручную# вручную
            if  self.ui.table_razr_MK.columnCount() != 21:
                return
            CQT.statusbar_text(self, 'Формирование списка вручную')
            s_vert = self.create_sp_dreva_ruchnoi()
            if s_vert == None:
                return
            tabl = self.ui.table_razr_MK
        else:  # xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml
            CQT.statusbar_text(self, 'Формирование списка по ХМЛ')
            s_vert = self.create_sp_dreva_po_xml()
            if s_vert == None:
                return
            tabl = self.ui.table_zayavk

        CQT.statusbar_text(self, 'Форматирование списка')
        nach_sod = len(s_vert[0])

        conn, cur = CSQ.connect_bd(self.db_dse)
        DICT_DSE = CSQ.zapros(self.db_dse,f'''SELECT * FROM dse''', conn = conn, cur=cur,rez_dict=True)
        CSQ.close_bd(conn,cur)
        self.DICT_DSE_save_mk = F.raskrit_dict(DICT_DSE,'Номенклатурный_номер')
        for i in range(1, len(s_vert)):
            ima = s_vert[i][F.nom_kol_po_im_v_shap(s_vert,'Наименование')].strip()
            nn = s_vert[i][F.nom_kol_po_im_v_shap(s_vert,'Обозначение')].strip()
            kol_det_vseg = s_vert[i][F.nom_kol_po_im_v_shap(s_vert,'Сумм.Количество')]

            if nn not in self.DICT_DSE_save_mk:
                CQT.msgbox('Не найден в БД ' + ima + ' ' + nn)
                return
            if self.DICT_DSE_save_mk[nn]['Номер_техкарты'] == '':
                CQT.msgbox('Не найдена техкарта ' + ima + ' ' + nn)
                return
            nom_tk = self.DICT_DSE_save_mk[nn]['Номер_техкарты']
            put_name_tk = F.scfg('add_docs') + os.sep + nom_tk + "_" + nn
            tk = F.otkr_f(put_name_tk + '.txt', False, "|", True, True)
            if tk == ['']:
                CQT.msgbox(f'Не найдена техкарта {put_name_tk}')
                return
            tk = CMS.grup_tk_po_rabc(tk, kol_det_vseg)
            self.ogran = nach_sod - 1
            for k in tk:
                if k[0] == "":
                    CQT.msgbox('Рабочий центр на ' + k[2] + ' операцию не назначен для ' + nn)
                    return
                print(k[0], k[1], k[2], i, self.ogran)
                s_vert = self.dob_etap(s_vert, k[0], k[1], k[2], i, self.ogran)
        ves, self.ves_res_list = self.raschet_vesa_dse()
        #if self.ui.tabWidget_2.currentIndex() == 1:               # вручную# вручную# вручную# вручную# вручную# вручную# вручную:
        self.ui.lineEdit_ves.setText(str(round(ves, 2)))
        #else:                                           # xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml# xml
        #    self.ui.lineEdit_ves.setText(str(round(self.raschet_vesa_dse()[0], 2)))
        CQT.statusbar_text(self, 'Оформление итоговой табицы')
        #s_vert = self.oformlenie_sp_pod_mk(s_vert)
        if s_vert == None:
            return

        for i in range(tabl.columnCount()):
            tabl.setColumnHidden(i, False)
        CQT.zapoln_wtabl(self, s_vert, tabl, 0, 0, "", "", 200, True, '', 90)
        tabl.setSelectionBehavior(1)
        CQT.statusbar_text(self, 'Раскрашивание')
        self.oformlenie_formi_mk(tabl, s_vert)

        self.ui.pushButton_ass_brak_to_mk.setEnabled(True)
        if self.ui.pushButton_save_MK.isEnabled() == False:
            self.ui.tabWidget.setCurrentIndex(3)
        CQT.statusbar_text(self)
        self.ui.btn_save_cust_drevo.setEnabled(False)
        self.ui.btn_load_cust_drevo.setEnabled(False)
        self.ui.pushButton_create_paralel.setEnabled(False)
        self.mk_file_founding = ''

    def oformlenie_formi_mk(self, tabl, s):
        for i in range(11, len(s[0]) - 1, 4):
            for j in range(0, len(s) - 1):
                # if tabl.item(j,i) == None:
                #    cellinfo = QtWidgets.QTableWidgetItem('')
                #    tabl.setItem(j,i, cellinfo)
                CQT.ust_color_wtab(tabl, j, i, 227, 227, 227)

                # tabl.item(j,i).setBackground(QtGui.QColor(227,227,227))
        for i in range(0, 11):
            for j in range(0, len(s) - 1):
                # if tabl.item(j,i) == None:
                #    cellinfo = QtWidgets.QTableWidgetItem('')
                #    tabl.setItem(j,i, cellinfo)
                # tabl.item(j,i).setBackground(QtGui.QColor(227,227,227))
                CQT.ust_color_wtab(tabl, j, i, 227, 227, 227)
    def show_file_founding_mk(self):
        if self.tabl_mk.currentRow() == -1:
            return
        row = self.tabl_mk.currentRow()
        tbl = self.tabl_mk
        nom_tek_mk = self.tabl_mk.item(row, CQT.nom_kol_po_imen(self.tabl_mk, 'Пномер')).text()
        nom_mk = int(nom_tek_mk)
        path = self.get_file_founding(nom_mk,F.put_po_umolch())
        if path == False:
            CQT.msgbox(f'Отсутствует файл')
            return
        F.zapust_file_os(path)

    def load_file_mk_founfing(self):
        path = CQT.f_dialog_name(self,'Выбрать СЗ',CMS.load_tmp_path('file_mk_founfing'),"PDF files (*.pdf)")
        if path == '.':
            return
        CMS.save_tmp_path('file_mk_founfing', path,True)
        self.mk_file_founding = F.load_file_convert_to_binary(path)
        if sys.getsizeof(self.mk_file_founding) > 1048576:
            self.mk_file_founding = ''
            CQT.msgbox(f'Размер файла должен быть не более 1 мб')
            return
        self.mk_file_founding = F.pack_byte_file(self.mk_file_founding)
        return

    def check_file_founding(self):
        if self.mk_file_founding == "":
            CQT.migat_obj(self,2,self.ui.btn_load_file_mk_founfing,'Файл - СЗ основание не выбран')
            return False
        return True
    def get_file_founding(self,nom_mk,path_save):
        file = CSQ.zapros(self.bd_files, f"""SELECT file FROM MK_founding WHERE Num_mk = {int(nom_mk)}""", one=True,
                          shapka=False, one_column=True)
        if file == False or file == [] or file == ['']:
            return False
        unpack = F.unpack_byte_file(file[0])
        F.save_binary_convert_to_file(unpack,
                                      path_save + F.sep() + f'{str(nom_mk)}.pdf')
        return path_save + F.sep() + f'{str(nom_mk)}.pdf'
    def save_mk(self):
        msg_proizv_err = 'Необходима СЗ .pdf подписанная нач. Производства'
        msg_pdo_err = 'Необходима СЗ .pdf подписанная нач. ПДО'
        msg_tehn_err = 'Необходима СЗ .pdf подписанная гл. Технологом'
        # nom_pu = self.ui.comboBox_PY
        nom_pr = self.ui.comboBox_np
        prim = self.ui.lineEdit_prim
        tab2 = self.ui.tabWidget_2

        if self.res == '':
            CQT.msgbox('Не создана ресурсная')
            return
        res_pickle = F.to_binary_pickle(self.res)
        if self.ui.comboBox_napravlenia.currentText() == '':
            CQT.msgbox('Не указано направление')
            return
        if self.ui.comboBox_vid.currentText() == '':
            CQT.msgbox('Не указан Вид изделия')
            return
        if self.ui.lineEdit_ves.text() == "" or F.is_numeric(self.ui.lineEdit_ves.text()) == False:
            CQT.msgbox("Не указан вес")
            return
        if nom_pr.currentText() == "" or nom_pr.currentText().count('$') < 2:
            CQT.msgbox('Не корректный номер проекта')
            return
        if self.ui.cmb_tip_mk.currentText() == '':
            CQT.migat_obj(self,2,self.ui.cmb_tip_mk,'Не выбран тип МК')
            return


        tip_mk = self.DICT_TIP_MK[self.ui.cmb_tip_mk.currentText()]['Пномер']
        if tip_mk == 4:
            CQT.msgbox(f'Отключено')
            return
        if tip_mk == 2:
            if self.ui.cmb_tip_dorez.currentText() == '':
                CQT.migat_obj(self, 2, self.ui.cmb_tip_dorez, 'Не выбран тип дорезки')
                return
            if not self.check_file_founding():
                if self.DICT_TIP_DOREZ[self.ui.cmb_tip_dorez.currentText()] in (1,2,3,11):
                    CQT.msgbox(msg_proizv_err)
                    return
                if self.DICT_TIP_DOREZ[self.ui.cmb_tip_dorez.currentText()] in (10,12):
                    CQT.msgbox(msg_pdo_err)
                    return
                if self.DICT_TIP_DOREZ[self.ui.cmb_tip_dorez.currentText()] in (4,5,6,7,8,9):
                    CQT.msgbox(msg_tehn_err)
                    return
        if tip_mk == 3:
            if not self.check_file_founding():
                CQT.msgbox(msg_tehn_err)
                return
        if tip_mk == 5:
            if not self.check_file_founding():
                CQT.msgbox(msg_proizv_err)
                return

        ves = F.valm(self.ui.lineEdit_ves.text())
        if ves <= 0:
            CQT.msgbox(f'Вес не может быть 0')
            return
        nom_pu_r = nom_pr.currentText().split('$')[1]
        nom_pr_r = nom_pr.currentText().split('$')[0]
        tablrazr_MK = self.ui.table_razr_MK
        tabl = self.ui.table_zayavk
        if tab2.currentIndex() == 1:

            if tablrazr_MK.rowCount() == 0:
                return
        if tab2.currentIndex() == 0:

            if tabl.rowCount() == 0:
                return


        project = ''
        spisok = ''
        if self.ui.tabWidget_2.currentIndex() == CQT.nom_tab_po_imen(self.ui.tabWidget_2,'Создание МК из *.XML'):
            spisok = CQT.spisok_iz_wtabl(tabl, '', True)
        if self.ui.tabWidget_2.currentIndex() == CQT.nom_tab_po_imen(self.ui.tabWidget_2,'Разработка МК'):
            spisok = CQT.spisok_iz_wtabl(tablrazr_MK, '', True)
        if spisok == '':
            return
        nk_ur = F.nom_kol_po_im_v_shap(spisok,'Уровень')
        nk_naim = F.nom_kol_po_im_v_shap(spisok,'Наименование')
        nk_oboz = F.nom_kol_po_im_v_shap(spisok, 'Обозначение')
        min_ur = 100
        for i in range(1, len(spisok)):
            if int(spisok[i][nk_ur]) < min_ur:
                min_ur = int(spisok[i][nk_ur])
                project = f'{spisok[i][nk_oboz]} {spisok[i][nk_naim]}'
        osnovanie = self.ui.label_ass.text()
        data_sozd = F.date(2)

        prim = prim.text().replace('\n', ' ')
        stroki_strok = [
            [data_sozd, 'Закрыта', project, nom_pu_r.strip(), nom_pr_r.strip(), self.ui.comboBox_vid.currentText(),
             prim, osnovanie, '', '9999',
             self.ui.comboBox_napravlenia.currentText(),
             ves, '', self.kol_izdeliy, '', '', '', 2, '', self.place,'',tip_mk, '']]
        self.xml_head = 1
        #---------------
        CONN, cur = CSQ.connect_bd(self.bd_naryad)
        #CSQ.dob_strok_v_bd_sql(self.bd_naryad, 'mk', stroki_strok, conn=CONN, cur = cur)
        CSQ.zapros(self.bd_naryad,f"""INSERT INTO mk(Дата
            , Статус
            , Номенклатура
            , Номер_заказа
            , Номер_проекта
            , Вид
            , Примечание
            , Основание
            , Прогресс
            , Приоритет
            , Направление
            , Вес
            , xml
            , Количество
            , Статус_ЧПУ
            , Ресурсная
            , Дата_завершения
            , Коэф_парал
            , Обеспечение
            , Место
            , Искл_план_рм
            , Тип
            , Ресурсная_дата) VALUES ({", ".join(['?']*len(stroki_strok[0]))});""",spisok_spiskov=stroki_strok, conn=CONN, cur = cur)

        nom = str(CSQ.posl_strok_bd(self.bd_naryad, 'mk', 'Пномер', ['Пномер'], conn=CONN, cur = cur)[0])

        #CSQ.dob_strok_v_bd_sql(self.bd_naryad, 'zagot', stroki_strok=[[int(nom), '']], s_pervoi=False, conn=CONN, cur =cur)
        CSQ.zapros(self.bd_naryad, """INSERT INTO  zagot(Ном_МК,Прим_резка,Вес_по_рес) VALUES (?,?,?);""", conn=CONN,
                   cur = cur, spisok_spiskov=[[int(nom), '',self.ves_res_list]])
        check_zagot = CSQ.zapros(self.bd_naryad,f"""SELECT * FROM zagot WHERE Ном_МК == {int(nom)}""")
        if len(check_zagot) == 1:
            CSQ.zapros(self.bd_naryad, """INSERT INTO  zagot(Ном_МК,Прим_резка,Вес_по_рес) VALUES (?,?,?);""", conn=CONN,
                       cur=cur, spisok_spiskov=[[int(nom), '',self.ves_res_list]])
            check_zagot = CSQ.zapros(self.bd_naryad, f"""SELECT * FROM zagot WHERE Ном_МК == {int(nom)}""")
            if len(check_zagot) == 1:
                CQT.msgbox(f'Ошибка загрузки МК, не внесена строка в журнал zagot нужно внести вручную')
        if tip_mk == 2:
            CSQ.zapros(self.bd_naryad, """INSERT INTO  дорезки_мк(Номер_мк,Причина) VALUES (?,?);""", conn=CONN,
                       cur=cur, spisok_spiskov=[[int(nom), self.DICT_TIP_DOREZ[self.ui.cmb_tip_dorez.currentText()]]])
        CSQ.close_bd(CONN, cur)

        CONN, cur = CSQ.connect_bd(self.db_resxml)
        #CSQ.dob_strok_v_bd_sql(self.db_resxml, 'xml', stroki_strok=[[int(nom), self.xml_file, self.xml_head]],
        #                       s_pervoi=True, conn=CONN, cur =cur)
        rez = CSQ.zapros(self.db_resxml, f"""SELECT Номер_мк FROM xml WHERE Номер_мк = {int(nom)}""", conn=CONN,
                   cur=cur,one=True)
        if len(rez) == 1:
            CSQ.zapros(self.db_resxml, """INSERT INTO  xml(Номер_мк,data,Head) VALUES (?,?,?);""", conn=CONN,
                       cur=cur, spisok_spiskov=[[int(nom), self.xml_file, self.xml_head]])
        else:
            CSQ.zapros(self.db_resxml, """UPDATE xml SET(Номер_мк,data,Head) = (?,?,?);""", conn=CONN,
                       cur=cur, spisok_spiskov=[[int(nom), self.xml_file, self.xml_head]])
        #CSQ.dob_strok_v_bd_sql(self.db_resxml, 'res', stroki_strok=[[int(nom), res_pickle]], s_pervoi=True,
        #                       conn=CONN, cur =cur)
        rez = CSQ.zapros(self.db_resxml, f"""SELECT Номер_мк FROM res WHERE Номер_мк = {int(nom)}""", conn=CONN,
                         cur=cur, one=True)
        if len(rez) == 1:
            CSQ.zapros(self.db_resxml, """INSERT INTO res(Номер_мк,data) VALUES (?,?);""", conn=CONN,
                       cur=cur, spisok_spiskov=[[int(nom), res_pickle]])
        else:
            CSQ.zapros(self.db_resxml, """UPDATE res SET(Номер_мк,data) = (?,?);""", conn=CONN,
                       cur=cur, spisok_spiskov=[[int(nom), res_pickle]])
        CSQ.close_bd(CONN, cur)

        CSQ.zapros(self.bd_files, """INSERT INTO  MK_founding(Num_mk,file,fio) VALUES (?,?,?);""",
                   spisok_spiskov=[[int(nom), self.mk_file_founding,F.user_name()]])

        # nom = str(CSQ.naiti_v_bd(self.bd_naryad, 'mk',{'Дата':data_sozd},['Пномер'])[0][0])
        if self.ui.tabWidget_2.currentIndex() == 0:
            spisok = CQT.spisok_iz_wtabl(tabl, '', True)
        else:
            spisok = CQT.spisok_iz_wtabl(tablrazr_MK, '', True)

        if F.nalich_file(F.scfg('mk_data') + os.sep + str(nom)) == False:
            F.sozd_dir(F.scfg('mk_data') + os.sep + str(nom))
        dir_tp = F.scfg('mk_data') + os.sep + str(nom)

        for i in range(0, len(spisok)):
            for j in range(9, len(spisok[0])):
                if '\n' in spisok[i][j]:
                    spisok[i][j] = spisok[i][j].replace('\n', '$')
        nom_kol_naim = F.nom_kol_po_im_v_shap(spisok, 'Наименование')
        nom_kol_nn = F.nom_kol_po_im_v_shap(spisok, 'Обозначение')
        for i in range(0, len(spisok)):
            if i > 0:
                naim_dse = spisok[i][nom_kol_naim].strip()
                nn_dse = spisok[i][nom_kol_nn].strip()
                if nn_dse not in self.DICT_DSE_save_mk:
                    CQT.msgbox(f'Не найден номер техкарты {nn_dse}')
                ntk = self.DICT_DSE_save_mk[nn_dse]['Номер_техкарты']
                #ntk = \
                #    CSQ.naiti_v_bd(self.bd_naryad, 'dse', {'Номенклатурный_номер': nn_dse, 'Наименование': naim_dse},
                #                   ['Номер_техкарты'], all=False, conn=CONN)[0]
                rez = F.skopir_file(F.scfg('add_docs') + os.sep + ntk + "_" + nn_dse + '.pickle',
                                    dir_tp + os.sep + ntk + "_" + nn_dse + '.pickle')
                if rez == False:
                    CQT.msgbox(f'Не удалось скопировать файл {ntk + "_" + nn_dse + ".pickle"}, не сохранено')
            spisok[i] = "|".join(spisok[i])
        if self.ui.tabWidget_2.currentIndex() == 1:
            self.clear_mk2()
        else:
            self.ui.table_zayavk.clear()
            self.ui.table_zayavk.setRowCount(0)
        try:
            msg = f"{F.user_full_namre()} СОЗДАЛ МК № {str(nom)}:\n{project} - {self.kol_izdeliy} шт.\n{nom_pu_r.strip()} Проект: {nom_pr_r.strip()}\n" \
                  f"Прим.: {prim} {osnovanie},\nТип: {self.ui.cmb_tip_mk.currentText()}  {self.ui.cmb_tip_dorez.currentText()}"
            self.send_info_mk_b24(msg,'chat41228')
        except:
            print('Ошибка отправки в Б24')
        CQT.msgbox('маршрутная карта ' + str(nom) + ' успешно сохранена')
        try:
            self.ui.tabWidget.setCurrentIndex(CQT.nom_tab_po_imen(self.ui.tabWidget, 'Маршрутные карты'))
            nk_ststus = CQT.nom_kol_po_imen(self.ui.tbl_filtr_mk, 'Статус')
            nk_dat_zav = CQT.nom_kol_po_imen(self.ui.tbl_filtr_mk, 'Дата_завершения')
            self.ui.tbl_filtr_mk.item(0, nk_ststus).setText('Закрыта')
            self.ui.tbl_filtr_mk.item(0, nk_dat_zav).setText('!*')
            CMS.primenit_filtr(self, self.ui.tbl_filtr_mk, self.ui.table_spis_MK)
            self.ui.table_spis_MK.selectRow(self.ui.table_spis_MK.rowCount() - 1)
            self.ui.table_spis_MK.scrollToBottom()
        except:
            pass

    def send_info_mk_b24(self,msg,id):
        conn = CB24.B24(id)
        conn.msg(msg)

    def summ_kol(self, s, i):
        naim = s[i][0].strip()
        nn = s[i][1].strip()
        summ = 0
        for j in range(1, len(s)):
            if s[j][0].strip() == naim and s[j][1].strip() == nn:
                summ += int(s[j][2])
        return summ

    def udal_kol(self, spisok, nom_kol):
        for i in range(0, len(spisok)):
            spisok[i].pop(nom_kol)
        return spisok

    def oformlenie_sp_pod_mk(self, s):
        nk_pki = F.nom_kol_po_im_v_shap(s,'ПКИ')
        nk_uroven = F.nom_kol_po_im_v_shap(s,'Уровень')
        if nk_pki == None:
            CQT.msgbox(f'Не найдена колонка ПКИ')
            return
        if nk_uroven == None:
            CQT.msgbox(f'Не найдена колонка Уровень')
            return
        for i in range(1, len(s)):
            if s[i][nk_pki] != '+':
                s[i][nk_pki] = s[i][nk_pki].replace('0', '')
                s[i][nk_pki] = s[i][nk_pki].replace('1', '+')
        #for j in range(nk_uroven, len(s[0])):
        #    s = self.udal_kol(s, nk_uroven)
        for j in s:
            for i in range(nk_uroven, len(s[0])):
                if '$' in str(j[i]):
                    vrem, oper = [x for x in j[i].split("$")]
                    j[i] = 'Время: ' + vrem + ' мин.' + '\n' + 'Операции:' + '\n' + oper
        i = 12
        while i:
            if i > len(s[0]):
                break
            for j in self.sp_ins:
                s = self.dob_kol(s, i, j)
                i += 1
            i += 1
        return s

    def summa_rc(self, rc):
        s = ''
        if F.is_numeric(rc):
            return int(rc)
        for i in rc:
            if F.is_numeric(i):
                s += str(i)
        s = int(s)
        return s

    def dob_kol(self, spis, nomer, ima):
        for i in range(0, len(spis)):
            if i == 0:
                spis[i].insert(nomer, ima)
            else:
                spis[i].insert(nomer, '')
        return spis

    def poporyadku(self, i, spis, stroka, oper):
        for j in range(i, len(spis[0])):
            if spis[stroka][j] != "":
                arr_tmp_nom = spis[stroka][j].split(';')
                arr_tmp_nom2 = arr_tmp_nom[-1].split('$')
                tmp_nom = arr_tmp_nom2[-1]
                sp_oper = oper.split(';')
                for k in range(len(sp_oper)):
                    if int(tmp_nom) < int(sp_oper[k]):
                        return False

        return True

    def dob_etap(self, spis, rc, vrem, oper, stroka, ogran):
        flag = 0
        for i in range(ogran + 1, len(spis[0])):
            if flag == 1:
                break
            if spis[0][i] == rc and self.poporyadku(i, spis, stroka, oper) == True:
                flag = 1
                spis[stroka][i] = str(vrem) + "$" + oper
                self.ogran = i - 1
                break
            if self.summa_rc(rc) < self.summa_rc(spis[0][i]):
                j = i - 1
                while j >= self.ogran:
                    if self.poporyadku(j + 1, spis, stroka, oper) == False:
                        if j + 2 >= len(spis[0]):
                            spis = self.dob_kol(spis, j + 2, rc)
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
                        self.ogran = j + 1
                        flag = 1
                        break
                    if self.summa_rc(rc) >= self.summa_rc(spis[0][j]):
                        if spis[0][j + 1] != rc:
                            spis = self.dob_kol(spis, j + 1, rc)
                        spis[stroka][j + 1] = str(vrem) + "$" + oper
                        self.ogran = j + 1
                        flag = 1
                        break
                    if j == self.ogran:
                        if spis[0][self.ogran] != rc:
                            spis = self.dob_kol(spis, self.ogran, rc)
                        spis[stroka][self.ogran] = str(vrem) + "$" + oper
                        flag = 1
                        break
                    j -= 1
            else:
                j = i + 1
                while j <= len(spis[0]) - 1:
                    if self.summa_rc(rc) == self.summa_rc(spis[0][j]) and self.poporyadku(j + 1, spis, stroka,
                                                                                          oper) == True:
                        spis[stroka][j] = str(vrem) + "$" + oper
                        self.ogran = j
                        flag = 1
                        break
                    if self.summa_rc(rc) < self.summa_rc(spis[0][j]) and self.poporyadku(j + 1, spis, stroka,
                                                                                         oper) == True:
                        spis = self.dob_kol(spis, j - 1, rc)
                        spis[stroka][j - 1] = str(vrem) + "$" + oper
                        self.ogran = j - 1
                        flag = 1
                        break
                    j += 1
                if flag == 0:
                    spis = self.dob_kol(spis, len(spis[0]), rc)
                    spis[stroka][len(spis[0]) - 1] = str(vrem) + "$" + oper
                    self.ogran = len(spis[0]) - 1
                    flag = 1
                    break

        if flag == 0:
            spis = self.dob_kol(spis, len(spis[0]), rc)
            spis[stroka][len(spis[0]) - 1] = str(vrem) + "$" + oper
            self.ogran = len(spis[0]) - 1
        return spis

    def dob_izd_k_bd(self):
        tree = self.ui.treeWidget
        spis_dse = CQT.spisok_dreva(tree,True)
        #bd = CSQ.spis_iz_bd_sql(self.bd_naryad, 'dse', False, shapka=True)
        n = 0
        m = 0
        nk_ima = F.nom_kol_po_im_v_shap(spis_dse, 'Наименование')
        nk_nn = F.nom_kol_po_im_v_shap(spis_dse, 'Обозначение полное')
        nk_mat = F.nom_kol_po_im_v_shap(spis_dse, 'Масса/М1,М2,М3')
        nk_adres = F.nom_kol_po_im_v_shap(spis_dse, 'Ссылка на объект DOCs')
        nk_klass = F.nom_kol_po_im_v_shap(spis_dse, 'Классификатор изделия')
        nk_kod_erp = F.nom_kol_po_im_v_shap(spis_dse, 'Код ERP')
        nk_prim = F.nom_kol_po_im_v_shap(spis_dse, 'Примечание')
        nk_razdel = F.nom_kol_po_im_v_shap(spis_dse, 'Раздел')
        nk_tip = F.nom_kol_po_im_v_shap(spis_dse, 'Тип')
        conn, cur = CSQ.connect_bd(self.db_dse)
        for i in range(1, len(spis_dse)):
            bd_tmp = []
            ima = F.ochist_strok_pod_ima_fila(spis_dse[i][nk_ima])
            nn = F.ochist_strok_pod_ima_fila(spis_dse[i][nk_nn])
            adres = spis_dse[i][nk_adres]
            mat = spis_dse[i][nk_mat]
            klass = spis_dse[i][nk_klass]
            kod_erp = spis_dse[i][nk_kod_erp]
            tip = spis_dse[i][nk_tip]
            if tip in self.TIP_NEGRUZ_DSE:
                continue
            if spis_dse[i][nk_prim] != '':
                prim = f'(ОГК: {spis_dse[i][nk_prim]})'
            else:
                prim = ''
            #zapros = f'SELECT * FROM dse WHERE Номенклатурный_номер = "{nn}" AND Наименование = "{ima}"'
            #query = CSQ.naiti_v_bd(self.db_dse, siroy_zapros=zapros, conn=conn, cur=cur, shapka=True, all=False)
            query = CSQ.zapros(self.db_dse,f'SELECT * FROM dse WHERE Номенклатурный_номер = "{nn}" '
                                           f'AND Наименование = "{ima}"', conn=conn, cur=cur, shapka=True, one=True)
            if len(query) == 1:
                bd_tmp.append([nn, ima, '', prim, adres, '', '', '', '', '', '', mat, kod_erp, klass])
                #CSQ.dob_strok_v_bd_sql(self.db_dse, 'dse', bd_tmp, conn=conn, cur=cur)
                CSQ.zapros(self.db_dse,f"""INSERT INTO dse(Номенклатурный_номер
                    ,Наименование
                    ,Номер_техкарты
                    ,Примечание
                    ,Путь_docs
                    ,Доступ
                    ,Процесс
                    ,Нормы
                    ,Материалы
                    ,Тех_заметки
                    ,Теги
                    ,Мат_кд
                    ,Код_ЕРП
                    ,Классификатор) VALUES ({', '.join(['?']*len(bd_tmp[-1]))});""", spisok_spiskov=bd_tmp, conn=conn, cur=cur)
                n += 1
            else:
                if prim != query[1][F.nom_kol_po_im_v_shap(query, 'Примечание')]:
                    #CSQ.update_bd_sql(self.db_dse, 'dse', {'Примечание': prim},
                    #                  {'Пномер': query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}, conn=conn, cur=cur)
                    CSQ.zapros(self.db_dse,f"""UPDATE dse SET Примечание = '{prim}' WHERE
                     Пномер = {query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}; """, conn=conn, cur=cur)
                    m += 1
                if adres != query[1][F.nom_kol_po_im_v_shap(query, 'Путь_docs')]:
                    #CSQ.update_bd_sql(self.db_dse, 'dse', {'Путь_docs': adres},
                    #                  {'Пномер': query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}, conn=conn, cur=cur)
                    CSQ.zapros(self.db_dse, f"""UPDATE dse SET Путь_docs = '{adres}' WHERE
                                         Пномер = {query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}; """, conn=conn,
                               cur=cur)
                    m += 1
                if mat != query[1][F.nom_kol_po_im_v_shap(query, "Мат_кд")]:
                    #CSQ.update_bd_sql(self.db_dse, 'dse', {'Мат_кд': mat},
                    #                  {'Пномер': query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}, conn=conn, cur=cur)
                    CSQ.zapros(self.db_dse, f"""UPDATE dse SET Мат_кд = '{mat}' WHERE
                                                             Пномер = {query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}; """,
                               conn=conn,
                               cur=cur)
                    m += 1
                if klass != query[1][F.nom_kol_po_im_v_shap(query, "Классификатор")]:
                    #CSQ.update_bd_sql(self.db_dse, 'dse', {'Классификатор': klass},
                    #                  {'Пномер': query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}, conn=conn, cur=cur)
                    CSQ.zapros(self.db_dse, f"""UPDATE dse SET Классификатор = '{klass}' WHERE
                                               Пномер = {query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}; """,
                               conn=conn,
                               cur=cur)
                    m += 1
                if kod_erp != query[1][F.nom_kol_po_im_v_shap(query, "Код_ЕРП")]:
                    #CSQ.update_bd_sql(self.db_dse, 'dse', {'Код_ЕРП': kod_erp},
                    #                  {'Пномер': query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}, conn=conn, cur=cur)
                    CSQ.zapros(self.db_dse, f"""UPDATE dse SET Код_ЕРП = '{kod_erp}' WHERE
                                                Пномер = {query[1][F.nom_kol_po_im_v_shap(query, 'Пномер')]}; """,
                               conn=conn,
                               cur=cur)
                    m += 1
        CSQ.close_bd(conn, cur)
        if n == 0:
            CQT.msgbox(f'Новых ДСЕ не добавлено, обновлено {m} значений')
        else:
            CQT.msgbox(f'Добавлено ' + str(n) + f' ед. ДСЕ, обновлено {m} значений')

    def viborXML(self):
        vklad = self.ui.tabWidget
        tabl = self.ui.table_zayavk
        tree = self.ui.treeWidget
        tab = self.ui.tabWidget


        if vklad.currentIndex() == CQT.nom_tab_po_imen(vklad,'Маршрутные карты') :
            xml = ''
            tbl = self.ui.table_spis_MK
            row = tbl.currentRow()
            if row == -1:
                return
            nk_pnom = CQT.nom_kol_po_imen(tbl,'Пномер')
            nom_mk =int(tbl.item(row,nk_pnom).text())
            try:
                query = f'''SELECT data, Head FROM xml 
                                   WHERE Номер_мк == {nom_mk}
                                               '''
                rez_xml = CSQ.zapros(self.db_resxml, query)
                xml = rez_xml[-1][0]
                xml_head = rez_xml[-1][1]
                if xml != '':
                    xml = XML.spisok_iz_xml(str_f=xml)
                    self.ui.tabWidget.setCurrentIndex(0)
                putt = ''

            except:
                pass
        else:
            if tab.currentIndex() > 2:
                self.ui.tabWidget.setCurrentIndex(0)
            tmp_putt = CMS.load_tmp_path("tmp_putt")
            putt = CQT.f_dialog_name(self, 'Выбрать XML', tmp_putt, "Файлы *.xml")
            if putt == '' or putt == '.':
                return

            CMS.save_tmp_path("tmp_putt", putt, True)

            xml = XML.spisok_iz_xml(putt)

        if vklad.currentIndex() == 0:
            spis_xml = self.podgotovka_xml(xml, show_negruz=True)
            if spis_xml == None:
                CQT.msgbox('Файл не корректный')
                return
            err_flag = False
            msg_text = ''
            for i in range(len(spis_xml)):
                if 'Тип' not in spis_xml[i]['data']:
                    err_flag = True
                    msg_text = f'Отсутствует поле Тип'
                if spis_xml[i]['data']['Наименование'] == "" and spis_xml[i]['data']['Обозначение полное'] == "":
                    err_flag = True
                    msg_text = f'Наименование  и  Обозначение полное ПУСТО'
                if spis_xml[i]['data']['Количество'] == "" or spis_xml[i]['data']['Количество на изделие'] == "":
                    msg_text = f'Количество  и  Количество на изделие ПУСТО'
                    err_flag = True
            if err_flag == True:
                CQT.msgbox(f'Файл XML {putt} имеет ошибки \n{msg_text}\n работать с ним нельзя!')
            if err_flag == True:
                self.ui.pushButton_add_v_bd.setEnabled(False)
                self.ui.pushButton_add_v_MK.setEnabled(False)
                return
            else:
                self.ui.pushButton_add_v_bd.setEnabled(True)
                self.ui.pushButton_add_v_MK.setEnabled(True)

            list_user = self.load_tree(spis_xml)
            self.zapoln_tree_spiskom(spis_xml,list_user)
            for _ in range(0, 8):
                tree.resizeColumnToContents(_)
        if vklad.currentIndex() == 1:
            spis_xml = self.podgotovka_xml(xml, show_negruz=False)
            if spis_xml == None:
                CQT.msgbox('Файл не корректный')
                return
            err_flag = False
            for i in range(len(spis_xml)):
                if spis_xml[i]['data']['Наименование'] == "" and spis_xml[i]['data']['Обозначение полное'] == "":
                    err_flag = True
                if spis_xml[i]['data']['Количество'] == "" or spis_xml[i]['data']['Количество на изделие'] == "":
                    err_flag = True
            if err_flag == True:
                CQT.msgbox(f'Файл XML {putt} имеет ошибки, работать с ним нельзя!')
            tabl.setSelectionBehavior(0)
            if err_flag == True:
                return
            self.dob_izd(spis_xml, putt)


    def podgotovka_xml(self,spis_xml:list,xml_head = '',show_negruz=False):
        if spis_xml == None:
            return
        rez = []
        if xml_head == '':

            self.xml_head = 0
        else:
            self.xml_head = xml_head

        for i in range(len(spis_xml)):
            if spis_xml[i]['data']['Покупное изделие'] == '1':
                if spis_xml[i]['data']['Обозначение полное'].strip() == '':
                    spis_xml[i]['data']['Обозначение полное'] = F.shifr(spis_xml[i]['data']['Наименование'])[:13]
            else:
                if spis_xml[i]['data']['Обозначение полное'].strip() == '':
                    CQT.msgbox(
                        f"Ошибка {spis_xml[i]['data']['Наименование']} {spis_xml[i]['data']['Обозначение полное']} не имеет Обозначение/не покупная")
                    return
            if 'Классификатор изделия' in spis_xml[i]['data']:
                if spis_xml[i]['data']['Классификатор изделия'] == None:
                    spis_xml[i]['data']['Классификатор изделия'] = ''
            if 'Код ERP' in spis_xml[i]['data']:
                if spis_xml[i]['data']['Код ERP'] == None:
                    spis_xml[i]['data']['Код ERP'] = ''

            mat = "/".join(
                (str(spis_xml[i]['data']['Масса']).replace(',', '.'), F.ochist_strok_pod_ima_fila(str(spis_xml[i]['data']['Материал'])),
                 F.ochist_strok_pod_ima_fila(str(spis_xml[i]['data']['Материал2'])),
                 F.ochist_strok_pod_ima_fila(str(spis_xml[i]['data']['Материал3']))))
            spis_xml[i]['data']['Масса/М1,М2,М3'] = mat
            spis_xml[i]['data'].pop('Материал', None)
            spis_xml[i]['data'].pop('Материал2', None)
            spis_xml[i]['data'].pop('Материал3', None)
            spis_xml[i]['data']['Наименование'] = F.ochist_strok_pod_ima_fila(spis_xml[i]['data']['Наименование'])
            spis_xml[i]['data']['Обозначение полное'] = F.ochist_strok_pod_ima_fila(spis_xml[i]['data']['Обозначение полное'])
            if 'Сводное наименование' in spis_xml[i]['data']:
                spis_xml[i]['data']['Сводное наименование'] = F.ochist_strok_pod_ima_fila(
                spis_xml[i]['data']['Сводное наименование'])


            if 'Тип' in spis_xml[i]['data']:
                if show_negruz:
                    rez.append(spis_xml[i])
                else:
                    if spis_xml[i]['data']['Тип'] not in self.TIP_NEGRUZ_DSE:
                        rez.append(spis_xml[i])
                    else:
                        tek_ur = spis_xml[i]['uroven']
                        if i == 0 :
                            pred_ur = -1
                        else:
                            pred_ur = spis_xml[i-1]['uroven']
                        delta_ur = tek_ur-pred_ur
                        for j in range(i+ 1 ,len(spis_xml)):
                            if spis_xml[j]['uroven']<= tek_ur:
                                break
                            spis_xml[j]['uroven'] -= delta_ur
            else:
                rez.append(spis_xml[i])
        return rez

    def nalich_dannih_v_tk(self, tk, nomer_st, conn='',nomenklatura = '',DICT_doc_reestr = ''):
        flag = 0
        for i in range(10, len(tk)):
            if len(tk[i]) == 21:
                if tk[i][20] == '0' and flag == 1:
                    return True
                if tk[i][20] == '0' and flag == 0:
                    flag = 1
                if tk[i][20] == '1' and nomer_st == 7:
                    if tk[i][nomer_st] == "" or F.is_numeric(tk[i][nomer_st]) == False:
                        return 'norma_vr'

                if tk[i][20] == '1':
                    if nomer_st == 0:
                        if tk[i][10] == '':
                            if tk[i][0] == 'Отрезка слесарная' and tk[i][8] == '19149':
                                return False
                            if tk[i][0] == 'Отрезка(гильотина)':
                                return False
                            if tk[i][0] == 'Отрезка(лентопил)':
                                return False


                    if nomer_st == 4:
                        if tk[i][15] != '':
                            if tk[i][15] not in DICT_doc_reestr:
                                return 'docs'
                        if tk[i][nomer_st] == "010101":
                            if tk[i][0] != 'Резка(ЧПУ)':
                                return "rc"
                            else:
                                if tk[i][15] == '':
                                    return 'dxf'
                                else:
                                    if '.dxf' not in tk[i][15]:
                                        return 'dxf'
                                if i < len(tk) - 2:
                                    if tk[i + 1][20] != '2':
                                        return 'seg'

                    if tk[i][nomer_st] == 'Резка(ЧПУ)' and nomer_st == 0:
                        if tk[i][10] == '':
                            return False

                        nk_nn_nom = F.nom_kol_po_im_v_shap(nomenklatura, 'Код')
                        nk_sort_nom = F.nom_kol_po_im_v_shap(nomenklatura, 'П5')
                        nk_kod_cam = F.nom_kol_po_im_v_shap(nomenklatura, 'П6')
                        kod_mat = ''
                        material = '?'
                        nn_mat = '?'
                        sp_mat = tk[i][10].split('{')
                        for material in sp_mat:
                            if kod_mat != '':
                                break
                            nn_mat = material.split('$')[0]
                            for k in range(1, len(nomenklatura)):
                                if nomenklatura[k][nk_nn_nom] == nn_mat:
                                    if nomenklatura[k][nk_sort_nom] == '1':
                                        kod_mat = nomenklatura[k][nk_kod_cam]
                                        break
                        if kod_mat == '':
                            CQT.msgbox(f'Не найден код CAM в номенклатуре для {material} на {nn_mat}')
                            return False
        if flag == 1:
            return True
        else:
            return False

    def rc_n_k(self, ntk, nn):
        tk = F.otkr_f(F.scfg('add_docs') + os.sep + ntk + "_" + nn + '.txt', False, "|")
        nachalo = ''
        konec = ''
        if len(tk) < 11:
            return [nachalo, konec, '', '']
        for i in range(10, len(tk)):
            if tk[i][20] == '0':
                if i == len(tk) - 1:
                    return [nachalo, konec, '', '']
                for j in range(i + 1, len(tk)):
                    if tk[j][20] == '1':
                        if nachalo == '':
                            nachalo = tk[j][4]
                            nach_op = tk[j][0]
                        konec = tk[j][4]
                        kon_op = tk[j][0]
                    if tk[j][20] == '0':
                        break
        return [nachalo, konec, nach_op, kon_op]

    def ispravit_koncovky_tehkart(self, n_dse, n_tk, new_rc):
        tk = F.otkr_f(F.scfg('add_docs') + os.sep + n_tk + "_" + n_dse + '.txt', False, "|")
        f_naid = 0
        nom_st_op = 0
        for i in range(len(tk)):
            if len(tk[i]) == 21:
                if tk[i][20] == '0' and f_naid == 1:
                    break
                if tk[i][20] == '0' and f_naid == 0:
                    f_naid = 1
                if tk[i][20] == '1' and f_naid == 1:
                    nom_st_op = i
        if nom_st_op == 0:
            CQT.msgbox(f'Не удалось найти последнюю операцию в техкарте {n_tk}')
            return
        tk[nom_st_op][4] = new_rc
        F.zap_f(F.scfg("add_docs") + os.sep + n_tk + "_" + n_dse + '.txt', tk, pickl=True, separ='|')
        return True

    def nalich_tk(self, spisok):
        s_bd = []
        spis_rc = CSQ.zapros(self.db_users,"""SELECT * FROM rab_c""")
        zapros = f'''SELECT * FROM nomen WHERE П5 == "1" '''
        nomenklatura = CSQ.zapros(self.bd_nomen, zapros)
        DICT_NN_NTK = CMS.load_dict_dse(self.db_dse)

        zapros = f""" SELECT Пномер, file_name FROM t_kards"""
        query = CSQ.zapros(self.bd_files, zapros, '', shapka=True, rez_dict=True)
        if query == False:
            CQT.msgbox(f'ОШИбка получения данных файлов с БД')
            return
        DICT_doc_reestr = F.raskrit_dict(query, 'file_name')

        for i in range(len(spisok)):
            print(f'{i} из {len(spisok)}')
            CQT.statusbar_text(self,f'{i} из {len(spisok)}')
            ima = spisok[i]['data']['Наименование']
            nn = spisok[i]['data']['Обозначение полное']
            flag_bd = 0
            flag_tk = 0
            flag_marsh = 0
            flag_vrema = 0
            flag_mat = 0
            flag_rc = 1
            flag_dxf = 1
            flag_seg = 1
            flag_docs = 1
            flag_rashodnik = 0
            if type(F.valm(spisok[i]['data']['Количество'])) == type(1.1):
                flag_rashodnik = 1
            nom_tk = ''
            if nn not in DICT_NN_NTK:
                CQT.msgbox(f'Не найдена {nn} в БД')
                return
            query_dse = DICT_NN_NTK[nn]
            if len(query_dse) > 1:
                nom_tk = query_dse["Номер_техкарты"]
                flag_bd = 1
                if nom_tk != '':
                    try:
                        tk = F.otkr_f(F.scfg('add_docs') + os.sep + nom_tk + "_" + nn + '.txt', False, separ='|',
                                      pickl=True)
                        flag_tk = 1
                        if self.nalich_dannih_v_tk(tk, 4,nomenklatura=nomenklatura, DICT_doc_reestr=DICT_doc_reestr) == True:
                            flag_marsh = 1
                        if self.nalich_dannih_v_tk(tk, 6,nomenklatura=nomenklatura, DICT_doc_reestr=DICT_doc_reestr) == True and \
                                self.nalich_dannih_v_tk(tk, 7,nomenklatura=nomenklatura, DICT_doc_reestr=DICT_doc_reestr) == True:
                            flag_vrema = 1
                        if self.nalich_dannih_v_tk(tk, 0,nomenklatura=nomenklatura, DICT_doc_reestr=DICT_doc_reestr) == True:
                            flag_mat = 1
                        rez = self.nalich_dannih_v_tk(tk, 4, conn='',nomenklatura=nomenklatura, DICT_doc_reestr=DICT_doc_reestr)
                        if rez == 'rc':
                            flag_rc = 0
                        if rez == 'dxf':
                            flag_dxf = 0
                        if rez == 'seg':
                            flag_seg = 0
                        if rez == 'docs':
                            flag_docs = 0
                    except:
                        CQT.msgbox(f'Что то не то с ТК {nom_tk}')
                        return
            if flag_bd == 0:
                s_bd.append('нет в базе ' + " " + nn + ' ' + ima)
            if flag_tk == 0:
                s_bd.append('нет техкарты ' + " " + nn + " " + ima)
            else:
                nachalo, konec, nach_op, kon_op = self.rc_n_k(nom_tk, nn)
                spisok[i]['tk'] = dict()
                spisok[i]['tk']['nach_op'] = nach_op
                spisok[i]['tk']['kon_op'] = kon_op
                spisok[i]['tk']['nachalo'] = nachalo
                spisok[i]['tk']['konec'] = konec
                spisok[i]['tk']['nom_tk'] = nom_tk
                #spisok[i][15] = nach_op
                #spisok[i][16] = kon_op
                #spisok[i][18] = nachalo
                #spisok[i][19] = konec
                #spisok[i][17] = nom_tk
            if flag_marsh == 0:
                s_bd.append('нет маршрутов в тк ' + " " + nn + " " + ima)
            if flag_vrema == 0:
                s_bd.append('нет/не корректное времени в тк ' + " " + nn + " " + ima)
            if flag_mat == 0:
                s_bd.append('не корректно занесен материал на операцию в тк ' + " " + nn + " " + ima)
            if flag_rc == 0:
                s_bd.append(
                    'не корректно занесено имя операции на РЦ 010101, должно быть Резка(чпу) в тк ' + " " + nn + " " + ima)
            if flag_dxf == 0:
                s_bd.append('не занесен DXF на РЦ 010101 где Резка(чпу) в тк ' + " " + nn + " " + ima)
            if flag_seg == 0:
                s_bd.append('не занесено число сегментов на РЦ 010101 где Резка(чпу) в тк ' + " " + nn + " " + ima)
            if flag_docs == 0:
                s_bd.append('отсутствует в бд файл вложения, прикрепелнный в тк ' + " " + nn + " " + ima)
            if flag_rashodnik == 1:
                s_bd.append(f'{nn} {ima} занесен как расходник, материалы в структуре не допустимы')


        for i in range(len(spisok)):
            ur = spisok[i]['uroven']
            ur2 = int(ur) + 1
            if i + 1 > len(spisok) - 1:
                break
            #print(f'{i}--')
            for j in range(i + 1, len(spisok)):
                if spisok[j]['uroven'] < ur2:
                    break
                if spisok[j]['uroven'] == ur2:
                    #print(f'{j}++')
                    if 'tk' in spisok[i] and 'tk' in spisok[j]:
                        
                        if spisok[i]['tk']['nachalo'] != spisok[j]['tk']['konec']:
                            frase = f'Не совдают концовки \n{spisok[i]["data"]["Наименование"]} {spisok[i]["data"]["Наименование"]} (Операция <<{spisok[i]["tk"]["nach_op"]}>> РЦ {spisok[i]["tk"]["nachalo"]}-{CMS.ima_rc_po_kod(spis_rc, spisok[i]["tk"]["nachalo"])})\n ' \
                                    f'и \n{spisok[j]["data"]["Наименование"]} {spisok[j]["data"]["Наименование"]} (Операция <<{spisok[j]["tk"]["kon_op"] }>> РЦ {spisok[j]["tk"]["konec"]}-{CMS.ima_rc_po_kod(spis_rc, spisok[j]["tk"]["konec"])})'
                            rez = CQT.msgboxgYN(frase + f"\n\nВыполнить корректировку концовки для {spisok[j]['data']['Наименование']} "
                                                        f"{spisok[j]['data']['Обозначение полное']}"
                                                        f" для операции <<{spisok[j]['tk']['kon_op']}>>?\n\n"
                                                        f"    РЦ {spisok[j]['tk']['konec']}-{CMS.ima_rc_po_kod(spis_rc, spisok[j]['tk']['konec'])} \n\nбудет заменен на \n\n"
                                                        f"        {spisok[i]['tk']['nachalo']}-{CMS.ima_rc_po_kod(spis_rc, spisok[i]['tk']['nachalo'])}")
                            if rez == True:
                                rez2 = self.ispravit_koncovky_tehkart(spisok[j]['data']['Обозначение полное'], spisok[j]['tk']['nom_tk'], spisok[i]['tk']['nachalo'])
                                if rez2 == None:
                                    s_bd.append(frase)
                            else:
                                s_bd.append(frase)
                    else:
                        s_bd.append(f' не сравнить концовки на'
                                    f' {spisok[i]["data"]["Наименование"]} {spisok[i]["data"]["Обозначение полное"]}'
                                    f' и {spisok[j]["data"]["Наименование"]} {spisok[j]["data"]["Обозначение полное"]}')
        CQT.statusbar_text(self)
        return s_bd

    def dob_izd(self, spisok, putt):
        sp_tk = self.nalich_tk(spisok)
        if sp_tk == None:
            return
        if len(sp_tk) > 0:
            viv = ''
            for i in sp_tk:
                viv += i + '\n'
            F.copy_bufer(viv)
            CQT.msgbox("Скопировано в буфер:" + '\n' + viv)
            return
        tabl = self.ui.table_zayavk
        if tabl.columnCount() > 5:
            tabl.clear()
            tabl.clearContents()
            shapka = ['Файл', 'Изделие', 'Количество']
            tabl.setColumnCount(3)
            tabl.setRowCount(0)
            tabl.setHorizontalHeaderLabels(shapka)
        s = CQT.spisok_iz_wtabl(tabl, '', True)

        kol_po_zayavke = ''
        s.append([putt,f"{spisok[0]['data']['Обозначение полное']} {spisok[0]['data']['Наименование']}" , kol_po_zayavke])
        edit = {2}
        CQT.zapoln_wtabl(self, s, tabl, 0, edit, (), (), 2200, True, "")

        try:
            if 'Отдел технолога\В работе' in putt:
                spis_put = putt.split('\\')
                for i, item in enumerate(spis_put):
                    if item == 'В работе':
                        np = spis_put[i + 2]
                        py = spis_put[i + 3]
                        break
        except:
            return
        try:
            for i in range(self.ui.comboBox_np.count()):
                if f'{np}${py}' in self.ui.comboBox_np.itemText(i):
                    self.ui.comboBox_np.setCurrentIndex(i)
                    self.vibor_pr()
                    break
        except:
            return




    def zapoln_tree_spiskom(self, spisok:list,list_user):
        tree = self.ui.treeWidget
        tree.clear()
        n = 0
        max_ur = 0
        for i in range(0, len(spisok)):
            if spisok[i]['uroven'] > max_ur:
                max_ur = spisok[i]['uroven']

        list_obj_lvls = ["" for _ in range(max_ur+1)]
        nk_naim = list_user.index('Наименование')
        nk_tip = list_user.index('Тип')

        for i in range(0, len(spisok)):
            ur = spisok[i]['uroven']
            print(i)
            if ur == 0:
                list_obj_lvls[ur] = QtWidgets.QTreeWidgetItem(tree)
            else:
                list_obj_lvls[ur] = QtWidgets.QTreeWidgetItem(list_obj_lvls[ur-1])
            root = list_obj_lvls[ur]

            for pole in range(0,len(list_user)):
                if list_user[pole] in spisok[i]['data'].keys():
                    root.setText(pole, str(spisok[i]['data'][list_user[pole]]))
                    #if list_user[pole] == 'Тип':
                    #    if str(spisok[i]['data'][list_user[pole]]) in self.TIP_NEGRUZ_DSE:
                    #        root.setTextColor(nk_naim, QtGui.QColor(222,111,111))

                else:
                    if list_user[pole] == 'Уровень':
                        root.setText(pole, str(ur))
                    else:
                        root.setText(pole, '')


            tree.addTopLevelItem(root)
            tree.expandItem(root)
            tree.setCurrentItem(root)
            n += 1
        CQT.cveta_v_drevo(tree,self.TIP_NEGRUZ_DSE,nk_tip, 222,111,111,255)


    def add_v_planetapi(self):
        if self.BD == None:
            CQT.msgbox('Не обновлен конфиг')
            return
        tbl = self.ui.table_spis_MK
        r = tbl.currentRow()

        if r == -1:
            return
        spis = []
        for i in range(tbl.columnCount()):
            try:
                if tbl.item(r, i) == None:
                    spis.append('')
                else:
                    spis.append(tbl.item(r, i).text())
            except:
                spis.append('')
        rez = [spis[1], spis[5], spis[4], spis[11], spis[6], spis[3], spis[3], int(spis[13]), '', '', 'К производству',
               F.user_name(), spis[1], spis[7], spis[0]]
        CMS.btn_oyp_add_project(self, [rez])


    def sl_mash_change(self, val):
        self.val_masht = val
        CMS.save_tmp_path('mk_val_masht',str(self.val_masht))
        if self.kpl_mode == 0:
            GKPL.oforml_table(self,self.ui.tbl_preview)
        else:
            GKPL.oforml_table(self, self.ui.tbl_pl_gaf, self.ui.tbl_pl_gaf_filtr)
            GVKPL.load_svod(self)

app = QtWidgets.QApplication([])

myappid = 'Powerz.BAG.SustControlWork.0.0.0'  # !!!

QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
app.setWindowIcon(QtGui.QIcon(os.path.join("icons", "icon.png")))

S = F.scfg('Stile').split(",")
app.setStyle(S[0])

application = mywindow()
# =============================================================
if CMS.kontrol_ver(application.versia,'МКарты') == False:
    quit()
# =============================================================

application.showMaximized()

sys.exit(app.exec())
# pyinstaller.exe --onefile --icon=1.ico --noconsole MKart.py
