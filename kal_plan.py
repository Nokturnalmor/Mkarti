import copy

import project_cust_38.Cust_Functions as F
import project_cust_38.Cust_SQLite as CSQ
import project_cust_38.Cust_Qt as CQT
import project_cust_38.Cust_mes as CMS
from PyQt5.QtCore import QDate
import gui_kal_plan as GPL
import gui_vol_plan as VPL


LIST_FREEZE_FIELDS = ['plan.Пномер','plan.Направление_деятельности','plan.local_graf']
LIST_HIDE_FIELDS = ['plan.local_graf']



def select_row(self):
    GPL.update_local_graf(self)


def load_gui(self):
    show_fr(self)
    load_table_db(self)
    self.kpl_mode = 0#объемный выключен

def btn_pl_add_poz_click(self):
    if self.regim == '':
        show_fr(self, 'tbl_add')
        load_tbl_add_new_poz(self)
        self.regim = 'add'
    else:
        show_fr(self)
        self.regim = ''

def btn_pl_mode(self):
    if self.kpl_mode == 0:#объемный выключен
        show_fr(self,graf=1)#объемный включаем
        self.kpl_mode = 1#объемный включен

        VPL.load_tbl_gant(self)#объемный загрузка

    else:
        load_gui(self)


def kal_pl_left(self):
    tbl = self.ui.tbl_pl_add_poz
    column = tbl.currentColumn()
    if column == None or column == -1 or column == 0:
        return
    spis = CQT.spisok_iz_wtabl(tbl,shapka=True)
    spis_new = copy.deepcopy(spis)
    spis_new[0].pop(column)
    spis_new[1].pop(column)
    spis_new[0].insert(column-1,spis[0][column])
    spis_new[1].insert(column-1,spis[1][column])
    fill_tbl_settings(self,spis_new)
    tbl.selectColumn(column - 1)

def kal_pl_right(self):
    tbl = self.ui.tbl_pl_add_poz
    column = tbl.currentColumn()
    spis = CQT.spisok_iz_wtabl(tbl, shapka=True)
    if column == None or column == -1 or column == len(spis[0])-1:
        return
    spis_new = copy.deepcopy(spis)
    spis_new[0].pop(column)
    spis_new[1].pop(column)
    spis_new[0].insert(column+1,spis[0][column])
    spis_new[1].insert(column+1,spis[1][column])
    fill_tbl_settings(self,spis_new)
    tbl.selectColumn(column + 1)

def fill_tbl_settings(self,list_conf):
    def check_val(self,checked,row,col):
        self.ui.tbl_pl_add_poz.item(row,col).setText(str(int(checked)))
    CQT.fill_wtabl(list_conf, self.ui.tbl_pl_add_poz)
    for j in range(self.ui.tbl_pl_add_poz.columnCount()):
        val = 1
        if list_conf[-1][j] != 1:
            val = 0
        CQT.add_check_box(self.ui.tbl_pl_add_poz, 0, j, val=val, conn_func_checked_row_col=check_val, self=self)

def btn_pl_settings(self):


    if self.regim == '':
        show_fr(self, 'tbl_add')
        self.ui.btn_kal_pl_left.setHidden(False)
        self.ui.btn_kal_pl_right.setHidden(False)
        db, list_conf = load_db(self)
        fill_tbl_settings(self,list_conf)
        self.regim = 'cnf'
    else:
        show_fr(self)
        self.regim = ''


def create_list_fields(self):
    'Загрузка всех полей с БД'
    list_tables = ['plan']
    tables = [_ for _ in CSQ.spis_tablic(self.db_kplan) if 'пл_' in _]

    for i in range(len(self.DICT_PODR)):
        for table in tables:
            if table in self.DICT_PODR:
                if self.DICT_PODR[table]['Порядок'] == i:
                    list_tables.append(table)
    for table in tables:
        if table not in self.DICT_PODR:
            list_tables.append(table)
    list_fields = []
    for table in list_tables:
        fields = CSQ.spisok_colonok(self.db_kplan, table)
        for field in fields:
            list_fields.append(f'{table}.{field}')
    tmp_list = [1 for _ in list_fields]
    return [list_fields,tmp_list]



def load_list_fields(self, all = False):
    """Приостановка отключенных полей из конфига"""
    path ='Config\\fields.pickle'
    table = create_list_fields(self)
    if F.nalich_file(path) and all== False:
        dict_cnf = F.load_file_pickle(path)
        tmp_list = [['n','fied']]
        for i in range(len(table[0])):
            if table[0][i] in dict_cnf:
                tmp_list.append([dict_cnf[table[0][i]]['order'],table[0][i]])
        tmp_list = F.sort_po_kol(tmp_list,'n')
        for i in range(len(table[0])):
            if table[0][i] not in dict_cnf:
                tmp_list.append([tmp_list[-1][0]+1,table[0][i]])
        table = [[],[]]
        for i in range(1,len(tmp_list)):
            table[0].append(tmp_list[i][1])
            if tmp_list[i][1] in dict_cnf:
                table[1].append(dict_cnf[tmp_list[i][1]]['hidden'])
            else:
                table[1].append(1)
        return table
    else:
        tmp_list = [['n', 'fied']]
        for i in range(len(table[0])):
            tmp_list.append([i, table[0][i]])
        table = [[], []]
        for i in range(1, len(tmp_list)):
            table[0].append(tmp_list[i][1])
            table[1].append(1)
    return table


def btn_pl_edit_poz_click(self):
    if self.regim == '':
        show_fr(self, 'tbl_edit')
        load_tbl_edit_poz(self)
        self.regim = 'edit'
    else:
        show_fr(self)
        self.regim = ''

def oform_table_editeble(self, tbl,name_field):
    for i in range(tbl.columnCount()):
        if 'дата' in tbl.horizontalHeaderItem(i).text().lower() or name_field.lower() == self.ui.tbl_pl_add_poz.horizontalHeaderItem(i).text().lower().split('.')[-1]:
            CQT.set_cell_editable(tbl,0,i,False)
            CQT.ust_color_wtab(tbl,0,i,230,230,230)


def load_tbl_edit_poz(self):
    list_podr = [_ for _ in CSQ.spis_tablic(self.db_kplan) if 'пл_' in _]
    list_podr.append('plan')
    list_podr.append('')
    list_podr.sort()
    self.ui.cmb_etap.clear()
    self.ui.cmb_etap.addItems(list_podr)
    CQT.clear_tbl(self.ui.tbl_pl_add_poz)

def select_etap_edit(self):
    def edit_tabel(self):
        month = self.ui.cmb_etap.currentText()
        if month == '':
            return
        list_month = CSQ.zapros(self.db_kplan,f"""SELECT * FROM {month}""")
        set_editeble_columns = set()
        for i in range(len(list_month[0])):
            if F.is_date(list_month[0][i],"d_%Y_%m_%d"):
                set_editeble_columns.add(i)
        CQT.fill_wtabl(list_month,self.ui.tbl_pl_add_poz,set_editeble_col_nomera = set_editeble_columns,colorful_edit=True)


    def edit_etap(self):
        podr = self.ui.cmb_etap.currentText()
        tbl_pl = self.ui.tbl_kal_pl
        row = tbl_pl.currentRow()
        if row == None or row == -1:
            return
        if podr == "":
            CQT.clear_tbl(self.ui.tbl_pl_add_poz)
            return
        name_field = 'НомПл'
        if podr == "plan":
            name_field = 'Пномер'
        nk_pnom = int(CQT.nom_kol_po_imen(tbl_pl, 'plan.Пномер'))
        pnom = tbl_pl.item(row, nk_pnom).text()
        list_itog = get_line_to_edit_podr(self, pnom)
        CQT.fill_wtabl(list_itog, self.ui.tbl_pl_add_poz, auto_type=False)
        for field in LIST_HIDE_FIELDS:
            if field.split('.')[0] == podr:
                nk = CQT.nom_kol_po_imen(self.ui.tbl_pl_add_poz, field.split('.')[1])
                if nk != None:
                    self.ui.tbl_pl_add_poz.setColumnHidden(nk, True)
        oform_table_editeble(self, self.ui.tbl_pl_add_poz, name_field)
        if podr == 'plan':
            list_napr_deyat = []
            for key in self.DICT_NAPR_DEYAT.keys():
                list_napr_deyat.append(self.DICT_NAPR_DEYAT[key]['Имя'])
            nk_napr_deyat = CQT.nom_kol_po_imen(self.ui.tbl_pl_add_poz, 'Направление_деятельности')
            CQT.add_combobox(self, self.ui.tbl_pl_add_poz, 0, nk_napr_deyat, list_napr_deyat, first_void=False,
                             conn_func=select_napr_deyat)
            try:
                self.ui.tbl_pl_add_poz.cellWidget(0, nk_napr_deyat).setCurrentText(
                    self.DICT_NAPR_DEYAT[int(self.ui.tbl_pl_add_poz.item(0, nk_napr_deyat).text())]['Имя'])
            except:
                pass
            list_status = []
            for key in self.DICT_STATUS_POZ.keys():
                list_status.append(self.DICT_STATUS_POZ[key]['Имя'])
            nk_status = CQT.nom_kol_po_imen(self.ui.tbl_pl_add_poz, 'Статус')
            CQT.add_combobox(self, self.ui.tbl_pl_add_poz, 0, nk_status, list_status, first_void=False,
                             conn_func=select_status)
            try:
                self.ui.tbl_pl_add_poz.cellWidget(0, nk_status).setCurrentText(
                    self.DICT_STATUS_POZ[int(self.ui.tbl_pl_add_poz.item(0, nk_status).text())]['Имя'])
            except:
                pass

            list_etapi_erp = []
            for key in self.DICT_STATUS_ETAPI_ERP.keys():
                list_etapi_erp.append(self.DICT_STATUS_ETAPI_ERP[key]['Имя'])
            nk_etapi_erp = CQT.nom_kol_po_imen(self.ui.tbl_pl_add_poz, 'Этапы_ЕРП')
            CQT.add_combobox(self, self.ui.tbl_pl_add_poz, 0, nk_etapi_erp, list_etapi_erp, first_void=False,
                             conn_func=select_etapi_erp)
            try:
                self.ui.tbl_pl_add_poz.cellWidget(0, nk_etapi_erp).setCurrentText(
                    self.DICT_STATUS_ETAPI_ERP[int(self.ui.tbl_pl_add_poz.item(0, nk_etapi_erp).text())]['Имя'])
            except:
                pass
        if podr == 'пл_топ':
            list_vid = []
            for key in self.DICT_VID_PO_NAPR.keys():
                list_vid.append(self.DICT_VID_PO_NAPR[key]['Имя'])
            nk_vid = CQT.nom_kol_po_imen(self.ui.tbl_pl_add_poz, 'Вид')
            CQT.add_combobox(self, self.ui.tbl_pl_add_poz, 0, nk_vid, list_vid, first_void=False,
                             conn_func=select_vid)
            list_tech = []
            for key in self.DICT_EMPLOEE_FULL.keys():
                if self.DICT_EMPLOEE_FULL[key]['Подразделение'] == 'Технологический отдел Производства':
                    list_tech.append(key)
            list_tech = sorted(list_tech)
            nk_otv_tech = CQT.nom_kol_po_imen(self.ui.tbl_pl_add_poz, 'Отв_технолог')
            CQT.add_combobox(self, self.ui.tbl_pl_add_poz, 0, nk_otv_tech, list_tech, first_void=False,
                             conn_func=select_tech)
            try:
                self.ui.tbl_pl_add_poz.cellWidget(0, nk_vid).setCurrentText(
                    self.DICT_VID_PO_NAPR[int(self.ui.tbl_pl_add_poz.item(0, nk_vid).text())]['Имя'])
            except:
                pass

    if self.edit_tabel_mode:
        edit_tabel(self)
    else:
        edit_etap(self)



def clck_cld(self):
    tbl = self.ui.tbl_pl_add_poz
    if not current_cell_is_data_type(tbl):
        return
    date = self.ui.calendarWidget.selectedDate()
    new_str = F.datetostr(QDate.toPyDate(date), "%Y-%m-%d")
    col = tbl.currentColumn()
    old_str = tbl.item(0, col).text()
    if CQT.msgboxgYN(f'Установить для {tbl.horizontalHeaderItem(col).text()} c \n {old_str} \n на \n {new_str} ?'):
        tbl.item(0, col).setText(new_str)

def select_vid(self, text,  row, col):
    nk_vid = col
    val = 0
    for key in self.DICT_VID_PO_NAPR.keys():
        if self.DICT_VID_PO_NAPR[key]['Имя'] == text:
            val = key
            break
    self.ui.tbl_pl_add_poz.item(row, nk_vid).setText(str(val))
    print(f'Выбран {val}')

def select_tech(self, text,  row, col):
    self.ui.tbl_pl_add_poz.item(row, col).setText(text)
    print(f'Выбран {text}')

def select_napr_deyat(self, text,  row, col):
    nk_napr_deyat = col
    val = 0
    for key in self.DICT_NAPR_DEYAT.keys():
        if self.DICT_NAPR_DEYAT[key]['Имя'] == text:
            val = key
            tooltip = f"{self.DICT_NAPR_DEYAT[key]['Примечание']} ({self.DICT_NAPR_DEYAT[key]['Псевдоним']})"
            break
    self.ui.tbl_pl_add_poz.item(row, nk_napr_deyat).setText(str(val))
    self.ui.tbl_pl_add_poz.cellWidget(row, nk_napr_deyat).setToolTip(tooltip)
    print(f'Выбран {val}')

def select_status(self, text,  row, col ):
    nk_ = col
    val = 0
    for key in self.DICT_STATUS_POZ.keys():
        if self.DICT_STATUS_POZ[key]['Имя'] == text:
            val = key
            break
    self.ui.tbl_pl_add_poz.item(row, nk_).setText(str(val))
    print(f'Выбран {val}')

def select_etapi_erp(self, text,  row, col ):
    nk_ = col
    val = 0
    for key in self.DICT_STATUS_ETAPI_ERP.keys():
        if self.DICT_STATUS_ETAPI_ERP[key]['Имя'] == text:
            val = key
            break
    self.ui.tbl_pl_add_poz.item(row, nk_).setText(str(val))
    print(f'Выбран {val}')

def load_tbl_add_new_poz(self):

    list_heads = CSQ.zapros(self.db_kplan,"""SELECT
plan.Дата_внесения,
plan.Позиция,
plan.Направление_деятельности,
plan.Приоритет,
пл_оуп.Дата_заявки,
пл_оуп.№проекта,
пл_оуп.№ERP,
пл_оуп.Дата_отгрузки_ПУ,
пл_оуп.Количество,
пл_оуп.ПКК,
пл_оуп.Номенклатура_ЕРП
FROM plan
INNER JOIN 
пл_оуп ON пл_оуп.НомПл = plan.Пномер
 LIMIT 1""",one=True,shapka=True )[0]
    list_itog = ['' for _ in list_heads]
    list_itog = [list_heads,list_itog]
    CQT.fill_wtabl(list_itog,self.ui.tbl_pl_add_poz)
    list_napr_deyat = []
    for key in self.DICT_NAPR_DEYAT.keys():
        list_napr_deyat.append(self.DICT_NAPR_DEYAT[key]['Имя'])
    nk_napr_deyat = CQT.nom_kol_po_imen(self.ui.tbl_pl_add_poz,'Направление_деятельности')
    CQT.add_combobox(self,self.ui.tbl_pl_add_poz,0,nk_napr_deyat,list_napr_deyat, first_void=False,conn_func= select_napr_deyat)
    name_field = 'Пномер'
    oform_table_editeble(self, self.ui.tbl_pl_add_poz,name_field)

def check_add_poz(self):
    def check_number(self,val,key,tbl):
        if F.is_numeric(val) == False:
            CQT.msgbox(f'{key} должно быть число')
            CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
            return False
        return True
    def check_db(self,val,key,tbl, dict):
        if F.valm(val) not in dict:
            CQT.msgbox(f'{key} должно быть по БД')
            CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
            return False
        return True
    def check_choose(self,val,key,tbl):
        if val == '1':
            CQT.msgbox(f'{key} должно быть выбрано')
            CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
            return False
        return True

    tbl = self.ui.tbl_pl_add_poz
    list_add = CQT.spisok_iz_wtabl(tbl,rez_dict=True)[0]
    for key in list_add.keys():
        val = list_add[key]
        if val == '':
            CQT.msgbox(f'{key} не может быть пусто')
            return False
        if key in ['plan.Позиция','plan.Направление_деятельности','plan.Приоритет','пл_оуп.Количество']:
            if not check_number(self,val,key,tbl):
                return False
        if key == 'plan.Направление_деятельности':
            if not check_db(self,val,key,tbl, self.DICT_NAPR_DEYAT):
                return False
            if not check_choose(self,val,key,tbl):
                return False
        if key == 'пл_оуп.№ERP':
            if 'ПУ00-0' not in val:
                CQT.msgbox(f'{key} Не корректная запись')
                CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
                return False
            if not F.is_numeric(val.replace('ПУ00-0','')):
                CQT.msgbox(f'{key} Не корректная запись')
                CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
                return False
        if key == 'пл_оуп.Дата_заявки':
            if not F.is_date(val,'%Y-%m-%d'):
                CQT.msgbox(f'{key} Не корректный формат даты')
                CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
                return False
        if key == 'пл_оуп.Дата_отгрузки_ПУ':
            if not F.is_date(val,'%Y-%m-%d'):
                CQT.msgbox(f'{key} Не корректный формат даты')
                CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
                return False
    return True

def check_edit_poz(self,old_list):
    def check_number(self,val,key,tbl):
        if F.is_numeric(val) == False:
            CQT.msgbox(f'{key} должно быть число')
            CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
            return False
        return True
    def check_db(self,val,key,tbl, dict):
        if F.valm(val) not in dict:
            CQT.msgbox(f'{key} должно быть по БД')
            CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
            return False
        return True
    def check_choose(self,val,key,tbl):
        if val == '1':
            CQT.msgbox(f'{key} должно быть выбрано')
            CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
            return False
        return True
    def check_date(self,val,key,tbl,dateformat='%Y-%m-%d'):
        if not F.is_date(val, dateformat):
            CQT.msgbox(f'{key} Не корректный формат даты')
            CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
            return False
        return True

    list_edit = CQT.spisok_iz_wtabl(self.ui.tbl_pl_add_poz,rez_dict=True)[0]
    tbl = self.ui.tbl_pl_add_poz
    podr = self.ui.cmb_etap.currentText()

    if podr == 'plan':
        for key in list_edit.keys():
            val = list_edit[key]
            if str(val) == str(old_list[key]):
                continue
            if key in ['plan.Позиция','plan.Направление_деятельности','plan.Статус','plan.МК','plan.Нчас_вспом',
                       'plan.Фчас_вспом','plan.Фчас_доп_раб','plan.Этапы_ЕРП','plan.Приоритет']:
                if not check_number(self,val, key,tbl):
                    return False
            if key in ['plan.Направление_деятельности']:
                if not check_choose(self,val, key,tbl):
                    return False
            if key == 'plan.Направление_деятельности':
                if not check_db(self,val,key,tbl, self.DICT_NAPR_DEYAT):
                    return False
            if key == 'plan.Статус':
                if not check_db(self, val, key, tbl, self.DICT_STATUS_POZ):
                    return False
            if key == 'plan.Этапы_ЕРП':
                if not check_db(self, val, key, tbl, self.DICT_STATUS_ETAPI_ERP):
                    return False
    if podr == 'пл_заг':
        for key in list_edit.keys():
            val = list_edit[key]
            if str(val) == str(old_list[key]):
                continue
            if key in ['пл_заг.Нчас_заг', 'пл_заг.Фчас_заг']:
                if not check_number(self, val, key, tbl):
                    return False
            if key in ['пл_заг.ПДата_нач_заг', 'пл_заг.ПДата_зав_заг', 'пл_заг.ФДата_нач_заг', 'пл_заг.ФДата_зав_заг',
                       'пл_заг.ФДата_раскладки', 'пл_заг.ФДата_резки', 'пл_заг.ФДата_г_ш']:
                if not check_date(self, val, key, tbl):
                    return False
    if podr == 'пл_ко':
        for key in list_edit.keys():
            val = list_edit[key]
            if str(val) == str(old_list[key]):
                continue
            if key in ['пл_ко.Вес_ВО', 'пл_ко.Вес_КД']:
                if not check_number(self, val, key, tbl):
                    return False
            if key in ['пл_ко.Пдата_КД', 'пл_ко.Фдата_КД']:
                if not check_date(self, val, key, tbl):
                    return False
            if key == 'пл_ко.Ссылка_КД':
                if 'docs://srv-docs.powerz.ru:' not in val:
                    CQT.msgbox(f'{key} не корректная ссылка')
                    CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
                    return False

    if podr == 'пл_компл':
        for key in list_edit.keys():
            val = list_edit[key]
            if str(val) == str(old_list[key]):
                continue

            if key in ['пл_компл.Дата_комплект_после_заг','пл_компл.Дата_компл_под_мех','пл_компл.Дата_комплект_под_сб',
                'пл_компл.Дата_комплект_под_покр','пл_компл.Дата_комплект_под_упак',
                       'пл_компл.ПДата_комплект_упаковки','пл_компл.ФДата_комплект_упаковки',]:
                if not check_date(self, val, key, tbl):
                    return False

    if podr == 'пл_мех':
        for key in list_edit.keys():
            val = list_edit[key]
            if str(val) == str(old_list[key]):
                continue
            if key in ['пл_мех.Нчас_мех', 'пл_мех.Фчас_мех']:
                if not check_number(self, val, key, tbl):
                    return False
            if key in ['пл_мех.Пдата_нач_мехобр', 'пл_мех.Пдата_зав_мехобр',
                       'пл_мех.Фдата_нач_мехобр', 'пл_мех.Фдата_зав_мехобр']:
                if not check_date(self, val, key, tbl):
                    return False
    if podr == 'пл_оуп':
        for key in list_edit.keys():
            val = list_edit[key]
            if str(val) == str(old_list[key]):
                continue
            if key in ['пл_оуп.Количество']:
                if not check_number(self, val, key, tbl):
                    return False
            if key in ['пл_оуп.Дата_заявки','пл_оуп.Дата_отгрузки_ПУ']:
                if not check_date(self, val, key, tbl):
                    return False
            if key in ['пл_оуп.№проекта','пл_оуп.ПКК','пл_оуп.Номенклатура_ЕРП']:
                if val == '':
                    CQT.msgbox(f'{key} Не может быть пусто')
                    CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
                    return False
            if key == 'пл_оуп.№ERP':
                if 'ПУ00-0' not in val:
                    CQT.msgbox(f'{key} Не корректная запись')
                    CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
                    return False
                if not F.is_numeric(val.replace('ПУ00-0', '')):
                    CQT.msgbox(f'{key} Не корректная запись')
                    CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
                    return False
    if podr == 'пл_оуп':
        for key in list_edit.keys():
            val = list_edit[key]
            if str(val) == str(old_list[key]):
                continue
            if key in ['пл_оуп.Нчас_покр','пл_оуп.Фчас_покр']:
                if not check_number(self, val, key, tbl):
                    return False
            if key in ['пл_оуп.Пдата_нач_покр', 'пл_оуп.Пдата_зав_покр','пл_оуп.Фдата_нач_покр',
                       'пл_оуп.Фдата_зав_покр']:
                if not check_date(self, val, key, tbl):
                    return False
    if podr == 'пл_сб':
        for key in list_edit.keys():
            val = list_edit[key]
            if str(val) == str(old_list[key]):
                continue
            if key in ['пл_сб.Нчас_сб', 'пл_сб.Фчас_сб']:
                if not check_number(self, val, key, tbl):
                    return False
            if key in ['пл_сб.Пдата_нач_сб', 'пл_сб.Пдата_зав_сб', 'пл_сб.Фдата_нач_сб', 'пл_сб.Фдата_зав_сб']:
                if not check_date(self, val, key, tbl):
                    return False
    if podr == 'пл_топ':
        for key in list_edit.keys():
            val = list_edit[key]
            if str(val) == str(old_list[key]):
                continue
            if key in ['пл_топ.Нчас_сб', 'пл_топ.Фчас_сб','пл_топ.Вид',
                       'пл_топ.Уд_вес_ВО','пл_топ.Нчас_сб_ВО','пл_топ.Число_ДСЕ']:
                if not check_number(self, val, key, tbl):
                    return False
            if key in ['пл_топ.Пдата_ТД', 'пл_топ.Фдата_ТД', 'пл_топ.Уд_вес_ТД','пл_топ.Дата_МК',
                       'пл_топ.Спецификация_дата']:
                if not check_date(self, val, key, tbl):
                    return False
            if key == 'пл_топ.Вид':
                if not check_db(self,val,key,tbl, self.DICT_VID_PO_NAPR):
                    return False
            if key == 'пл_топ.Спецификация_ЕРП':
                nk_npoz = CQT.nom_kol_po_imen(self.ui.tbl_pl_add_poz,'НомПл')
                npoz = int(self.ui.tbl_pl_add_poz.item(0,nk_npoz).text())
                oyp_nomenkl = CSQ.zapros(self.db_kplan,
                                         f"""SELECT Номенклатура_ЕРП FROM пл_оуп WHERE НомПл == {npoz};""",one_column=True)
                if len(oyp_nomenkl) != 2:
                    return False
                if val != oyp_nomenkl[-1]:
                    CQT.msgbox(f'{key} Наименование должно совпадать с номенклатурой: {oyp_nomenkl[-1]}')
                    CQT.migat(self, tbl, 0, CQT.nom_kol_po_imen(tbl, key), 1)
                    return False
    return True


def btn_pl_ok_add_poz_click(self):

    def check_edit_tabel(self):
        month = self.ui.cmb_etap.currentText()
        if month == '':
            return
        list_month = CSQ.zapros(self.db_kplan, f"""SELECT * FROM {month}""")
        list_new = CQT.spisok_iz_wtabl(self.ui.tbl_pl_add_poz,'',True)
        if len(list_month) != len(list_new):
            CQT.msgbox(f'Что то пошло не так')
            return
        if len(list_month[0]) != len(list_new[0]):
            CQT.msgbox(f'Что то пошло не так')
            return
        list_changes = []
        list_sql = []
        for i in range(3, len(list_new)):
            for j in range(3,len(list_new[0])):
                if list_month[i][j] != list_new[i][j]:
                    list_changes.append(f'Для {list_month[i][1]} от {list_month[0][j]} было:{list_month[i][j]}, стало:{list_new[i][j]}')
                    list_sql.append(f"""UPDATE {month} SET {list_month[0][j]} = {list_new[i][j]} WHERE Подразделение = '{list_month[i][1]}'""")
        if list_changes == []:
            CQT.msgbox(f'Изменений не найдено')
            return False

        msg_str = 'Внести изменения?\n\n' + "\n".join(list_changes)
        if CQT.msgboxgYN(msg=msg_str):
            return list_sql
        return False
    def apply_edit_tabel(self,list_sql):
        for zapros in list_sql:
            CSQ.zapros(self.db_kplan,zapros)
        CQT.msgbox(f'Успешно')

    def add_new_poz(self):
        if not check_add_poz(self):
            return
        show_fr(self)

        list_add = CQT.spisok_iz_wtabl(self.ui.tbl_pl_add_poz, rez_dict=True)[0]

        CSQ.zapros(self.db_kplan, f"""INSERT INTO plan(
            Позиция,
            Направление_деятельности,
            Приоритет
            )
            VALUES (?,?,?);""", spisok_spiskov=[[list_add['Позиция'],
                                                   list_add['Направление_деятельности'],
                                                   list_add['Приоритет']]])
        pnom = CSQ.posl_strok_bd(self.db_kplan, 'plan', 'Пномер', ['Пномер'])[0]

        list_podr = [_ for _ in CSQ.spis_tablic(self.db_kplan) if 'пл_' in _]
        for podr in list_podr:
            CSQ.zapros(self.db_kplan, f"""INSERT INTO {podr}(
                НомПл
                )
                VALUES (?);""", spisok_spiskov=[[pnom]])
        vals = [list_add['Дата_заявки'],
                list_add['№проекта'],
                list_add['№ERP'],
                list_add['Дата_отгрузки_ПУ'],
                list_add['Количество'],
                list_add['ПКК'],
                list_add['Номенклатура_ЕРП']]
        CSQ.zapros(self.db_kplan, f"""UPDATE пл_оуп SET(
               Дата_заявки,
               №проекта,
               №ERP,
               Дата_отгрузки_ПУ,
               Количество,
               ПКК,
               Номенклатура_ЕРП
               ) =
                ({"?, ".join([""] * len(vals)) + "?"}) WHERE НомПл == {pnom};""", spisok_spiskov=vals)
        CQT.msgbox(f'Успешно')
        return True

    def edit_poz(self):
        tbl = self.ui.tbl_pl_add_poz
        podr = self.ui.cmb_etap.currentText()
        if podr == '':
            return
        if podr == 'plan':
            nk_nom = CQT.nom_kol_po_imen(tbl,'Пномер')
            name_field = 'Пномер'
        else:
            nk_nom = CQT.nom_kol_po_imen(tbl, 'НомПл')
            name_field = 'НомПл'
        pnom = int(tbl.item(0,nk_nom).text())
        old_list = get_line_to_edit_podr(self,pnom)
        old_list = F.list_to_dict(old_list)[0]
        if not check_edit_poz(self,old_list):
            return

        new_list = CQT.spisok_iz_wtabl(tbl,shapka=True,rez_dict=True)[0]
        old_list.pop(name_field)
        new_list.pop(name_field)
        for key in new_list.keys():
            if str(new_list[key]) != str(old_list[key]):
                pass

        list_fields = list(new_list.keys())
        list_vals = list(new_list.values())
        CSQ.zapros(self.db_kplan,f"""UPDATE {podr} SET({','.join(list_fields)}) =
         ({'?,'.join(['' for _ in list_fields]) + '?'}) WHERE {name_field} = {pnom};""",spisok_spiskov=list_vals)
        CQT.msgbox(f'Успешно')

        show_fr(self)
        return True

    def save_cnf(self):
        if 'shift' in  CQT.get_key_modifiers(self):
            path = 'Config\\fields.pickle'
            F.udal_file(path)
            return True
        spis = CQT.spisok_iz_wtabl(self.ui.tbl_pl_add_poz,shapka=True)
        rez_dict = dict()
        for j in range(0,len(spis[-1])):
            spis[1][j] = int(spis[1][j])
            hid = 1
            if spis[1][j] != 1:
                if spis[0][j] in LIST_FREEZE_FIELDS:
                    hid = 2
                else:
                    hid = 0
            rez_dict[spis[0][j]] = {'hidden' : hid, 'order' : j+1}
        path = 'Config\\fields.pickle'
        F.save_file_pickle(path,rez_dict)
        return True

    if self.edit_tabel_mode:
        list_sql = check_edit_tabel(self)
        if list_sql:
            apply_edit_tabel(self,list_sql)
    else:
        rez = None
        if self.regim == 'add':
           rez = add_new_poz(self)
        if self.regim == 'edit':
            rez = edit_poz(self)
        if self.regim == 'cnf':
            rez = save_cnf(self)
        if rez == None:
            return
        self.regim = ''
        load_table_db(self)
        show_fr(self)

def get_line_to_edit_podr(self,pnom):
    podr = self.ui.cmb_etap.currentText()
    if podr == "":
        CQT.clear_tbl(self.ui.tbl_pl_add_poz)
        return
    name_field = podr + '.НомПл'
    if podr == "plan":
        name_field = 'plan.Пномер'

    list_itog = CSQ.zapros(self.db_kplan, f"""SELECT * FROM 
                    {podr} WHERE {name_field} == {pnom}
                     """, one=True, shapka=True)
    return  list_itog



def show_fr(self,fr='',graf=0):
    self.ui.btn_kal_pl_left.setHidden(True)
    self.ui.btn_kal_pl_right.setHidden(True)
    if graf == 0:#объемный выключаем
        self.ui.fr_pl_graf.setHidden(True)
        self.ui.fr_pl_tables.setHidden(False)

        if fr == '':
            self.ui.fr_pl_cal.setHidden(True)
            self.ui.fr_pl_add_poz.setHidden(True)
            self.ui.fr_pl_etap.setHidden(True)
        if fr == 'tbl_add':
            self.ui.fr_pl_cal.setHidden(False)
            self.ui.fr_pl_add_poz.setHidden(False)
            self.ui.fr_pl_etap.setHidden(True)
        if fr == 'tbl_edit':
            self.ui.fr_pl_cal.setHidden(False)
            self.ui.fr_pl_add_poz.setHidden(False)
            self.ui.fr_pl_etap.setHidden(False)

    else:#объемный включаем
        self.ui.fr_pl_graf.setHidden(False)
        self.ui.fr_pl_tables.setHidden(True)
        self.ui.fr_svod.setHidden(True)
        self.ui.fr_pl_gaf.setHidden(False)




def check_db(self):
    return  F.nalich_file(self.db_kplan)

def current_cell_is_data_type(tbl):
    try:
        column = tbl.currentColumn()
        if 'дата' in tbl.horizontalHeaderItem(column).text().lower():
            return True
        return False
    except:
        return False


def dbl_clk_tbl_add_poz(self):
    if current_cell_is_data_type(self.ui.tbl_pl_add_poz):
        CQT.migat_obj(self,2,self.ui.calendarWidget,'Выбрать дату в календаре')

def select_field_from_kgui(self):
    tbl = self.ui.tbl_preview
    r = tbl.currentRow()
    c = tbl.currentColumn()
    if self.list_tbl_info[r + 1][c] == '':
        return
    dict_obj = copy.deepcopy(self.list_tbl_info[r + 1][c])[0]
    if 'shift' in CQT.get_key_modifiers(self):
        name_field = dict_obj['Имя_нз'][1]
    else:
        name_field = dict_obj['Имя_нз'][0]
    nk = CQT.nom_kol_po_imen(self.ui.tbl_kal_pl,name_field)
    self.ui.tbl_kal_pl.setCurrentCell(self.ui.tbl_kal_pl.currentRow(),nk)


def create_db(self):
    frase_tmp = """
PRAGMA foreign_keys = off;
BEGIN TRANSACTION;

-- Таблица: napravlenie
CREATE TABLE napravlenie (name TEXT NOT NULL PRIMARY KEY UNIQUE, val INTEGER NOT NULL);

-- Таблица: plan
CREATE TABLE "plan" (Пномер INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL UNIQUE ON CONFLICT ROLLBACK, Дата_заявки TEXT DEFAULT "", №проекта TEXT DEFAULT "", №ERP TEXT DEFAULT "", Позиция TEXT DEFAULT "", Вид TEXT DEFAULT "", Количество INTEGER DEFAULT (1), Вес_ВО INTEGER DEFAULT (1), Дата_отгрузки_ПУ TEXT DEFAULT "", Уд_вес_ВО INTEGER DEFAULT (1), Нчас_сб_ВО INTEGER DEFAULT (1), ПКК TEXT DEFAULT "", Отв_технолог INTEGER DEFAULT (1), Статус TEXT DEFAULT "", Пдата_КД TEXT DEFAULT "", Фдата_КД TEXT DEFAULT "", Вес_КД INTEGER DEFAULT (1), Направление TEXT DEFAULT "", Пдата_ТД TEXT DEFAULT "", Фдата_ТД TEXT DEFAULT "", Уд_вес_ТД INTEGER DEFAULT (1), Число_ДСЕ INTEGER DEFAULT (1), МК INTEGER DEFAULT (1), Нчас_заг INTEGER DEFAULT (1), Нчас_мех INTEGER DEFAULT (1), Нчас_сб INTEGER DEFAULT (1), Нчас_покр INTEGER DEFAULT (1), Нчас_вспом INTEGER DEFAULT (1), Этапы_ЕРП INTEGER DEFAULT (0), Примечание TEXT DEFAULT "", Приоритет INTEGER DEFAULT (1));

-- Таблица: vid_pr
CREATE TABLE vid_pr (name TEXT DEFAULT "", koef_proizv INTEGER DEFAULT (1));

COMMIT TRANSACTION;
PRAGMA foreign_keys = on;

"""
    CSQ.sozd_bd_sql(self.db_kplan,frase_tmp)

def load_db(self,pnom=False):
    dict_inner = {'plan.Направление_деятельности as "plan.Направление_деятельности"': 'napravl_deyat.Имя as "plan.Направление_деятельности"',
                  'plan.Статус as "plan.Статус"': 'status_poz.Имя as "plan.Статус"',
                  'plan.Этапы_ЕРП as "plan.Этапы_ЕРП"': 'status_etapi_erp.Имя as "plan.Этапы_ЕРП"',
                  'пл_топ.Вид as "пл_топ.Вид"': 'Виды_по_напр.Имя as "пл_топ.Вид"',
                  }
    if check_db(self) == False:
        create_db(self)


    rez_list_tabels = ['napravlenie.name as "Направление"', 'napravl_deyat.Псевдоним as "Псевдоним"']
    postfix = ''
    if pnom:
        postfix = f'WHERE plan.Пномер == {int(pnom)}'
        list_conf = load_list_fields(self,True)
        for i in range(len(list_conf[0])):
            rez_list_tabels.append(f'{list_conf[0][i]} as "{list_conf[0][i]}"')
    else:
        list_conf = load_list_fields(self)
        for i in range(len( list_conf[0])):
            if list_conf[1][i]:
                rez_list_tabels.append(f'{list_conf[0][i]} as "{list_conf[0][i]}"')
    str_field = ', '.join(rez_list_tabels)
    for key in dict_inner.keys():
        str_field = str_field.replace(key, dict_inner[key])



    list = CSQ.zapros(self.db_kplan, f"""SELECT
    {str_field}
    FROM plan
    INNER JOIN 
    пл_оуп ON пл_оуп.НомПл = plan.Пномер,
    пл_ко ON пл_ко.НомПл = plan.Пномер,
    пл_топ ON пл_топ.НомПл = plan.Пномер,
    пл_заг ON пл_заг.НомПл = plan.Пномер,
    пл_компл ON пл_компл.НомПл = plan.Пномер,
    пл_мех ON пл_мех.НомПл = plan.Пномер,
    пл_сб ON пл_сб.НомПл = plan.Пномер,
    пл_покр ON пл_покр.НомПл = plan.Пномер,
    napravl_deyat ON napravl_deyat.Пномер = plan.Направление_деятельности,
    status_poz ON status_poz.Пномер = plan.Статус,
    status_etapi_erp ON status_etapi_erp.Пномер = plan.Этапы_ЕРП,
    Виды_по_напр ON Виды_по_напр.Пномер = пл_топ.Вид,
    napravlenie ON napravlenie.Пномер = napravl_deyat.Направление
    {postfix};
    """)

    return list, list_conf

def load_table_db(self):
    def oforml_table(self):
        tbl = self.ui.tbl_kal_pl
        nk_pseudo = CQT.nom_kol_po_imen(tbl,'Псевдоним')
        nk_napr = CQT.nom_kol_po_imen(tbl,'plan.Направление_деятельности')
        nk_nom_pr = CQT.nom_kol_po_imen(tbl,'пл_оуп.№проекта')
        nk_local_graf = CQT.nom_kol_po_imen(tbl, 'plan.local_graf')
        self.ui.tbl_kal_pl.setColumnHidden(nk_local_graf, True)
        for i in range(tbl.rowCount()):
            r, g, b = self.DICT_NAPR_DEYAT_NAME[tbl.item(i,nk_napr).text()]['Цвет'].split(';')
            CQT.ust_color_wtab(tbl,i,nk_pseudo,r,g,b)
            CQT.font_cell_size_format(tbl,i,nk_nom_pr,underline=True)


    list_from_db, list_conf = load_db(self)
    #if len(list) == 1:
    #    list = add_excell(list)
    #list = CSQ.fix_types_table(list)
    CQT.fill_wtabl(list_from_db,self.ui.tbl_kal_pl, auto_type=True,set_editeble_col_nomera=[])
    for i in range(len(list_conf[0])):
        if CQT.nom_kol_po_imen(self.ui.tbl_kal_pl, list_conf[0][i]) != None:
            if list_conf[1][i] == 2:
                self.ui.tbl_kal_pl.setColumnHidden(CQT.nom_kol_po_imen(self.ui.tbl_kal_pl, list_conf[0][i]),True)
            else:
                self.ui.tbl_kal_pl.setColumnHidden(CQT.nom_kol_po_imen(self.ui.tbl_kal_pl, list_conf[0][i]), False)
    CMS.zapolnit_filtr(self,self.ui.tbl_filtr_kal_pl,self.ui.tbl_kal_pl,hidden_scroll=True)

    oforml_table(self)
    CMS.load_column_widths(self,self.ui.tbl_kal_pl)
    CMS.update_width_filtr(self.ui.tbl_kal_pl, self.ui.tbl_filtr_kal_pl)



@CQT.onerror
def clck_tbl_kal_pl_tbl(self,tbl):
    summ_selct_tbl(self,tbl)
    select_row(self)

def show_tabel(self):
    if self.ui.fr_pl_add_poz.isHidden():
        self.edit_tabel_mode = True
        self.ui.fr_pl_etap.setHidden(False)
        self.ui.fr_pl_add_poz.setHidden(False)
        self.ui.tbl_pl_add_poz.setMaximumHeight(370)
        self.ui.fr_pl_add_poz.setMaximumHeight(470)
        self.ui.fr_gant_local.setHidden(True)
        self.ui.cmb_etap.clear()
        self.ui.tbl_pl_add_poz.clear()
        list_month = [_ for _ in CSQ.spis_tablic(self.db_kplan) if 'm_cld_' in _]
        list_month.insert(0,'')
        self.ui.cmb_etap.addItems(list_month)
    else:
        self.ui.fr_pl_add_poz.setHidden(True)
        self.ui.fr_gant_local.setHidden(False)
        self.edit_tabel_mode = False
        self.ui.tbl_pl_add_poz.setMaximumHeight(70)
        self.ui.fr_pl_add_poz.setMaximumHeight(170)
        self.ui.tbl_pl_add_poz.clear()

@CQT.onerror
def clck_tbl_preview(self,tbl):
    summ_selct_tbl(self,tbl)
    load_info_select_block(self,tbl)

def load_info_select_block(self,tbl):
    r = tbl.currentRow()
    c = tbl.currentColumn()
    msg =  self.statusBar().currentMessage()
    try:
        if self.list_tbl_info[r+1][c] == '':
            CQT.statusbar_text(self,'')
            tbl.setToolTip('')
    except:
        return
    list = copy.deepcopy(self.list_tbl_info[r+1][c])
    info = ''
    if type(list) == type([]):
        tmp = []
        for item in list:
            item.pop("Имя_нз",list)
            tmp.append(str(item))
        info = ('\n'.join(tmp))
        tbl.setToolTip(info)
    CQT.statusbar_text(self,
                       f'{msg} |  {info}')



@CQT.onerror
def summ_selct_tbl(self,tbl):
    CMS.on_section_resized(self)
    summ= 0
    sch = 1
    try:
        #tbl = self.ui.tbl_kal_pl
        #col = tbl.currentColumn()
        #row = tbl.currentRow()
        summ = 0
        sch = 0
        for ix in tbl.selectedIndexes():
            #if col == ix.column():
            try:
                if F.is_numeric(ix.data()):
                    summ+=F.valm(ix.data())
                    sch+=1
            except:
                pass
        CQT.statusbar_text(self,f'                                         Сумма: {summ},  Среднее: {round(summ/sch,3)}')
    except:
        CQT.statusbar_text(self,
                           f'                                         ')
        pass
    try:
        CQT.statusbar_text(self, f'                                         Сумма: {summ},  Среднее: {round(summ / sch,3)}')
    except:
        CQT.statusbar_text(self,
                           f'                                         ')
        pass


def add_excell(list):
    list_of_pr = F.otkr_f('O:\Производство Powerz\Отдел технолога\ТД\Учет табель\бд_проекты\BD_Proect.txt',separ ='|',utf8=False)
    for i, item in enumerate(list_of_pr):
        if i < 5:
            continue
        nompr = str(item[0])
        nompy = str(item[1])
        vid = item[2]
        status = item[3]
        datakd = item[4]
        sb = item[6].replace('&','')
        kolvo = item[8]
        napravl_deyat = ''
        nomen_erp = ''
        dataotgr = item[9]
        ves = F.valm(item[10].replace('&',''))
        prim = item[11]
        tehnolog = item[16]
        datazayavk = item[17]
        datakd_plan = item[19]
        pkk = item[20]
        datatd = item[21]
        napravl = item[22]
        prioritet = item[23]

        if F.is_numeric(sb):
            sb = F.valm(sb)
            udvestd = round(sb*2/8*102)
        else:
            sb = 1
            udvestd = 1
        if item[7] == '?':
            etap = 0
        else:
            etap = 1
        list.append(['',datazayavk,nompr,nompy,'00',vid,napravl_deyat,nomen_erp,kolvo,1,dataotgr,'','',1,1,pkk,
                     tehnolog,status,datakd_plan,datakd,ves,napravl,"",datatd,"",udvestd,1,1,1,"",1,"",round(sb*2,2),"",1,"",1,"","",etap,prim,prioritet,''])
    return list

def load_table_add(self):
    pass