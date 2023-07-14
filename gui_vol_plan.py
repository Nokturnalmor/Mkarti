from __future__ import annotations

import copy
import datetime

import project_cust_38.Cust_Qt as CQT
import project_cust_38.Cust_Functions as F
import project_cust_38.Cust_SQLite  as CSQ
import kal_plan as KPL
import gui_kal_plan as GKPL
import project_cust_38.Cust_mes as CMS

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from MKart import mywindow

@CQT.onerror
def save_kpl_plan(self):
    if self.dict_form_kpl == "":
        CQT.msgbox(f'План не сформирован')
        return
    path = CMS.load_tmp_path('kpl_vol_plan')
    pathf = path + F.sep() + F.now("KPL%Y_%m_%d") + '.pickle'
    pathf = CQT.f_dialog_save(self,'Сохранить план', pathf, "*.pickle")
    if pathf == '.':
        return
    F.save_file_pickle(pathf,self.dict_form_kpl)
    F.otkr_papky(CMS.save_tmp_path('kpl_vol_plan',pathf,True))



def hover_tbl_pl_gaf_svod(self, event):
    tbl = self.ui.tbl_pl_gaf_svod
    row, column = CQT.get_hover_row_col(self, tbl, event)
    if row == False or column == False:
        return
    val = tbl.item(row, column).text()
    if val != '':
        set_tooltip_val(self, tbl, row, column)
    else:
        set_tooltip_val(self, tbl, row, column, True)


@CQT.onerror
@F.vrem_vip_cls_func_args
def load_tbl_gant(self:mywindow):
    def load_tabels(self:mywindow) -> dict:
        tbls = []
        tbl = self.ui.tbl_kal_pl
        nk_pnom = CQT.nom_kol_po_imen(tbl,'plan.Пномер')
        nk_local_graf = CQT.nom_kol_po_imen(tbl, 'plan.local_graf')
        for i in range(tbl.rowCount()):
            if not tbl.isRowHidden(i):
                tbls.append(tbl.item(i,nk_pnom).text())
        query = CSQ.zapros(self.db_kplan,f"""SELECT plan.Пномер, plan.Позиция, plan.local_graf, plan.Приоритет, 
        пл_оуп.№проекта, пл_оуп.№ERP, napravl_deyat.Псевдоним as Направление FROM plan INNER JOIN
        пл_оуп ON пл_оуп.НомПл = plan.Пномер,
        napravl_deyat ON napravl_deyat.Пномер = plan.Направление_деятельности
         WHERE plan.Пномер in ({','.join(tbls)}) ORDER BY plan.Приоритет DESC""",rez_dict=True)
        if query == False or len(query) == 0:
            return False
        for item in query:
            if item['local_graf'] == '':
                datat_bin = GKPL.update_local_graf(self, True, int(item['Пномер']), False)
                print(f"Создан локальный график на {item['Пномер']}")
                #tbl.item(i, nk_local_graf).setText(str(datat_bin))
        return query

    def generate_full_table(self,query:dict):
        set_dates = set()
        dict_dates_vals = dict()
        for item in query:
            tbl_gant = F.from_binary_pickle(item['local_graf'])
            if tbl_gant == None:
                CQT.msgbox(f'Пномер {item["Пномер"]}, №проекта {item["№проекта"]} - Не сформирован локальный график')
            try:
                for date in tbl_gant[0]['data'].keys():
                    set_dates.add(date)
                    dict_dates_vals[date] = {'Выходные':tbl_gant[0]['data'][date]['Выходные'],'День недели':tbl_gant[0]['data'][date]['День недели']}
            except:
                CQT.msgbox(f'Ошибка генерации{str(tbl_gant[0])}')
        list_dates = list(set_dates)
        list_dates = sorted(list_dates)
        dict_form = []
        for item in query:
            tbl_gant = F.from_binary_pickle(item['local_graf'])
            free_shablon = dict()
            for key in tbl_gant[0]['data'].keys():
                free_shablon = tbl_gant[0]['data'][key]['podr']
                break
            for key in free_shablon.keys():
                free_shablon[key] = ''

            dict_tmp_table = dict()
            for date in list_dates:
                if date not in tbl_gant[0]['data'].keys():
                    dict_tmp_table[date] = {'Выходные':dict_dates_vals[date]['Выходные'],
                                            'День недели':dict_dates_vals[date]['День недели'],
                                            'podr':free_shablon}
                else:
                    dict_tmp_table[date] = tbl_gant[0]['data'][date]
            tmp_list = []
            #query[i]['local_graf'] = dict_tmp_table
            dict_form.append({'pnom':item['Пномер'],'proj':f"{item['№проекта']} {item['№ERP']}",'poz':item['Позиция'],'napr':item['Направление'] ,'data':dict_tmp_table})

        return dict_form


    list_of_tbls = load_tabels(self)
    if not list_of_tbls:
        CQT.msgbox(f'Ошибка')
        return
    dict_form = generate_full_table(self,list_of_tbls)
    GKPL.fill_gant_table(self, self.ui.tbl_pl_gaf, self.ui.tbl_pl_gaf_filtr, dict_form)


def show_svod(self):
    if self.ui.fr_pl_gaf.isHidden():
        self.ui.fr_pl_gaf.setHidden(False)
        self.ui.fr_svod.setHidden(True)
    else:
        self.ui.fr_pl_gaf.setHidden(True)
        self.ui.fr_svod.setHidden(False)
        load_svod(self)
        load_svod_percent(self)



def set_tooltip_val(self,tbls='',r='',c='',clear=False):
    if tbls == "":
        tbls = self.ui.tbl_pl_gaf_svod
    if clear:
        CQT.statusbar_text(self, '')
        tbls.setToolTip('')
        return
    KPL.summ_selct_tbl(self,tbls)
    max_mosh, podr, date_tmp = get_max_mosh_frow_tbl(self,r,c)
    info = f'Максимальная мощность {podr} на {date_tmp} : {max_mosh} н-ч.'
    tbls.setToolTip(info)
    CQT.statusbar_text(self,
                       f'{self.glob_kpl_summ_selct_tbl} |  {info}')

def get_max_mosh_frow_tbl(self,i='',j=''):
    tbls = self.ui.tbl_pl_gaf_svod
    if i == '':
        i = tbls.currentRow()+1
    if j == '':
        j = tbls.currentColumn()
    if i == -1 or j == -1:
        return None,None,None
    rez_list = CQT.spisok_iz_wtabl(tbls, shapka=True)

    date_tmp = ".".join(rez_list[0][j].split('\n')[:-1])
    podr = rez_list[i][0].replace('факт_', '').replace('план_', '')
    max_mosh = 0
    try:
        max_mosh = self.KPLAN_max_mosh[F.strtodate(date_tmp, f"%d.%m.%y")][podr]
    except:
        pass
    return max_mosh, podr, date_tmp

@CQT.onerror
def oform_tbl_svod_percent(self:mywindow, rez_list:list =''):
    tbls = self.ui.tbl_pl_gaf_svod_percent
    if rez_list == '':
        rez_list = CQT.spisok_iz_wtabl(tbls, shapka=True)
    count_tbl_field = 5
    tbl = self.ui.tbl_pl_gaf
    for j in range(1, count_tbl_field):  # HORIZONTAL HEADER TABLE FIELDS
        CQT.ust_color_text_header_wtab_horisontal(tbls, j, 11, 11, 11, self.val_masht * 0.9, False)
        for i in range(1, len(rez_list)):
            CQT.font_cell_size_format(tbls, i - 1, j, self.val_masht)
    for j in range(count_tbl_field, len(rez_list[0])):  # HORIZONTAL HEADER GANT FIELDS
        if self.list_tbl_info[1][j] == 1:
            CQT.ust_color_text_header_wtab_horisontal(tbls, j, 200, 11, 11, self.val_masht * 0.8, True)
        else:
            CQT.ust_color_text_header_wtab_horisontal(tbls, j, 11, 11, 11, self.val_masht * 0.7, False)
            # CQT.ust_color_text_header_wtab_horisontal(tbls, j, 11, 11, 11, self.val_masht * 0.8, False)
        for i in range(1, len(rez_list)):
            CQT.font_cell_size_format(tbls, i - 1, j, self.val_masht * 0.8)

    for i in range(1, len(rez_list)):
        napr = rez_list[i][0]
        val = 0
        r, g, b = 10,10,10
        for key in self.data_kpl.DICT_NAPRAVLENIE:
            if self.data_kpl.DICT_NAPRAVLENIE[key]['name'] == napr:
                val = self.data_kpl.DICT_NAPRAVLENIE[key]['val']
                r, g ,b = self.data_kpl.DICT_NAPRAVLENIE[key]['Цвет'].split(';')
                break
        for j in range(count_tbl_field, len(rez_list[0])):
            if rez_list[i][j] > (val + 20) or rez_list[i][j] < (val - 20):
                CQT.ust_font_color_wtab(tbls, i - 1, j, 244, 244, 244)
                CQT.ust_color_wtab(tbls, i - 1, j, 233, 33, 33)
        CQT.ust_color_text_header_wtab_vertical(tbls, i - 1, r, g, b, self.val_masht * 0.8, True)
        if rez_list[i][4] > (val + 20) or rez_list[i][4] < (val - 20):
            CQT.ust_font_color_wtab(tbls, i - 1, 4, 244, 244, 244)
            CQT.ust_color_wtab(tbls, i - 1, 4, 233, 33, 33)
        CQT.font_cell_size_format(tbls, i - 1, 4,0,True)

    CMS.update_width_filtr(tbl, tbls)
    fields_hide = ['Этап', 'Пномер', "Проект", "Поз."]
    for field in fields_hide:
        try:
            tbls.setColumnHidden(CQT.nom_kol_po_imen(tbls, field), True)
        except:
            pass



def oform_tbl_svod(self,rez_list:list =''):
    tbls = self.ui.tbl_pl_gaf_svod

    if rez_list == '':
        rez_list = CQT.spisok_iz_wtabl(tbls,shapka=True)
    count_tbl_field = 5
    tbl = self.ui.tbl_pl_gaf
    for j in range(1, count_tbl_field):#HORIZONTAL HEADER TABLE FIELDS
        CQT.ust_color_text_header_wtab_horisontal(tbls, j, 11, 11, 11, self.val_masht * 0.9, False)
        for i in range(1,len(rez_list)):
            CQT.font_cell_size_format(tbls, i - 1, j, self.val_masht)
    for j in range(count_tbl_field, len(rez_list[0])):#HORIZONTAL HEADER GANT FIELDS
        if self.list_tbl_info[1][j] == 1:
            CQT.ust_color_text_header_wtab_horisontal(tbls, j, 200, 11, 11, self.val_masht*0.8, True)
        else:
            CQT.ust_color_text_header_wtab_horisontal(tbls, j, 11, 11, 11, self.val_masht*0.7, False)
            #CQT.ust_color_text_header_wtab_horisontal(tbls, j, 11, 11, 11, self.val_masht * 0.8, False)
        for i in range(1,len(rez_list)):
            CQT.font_cell_size_format(tbls, i - 1, j, self.val_masht*0.8)
            if rez_list[i][j] > 0:
                max_mosh, podr, date_tmp =  get_max_mosh_frow_tbl(self,i,j)
                if rez_list[i][j]>max_mosh:
                    CQT.ust_font_color_wtab(tbls, i - 1, j, 244, 244, 244)
                    CQT.font_cell_size_format(tbls,i - 1, j, 0, True)
                    CQT.ust_color_wtab(tbls, i - 1, j, 233, 33, 33)
                    CQT.ust_color_text_header_wtab_horisontal(tbls, j, 230, 233, 233, self.val_masht * 0.8, True)
                podr = rez_list[i][0].replace('факт_', '').replace('план_', '')
                r = 233
                g = 233
                b = 233
                if podr in self.data_kpl.DICT_PODR:
                    r, g, b = self.data_kpl.DICT_PODR[podr]['Цвет'].split(";")
                CQT.dob_color_wtab(tbls, i - 1, j, int(r), int(g), int(b))
            else:
                CQT.ust_font_color_wtab(tbls, i - 1, j, 233, 233, 233)

    for i in range(1, len(rez_list)):
        podr = rez_list[i][0].replace('факт_', '').replace('план_', '')
        r = 233
        g = 233
        b = 233
        if podr in self.data_kpl.DICT_PODR:
            r, g, b = self.data_kpl.DICT_PODR[podr]['Цвет'].split(";")
        CQT.ust_color_text_header_wtab_vertical(tbls, i - 1, r, g, b, self.val_masht * 0.8, True)


    CMS.update_width_filtr(tbl, tbls)
    fields_hide = ['Этап', 'Пномер', "Проект", "Поз.", "Напр."]
    for field in fields_hide:
        try:
            tbls.setColumnHidden(CQT.nom_kol_po_imen(tbls, field), True)
        except:
            pass



def dbl_clk_select_etap(self):
    def get_down_to_local():
        try:
            cell = self.list_tbl_info[r + 1][c]
            if cell == '':
                return
            name_kol = cell[0]['Имя_нз'][0]
        except:
            return
        # date = F.strtodate(".".join(tbl.horizontalHeaderItem(c).text().split('\n')[:-1]), f"%d.%m.%y")
        date = tbl.horizontalHeaderItem(c).text()

        table_new = CQT.spisok_iz_wtabl(self.ui.tbl_kal_pl, '', True)
        new_shapka = ['' for _ in table_new[0]]
        nk_pnom = CQT.nom_kol_po_imen(self.ui.tbl_kal_pl, 'plan.Пномер')
        new_shapka[nk_pnom] = pnom
        CMS.zapolnit_filtr(self, tbl_filtr, self.ui.tbl_kal_pl, [new_shapka])
        CMS.update_width_filtr(self.ui.tbl_kal_pl, tbl_filtr)
        CMS.primenit_filtr(self, tbl_filtr, self.ui.tbl_kal_pl)
        not_hidden_row = 0
        for i in range(self.ui.tbl_kal_pl.rowCount()):
            if not self.ui.tbl_kal_pl.isRowHidden(i):
                not_hidden_row = i
                break
        # self.ui.tbl_kal_pl.setCurrentCell(not_hidden_row,CQT.nom_kol_po_imen(self.ui.tbl_kal_pl,name_kol))
        KPL.btn_pl_mode(self)
        CQT.select_cell(self.ui.tbl_kal_pl, not_hidden_row, CQT.nom_kol_po_imen(self.ui.tbl_kal_pl, name_kol))
        # tbl.clearSelection()

        list_local_gant = CQT.spisok_iz_wtabl(self.ui.tbl_preview, '', True)
        nk_etap = CQT.nom_kol_po_imen(self.ui.tbl_preview, 'Этап')
        for i in range(len(list_local_gant)):
            if list_local_gant[i][nk_etap] == etap:
                for j in range(len(list_local_gant[0])):
                    if list_local_gant[0][j] == date:
                        CQT.select_cell(self.ui.tbl_preview, i - 1, j)
                        return

    list_for_copy_filtr = ['Этап'
,'Пномер'
,'Проект'
,'Поз.'
,'Напр.']
    tbl = self.ui.tbl_pl_gaf
    tbl_filtr = self.ui.tbl_filtr_kal_pl
    r = tbl.currentRow()
    c = tbl.currentColumn()
    pnom = tbl.item(r,CQT.nom_kol_po_imen(tbl,'Пномер')).text()
    etap = tbl.item(r,CQT.nom_kol_po_imen(tbl,'Этап')).text()

    if self.list_tbl_info[0][c] in list_for_copy_filtr:
        self.ui.tbl_pl_gaf_filtr.item(0,c).setText(self.ui.tbl_pl_gaf.item(r,c).text())
        return
    get_down_to_local()



def dbl_clk_svod_select_etap(self):
    tbls = self.ui.tbl_pl_gaf_svod
    tbl_filtr = self.ui.tbl_pl_gaf_filtr
    r = tbls.currentRow()
    c = tbls.currentColumn()
    etap = tbls.item(r,0).text()
    date = F.strtodate(".".join(tbls.horizontalHeaderItem(c).text().split('\n')[:-1]), f"%d.%m.%y")
    self.ui.fr_svod.setHidden(True)
    self.ui.fr_pl_gaf.setHidden(False)
    table_new = CQT.spisok_iz_wtabl(self.ui.tbl_pl_gaf,'',True)
    new_shapka = ['' for _ in table_new[0]]
    nk_etap = CQT.nom_kol_po_imen(self.ui.tbl_pl_gaf,'Этап')
    new_shapka[nk_etap] = etap
    CMS.zapolnit_filtr(self,tbl_filtr,self.ui.tbl_pl_gaf,[new_shapka])
    tbl_filtr.setRowHeight(0, 25)
    for j in range(1, len(self.list_tbl_info[0])):
        if self.list_tbl_info[1][j] == 1:
            CQT.ust_color_text_header_wtab_horisontal(tbl_filtr, j, 200, 11, 11, self.val_masht * 0.5, False)
        else:
            CQT.ust_color_text_header_wtab_horisontal(tbl_filtr, j, 11, 11, 11, self.val_masht * 0.5, False)
    CMS.update_width_filtr(self.ui.tbl_pl_gaf, tbl_filtr)
    CMS.primenit_filtr(self,self.ui.tbl_pl_gaf_filtr,self.ui.tbl_pl_gaf)
    not_hidden_row = 0
    for i in range(self.ui.tbl_pl_gaf.rowCount()):
        if not self.ui.tbl_pl_gaf.isRowHidden(i):
            not_hidden_row= i
            break
    #self.ui.tbl_pl_gaf.setCurrentCell(not_hidden_row,CQT.nom_kol_po_imen(self.ui.tbl_pl_gaf,tbls.horizontalHeaderItem(c).text()))
    CQT.select_cell(self.ui.tbl_pl_gaf,not_hidden_row,CQT.nom_kol_po_imen(self.ui.tbl_pl_gaf,tbls.horizontalHeaderItem(c).text()))
    #tbls.clearSelection()
def get_max_mosh_from_db(self):
    list_tables = CSQ.spis_tablic(self.db_kplan)
    dict_days = dict()
    for tbl_name in list_tables:
        if F.is_date(tbl_name,'m_cld_%Y_%m_%d'):
            dict_month = CSQ.zapros(self.db_kplan,f"""SELECT * FROM {tbl_name}""")
            for i in range(3,len(dict_month[0])):
                if F.is_date(dict_month[0][i],'d_%Y_%m_%d'):
                    dict_day = dict()
                    for j in range(3,len(dict_month)):
                        dict_day[dict_month[j][1]] = dict_month[j][i]
                    dict_days[F.strtodate(dict_month[0][i],'d_%Y_%m_%d')] = dict_day
    self.KPLAN_max_mosh = dict_days

@CQT.onerror
def load_svod_percent(self:mywindow):
    tbl = self.ui.tbl_pl_gaf
    tbls = self.ui.tbl_pl_gaf_svod_percent
    rez_list = [copy.deepcopy(self.list_tbl[0])]
    set_napr = set()
    nk_napr = F.nom_kol_po_im_v_shap(self.list_tbl, 'Напр.')
    for i in range(3, len(self.list_tbl)):
        if not self.ui.tbl_pl_gaf.isRowHidden(i - 1):
            set_napr.add(self.list_tbl[i][nk_napr])
    list_napr = sorted(set_napr)
    for napr in list_napr:
        tmp_row = [napr, '', '', '', '']
        for j in range(5, len(self.list_tbl[0])):
            summ_chas = 0
            for i in range(3, len(self.list_tbl)):
                if self.list_tbl[i][nk_napr] == napr:
                    if not self.ui.tbl_pl_gaf.isRowHidden(i - 1):
                        if self.list_tbl[i][j] != '':
                            summ_vol = 0
                            for oper in self.list_tbl_info[i][j]:
                                vol = oper['Время_час']
                                summ_vol += vol
                            summ_chas += summ_vol
            tmp_row.append(round(summ_chas))
        rez_list.append(tmp_row)
    rez_list[0][4]='Среднее'
    for i in range(1, len(rez_list)):
        rez_list[i][4] = 0
    for j in range(5,len(rez_list[0])):
        summ = 0
        for i in range(1,len(rez_list)):
            summ+= F.valm(rez_list[i][j])
        for i in range(1,len(rez_list)):
            if not summ == 0:
                rez_list[i][4] += rez_list[i][j]
                rez_list[i][j] = round(F.valm(rez_list[i][j])/summ*100)
    summ = 0
    for i in range(1, len(rez_list)):
        summ += F.valm(rez_list[i][4])
    for i in range(1,len(rez_list)):
        rez_list[i][4] =round(F.valm(rez_list[i][4])/summ*100)

    CQT.fill_wtabl(rez_list, tbls, min_width_col=int(4 * 0.8),
                   height_row=self.val_masht * 2, colorful_edit=False, auto_type=False, head_column=0,
                   set_editeble_col_nomera={}, hide_head_column=False)
    oform_tbl_svod_percent(self, rez_list)


def load_svod(self):
    tbl = self.ui.tbl_pl_gaf
    tbls =  self.ui.tbl_pl_gaf_svod
    rez_list = [self.list_tbl[0]]
    set_etapov = set()
    nk_etap = F.nom_kol_po_im_v_shap(self.list_tbl,'Этап')
    for i in range(3,len(self.list_tbl)):
        if not self.ui.tbl_pl_gaf.isRowHidden(i-1):
            set_etapov.add(self.list_tbl[i][nk_etap])
    list_etapov = sorted(set_etapov)
    for etap in list_etapov:
        tmp_row = [etap,'','','','']
        for j in range(5,len(self.list_tbl[0])):
            summ_chas = 0
            for i in range(3,len(self.list_tbl)):
                if self.list_tbl[i][nk_etap] == etap:
                    if self.list_tbl[i][j] != '':
                        summ_vol = 0
                        for oper in self.list_tbl_info[i][j]:
                            vol = oper['Время_час']
                            summ_vol+=vol
                        summ_chas += summ_vol
            tmp_row.append(round(summ_chas))
        rez_list.append(tmp_row)
    CQT.fill_wtabl(rez_list, tbls, min_width_col= int( 4 * 0.8),
                   height_row=self.val_masht * 2, colorful_edit=False, auto_type=False, head_column=0,
                   set_editeble_col_nomera={}, hide_head_column=False)
    oform_tbl_svod(self,rez_list)


def save_diapazon_month(self: mywindow):
    str_d = F.datetostr(self.ui.de_vol_pl.date().toPyDate()) + ';' + F.datetostr(self.ui.de_vol_pl_end.date().toPyDate())
    CMS.save_tmp_path('pl_diapazon_month',str_d)

def load_diapazon_month(self: mywindow):
    try:
        list_str_d =  CMS.load_tmp_path('pl_diapazon_month').split(';')
        self.ui.de_vol_pl.setDate(F.strtodate(list_str_d[0]))
        self.ui.de_vol_pl_end.setDate(F.strtodate(list_str_d[1]))
    except:
        pass

#TODO двойным кликом тянуть в фильтр


