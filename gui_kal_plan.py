from __future__ import annotations

import copy
import project_cust_38.Cust_Qt as CQT
import project_cust_38.Cust_Functions as F
import project_cust_38.Cust_SQLite  as CSQ
import kal_plan as KPL
import datetime
from dateutil.rrule import rrule, MONTHLY
from  copy import deepcopy
import project_cust_38.Cust_mes as CMS

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from MKart import mywindow


def hover_tbl_pl_gaf(self, event):
    tbl = self.ui.tbl_pl_gaf
    row, column = CQT.get_hover_row_col(self, tbl, event)
    if row == False or column == False:
        return
    val = tbl.item(row, column).text()
    if val != '':
        load_info_select_block(self,tbl,row,column)
    else:
        load_info_select_block(self, tbl, row, column,True)

def hover_tbl_preview(self, event):
    tbl = self.ui.tbl_preview
    row, column = CQT.get_hover_row_col(self, tbl, event)
    if row == False or column == False:
        return
    val = tbl.item(row, column).text()
    if val != '':
        load_info_select_block(self,tbl,row,column)
    else:
        load_info_select_block(self, tbl, row, column,True)
def max_mosh(self:mywindow, day, podr:str):
    podr = podr.split('план_')[-1]
    podr = podr.split('факт_')[-1]
    try:
        return self.KPLAN_max_mosh[day][podr]
    except:
        return 'err'

def load_info_select_block(self,tbl,r = '',c = '',clear=False):
    if clear:
        CQT.statusbar_text(self, '')
        tbl.setToolTip('')
        return
    if r =="":
        r = tbl.currentRow()
    if c == "":
        c = tbl.currentColumn()
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
    mosh = ''
    try:
        day = self.list_tbl_info[0][c]
        podr = self.list_tbl_info[r+1][0]
        mosh = f'Мощность: {max_mosh(self,day,podr)} н-час.'
    except:
        pass
    CQT.statusbar_text(self,
                       f'{self.glob_kpl_summ_selct_tbl} |  {info} | {mosh}' )



def update_local_graf(self,update=False,pnom:int = 0,fill_gant=True):
    if pnom == 0:
        tbl = self.ui.tbl_kal_pl
        r = tbl.currentRow()
        if r == None or r == -1:
            return
        nk_pnom = CQT.nom_kol_po_imen(tbl, 'plan.Пномер')
        pnom = int(tbl.item(r, nk_pnom).text())

    if 'shift' in  CQT.get_key_modifiers(self):
        update = True


    def load_min_max_date(self,pnom,list):
        min_date = ''
        max_date = ''
        for item in list[1]:
            if F.is_date(item,"%Y-%m-%d"):
                dt_date = F.strtodate(item,"%Y-%m-%d")
                if min_date == '':
                    min_date = dt_date
                if dt_date < min_date:
                    min_date = dt_date

                if max_date == '':
                    max_date = dt_date
                if dt_date > max_date:
                    max_date = dt_date
        if min_date < F.strtodate("2022-09-01","%Y-%m-%d"):
            min_date = F.strtodate("2022-09-01","%Y-%m-%d")
        return min_date,max_date


    def load_dict_form(self,min_date,max_date):
        def load_list_of_month(self,min_date,max_date):
            list_of_month = [F.datetostr(dt,"%Y-%m") for dt in rrule(MONTHLY, dtstart=F.datetime_to_date(min_date), until=F.datetime_to_date(max_date))]
            return  list_of_month
        def genetrate_cld(self,list_of_month):
            rez = dict()
            list_days = sorted([k for  k in  self.data_kpl.DICT_CLD.keys()])
            for day in list_days:
                if F.datetostr(day,'%Y-%m') in list_of_month:
                    rez[day] = self.data_kpl.DICT_CLD[day]
                    rez[day]['podr'] = dict()
                    for podr in self.data_kpl.DICT_PODR.keys():
                        if self.data_kpl.DICT_PODR[podr]['Порядок'] >= 0:
                            rez[day]['podr']['план_' + podr] = ""
                            rez[day]['podr']['факт_'+ podr] = ""
            return rez


        list_of_month = load_list_of_month(self,min_date,max_date)

        dict_cld = genetrate_cld(self,list_of_month)

        return dict_cld


    def save_form_db(self,dict_form,pnom):
        tbl = self.ui.tbl_kal_pl
        data = F.to_binary_pickle(dict_form)
        CSQ.zapros(self.db_kplan,f"""UPDATE plan SET local_graf = ? WHERE Пномер == ?;""",spisok_spiskov=[data,pnom])
        nk_pnom = CQT.nom_kol_po_imen(tbl,'plan.Пномер')
        nk_graf = CQT.nom_kol_po_imen(tbl,'plan.local_graf')
        for i in range(tbl.rowCount()):
            if tbl.item(i,nk_pnom).text() == str(pnom):
                tbl.item(i, nk_graf).setText(str(data))
                break
        return data

    def fill_date(self,dict_form,list,pnom,proj,poz, napr):


        def search_norma(name,list,podr):
            # =================
            capacity = 0
            vid_etap = name.split("__")[-1]
            for j in range(len(list[0])):
                field = list[0][j]
                if podr == field.split('.')[0]:
                    left_str = "Нчас_" + vid_etap
                    if left_str.lower() == field.split('.')[1].lower():
                        capacity = list[1][j]
                        break
                    left_str = "Нмин_" + vid_etap
                    if left_str.lower() == field.split('.')[1].lower():
                        capacity = round(list[1][j]/60,2)
                        break
            # =================
            return  capacity


        def fill_date_to_form(dict_form, podr, date_nach,date_zav,etap,capacity,name_nach,name_zav):
            fl_rab_dn = True
            rab_dn = 0
            for date in dict_form.keys():
                if date >= date_nach and date <= date_zav:
                    if dict_form[date]['Выходные'] == 0:
                        rab_dn +=1
            if rab_dn == 0:
                fl_rab_dn = False
                for date in dict_form.keys():
                    if date >= date_nach and date <= date_zav:
                        rab_dn += 1
            if rab_dn == 0:
                mosh = 0
            else:
                mosh = round(capacity / (rab_dn),3)

            for date in dict_form.keys():
                if date >= date_nach and date <= date_zav:
                    if date > date_zav:
                        break
                    if dict_form[date]['Выходные'] == 0 or not fl_rab_dn:
                        data_et = {"Время_час" : mosh, 'Этап' : etap,
                                                         "Начало" : F.datetostr(date_nach,"%d.%m.%y"),
                                                         "Конец" : F.datetostr(date_zav,"%d.%m.%y"),
                                                         "Имя_нз" : [name_nach,name_zav]}

                        if dict_form[date]['podr'][podr] != '':
                            dict_form[date]['podr'][podr].append(data_et)
                        else:
                            dict_form[date]['podr'][podr] = [data_et]

            return dict_form
        rez = ''
        dict_process = dict()
        for podr in self.data_kpl.DICT_PODR.keys():
            if self.data_kpl.DICT_PODR[podr]['Порядок'] >= 0:
                if podr not in dict_process:
                    dict_process['план_' + podr] = dict()
                    dict_process['факт_' + podr] = dict()
                for i in range(len(list[0])):
                    field = list[0][i]
                    if podr == field.split('.')[0]:
                        if "дата" in field.lower():
                            current_vid_pf = 'факт_' + podr
                            if "пдата" in field.lower():
                                current_vid_pf = 'план_' + podr
                            if "нач" in field.lower() or "зав" in field.lower():
                                name = field.lower().replace("нач",'').replace("зав",'')
                                capacity  = search_norma(name,list,podr)
                                if name not in dict_process[current_vid_pf]:
                                    dict_process[current_vid_pf][name] = dict()
                                dict_process[current_vid_pf][name]["Норм"] = capacity
                                if "нач" in field.lower():
                                    dict_process[current_vid_pf][name]["нач"] = dict()
                                    dict_process[current_vid_pf][name]["нач"]['val'] = ''
                                    if F.is_date(list[1][i], "%Y-%m-%d"):
                                        dict_process[current_vid_pf][name]["нач"]['val'] = F.strtodate(list[1][i],"%Y-%m-%d")
                                    dict_process[current_vid_pf][name]["нач"]['field'] = field
                                if "зав" in field.lower():
                                    dict_process[current_vid_pf][name]["зав"] = dict()
                                    dict_process[current_vid_pf][name]["зав"]['val'] = ''
                                    if F.is_date(list[1][i], "%Y-%m-%d"):
                                        dict_process[current_vid_pf][name]["зав"]['val'] = F.strtodate(list[1][i], "%Y-%m-%d")
                                    dict_process[current_vid_pf][name]["зав"]['field'] = field
                            else:
                                if field.lower() not in dict_process[current_vid_pf]:
                                    dict_process[current_vid_pf][field.lower()] = dict()
                                if "ед" not in dict_process[current_vid_pf][field.lower()]:
                                    dict_process[current_vid_pf][field.lower()]["ед"] = dict()
                                dict_process[current_vid_pf][field.lower()]["ед"]['val'] = ''
                                if F.is_date(list[1][i], "%Y-%m-%d"):
                                    dict_process[current_vid_pf][field.lower()]["ед"]['val'] = F.strtodate(list[1][i],"%Y-%m-%d")
                                dict_process[current_vid_pf][field.lower()]["ед"]['field'] = field
        for podr in dict_process.keys():
            for etap in dict_process[podr].keys():
                date_nach = ''
                date_zav = ''
                capacity = 0
                for vid in dict_process[podr][etap].keys():
                    if vid == 'Норм':
                        capacity = dict_process[podr][etap][vid]
                    if vid == 'ед':
                        date_nach = date_zav = dict_process[podr][etap][vid]['val']
                        name_nach = name_zav = dict_process[podr][etap][vid]['field']
                    if vid == 'нач':
                        date_nach = dict_process[podr][etap][vid]['val']
                        name_nach = dict_process[podr][etap][vid]['field']
                    if vid == 'зав':
                        date_zav = dict_process[podr][etap][vid]['val']
                        name_zav = dict_process[podr][etap][vid]['field']
                if date_nach == "" or date_zav == '':
                    print('')
                else:
                    dict_form = fill_date_to_form(dict_form,podr,date_nach,date_zav,etap,capacity,name_nach,name_zav)
        dict_form = [{'pnom':pnom,'proj':proj,'poz':poz,'napr':napr,'data':dict_form}]
        return dict_form

    dict_poz = load_dict_poz_from_sql(self,pnom)
    if dict_poz == False:
        return
    self.pnom_kplan_select = dict_poz['Пномер']
    fl_upd = True
    dict_form = []
    if update == False:
        data = dict_poz['local_graf']
        if data != '' and data != 'None':
            dict_form = F.from_binary_pickle(data)
            if dict_form != None:
                fl_upd = False
    data_bin = None
    if fl_upd:

        list, list_conf = KPL.load_db(self, dict_poz['Пномер'])

        min_date, max_date = load_min_max_date(self,dict_poz['Пномер'],list)

        dict_form = load_dict_form(self,min_date,max_date)

        dict_form = fill_date(self,dict_form,list,dict_poz['Пномер'],f"{dict_poz['№проекта']} {dict_poz['№ERP']}",dict_poz['Позиция'],dict_poz['Направление'])

        data_bin = save_form_db(self,dict_form,dict_poz['Пномер'])
    if fill_gant:
        fill_gant_table(self, self.ui.tbl_preview,'', dict_form, pnom)
    if fl_upd:
        return data_bin
    return

def move_left(self):
    move(self, -1)

def move_right(self):
    move(self, 1)

def move(self, direction = 1):
    def update_db(self,name_field,delta_nach,direction):
        table, field = name_field.split('.')
        name_field = 'НомПл'
        if table == "plan":
            name_field = 'Пномер'
        list_old_date = CSQ.zapros(self.db_kplan,
                              f"""SELECT {field} FROM {table} WHERE {name_field} == {self.pnom_kplan_select};""")
        if list_old_date == False or list_old_date == None or len(list_old_date) != 2:
            CQT.msgbox(f'ОШибка загрузки даты')
            return False
        old_date = list_old_date[-1][0]
        if not F.is_date(old_date,"%Y-%m-%d"):
            CQT.msgbox(f'ОШибка распознавания дат {old_date}')
            return False
        new_date = F.date_add_days(old_date, delta_nach*direction, "%Y-%m-%d", "%Y-%m-%d")
        CSQ.zapros(self.db_kplan,
                   f"""UPDATE {table} SET {field} = "{new_date}"  WHERE {name_field} == {self.pnom_kplan_select};""")



    r = self.ui.tbl_preview.currentRow()
    c = self.ui.tbl_preview.currentColumn()
    if self.ui.le_edit_local_gant_nach.text() == '':
        self.ui.le_edit_local_gant_nach.setText('0')
    if self.ui.le_edit_local_gant_kon.text() == '':
        self.ui.le_edit_local_gant_kon.setText('0')
    if not F.is_numeric(self.ui.le_edit_local_gant_nach.text()) or not F.is_numeric(self.ui.le_edit_local_gant_kon.text()):
        CQT.msgbox(f'Смещение не число')
        return
    if type(self.list_tbl_info[r + 1][c]) != list:
        return
    else:
        if "Имя_нз" not in self.list_tbl_info[r + 1][c][0]:
            return
    name_nach = deepcopy(self.list_tbl_info[r + 1][c][0])["Имя_нз"][0]
    name_zav =  deepcopy(self.list_tbl_info[r + 1][c][0])["Имя_нз"][1]
    delta_nach = F.valm(self.ui.le_edit_local_gant_nach.text())
    delta_kon = F.valm(self.ui.le_edit_local_gant_kon.text())
    fl = False
    if delta_nach != 0:
        if update_db(self,name_nach,delta_nach,direction) == False:
            return
        c = c + delta_nach * direction
        if c >= self.ui.tbl_preview.columnCount():
            c = self.ui.tbl_preview.columnCount()-1
        if c < 1:
            c = 1
        fl = True
    if delta_kon != 0:
        if update_db(self,name_zav,delta_kon,direction) == False:
            return
        c = c + delta_nach * direction
        if c >= self.ui.tbl_preview.columnCount():
            c = self.ui.tbl_preview.columnCount()-1
        if c < 1:
            c = 1
        fl = True
    if fl:
        update_local_graf(self,True)
        self.ui.tbl_preview.setCurrentCell(r,c)
        hide_free_columns(self,self.ui.tbl_preview)


def load_dict_poz_from_sql(self,pnom:int):
    query = CSQ.zapros(self.db_kplan, f"""SELECT plan.Пномер, plan.Позиция, plan.local_graf, plan.Приоритет, 
            пл_оуп.№проекта, пл_оуп.№ERP, napravl_deyat.Псевдоним as Направление FROM plan INNER JOIN
            пл_оуп ON пл_оуп.НомПл = plan.Пномер,
            napravl_deyat ON napravl_deyat.Пномер = plan.Направление_деятельности
             WHERE plan.Пномер == {pnom}""", rez_dict=True)
    if query == False or len(query) == 0:
        return False
    return query[0]

def load_form_db(self,pnom):
    rez = CSQ.zapros(self.db_kplan,f"""SELECT local_graf FROM plan WHERE Пномер == {pnom};""",)
    data = F.from_binary_pickle(rez)
    return data



@F.vrem_vip_cls_func_args
def fill_gant_table(self: mywindow , tbl, tbl_filtr='', dict_form='', pnom=0):

    def generate_list(dict_form_list,min_date,max_date):
        list_tbl = []
        DICT_DAY_NAME = {1:'Пн',2:'Вт',3:'Ср',4:'Чт',5:'Пт',6:'Сб',7:'Вс'}
        set_podr = set()
        list_sablon = ['','','','','']
        list_date = ['Этап','Пномер','Проект','Поз.','Напр.']
        list_vih = copy.deepcopy(list_sablon)
        list_dned = copy.deepcopy(list_sablon)
        for dict_form_item in dict_form_list:
            for date in dict_form_item['data'].keys():
                if max_date > date >= min_date:
                    list_date.append(date)
                    list_vih.append(dict_form_item['data'][date]['Выходные'])
                    list_dned.append(dict_form_item['data'][date]['День недели'])
            list_tbl.append(list_date)
            list_tbl.append(list_vih)
            list_tbl.append(list_dned)
            break
        list_tbl_info = deepcopy(list_tbl)
        set_rows_to_add = {_ for _ in range(len(list_tbl))}
        for dict_form_item in dict_form_list:
            for date in dict_form_item['data'].keys():
                for podr in dict_form_item['data'][date]['podr'].keys():
                    set_podr.add(podr)
            tmp_list_podr = list(set_podr)
            tmp_list_podr.sort()
            list_podr = []
            for i in range(len(self.data_kpl.DICT_PODR)):
                for podr in tmp_list_podr:
                    podr_cut ="_".join(podr.split("_")[1:])
                    if podr_cut in self.data_kpl.DICT_PODR:
                        if self.data_kpl.DICT_PODR[podr_cut]['Порядок'] == i:
                            list_podr.append(podr)
            start_row = len(list_tbl)
            for podr in list_podr:
                list_tbl.append([podr,dict_form_item['pnom'],dict_form_item['proj'],dict_form_item['poz'],dict_form_item['napr']])
                list_tbl_info.append([podr,dict_form_item['pnom'],dict_form_item['proj'],dict_form_item['poz'],dict_form_item['napr']])

            for i in range(len(list_sablon),len(list_tbl[0])):
                for j in range(start_row,len(list_tbl)):
                    podr = list_tbl[j][0]
                    day = list_tbl[0][i]
                    if podr in dict_form_item['data'][day]['podr']:
                        list_vals = dict_form_item['data'][day]['podr'][podr]
                        time_rab = ''
                        if list_vals != '':
                            time_rab = 0
                            for val in list_vals:
                                time_rab += round(val['Время_час'])
                        list_tbl[j].append(time_rab)
                        list_tbl_info[j].append(list_vals)
                        if time_rab != '':
                            set_rows_to_add.add(j)
                    else:
                        CQT.msgbox(f'{podr} отсутствует в локальном графике, нужно обновить Пномер{str(dict_form_item["pnom"])}')
                        return None, None

        list_tbl = [val for _,val  in enumerate(list_tbl) if _ in set_rows_to_add]
        list_tbl_info = [val for _,val  in enumerate(list_tbl_info) if _ in set_rows_to_add]


        for i in range(len(list_sablon), len(list_tbl[0])):
            list_tbl[0][i] = F.datetostr(list_tbl[0][i], f"%d\n%m\n%y\n{DICT_DAY_NAME[int(list_tbl[2][i])]}")
        return list_tbl,list_tbl_info
    if dict_form == '':
        if pnom == 0:
            tbl = self.ui.tbl_kal_pl
            r = tbl.currentRow()
            if r == None or r == -1:
                return
            nk_pnom = CQT.nom_kol_po_imen(tbl, 'plan.Пномер')
            pnom = int(tbl.item(r, nk_pnom).text())
        dict_poz = load_dict_poz_from_sql(self,pnom)
        if dict_poz == False:
            return
        dict_form = F.from_binary_pickle(dict_poz['local_graf'])

    min_date = F.strtodate('2020-01-01 00:00:01', "%Y-%m-%d %H:%M:%S")
    max_date = F.strtodate('2220-01-01 00:00:01', "%Y-%m-%d %H:%M:%S")
    if self.kpl_mode == 1:
        month = self.ui.de_vol_pl.date().toPyDate()
        month_end = self.ui.de_vol_pl_end.date().toPyDate()
        min_date = F.nach_kon_date(F.date_to_datetime(month,0,0,1),'','m','')[0]
        max_date = F.nach_kon_date(F.date_to_datetime(month_end,23,59,59),'','m','')[1]

    self.list_tbl,self.list_tbl_info = generate_list(dict_form,min_date,max_date)
    self.dict_form_kpl = dict_form
    if self.list_tbl_info == None:
        return
    oforml_table(self, tbl, tbl_filtr)
    pass

@F.vrem_vip_cls_func_args
def hide_free_columns(self,tbl):
    count_tbl_field = 5
    for j in range(count_tbl_field, tbl.columnCount()):
        self.ui.tbl_preview.setColumnHidden(j, False)
    for j in range(count_tbl_field, tbl.columnCount()):
        fl_hide = True
        for i in range(2,tbl.rowCount()):
            if tbl.item(i,j).text() != '':
                fl_hide = False
                break
        if fl_hide:
            tbl.setColumnHidden(j,True)
        else:
            break

    for j in range(tbl.columnCount()-1,-1,count_tbl_field):
        fl_hide = True
        for i in range(2,tbl.rowCount()):
            if tbl.item(i,j).text() != '':
                fl_hide = False
                break
        if fl_hide:
            tbl.setColumnHidden(j,True)
        else:
            break
    tbl.resizeColumnsToContents()

@F.vrem_vip_cls_func_args
def oforml_table(self,tbl, tbl_filtr= ''):
    count_tbl_field = 5
    CQT.fill_wtabl(self.list_tbl,tbl,min_width_col= int(4*0.8),
                   height_row=self.val_masht*2, colorful_edit=False,auto_type= False,head_column=0,set_editeble_col_nomera={},hide_head_column=False)

    for j in range(1,count_tbl_field):
        CQT.ust_color_text_header_wtab_horisontal(tbl, j, 11, 11, 11, self.val_masht*0.7, False)
        for i in range(3, len(self.list_tbl_info)):
            CQT.font_cell_size_format(tbl, i - 1, j, self.val_masht)
    for j in range(count_tbl_field,len(self.list_tbl_info[0])):
        if self.list_tbl_info[1][j] == 1:
            CQT.ust_color_text_header_wtab_horisontal(tbl, j, 200, 11, 11, self.val_masht*0.8, True)
        else:
            CQT.ust_color_text_header_wtab_horisontal(tbl, j, 11, 11, 11, self.val_masht*0.7, False)
        for i in range(3,len(self.list_tbl_info)):
            if self.list_tbl_info[i][j] != "":
                podr = self.list_tbl_info[i][0].replace('факт_','').replace('план_','')
                r = 233
                g = 233
                b = 233
                if podr in self.data_kpl.DICT_PODR:
                    r, g, b = self.data_kpl.DICT_PODR[podr]['Цвет'].split(";")
                for item in self.list_tbl_info[i][j]:
                    CQT.dob_color_wtab(tbl,i-1,j,int(r),int(g),int(b))
                CQT.font_cell_size_format(tbl,i-1,j,self.val_masht)
    for i in range(3,len(self.list_tbl_info)):
        podr = self.list_tbl_info[i][0].replace('факт_', '').replace('план_', '')
        r = 233
        g = 233
        b = 233
        if podr in self.data_kpl.DICT_PODR:
            r, g, b = self.data_kpl.DICT_PODR[podr]['Цвет'].split(";")
        CQT.ust_color_text_header_wtab_vertical(tbl, i-1, r, g, b, self.val_masht*0.8, True)
    if self.kpl_mode == 0:
        hide_free_columns(self,tbl)

    #self.ui.tbl_preview.setColumnWidth(0, self.val_masht*7.5)

    tbl.setRowHidden(0, True)
    tbl.setRowHidden(1, True)

    if tbl_filtr != '':
        fields_hide = ['Пномер']
        for field in fields_hide:
            try:
                tbl.setColumnHidden(CQT.nom_kol_po_imen(tbl, field), True)
            except:
                pass

        CMS.zapolnit_filtr(self, tbl_filtr, tbl,hidden_scroll=True)
        tbl_filtr.setVerticalHeaderLabels(['план_факт_подр'])
        tbl_filtr.setRowHeight(0, 25)
        for j in range(1, len(self.list_tbl_info[0])):
            if self.list_tbl_info[1][j] == 1:
                CQT.ust_color_text_header_wtab_horisontal(tbl_filtr, j, 200, 11, 11, self.val_masht * 0.5, False)
            else:
                CQT.ust_color_text_header_wtab_horisontal(tbl_filtr, j, 11, 11, 11, self.val_masht * 0.5, False)
        CMS.update_width_filtr(tbl,tbl_filtr)
    else:
        fields_hide = ['Этап','Пномер',"Проект","Поз.","Напр."]
        for field in fields_hide:
            try:
                tbl.setColumnHidden(CQT.nom_kol_po_imen(tbl, field), True)
            except:
                pass