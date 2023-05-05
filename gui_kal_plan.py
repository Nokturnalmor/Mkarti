import project_cust_38.Cust_Qt as CQT
import project_cust_38.Cust_Functions as F
import project_cust_38.Cust_SQLite  as CSQ
import kal_plan as KPL
import datetime
from dateutil.rrule import rrule, MONTHLY
from  copy import deepcopy
def update_local_graf(self,update=False):
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
            return  [F.datetostr(dt,"%Y-%m") for dt in rrule(MONTHLY, dtstart=F.datetime_to_date(min_date), until=F.datetime_to_date(max_date))]
        def genetrate_cld(self,list_of_month):
            rez = dict()
            for day in self.DICT_CLD.keys():
                if F.datetostr(day,'%Y-%m') in list_of_month:
                    rez[day] = self.DICT_CLD[day]
                    rez[day]['podr'] = dict()
                    for podr in self.DICT_PODR.keys():
                        if self.DICT_PODR[podr]['Порядок'] > 1:
                            rez[day]['podr']['план_' + podr] = ""
                            rez[day]['podr']['факт_'+ podr] = ""
            return rez


        list_of_month = load_list_of_month(self,min_date,max_date)

        dict_cld = genetrate_cld(self,list_of_month)

        return dict_cld

    def save_form_db(self,dict_form,pnom):
        data = F.to_binary_pickle(dict_form)
        CSQ.zapros(self.db_kplan,f"""UPDATE plan SET local_graf = ? WHERE Пномер == ?;""",spisok_spiskov=[data,pnom])



    def fill_date(self,dict_form,list):

        def search_norma(name,list,podr):
            # =================
            capacity = 8
            vid_etap = name.split("__")[-1]
            for j in range(len(list[0])):
                field = list[0][j]
                if podr == field.split('.')[0]:
                    if "Нчас_" + vid_etap == field.split('.')[1]:
                        capacity = list[1][j]
                        break
                    if "Нмин_" + vid_etap == field.split('.')[1]:
                        capacity = round(list[1][j]/60,2)
                        break
            # =================
            return  capacity
        def fill_date_to_form(dict_form, podr, date_nach,date_zav,etap,capacity):
            rab_dn = 0
            for date in dict_form.keys():
                if date >= date_nach and date <= date_zav:
                    if dict_form[date]['Выходные'] == 0:
                        rab_dn +=1
            if rab_dn == 0:
                return dict_form

            mosh = round(capacity / (rab_dn),3)
            for date in dict_form.keys():
                if date >= date_nach and date <= date_zav:
                    if date > date_zav:
                        break
                    if dict_form[date]['Выходные'] == 0:
                        dict_form[date]['podr'][podr] = {"Время_час" : mosh, 'Этап' : etap,
                                                         "Начало" : F.datetostr(date_nach,"%d.%m.%y"),
                                                         "Конец" : F.datetostr(date_zav,"%d.%m.%y") }
            return dict_form
        rez = ''
        dict_process = dict()
        for podr in self.DICT_PODR.keys():
            if self.DICT_PODR[podr]['Порядок'] > 1:
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
                capacity = 8
                for vid in dict_process[podr][etap].keys():
                    if vid == 'Норм':
                        capacity = dict_process[podr][etap][vid]
                    if vid == 'ед':
                        date_nach = date_zav = dict_process[podr][etap][vid]['val']
                    if vid == 'нач':
                        date_nach = dict_process[podr][etap][vid]['val']
                    if vid == 'зав':
                        date_zav = dict_process[podr][etap][vid]['val']
                if date_nach == "" or date_zav == '':
                    print('')
                else:
                    dict_form = fill_date_to_form(dict_form,podr,date_nach,date_zav,etap,capacity)
        return dict_form

    pnom = load_pnom_from_tbl(self)

    fl_upd = True
    if update == False:
        data =  CSQ.zapros(self.db_kplan,f"""SELECT local_graf FROM plan WHERE Пномер == {pnom}""")
        if data[1][0] != '' and data[1][0] != 'None':
            dict_form = F.from_binary_pickle(data[1][0])
            if dict_form != None:
                fl_upd = False
    if fl_upd:

        list, list_conf = KPL.load_db(self, pnom)

        min_date, max_date = load_min_max_date(self,pnom,list)

        dict_form = load_dict_form(self,min_date,max_date)

        dict_form = fill_date(self,dict_form,list)

        save_form_db(self,dict_form,pnom)

    fill_local_table(self,dict_form)

def load_pnom_from_tbl(self):
    tbl = self.ui.tbl_kal_pl

    r = tbl.currentRow()
    if r == None or r == -1:
        return
    nk_pnom = CQT.nom_kol_po_imen(tbl, 'plan.Пномер')
    pnom = int(tbl.item(r, nk_pnom).text())
    return pnom

def load_form_db(self,pnom):
    rez = CSQ.zapros(self.db_kplan,f"""SELECT local_graf FROM plan WHERE Пномер == {pnom};""",)
    data = F.from_binary_pickle(rez)
    return data


def fill_local_table(self,dict_form=''):
    def generate_list(dict_form):
        list_tbl = []

        set_podr = set()
        list_date = ['']
        list_vih = ['']
        list_dned = ['']

        for date in dict_form.keys():
            list_date.append(date)
            list_vih.append(dict_form[date]['Выходные'])
            list_dned.append(dict_form[date]['День недели'])
            for podr in dict_form[date]['podr'].keys():
                set_podr.add(podr)
        list_tbl.append(list_date)
        list_tbl.append(list_vih)
        list_tbl.append(list_dned)
        tmp_list_podr = list(set_podr)
        tmp_list_podr.sort()
        list_podr = []
        for i in range(len(self.DICT_PODR)):
            for podr in tmp_list_podr:
                podr_cut ="_".join(podr.split("_")[1:])
                if podr_cut in self.DICT_PODR:
                    if self.DICT_PODR[podr_cut]['Порядок'] == i:
                        list_podr.append(podr)
        for podr in list_podr:
            list_tbl.append([podr])

        list_tbl_info = deepcopy(list_tbl)
        for i in range(1,len(list_tbl[0])):
            for j in range(3,len(list_tbl)):
                podr = list_tbl[j][0]
                day = list_tbl[0][i]
                val = dict_form[day]['podr'][podr]
                time_rab = ''
                if val != '':
                    time_rab = round(val['Время_час'])
                list_tbl[j].append(time_rab)
                list_tbl_info[j].append(val)
            list_tbl[0][i] = F.datetostr(list_tbl[0][i],"%d\n%m\n%y")
        return list_tbl,list_tbl_info

    if dict_form == '':
        pnom = load_pnom_from_tbl(self)
        dict_form = load_form_db(self,pnom)

    list_tbl,list_tbl_info = generate_list(dict_form)
    self.list_tbl_info = list_tbl_info
    CQT.fill_wtabl(list_tbl,self.ui.tbl_preview,ogr_maxshir_kol=29, height_row=22, colorful_edit=False,auto_type= False,head_column=0)
    self.ui.tbl_preview.setColumnWidth(0, 150)
    self.ui.tbl_preview.setRowHidden(0,True)
    self.ui.tbl_preview.setRowHidden(1, True)
    oforml_table(self, self.ui.tbl_preview)
    pass

def oforml_table(self,tbl):
    for j in range(tbl.columnCount()):
        if tbl.item(0,j).text() == "1":
            CQT.ust_color_text_header_wtab_horisontal(tbl, j, 200, 11, 11, 10, True)
        else:
            CQT.ust_color_text_header_wtab_horisontal(tbl, j, 11, 11, 11, 10, False)
        for i in range(1,tbl.rowCount()):
            if tbl.item(i, j).text() != "":
                podr = tbl.item(i, 0).text().replace('факт_','').replace('план_','')
                r = 233
                g = 233
                b = 233
                if podr in self.DICT_PODR:
                    r, g, b = self.DICT_PODR[podr]['Цвет'].split(";")
                CQT.ust_color_wtab(tbl,i,j,r,g,b)
    for i in range(1, tbl.rowCount()):
        podr = tbl.verticalHeaderItem(i).text().replace('факт_', '').replace('план_', '')
        r = 233
        g = 233
        b = 233
        if podr in self.DICT_PODR:
            r, g, b = self.DICT_PODR[podr]['Цвет'].split(";")
        CQT.ust_color_text_header_wtab_vertical(tbl, i, r, g, b,10, True)
    pass

def load_tbl_gant(self):
    def generate_full_table(self):
        pass

    def oform_tbl(self):
        pass

    list = generate_full_table(self)
    CQT.fill_wtabl(list,self.ui.tbl_pl_gaf,height_row=10)

    oform_tbl(self)

