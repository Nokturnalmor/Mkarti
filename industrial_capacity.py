import project_cust_38.Cust_Functions as F
import project_cust_38.Cust_SQLite as CSQ
import project_cust_38.Cust_Qt as CQT
import project_cust_38.Cust_mes as CMS
from PyQt5 import QAxContainer
from PyQt5.QtCore import Qt
from pathlib import Path
from PyQt5.QtGui import QPainter, QBrush, QPen, QPixmap, QColor
import project_cust_38.Cust_QtGui as CGUI

def load_tabel_in_db(self):
    if not check_load_tabel_in_db(self):
        return
    if not CQT.msgboxgYN(f'Точно обновить таблицу на {self.ui.cmb_tabeli.currentText()}?'):
        return
    if not CMS.user_access(self.bd_naryad,'update_tabel',F.user_name(),'Нет прав'):
        return
    list_of_tabel  = CQT.spisok_iz_wtabl(self.ui.tbl_tabeli,sep='',shapka=True,rez_dict=True)
    name_tbl_db = F.datetostr(F.strtodate(self.ui.cmb_tabeli.currentText().split(' ')[0], '%d.%m.%Y'), 'mtdz_%Y_%m_%d')
    tabel_db = CSQ.zapros(self.db_users,f"""SELECT * FROM {name_tbl_db}""",shapka=True,rez_dict=True)
    change_count = 0
    for item in list_of_tabel:
        flag_naid = False
        for item_db in tabel_db:
            if item['ФИО'] ==  ' '.join(item_db['ФИО'].split(' ')[:3]):
                flag_naid = True
                change_count = compare_user(self, item, item_db, name_tbl_db, item_db['ФИО'],change_count)
        if flag_naid == False:
            CQT.msgbox(f'Не найден {item["ФИО"]} в МЕС, нужно перезагрузить сотрудников через рейтинг.')
    CQT.msgbox(f'Обновлено успешно {change_count} изменений')


def compare_user(self, item, item_db, name_tbl, fio,change_count):
    for day in item.keys():
        if day not in ('ФИО','ИТОГ'):
            for dat_db in item_db.keys():
                if F.is_date(dat_db,'d_%Y_%m_%d'):
                    day_db =  F.strtodate(dat_db,'d_%Y_%m_%d').day
                    if day_db == int(day):
                        if F.is_numeric(item[day]):
                            item[day] = F.valm(item[day])
                        else:
                            item[day] = 0
                        if item[day] != item_db[dat_db]:
                            rez = CSQ.zapros(self.db_users,f"""UPDATE {name_tbl} SET {dat_db} = {item[day]} WHERE ФИО == '{fio}';""")
                            change_count+=1
                            if rez == None:
                                CQT.msgbox(f'ОШибка обновления попробуй позже')
                                return
    return change_count


def check_load_tabel_in_db(self):
    if self.ui.cmb_tabeli.currentText() == '':
        CQT.migat_obj(self,2,self.ui.cmb_tabeli,'Не выбран месяц')
        return False
    if self.ui.tbl_tabeli.columnCount() < 3:
        CQT.migat_obj(self,2,self.ui.tbl_tabeli,'Не корректно заполнена таблица')
        return
    try:
        name_tbl_db =F.datetostr(F.strtodate(self.ui.cmb_tabeli.currentText().split(' ')[0],'%d.%m.%Y'),'mtdz_%Y_%m_%d')
    except:
        CQT.migat_obj(self,2,self.ui.cmb_tabeli,'Не распознать месяц')
        return False
    return True



def load_from_bufer():
    COUNT_TRY = 3
    n_try = 1
    buf = False
    while n_try < COUNT_TRY:
        try:
            buf = F.paste_bufer('j')
            break
        except:
            n_try+=1
            pass
            return
    if not buf:
        CQT.msgbox('Что-то пошло не так, попробуй еще')
        return
    if buf != '':
        try:
            list = [[i.replace('\r', '') for i in _.split('\t')] for _ in buf.split('\n')]
        except:
            CQT.msgbox(f'Не подходящий формат')
            return
        list[0][0] = list[0][0].replace('j', '')
        return list


def add_list_of_months_to_cmb(self):
    cmb = self.ui.cmb_tabeli
    list_of_month = CSQ.zapros(self.db_users, """SELECT name from sqlite_master where type = 'table';""")
    cmb.clear()
    cmb.addItem('')
    for item in list_of_month:
        if item[0][:5] == 'mtdz_':
            if F.now('').year == F.strtodate(item[0], "mtdz_%Y_%m_%d").year:
                date = F.strtodate(item[0], "mtdz_%Y_%m_%d")
                read_name = F.datetostr(date, f'%d.%m.%Y ({F.month_rus_from_date(date, "", False)})')
                cmb.addItem(read_name)

@CQT.onerror
def load_tabel_from_bufer(self):
    list = load_from_bufer()
    if list == None:
        return
    CQT.fill_wtabl(list, self.ui.tbl_tabeli)
    list = CQT.spisok_iz_wtabl(self.ui.tbl_tabeli, '', True)
    dict_users = dict()
    index = 1

    nk_empl = F.nom_kol_po_im_v_shap(list, 'Сотрудник')
    nk_nom = F.nom_kol_po_im_v_shap(list, '№')
    nk_first_day = F.nom_kol_po_im_v_shap(list, '01')
    if nk_empl == None:
        return
    for i, item in enumerate(list):
        if len(item) >= nk_empl:
            if item[nk_nom] != '' and F.is_numeric(item[nk_nom]):
                if F.valm(item[nk_nom]) == index:
                    fio = item[nk_empl]
                    #print(fio)
                    dict_days = dict()
                    for j in range(i + 1, len(list)):
                        if list[j][nk_nom] != '' and F.is_numeric(list[j][nk_nom]):
                            break
                        for col in range(nk_first_day, len(list[j])):
                            if list[j][col] != '' and F.is_numeric(list[0][col]):
                                dict_days[list[0][col]] = list[j][col]
                    dict_users[fio] = dict_days
                    index += 1
    oform_tabel_to_table(self,dict_users)
    CMS.zapolnit_filtr(self,self.ui.tbl_tabeli_filtr,self.ui.tbl_tabeli)


def oform_tabel_to_table(self, dict_users: dict):
    set_days = set()
    for user in dict_users.keys():
        for day in dict_users[user].keys():
            set_days.add(day)
    list_of_days = sorted(list(set_days))
    rez = [['ФИО']]
    for day in list_of_days:
        rez[0].append(day)
    rez[0].append('ИТОГ')
    list_of_days.append('ИТОГ')
    for user in dict_users.keys():
        rez.append([user])
        summ = 0
        for day in list_of_days:
            if day == 'ИТОГ':
                rez[-1].append(round(summ, 2))
                break
            if day in dict_users[user]:
                rez[-1].append(dict_users[user][day])
                if F.is_numeric(dict_users[user][day]):
                    summ += F.valm(dict_users[user][day])
            else:
                rez[-1].append(0)
    CQT.zapoln_wtabl(self, rez, self.ui.tbl_tabeli, separ='', isp_shapka=True, ogr_maxshir_kol=600,min_shir_col=60)
    for i in range(self.ui.tbl_tabeli.rowCount()):
        for j in range(1, self.ui.tbl_tabeli.columnCount()):
            if not F.is_numeric(self.ui.tbl_tabeli.item(i, j).text()):
                if self.ui.tbl_tabeli.item(i, j).text() == 'В':
                    CQT.ust_color_wtab(self.ui.tbl_tabeli, i, j, 250, 235, 215)
                else:
                    CQT.ust_color_wtab(self.ui.tbl_tabeli, i, j, 255, 250, 205)
            else:
                if F.valm(self.ui.tbl_tabeli.item(i, j).text()) == 0:
                    CQT.ust_color_wtab(self.ui.tbl_tabeli, i, j, 255, 250, 205)
                else:
                    if F.valm(self.ui.tbl_tabeli.item(i, j).text()) <8:
                        CQT.ust_color_wtab(self.ui.tbl_tabeli, i, j, 255, 230, 235)


def check_time(txt):
    if F.is_date(txt,"%H:%M"):
        return True
    CQT.msgbox('Не верный формат времени')
    return False

def check_val(txt):
    if F.is_numeric(txt):
        return True
    CQT.msgbox('Не верный формат числа')
    return False

def cellChanged(self, row, col):
    if CMS.user_access(self.bd_naryad,'rab_mesta_edit',F.user_name()) == False:
        return

    ima_col = self.ui.tbl_rc.horizontalHeaderItem(col).text()
    if ima_col == None:
        CQT.msgbox('Ошибка уточнения полей')
        return
    conn, cur = CSQ.connect_bd(self.db_users)
    znach = self.ui.tbl_rc.item(row, col).text()
    nk_pnom = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Пномер')
    pnom = int(self.ui.tbl_rc.item(row, nk_pnom).text())

    if ima_col == 'Время_начала_1' or ima_col == 'Время_конца_1' or ima_col == 'Время_начала_2' or \
        ima_col == 'Время_конца_2' or ima_col == 'Время_начала_3' or ima_col == 'Время_конца_3':
        if check_time(znach) == False:
            old_znach = CSQ.zapros(self.db_users,f"""SELECT {ima_col} FROM rab_mesta WHERE Пномер = {pnom}""",conn=conn)[-1][0]
            self.ui.tbl_rc.item(row,col).setText(str(old_znach))
            CSQ.close_bd(conn)
            return
    if ima_col == 'Нераб_мин1' or ima_col == 'Между_нар_мин1' or ima_col == 'Коэфф_производит1' or \
            ima_col == 'Нераб_мин2' or ima_col == 'Между_нар_мин2' or ima_col == 'Коэфф_производит2' or \
            ima_col == 'Нераб_мин3' or ima_col == 'Между_нар_мин3' or ima_col == 'Коэфф_производит3':
        if check_val(znach) == False:
            old_znach = CSQ.zapros(self.db_users,f"""SELECT {ima_col} FROM rab_mesta WHERE Пномер = {pnom}""",conn=conn)[-1][0]
            self.ui.tbl_rc.item(row,col).setText(str(old_znach))
            CSQ.close_bd(conn)
            return
    CSQ.zapros(self.db_users, f"""UPDATE rab_mesta SET {ima_col} = "{znach}" WHERE Пномер = {pnom}""",conn=conn)
    CSQ.close_bd(conn)

def load_deficit_emploee(self):
    rez = [['Пномер_РМ','Расположение','Прозвище','Профессия','Смена']]
    dict_rez_itog = dict()
    dict_rez_all = dict()
    conn, cur = CSQ.connect_bd(self.db_users)
    zapros = """SELECT rab_mesta.Пномер, places.adress as Расположение, rab_mesta.Прозвище, professions.имя , rab_mesta.ФИО_1, rab_mesta.ФИО_2, rab_mesta.ФИО_3
            FROM rab_mesta 
            INNER JOIN professions ON professions.код == rab_mesta.Код_профессии
            INNER JOIN places ON places.serial == rab_mesta.Расположение
             """
    spis_rm = CSQ.zapros(self.db_users, zapros, shapka=True, conn=conn)
    nk_r_fio1 = F.nom_kol_po_im_v_shap(spis_rm, 'ФИО_1')
    nk_r_fio2 = F.nom_kol_po_im_v_shap(spis_rm, 'ФИО_2')
    nk_r_fio3 = F.nom_kol_po_im_v_shap(spis_rm, 'ФИО_3')
    nk_r_pnom = F.nom_kol_po_im_v_shap(spis_rm, 'Пномер')
    nk_r_rasp = F.nom_kol_po_im_v_shap(spis_rm, 'Расположение')
    nk_r_proz = F.nom_kol_po_im_v_shap(spis_rm, 'Прозвище')
    nk_r_profes = F.nom_kol_po_im_v_shap(spis_rm, 'имя')
    for item in spis_rm[1:]:
        if item[nk_r_fio1] == 1:
            deficit_cm = '1'
            rez.append([item[nk_r_pnom], item[nk_r_rasp], item[nk_r_proz], item[nk_r_profes], deficit_cm])
            if item[nk_r_profes] not in dict_rez_itog:
                dict_rez_itog[item[nk_r_profes]] = 1
            else:
                dict_rez_itog[item[nk_r_profes]] += 1
        if item[nk_r_fio2] == 1:
            deficit_cm = '2'
            rez.append([item[nk_r_pnom], item[nk_r_rasp], item[nk_r_proz], item[nk_r_profes], deficit_cm])
            if item[nk_r_profes] not in dict_rez_itog:
                dict_rez_itog[item[nk_r_profes]] = 1
            else:
                dict_rez_itog[item[nk_r_profes]] += 1
        if item[nk_r_fio3] == 1:
            deficit_cm = '3'
            rez.append([item[nk_r_pnom], item[nk_r_rasp], item[nk_r_proz], item[nk_r_profes], deficit_cm])
            if item[nk_r_profes] not in dict_rez_itog:
                dict_rez_itog[item[nk_r_profes]] = 1
            else:
                dict_rez_itog[item[nk_r_profes]] += 1
        if item[nk_r_fio1] != 1 and item[nk_r_fio1] != 2 and item[nk_r_fio1] != 838:
            if item[nk_r_profes] not in dict_rez_all:
                dict_rez_all[item[nk_r_profes]] = 1
            else:
                dict_rez_all[item[nk_r_profes]] += 1
        if item[nk_r_fio2] != 1 and item[nk_r_fio2] != 2 and item[nk_r_fio2] != 838:
            if item[nk_r_profes] not in dict_rez_all:
                dict_rez_all[item[nk_r_profes]] = 1
            else:
                dict_rez_all[item[nk_r_profes]] += 1
        if item[nk_r_fio3] != 1 and item[nk_r_fio3] != 2 and item[nk_r_fio3] != 838:
            if item[nk_r_profes] not in dict_rez_all:
                dict_rez_all[item[nk_r_profes]] = 1
            else:
                dict_rez_all[item[nk_r_profes]] += 1

    rez.append(["====" for _ in rez[0]])
    rez.append(['Профессия','Количество по РМ','Дефицит','Дефицит,%',''])
    for key in dict_rez_all.keys():
        deficit = 0
        procent = 0
        if key in dict_rez_itog:
            deficit = dict_rez_itog[key]
            procent = round(deficit*100/dict_rez_all[key])
        rez.append([key,dict_rez_all[key],deficit,procent,''])
    CQT.zapoln_wtabl(self, rez, self.ui.tbl_vacant, isp_shapka=True, separ='')
    CMS.zapolnit_filtr(self, self.ui.tbl_vacant_filtr, self.ui.tbl_vacant)
    CSQ.close_bd(conn)


def load_emploee(self):
    conn, cur = CSQ.connect_bd(self.db_users)
    zapros = """SELECT employee.Пномер, employee.ФИО, employee.Должность, employee.Статус
    , "" as Номер_РМ
        FROM employee 
        WHERE employee.ФИО != "" and employee.ФИО != "-" AND Статус != 'Увольнение' 
         """
    spis_empl = CSQ.zapros(self.db_users, zapros, shapka=True, conn=conn)
    nk_e_pnom = F.nom_kol_po_im_v_shap(spis_empl, 'Пномер')
    nk_e_nrm = F.nom_kol_po_im_v_shap(spis_empl,'Номер_РМ')
    zapros = """SELECT rab_mesta.Пномер, rab_mesta.ФИО_1, rab_mesta.ФИО_2, rab_mesta.ФИО_3
        FROM rab_mesta
         """
    spis_rm = CSQ.zapros(self.db_users, zapros, shapka=True, conn=conn)
    nk_r_fio1 = F.nom_kol_po_im_v_shap(spis_rm,'ФИО_1')
    nk_r_fio2 = F.nom_kol_po_im_v_shap(spis_rm, 'ФИО_2')
    nk_r_fio3 = F.nom_kol_po_im_v_shap(spis_rm, 'ФИО_3')
    nk_r_pnom = F.nom_kol_po_im_v_shap(spis_rm,'Пномер')

    for i in range(1,len(spis_empl)):
        for rm in spis_rm:
            if spis_empl[i][nk_e_pnom] == rm[nk_r_fio1]:
                spis_empl[i][nk_e_nrm] += f' {str(rm[nk_r_pnom])};'
            if spis_empl[i][nk_e_pnom] == rm[nk_r_fio2]:
                spis_empl[i][nk_e_nrm] += f' {str(rm[nk_r_pnom])};'
            if spis_empl[i][nk_e_pnom] == rm[nk_r_fio3]:
                spis_empl[i][nk_e_nrm] += f' {str(rm[nk_r_pnom])};'

    CQT.zapoln_wtabl(self, spis_empl, self.ui.tbl_emploee, isp_shapka=True, separ='')
    CMS.zapolnit_filtr(self,self.ui.tbl_emploee_filtr,self.ui.tbl_emploee)
    CSQ.close_bd(conn)

    return
def zagruzka_rc(self):
    conn,cur = CSQ.connect_bd(self.db_users)
    zapros = """SELECT rab_mesta.Пномер, places.adress as Расположение, rab_mesta.Прозвище, rab_c.Имя as РЦ, equipment.Наименование  || ' ' || equipment.Инв_номер as Оборудование,
    professions.имя as Профессия_рм,
     s1.ФИО as ФИО_1см, s1.Должность as Должность_1см, Время_начала_1, Время_конца_1, Нераб_мин1, Между_нар_мин1, Коэфф_производит1,
     s2.ФИО as ФИО_2см, s2.Должность as Должность_2см, Время_начала_2, Время_конца_2, Нераб_мин2, Между_нар_мин2, Коэфф_производит2,
     s3.ФИО as ФИО_3см, s3.Должность as Должность_3см, Время_начала_3, Время_конца_3, Нераб_мин3, Между_нар_мин3, Коэфф_производит3, 
     rab_mesta.Примечание, s1.Пномер as Пномер_emp1, s2.Пномер as Пномер_emp2, s3.Пномер as Пномер_emp3
     FROM rab_mesta
     INNER JOIN rab_c ON rab_c.Код == rab_mesta.Код_РЦ 
     INNER JOIN professions ON professions.код == rab_mesta.Код_профессии
     INNER JOIN equipment ON equipment.Пномер == rab_mesta.Номер_осн_оборуд
     INNER JOIN places ON places.serial == rab_mesta.Расположение
     INNER JOIN employee s1 ON s1.Пномер == rab_mesta.ФИО_1
     INNER JOIN employee s2 ON s2.Пномер == rab_mesta.ФИО_2
     INNER JOIN employee s3 ON s3.Пномер == rab_mesta.ФИО_3"""
    spis = CSQ.zapros(self.db_users,zapros,shapka=True,conn=conn, rez_dict=True)
    spis_prof_fio_status_emploee = CSQ.zapros(self.db_users,"""SELECT Должность, ФИО ,Статус, Пномер FROM employee""",shapka=False,conn=conn, rez_dict=True)
    spis_prof_emploee = sorted(list(set([_['Должность'] for _ in spis_prof_fio_status_emploee])))
    spis_fio_uvol_emploee = [[_['ФИО'], _['Пномер']] for _ in spis_prof_fio_status_emploee if _['Статус']=="Увольнение"]


    for i in range(len(spis)):
        if [spis[i]['ФИО_1см'], spis[i]['Пномер_emp1']] in spis_fio_uvol_emploee:
            spis[i]['ФИО_1см'] += ' УВОЛЕН'
        if [spis[i]['ФИО_2см'], spis[i]['Пномер_emp2']] in spis_fio_uvol_emploee:
            spis[i]['ФИО_2см'] += ' УВОЛЕН'
        if [spis[i]['ФИО_3см'], spis[i]['Пномер_emp3']] in spis_fio_uvol_emploee:
            spis[i]['ФИО_3см'] += ' УВОЛЕН'

    spis_rc = CSQ.zapros(self.db_users, """SELECT Имя FROM rab_c""", shapka=False, conn=conn)
    spis_rc = sorted(list(set([_[0] for _ in spis_rc])))

    spis_oborud = CSQ.zapros(self.db_users, """SELECT Наименование || ' ' || Инв_номер FROM equipment""", shapka=False, conn=conn)
    spis_oborud = sorted(list(set([_[0] for _ in spis_oborud])))

    spis_prof = CSQ.zapros(self.db_users, """SELECT имя FROM professions""", shapka=False, conn=conn)
    spis_prof = sorted(list(set([_[0] for _ in spis_prof])))

    spis_places = CSQ.zapros(self.db_users, """SELECT adress FROM places""", shapka=False, conn=conn)
    spis_places = sorted(list(set([_[0] for _ in spis_places])))

    set_edit = {'Прозвище',
                'Примечание',
                'Время_начала_1',
                'Время_конца_1',
                'Нераб_мин1',
                'Между_нар_мин1',
                'Коэфф_производит1',

                 'Время_начала_2',
                 'Время_конца_2',
                 'Нераб_мин2',
                 'Между_нар_мин2',
                 'Коэфф_производит2',

                 'Время_начала_3',
                 'Время_конца_3',
                 'Нераб_мин3',
                 'Между_нар_мин3',
                 'Коэфф_производит3',
                }

    CQT.zapoln_wtabl(self,spis,self.ui.tbl_rc,isp_shapka=True,separ='',set_editeble_col_nomera=set_edit)
    CQT.cvet_cell_wtabl(self.ui.tbl_rc,'ФИО_1см','УВОЛЕН',r=110)
    CQT.cvet_cell_wtabl(self.ui.tbl_rc, 'ФИО_2см', 'УВОЛЕН', r=110)
    CQT.cvet_cell_wtabl(self.ui.tbl_rc, 'ФИО_3см', 'УВОЛЕН', r=110)
    nk_dolg1 = CQT.nom_kol_po_imen(self.ui.tbl_rc,'Должность_1см')
    nk_dolg2 = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Должность_2см')
    nk_dolg3 = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Должность_3см')
    nk_raspolog = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Расположение')
    nk_rc = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'РЦ')
    nk_oborud = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Оборудование')
    nk_prof = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Профессия_рм')

    for i in range(self.ui.tbl_rc.rowCount()):
        CQT.add_combobox(self, self.ui.tbl_rc, i, nk_dolg1, spis_prof_emploee, False, select_prof_emploee)
        CQT.add_combobox(self, self.ui.tbl_rc, i, nk_dolg2, spis_prof_emploee, False, select_prof_emploee)
        CQT.add_combobox(self, self.ui.tbl_rc, i, nk_dolg3, spis_prof_emploee, False, select_prof_emploee)
        CQT.add_combobox(self, self.ui.tbl_rc, i, nk_raspolog, spis_places, False, select_rasp)
        CQT.add_combobox(self, self.ui.tbl_rc, i, nk_rc, spis_rc, False, select_rc)
        CQT.add_combobox(self, self.ui.tbl_rc, i, nk_oborud, spis_oborud, False, select_oborud)
        CQT.add_combobox(self, self.ui.tbl_rc, i, nk_prof, spis_prof, False, select_prof)
    CMS.zapolnit_filtr(self,self.ui.tbl_rc_filtr,self.ui.tbl_rc)
    CSQ.close_bd(conn)
    self.ui.tbl_rc.setToolTip(f'"|_|" вакант = нет     "-" не нужен = нет     "+" не нужен = есть')

    return

def add_rm(self):
    modifiers = CQT.get_key_modifiers()
    if CMS.user_access(self.bd_naryad,'rab_mesta_edit',F.user_name()) == False:
        return
    rez = CQT.msgboxgYN('Добавить новое рабочее место?')
    if rez != True:
        return
    if modifiers == ['shift']:
        nom_rm = self.ui.tbl_rc.item(self.ui.tbl_rc.currentRow(),CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Пномер')).text()
        conn, cur = CSQ.connect_bd(self.db_users)
        zapros = f"""SELECT rab_mesta.Расположение, rab_mesta.Код_РЦ as РЦ, rab_mesta.Прозвище, rab_mesta.Номер_осн_оборуд as Оборудование,
            rab_mesta.Код_профессии as Профессия_рм,
             Время_начала_1, Время_конца_1, Нераб_мин1, Между_нар_мин1, Коэфф_производит1,
             Время_начала_2, Время_конца_2, Нераб_мин2, Между_нар_мин2, Коэфф_производит2,
             Время_начала_3, Время_конца_3, Нераб_мин3, Между_нар_мин3, Коэфф_производит3, 
             rab_mesta.Примечание
             FROM rab_mesta WHERE rab_mesta.Пномер = {nom_rm}"""
        spis = CSQ.zapros(self.db_users, zapros, shapka=True, conn=conn, rez_dict=True)[0]
        CSQ.close_bd(conn)
        rasp = spis['Расположение']
        kod_rc = spis['РЦ']
        Прозвище = spis['Прозвище']
        Номер_осн_оборуд =  spis['Оборудование']
        Код_профессии =   spis['Профессия_рм']
        Время_начала_1 =   spis['Время_начала_1']
        Время_конца_1 =  spis['Время_конца_1']
        Нераб_мин1 =  spis['Нераб_мин1']
        Между_нар_мин1 =    spis['Между_нар_мин1']
        Коэфф_производит1 =  spis['Коэфф_производит1']
        Время_начала_2 =   spis['Время_начала_2']
        Время_конца_2 =  spis['Время_конца_2']
        Нераб_мин2 =   spis['Нераб_мин2']
        Между_нар_мин2 =   spis['Между_нар_мин2']
        Коэфф_производит2 =  spis['Коэфф_производит2']
        Время_начала_3 =  spis['Время_начала_3']
        Время_конца_3 =    spis['Время_конца_3']
        Нераб_мин3 =  spis['Нераб_мин3']
        Между_нар_мин3 =   spis['Между_нар_мин3']
        Коэфф_производит3 =  spis['Коэфф_производит3']
        Примечание =  spis['Примечание']
        zapros = f"""INSERT INTO rab_mesta
                                      (Расположение, Код_РЦ, Прозвище, Номер_осн_оборуд, Код_профессии,ФИО_1,Время_начала_1,
                                      Время_конца_1,Нераб_мин1,Между_нар_мин1,Коэфф_производит1,ФИО_2,Время_начала_2,
                                      Время_конца_2,Нераб_мин2,Между_нар_мин2,Коэфф_производит2,ФИО_3,Время_начала_3,
                                      Время_конца_3,Нераб_мин3,Между_нар_мин3,Коэфф_производит3,Примечание)
                                      VALUES ({rasp}, "{kod_rc}", "{Прозвище}", {Номер_осн_оборуд}, "{Код_профессии}",
                                      1, "{Время_начала_1}", "{Время_конца_1}",{Нераб_мин1},{Между_нар_мин1},{Коэфф_производит1},
                                      1, "{Время_начала_2}", "{Время_конца_2}",{Нераб_мин2},{Между_нар_мин2},{Коэфф_производит2},
                                      1, "{Время_начала_3}", "{Время_конца_3}",{Нераб_мин3},{Между_нар_мин3},{Коэфф_производит3},
                                      "{Примечание}");"""
    else:
        zapros = """INSERT INTO rab_mesta
                              (Расположение, Код_РЦ, Прозвище, Номер_осн_оборуд, Код_профессии,ФИО_1,Время_начала_1,
                              Время_конца_1,Нераб_мин1,Между_нар_мин1,Коэфф_производит1,ФИО_2,Время_начала_2,
                              Время_конца_2,Нераб_мин2,Между_нар_мин2,Коэфф_производит2,ФИО_3,Время_начала_3,
                              Время_конца_3,Нераб_мин3,Между_нар_мин3,Коэфф_производит3,Примечание)
                              VALUES (0, "010101", "", 1, "10371",
                              1, "07:00","15:30",75,40,1,
                              1, "15:30","23:59",75,40,0.9,
                              1, "00:01","07:00",75,40,0.8,
                              "");"""
    if CSQ.zapros(self.db_users, zapros):
        CQT.msgbox('Успешно')
        zagruzka_rc(self)
    else:
        CQT.msgbox('Ошибка')


def select_prof(self, text,  row, col):
    if CMS.user_access(self.bd_naryad,'rab_mesta_edit',F.user_name()) == False:
        return
    nk_pnom = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Пномер')
    pnom = int(self.ui.tbl_rc.item(row, nk_pnom).text())
    conn, cur = CSQ.connect_bd(self.db_users)
    rez = CSQ.zapros(self.db_users,f"""SELECT код FROM professions WHERE имя == '{text}'""",conn=conn)
    kod_prof = rez[-1][0]
    CSQ.zapros(self.db_users, f"""UPDATE rab_mesta SET Код_профессии = "{kod_prof}" WHERE Пномер = {pnom}""", shapka=False)
    CSQ.close_bd(conn)

def select_oborud(self, text,  row, col):
    if CMS.user_access(self.bd_naryad,'rab_mesta_edit',F.user_name()) == False:
        return
    nk_pnom = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Пномер')
    pnom = int(self.ui.tbl_rc.item(row, nk_pnom).text())
    conn, cur = CSQ.connect_bd(self.db_users)
    rez = CSQ.zapros(self.db_users,f"""SELECT Пномер FROM equipment WHERE Наименование || ' ' || Инв_номер == '{text}'""",conn=conn)
    kod_oborud = rez[-1][0]
    CSQ.zapros(self.db_users, f"""UPDATE rab_mesta SET Номер_осн_оборуд = "{kod_oborud}" WHERE Пномер = {pnom}""", shapka=False)
    CSQ.close_bd(conn)

def select_rc(self, text,  row, col):
    if CMS.user_access(self.bd_naryad,'rab_mesta_edit',F.user_name()) == False:
        return
    nk_pnom = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Пномер')
    pnom = int(self.ui.tbl_rc.item(row, nk_pnom).text())
    conn, cur = CSQ.connect_bd(self.db_users)
    rez = CSQ.zapros(self.db_users,f"""SELECT Код FROM rab_c WHERE Имя == '{text}'""",conn=conn)
    kod_rc = rez[-1][0]
    CSQ.zapros(self.db_users, f"""UPDATE rab_mesta SET Код_РЦ = "{kod_rc}" WHERE Пномер = {pnom}""", shapka=False)
    CSQ.close_bd(conn)

def select_rasp(self, text,  row, col):
    if CMS.user_access(self.bd_naryad,'rab_mesta_edit',F.user_name()) == False:
        return
    nk_pnom = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Пномер')
    pnom = int(self.ui.tbl_rc.item(row, nk_pnom).text())
    conn, cur = CSQ.connect_bd(self.db_users)
    rez = CSQ.zapros(self.db_users, f"""SELECT serial FROM places WHERE adress == '{text}'""", conn=conn)
    kod_place = rez[-1][0]
    CSQ.zapros(self.db_users, f"""UPDATE rab_mesta SET Расположение = {kod_place} WHERE Пномер = {pnom}""", shapka=False)
    CSQ.close_bd(conn)


def clck_tbl_rc(self):
    #self.ui.lbl_info_rc.setText('')
    CQT.statusbar_text(self)
    tbl = self.ui.tbl_rc
    r = tbl.currentRow()
    col = tbl.currentColumn()
    if r == -1 or col == -1:
        return
    nk_fio = False
    if tbl.horizontalHeaderItem(col).text() == 'ФИО_1см':
        nk_fio = CQT.nom_kol_po_imen(tbl,'ФИО_1см')
    if tbl.horizontalHeaderItem(col).text() == 'ФИО_2см':
        nk_fio = CQT.nom_kol_po_imen(tbl,'ФИО_2см')
    if tbl.horizontalHeaderItem(col).text() == 'ФИО_3см':
        nk_fio = CQT.nom_kol_po_imen(tbl,'ФИО_3см')
    if nk_fio:
        fio = tbl.item(r,col).text()
        for id in self.DICT_EMPLOEE_FULL.keys():
            if fio in self.DICT_EMPLOEE_FULL:
                strok = str(self.DICT_EMPLOEE_FULL[fio])
                #self.ui.lbl_info_rc.setText(strok)
                CQT.statusbar_text(self,strok)

def select_prof_emploee(self, text, row, col):
    if CMS.user_access(self.bd_naryad,'rab_mesta_edit',F.user_name()) == False:
        return
    nk_fio = None
    if self.ui.tbl_rc.horizontalHeaderItem(col).text() == 'Должность_1см':
        nk_fio = CQT.nom_kol_po_imen(self.ui.tbl_rc,'ФИО_1см')
    if self.ui.tbl_rc.horizontalHeaderItem(col).text() == 'Должность_2см':
        nk_fio = CQT.nom_kol_po_imen(self.ui.tbl_rc,'ФИО_2см')
    if self.ui.tbl_rc.horizontalHeaderItem(col).text() == 'Должность_3см':
        nk_fio = CQT.nom_kol_po_imen(self.ui.tbl_rc,'ФИО_3см')
    if nk_fio == None:
        return
    #spis_fio = CSQ.zapros(self.db_users, f"""SELECT ФИО || "$" || Должность || "$" || Режим  FROM employee WHERE Должность == '{text}' AND Статус != 'Увольнение' """, shapka=False)
    #spis_fio = [_[0] for _ in spis_fio]
    spis_fio = []
    for fio in self.DICT_EMPLOEE_FULL.keys():
        if self.DICT_EMPLOEE_FULL[fio]['Должность'] == text:
            spis_fio.append('$'.join([fio,self.DICT_EMPLOEE_FULL[fio]['Подразделение'],self.DICT_EMPLOEE_FULL[fio]['Режим']]))
    spis_fio = sorted(spis_fio)
    CQT.add_combobox(self, self.ui.tbl_rc, row, nk_fio, spis_fio, True, select_fio)

def select_fio(self, text,  row, col):
    if CMS.user_access(self.bd_naryad,'rab_mesta_edit',F.user_name()) == False:
        return
    nk_fio = None
    if self.ui.tbl_rc.horizontalHeaderItem(col).text() == 'ФИО_1см':
        nk_fio = "ФИО_1"
    if self.ui.tbl_rc.horizontalHeaderItem(col).text() == 'ФИО_2см':
        nk_fio = "ФИО_2"
    if self.ui.tbl_rc.horizontalHeaderItem(col).text() == 'ФИО_3см':
        nk_fio = "ФИО_3"
    if nk_fio == None:
        return
    nk_pnom = CQT.nom_kol_po_imen(self.ui.tbl_rc, 'Пномер')
    pnom = int(self.ui.tbl_rc.item(row,nk_pnom).text())
    conn,cur = CSQ.connect_bd(self.db_users)
    #rez = CSQ.zapros(self.db_users,f"""SELECT Пномер FROM employee WHERE ФИО == '{text}' AND Статус != 'Увольнение' """,conn=conn)
    pnom_emploe = 2
    if len(text.split("$")) == 3:
        sel_fio, sel_podr, sel_rej = text.split("$")
        pnom_emploe = self.DICT_EMPLOEE_FULL[sel_fio]['Пномер']
    CSQ.zapros(self.db_users, f"""UPDATE rab_mesta SET {nk_fio} = {pnom_emploe} WHERE Пномер = {pnom}""", shapka=False)
    CSQ.close_bd(conn)


def select_schema_dbl_clk(self):
    r = self.ui.tbl_rc.currentRow()
    column = self.ui.tbl_rc.currentColumn()
    if CQT.nom_kol_po_imen(self.ui.tbl_rc,'Пномер') == column:
        nk_rasp = CQT.nom_kol_po_imen(self.ui.tbl_rc,'Расположение')
        CMS.dict_rab_mesta(self, self.db_users)
        img = self.DICT_PLACES[self.ui.tbl_rc.item(r,nk_rasp).text()]['img_name']
        path = F.scfg('mk_data') + F.sep() + 'schems' + F.sep()
        coords = self.DICT_RM[int(self.ui.tbl_rc.item(r,column).text())]['coord']
        x = y = ''
        if coords != None and coords != '':
            x , y = coords.split(";")
            x = int(x)
            y = int(y)
        load_schema_jpg(self, path + img,x,y)
        self.ui.tabWidget_6.setCurrentIndex(CQT.nom_tab_po_imen(self.ui.tabWidget_6,'Схема'))

def select_schema(self):
    path = F.scfg('mk_data') + F.sep() + 'schems' + F.sep()
    if F.nalich_file(path):
        list_files = F.spis_files(path)
        for file in list_files[0][2]:
            if F.ostavit_rasshir(file) == '.jpg':
                if F.ubrat_rasshir(file) == self.ui.cmb_schems.currentText():
                    load_schema_jpg(self,list_files[0][0] + file)
                    return
        CQT.msgbox(f'Не найден файл {self.ui.cmb_schems.text()}')
        return
    CQT.msgbox(f'Не найден путь {path}')
    return

def load_schema_jpg(self, path, x_sel = '', y_sel = ''):
    lbl = self.ui.lbl_shema

    fon = QPixmap(path)
    #koef = fon.height()/lbl.height()
    #lbl.setFixedHeight(fon.width()/koef)
    old_w = fon.width()
    old_h = fon.height()
    #lbl.setFixedWidth(fon.height()/koef)
    lbl.clear()
    self.ui.lbl_shema.setFixedSize(self.SIZE_SCHEMA_LBL)
    pixmap = fon.scaled(lbl.size(), Qt.KeepAspectRatio)
    wind_k_width = old_w / pixmap.width()
    wind_k_height = old_h / pixmap.height()
    if x_sel != '':
        path_pic = F.scfg('mk_data') + F.sep() + 'schems' + F.sep() + "Shape_1.png"
        #painter = QPainter(self)
        pic = QPixmap(path_pic)
        painter = QPainter(pixmap)
        pic_h = pic.height()/wind_k_height
        pic_w = pic.width()/wind_k_width
        painter.drawPixmap(int(round(x_sel-pic_w/2)), int(round(y_sel-pic_h/2)), int(pic_w), int(pic_h), pic)
        painter.end()
    self.ui.lbl_shema.setPixmap(pixmap)