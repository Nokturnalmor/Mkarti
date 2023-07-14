from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QStyle
from form_1 import Ui_Form_first  # импорт нашего сгенерированного файла
import project_cust_38.Cust_Qt as CQT
import project_cust_38.Cust_mes as CMS
import project_cust_38.Cust_Functions as F
from copy import deepcopy
import project_cust_38.operacii as operacii

class mywindow2(QtWidgets.QDialog):  # диалоговое окно
    def __init__(self, parent):
        self.myparent = parent
        super(mywindow2, self).__init__()
        self.ui2 = Ui_Form_first()
        self.ui2.setupUi(self)
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setWindowModality(QtCore.Qt.ApplicationModal)
        self.setWindowTitle("Заголовок")
        self.ui2.tbtn_max.clicked.connect(self.full_screen)
        self.ui2.tbtn_min.clicked.connect(self.minized)
        self.ui2.tbtn_exit.clicked.connect(self.exit_form)
        self.ui2.btn_calc.clicked.connect(self.calc_var_from_gvar)
        self.ui2.btn_ok.clicked.connect(self.create_mk)
        self.ui2.tbl_var_tk.clicked.connect(self.take_name_column_tk)
        self.ui2.tbl_var_tk_mat.clicked.connect(self.take_name_column_tk)
        self.ui2.tbl_var_vo.clicked.connect(self.take_name_column_vo)
        self.ui2.btn_reload_glob_vars.clicked.connect(self.load_glob_vars)
        self.ui2.btn_load_csv_gl_var.clicked.connect(self.load_glob_vars_csv)
        self.load_parametrs_vo()
        self.app_icons()
        self.showMaximized()
        self.dragPos = QtCore.QPoint()
        CQT.load_css(self)
        CQT.load_icons(self,24)

    # вызывается при нажатии кнопки мыши по форме
    @CQT.onerror
    def mousePressEvent(self, event):
        # Если нажата левая кнопка мыши
        if event.button() == QtCore.Qt.LeftButton:
            # получаем координаты окна относительно экрана
            x_main = self.geometry().x()
            y_main = self.geometry().y()
            # получаем координаты курсора относительно окна нашей программы
            cursor_x = QtGui.QCursor.pos().x()
            cursor_y = QtGui.QCursor.pos().y()
            # проверяем условием позицию курсора на нужной области программы(у нас это верхний бар)
            # если всё ок - перемещаем
            # иначе игнорируем
            if x_main <= cursor_x <= x_main + self.geometry().width():
                if y_main <= cursor_y <= y_main + self.ui2.line.geometry().y():
                    self.old_pos = event.pos()
                else:
                    self.old_pos = None
        elif event.button() == QtCore.Qt.RightButton:
            self.old_pos = None

    # вызывается при отпускании кнопки мыши
    @CQT.onerror
    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.old_pos = None

    # вызывается всякий раз, когда мышь перемещается
    @CQT.onerror
    def mouseMoveEvent(self, event):
        if not self.old_pos:
            return
        delta = event.pos() - self.old_pos
        self.move(self.pos() + delta)


    def take_name_column_tk(self):
        modifiers = CQT.get_key_modifiers(self)
        if modifiers == ['alt']:
            if self.ui2.tbl_var_tk.hasFocus():
                pole = self.ui2.tbl_var_tk.horizontalHeaderItem(self.ui2.tbl_var_tk.currentColumn()).text()
                kod = self.ui2.tbl_var_tk.item(self.ui2.tbl_var_tk.currentRow(),CQT.nom_kol_po_imen(self.ui2.tbl_var_tk,'Код')).text()
                F.copy_bufer(f'{{tblv("{pole}","{kod}")}}')
            if self.ui2.tbl_var_tk_mat.hasFocus():
                kod = self.ui2.tbl_var_tk_mat.item(self.ui2.tbl_var_tk_mat.currentRow(),CQT.nom_kol_po_imen(self.ui2.tbl_var_tk_mat,'Код')).text()
                F.copy_bufer(f'{{tblm("{kod}")}}')
        if self.ui2.tbl_var_tk.hasFocus():
            self.ui2.lbl_glob_var.setText(str(self.glob_dict_from_tk[CQT.znach_vib_strok_po_kol(self.ui2.tbl_var_tk,'ДСЕ')]))
        if self.ui2.tbl_var_tk_mat.hasFocus():
            self.ui2.lbl_glob_var.setText(str(self.glob_dict_from_tk[CQT.znach_vib_strok_po_kol(self.ui2.tbl_var_tk_mat,'ДСЕ')]))


    def take_name_column_vo(self):
        modifiers = CQT.get_key_modifiers(self)
        if modifiers == ['alt']:
            F.copy_bufer(self.ui2.tbl_var_vo.horizontalHeaderItem(self.ui2.tbl_var_vo.currentColumn()).text())

    def create_mk(self):
        list_vars_vo =CQT.spisok_iz_wtabl(self.ui2.tbl_rez_tk, shapka=True)
        list_mat_vo = CQT.spisok_iz_wtabl(self.ui2.tbl_rez_tk_mat, shapka=True)
        dict_mat_vo = F.list_to_dict(list_mat_vo)
        if list_vars_vo == [[]] or list_mat_vo == [[]]:
            return
        for row in list_vars_vo:
            for item in row:
                if item == 'ERR':
                    CQT.msgbox(f'Ошибки в результатах')
                    return
        dict_mat = dict()
        for item in dict_mat_vo:
            if item['Норма'] == 0:
                continue
            key = f"{item['ДСЕ']}${item['Код']}"
            val ='$'.join([item['НН'],item['Материал'],item['Ед.Изм'],'{:.8f}'.format(round(F.valm(item['Норма']), 8))])
            if key in dict_mat:
                dict_mat[key].append(val)
            else:
                dict_mat[key] = [val]
        for key in dict_mat.keys():
            dict_mat[key] = '{'.join(dict_mat[key])
        self.myparent.list_vars_vo = list_vars_vo
        self.myparent.dict_mat_vo = dict_mat
        self.hide()

    def check_insert_gvars(self,gvars):
        for item in list(gvars.values()):
            if item == '':
                return False
            if F.is_numeric(item)== False:
                return False
        return True

    def tblv(self,var_name, kod = ''):
        if kod == '':
            kod = self.current_calc_kod
        list = CQT.spisok_iz_wtabl(self.ui2.tbl_var_tk,shapka=True,rez_dict=True)
        for item in list:
            if item['Код'] == kod:
                if var_name in item:
                    tmp = item[var_name]
                    tmp = str(self.cust_eval(tmp, 4))
                    tmp = self.apply_vars(tmp, self.dict_form_tk, self.list_gvars)
                    return self.cust_eval(tmp,4)
        return

    def tblm(self,kod):
        list = CQT.spisok_iz_wtabl(self.ui2.tbl_var_tk_mat,shapka=True,rez_dict=True)
        for item in list:
            if item['Код'] == kod:
                tmp = item['Норма']
                tmp = str(self.cust_eval(tmp, 4))
                tmp = self.apply_vars(tmp, self.dict_form_tk, self.list_gvars)
                return self.cust_eval(tmp,4)
        return

    def arr(self,rval,cval,arr):
        try:
            arr = eval(arr)
        except:
            pass
        if type(arr[1][0]) == type(1):
            row = len(arr)-1
            for i in range(1,len(arr)):
                if arr[i][0] >= rval:
                    row = i
                    break
        else:
            row = len(arr) - 1
            for i in range(1,len(arr)):
                if arr[i][0] == rval:
                    row = i
                    break
        if type(arr[0][1]) == type(1):
            col = len(arr[0])-1
            for j in range(1,len(arr[i])):
                if arr[0][j] >= cval:
                    col = j
                    break
        else:
            col = len(arr[0]) - 1
            for j in range(1,len(arr[i])):
                if arr[0][j] == cval:
                    col = j
                    break
        return arr[row][col]



    def calc_var_time(self):
        list_tk = CQT.spisok_iz_wtabl(self.ui2.tbl_var_tk, shapka=True)
        list_tk_old = deepcopy(list_tk)
        list_tk = self.claculate_cells(list_tk, self.list_gvars)
        set_edit = {_ for _ in range(len(list_tk[0])) if _ not in (0, 1, 2, 3)}
        CQT.fill_wtabl(list_tk, self.ui2.tbl_rez_tk, set_editeble_col_nomera=set_edit)
        for i in range(1, len(list_tk)):
            for j in range(4, len(list_tk[i])):
                if list_tk[i][j] == '_':
                    CQT.ust_color_wtab(self.ui2.tbl_rez_tk, i - 1, j, 245, 245, 245)
                else:
                    if "ERR" == list_tk[i][j] or list_tk[i][j] == '0':
                        CQT.ust_color_wtab(self.ui2.tbl_rez_tk, i - 1, j, 253, 200, 200)
                    else:
                        if type(list_tk_old[i][j]) == str and "f'" in list_tk_old[i][j]:
                            CQT.ust_color_wtab(self.ui2.tbl_rez_tk, i - 1, j, 200, 250, 200)



    def calc_var_mat(self):
        list_oper = CQT.spisok_iz_wtabl(self.ui2.tbl_var_tk_mat, shapka=True)
        list_oper_old = deepcopy(list_oper)
        list_tk = self.claculate_cells(list_oper, self.list_gvars)
        set_edit = {_ for _ in range(len(list_tk[0])) if _ not in (0, 1, 2, 3,4,5)}
        CQT.fill_wtabl(list_tk, self.ui2.tbl_rez_tk_mat, set_editeble_col_nomera=set_edit)
        for i in range(1, len(list_tk)):
            for j in range(4, len(list_tk[i])):
                if list_tk[i][j] == '_':
                    CQT.ust_color_wtab(self.ui2.tbl_rez_tk_mat, i - 1, j, 245, 245, 245)
                else:
                    if "ERR" == list_tk[i][j] or list_tk[i][j] == '0':
                        CQT.ust_color_wtab(self.ui2.tbl_rez_tk_mat, i - 1, j, 253, 200, 200)
                    else:
                        if type(list_oper_old[i][j]) == str and "f'" in list_oper_old[i][j]:
                            CQT.ust_color_wtab(self.ui2.tbl_rez_tk_mat, i - 1, j, 200, 250, 200)


    def calc_var_from_gvar(self):
        self.list_gvars = CQT.spisok_iz_wtabl(self.ui2.tbl_var_vo, shapka=True, rez_dict=True)[0]
        if self.check_insert_gvars(self.list_gvars) == False:
            CQT.msgbox(f'Не корректно внесены глобальные переменные')
            return
        self.calc_var_time()
        self.calc_var_mat()
        self.claculate_struct()
        self.claculate_mass()

    def calc_dse_ves(self) -> dict:
        list_dse_mat = CQT.spisok_iz_wtabl(self.ui2.tbl_rez_tk_mat, shapka=True,rez_dict=True)
        dict_dse = dict()
        for item in list_dse_mat:
            try:
                if item['ДСЕ'] not in dict_dse:
                    dict_dse[item['ДСЕ']] = F.valm(item['Норма'])
                else:
                    dict_dse[item['ДСЕ']] += F.valm(item['Норма'])
            except:
                CQT.msgbox(f"ОШибка суммирования массы в {item['ДСЕ']} {item['Код']} не учтено!")

        return dict_dse

    def claculate_mass(self):
        dict_dse = self.calc_dse_ves()
        list_struct = CQT.spisok_iz_wtabl(self.myparent.ui.table_razr_MK, shapka=True)
        nk_mass = F.nom_kol_po_im_v_shap(list_struct, 'Масса/М1,М2,М3')
        nk_naim = F.nom_kol_po_im_v_shap(list_struct, 'Обозначение')
        for i in range(1, len(list_struct)):
            arr_mass = list_struct[i][nk_mass].split('/М1')
            dse = list_struct[i][nk_naim]
            try:
                if dse in dict_dse:
                    arr_mass[0] = str(round(dict_dse[dse],3))
                else:
                    arr_mass[0] = ''
            except:
                CQT.msgbox(f'Ошибка {dse} не рассчитана')
            list_struct[i][nk_mass] = '/М1'.join(arr_mass)
        CQT.fill_wtabl(list_struct, self.myparent.ui.table_razr_MK,
                       set_editeble_col_nomera=self.myparent.edit_cr_mk_ruch, auto_type=False)

    def claculate_struct(self):
        list_struct = CQT.spisok_iz_wtabl(self.myparent.ui.table_razr_MK,shapka=True)
        nk_kolich = F.nom_kol_po_im_v_shap(list_struct,'Количество')
        for i in range(1,len(list_struct)):
            self.dict_form_tk = self.use_global_var_form_tk(list_struct[i][1])
            if "f'" in str(list_struct[i][nk_kolich]):
                try:
                    list_struct[i][nk_kolich] = self.calc_funcs(list_struct[i][nk_kolich])
                    list_struct[i][nk_kolich] = self.apply_vars(list_struct[i][nk_kolich], self.dict_form_tk, self.list_gvars)
                    list_struct[i][nk_kolich] = self.cust_eval(list_struct[i][nk_kolich],4)
                except:
                    pass
        CQT.fill_wtabl(list_struct, self.myparent.ui.table_razr_MK, set_editeble_col_nomera=self.myparent.edit_cr_mk_ruch,auto_type=False)

    def calc_funcs(self,func:str):
        try:
            func = func.replace("arr(", "self.arr(")
            func = func.replace("tblv(", "self.tblv(")
            func = func.replace("tblm(", "self.tblm(")
            return func
        except:
            return func

    def claculate_cells(self,list_tk,list_gvars):
        for i in range(1,len(list_tk)):
            self.dict_form_tk = self.use_global_var_form_tk(list_tk[i][0])
            for j in range(4, len(list_tk[i])):
                if list_tk[i][j] == '_':
                    continue
                else:
                    list_tk[i][j] = self.calc_funcs(list_tk[i][j])
                if type(list_tk[i][j]) == str and ("f'" in list_tk[i][j] or 'tbl' in list_tk[i][j] or 'arr' in list_tk[i][j]):
                    list_tk[i][j] = self.apply_vars(list_tk[i][j], self.dict_form_tk, list_gvars)
                    try:
                        self.current_calc_nn = list_tk[i][0]
                        self.current_calc_nn = list_tk[i][3]
                        if "tbl" in str(list_tk[i][j]) or "arr" in str(list_tk[i][j]):
                            list_tk[i][j] = str(self.cust_eval(list_tk[i][j],44))
                            list_tk[i][j] = self.apply_vars(list_tk[i][j], self.dict_form_tk, list_gvars)
                        list_tk[i][j] = self.cust_eval(list_tk[i][j],4)
                    except:
                        list_tk[i][j] = 'ERR'
                else:
                    try:
                        list_tk[i][j] = self.cust_eval(list_tk[i][j], 4)
                    except:
                        list_tk[i][j] = 'ERR'
                try:
                    if F.is_numeric(list_tk[i][j]):
                        list_tk[i][j] = str(round(F.valm(list_tk[i][j]),4))
                except:
                    pass
        return list_tk

    def cust_eval(self,item,count):
        for _ in range(count):
            try:
                item = str(eval(item))
            except:
                pass
        return item

    def apply_glob_tk(self,item, dict_form_tk):
        for key in dict_form_tk.keys():
            item = item.replace(key, str(dict_form_tk[key]))
        return item

    def apply_list_gvars(self,item, list_gvars):
        for key in list_gvars.keys():
            item = item.replace(key, str(list_gvars[key]))
        return item

    def apply_vars(self,item:str, dict_form_tk, list_gvars):
        item = self.apply_glob_tk(item, dict_form_tk)
        item = self.apply_list_gvars(item, list_gvars)
        return item


    def use_global_var_form_tk(self,nn):
        if nn not in self.glob_dict_from_tk:
            return dict()
        if '{' in self.glob_dict_from_tk[nn] and '}' in self.glob_dict_from_tk[nn]:
            try:
                tmp_dict = eval(self.glob_dict_from_tk[nn])
            except:
                return dict()
            if type(tmp_dict) == type(dict()):
                return tmp_dict
            else:
                return dict()
        else:
            return dict()


    def exit_form(self):
        self.close()

    def full_screen(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()


    def minized(self):
        if self.isMinimized():
            self.showNormal()
        else:
            self.showMinimized()

    def vars_oper_from_db(self,dict_of_dse):
        set_var = set()
        for dse in dict_of_dse.keys():
            for oper in dict_of_dse[dse]:
                if oper[2] == '1':
                    if oper[0] in self.myparent.DICT_VAR_OPER:
                        list_var = self.myparent.DICT_VAR_OPER[oper[0]][0]['Vars'].split(';')
                        for var in list_var:
                            set_var.add(var)
                    else:
                        CQT.msgbox(f'{oper[0]} не найден в списке операций')
                else:
                    if oper[2] in self.myparent.DICT_VAR_OPER:
                        dict_pereh = self.myparent.DICT_VAR_OPER[oper[2]][1]
                        if oper[0] in dict_pereh:
                            list_pereh_var = dict_pereh[oper[0]]
                            for var in list_pereh_var:
                                set_var.add(var)

                    else:
                        CQT.msgbox(f'{oper[0]} не найден в списке операций')
        return set_var


    def load_mater(self):
        dict_of_dse = self.calc_list_of_operations()
        list_rez = [['ДСЕ', 'Операция', 'Код', 'НН','Материал','Ед.Изм','Норма']]

        for dse in dict_of_dse.keys():
            for oper in dict_of_dse[dse]:
                if oper[2] == '1' and oper[4] != [['']]:
                    for m in range(len(oper[4])):
                        tmp = []
                        tmp.append(dse)
                        tmp.append(oper[0])
                        tmp.append(oper[3])
                        tmp.append(oper[4][m][0])
                        tmp.append(oper[4][m][1])
                        tmp.append(oper[4][m][2])
                        tmp.append(oper[4][m][3])


                        list_rez.append(tmp)
        set_edit = {_ for _ in range(len(list_rez[0])) if _ not in (0, 1, 2, 3, 4, 5)}
        CQT.fill_wtabl(list_rez, self.ui2.tbl_var_tk_mat, set_editeble_col_nomera=set_edit, height_row=18)
        for i in range(1, len(list_rez)):
            if "f'" in list_rez[i][6]:
                CQT.ust_color_wtab(self.ui2.tbl_var_tk_mat, i - 1, 6, 253, 200, 200)

    def load_time(self):
        dict_of_dse = self.calc_list_of_operations()
        set_var = self.vars_oper_from_db(dict_of_dse)
        list_rez = [['ДСЕ', 'Операция', 'Переход', 'Код']]
        for var in sorted(list(set_var)):
            list_rez[0].append(var)

        for dse in dict_of_dse.keys():
            for oper in dict_of_dse[dse]:
                list_var = oper[1].split('$')
                list_var_name = ['']
                tmp = ['_' for _ in list_rez[0]]
                tmp[0] = dse
                tmp[3] = oper[3]
                if oper[2] == '1':
                    tmp[1] = oper[0]
                    tmp[2] = ''
                    if oper[0] in self.myparent.DICT_VAR_OPER:
                        list_var_name = self.myparent.DICT_VAR_OPER[oper[0]][0]['Vars'].split(';')
                else:
                    tmp[1] = oper[2]
                    tmp[2] = oper[0]
                    if oper[2] in self.myparent.DICT_VAR_OPER:
                        if oper[0] in self.myparent.DICT_VAR_OPER[oper[2]][1]:
                            list_var_name = self.myparent.DICT_VAR_OPER[oper[2]][1][oper[0]]


                if len(list_var_name) != len(list_var):
                    CQT.msgbox(f'Не совпадают переменные в {dse} {oper} не учтена!')
                else:
                    for i in range(len(list_var_name)):
                        if list_var_name[i] != '':
                            nk = F.nom_kol_po_im_v_shap(list_rez, list_var_name[i])
                            if nk == None:
                                print(f' {list_var_name[i]} не найден в {tmp}')
                            tmp[nk] = list_var[i]
                list_rez.append(tmp)
        set_edit = {_ for _ in range(len(list_rez[0])) if _ not in (0, 1, 2, 3)}
        CQT.fill_wtabl(list_rez, self.ui2.tbl_var_tk, set_editeble_col_nomera=set_edit, height_row=18)
        for i in range(1, len(list_rez)):
            for j in range(4, len(list_rez[i])):
                if list_rez[i][j] == '_':
                    CQT.ust_color_wtab(self.ui2.tbl_var_tk, i - 1, j, 245, 245, 245)
                if "f'" in list_rez[i][j]:
                    CQT.ust_color_wtab(self.ui2.tbl_var_tk, i - 1, j, 253, 200, 200)


    def load_parametrs_vo(self):
        self.set_glob_var = set()
        self.load_mater()
        self.load_time()
        self.load_glob_vars()

    def load_glob_vars_csv(self):
        try:
            dict_csv = self.load_csv()
            list_vars = CQT.spisok_iz_wtabl(self.ui2.tbl_var_vo,sep='',shapka=True,rez_dict=True)[0]
            for key in dict_csv.keys():
                if key in list_vars:
                    list_vars[key] = dict_csv[key]
            list_vars = F.dict_to_list(list_vars,transponir=True)
            CQT.fill_wtabl(list_vars,self.ui2.tbl_var_vo,auto_type=False)
        except:
            CQT.msgbox(f'Что то пошло не так')
            return

    def load_csv(self):
        defolt_path = 'O:\Служба главного конструктора\Временная\CSV'
        if not F.nalich_file(defolt_path):
            defolt_path = F.put_po_umolch()
        file = CQT.f_dialog_name(self,'Выбор',defolt_path,'*.csv;*.txt')
        if file == ['.']:
            return
        list = F.load_file(file)
        return dict(list)

    def load_glob_vars(self):
        self.set_glob_var = set()
        list_tk = CQT.spisok_iz_wtabl(self.ui2.tbl_var_tk,shapka=True,rez_dict=True)
        list_tk_mat = CQT.spisok_iz_wtabl(self.ui2.tbl_var_tk_mat, shapka=True, rez_dict=True)

        for i in range(len(list_tk)):
            for key in list_tk[i].keys():
                if key not in ('ДСЕ', 'Операция', 'Переход', 'Код',''):
                    if "f'" in str(list_tk[i][key]):
                        list_glob_var = self.catch_glob_var(list_tk[i][key])
                        for _ in list_glob_var:
                            self.set_glob_var.add(_)
        for i in range(len(list_tk_mat)):
            for key in list_tk_mat[i].keys():
                if key in ('Норма'):
                    if "f'" in str(list_tk_mat[i][key]):
                        list_glob_var = self.catch_glob_var(list_tk_mat[i][key])
                        for _ in list_glob_var:
                            self.set_glob_var.add(_)
        rez_list_glob_var = [list(self.set_glob_var), ['' for _ in list(self.set_glob_var)]]
        CQT.fill_wtabl(rez_list_glob_var, self.ui2.tbl_var_vo)

    def catch_glob_var(self,text:str):
        list_var = []
        list_1 = text.split('{')
        for item in list_1[1:]:
            var = item.split('}')[0]
            if var[:3] != 'tbl' and var[:3] != 'arr':
                list_var.append(var)
        return list_var

    def calc_list_of_operations(self):
        DICT_NN_NTK = CMS.load_dict_dse(self.myparent.db_dse)
        list_of_operations = dict()
        list_pre_mk = CQT.spisok_iz_wtabl(self.myparent.ui.table_razr_MK ,'',True,False,True)
        self.glob_dict_from_tk = dict()
        for tk in list_pre_mk:
            operations = self.operations_tk(tk['Обозначение'],DICT_NN_NTK)
            list_of_operations[tk['Обозначение']] = operations
        return  list_of_operations


    def operations_tk(self, nn,DICT_NN_NTK):
        nom_tk = DICT_NN_NTK[nn]['Номер_техкарты']
        put_name_tk = F.scfg('add_docs') + F.sep() + nom_tk + "_" + nn
        tk = F.otkr_f(put_name_tk + '.txt', False, "|", True, True)
        if tk == ['']:
            CQT.msgbox(f'Не найдена техкарта {put_name_tk}')
            return
        list_oper = []
        fl = False
        for item in tk:
            if len(item) == 21:
                if fl and item[20] == '0':
                    break
                if item[20] == '0':
                    fl = True
                    tmp_oper_name = ''
                    self.glob_dict_from_tk[nn] = item[7]
                    if '[' in self.glob_dict_from_tk[nn] and ']' in self.glob_dict_from_tk[nn]:
                        self.glob_dict_from_tk[nn] = self.glob_dict_from_tk[nn].replace("[", '{')
                        self.glob_dict_from_tk[nn] = self.glob_dict_from_tk[nn].replace("]", '}')
                    else:
                        self.glob_dict_from_tk[nn] = ''
                if fl == True and item[20] == '1':
                    mats = [_.split('$') for _ in [m for m in item[10].split('{')]]
                    for i in range(len(mats)):
                        if len(mats[i]) >=4:
                            mats[i][3] = mats[i][3].replace("£", "{")
                            mats[i][3] = mats[i][3].replace("¢","}")
                    list_oper.append([item[0],item[14],item[20],item[3],mats])
                    tmp_oper_name = item[0]
                if fl == True and item[20] == '2':
                    list_oper.append([item[0],item[14],tmp_oper_name,item[3],[]])
        return  list_oper

    def app_icons(self):
        # from PyQt5.QtGui import QIcon
        # from PyQt5.QtWidgets import QApplication, QStyle
        self.ui2.tbtn_exit.setIcon(QtGui.QIcon(QApplication.style().standardIcon(QStyle.SP_TitleBarCloseButton)))
        self.ui2.tbtn_exit.setIconSize(QtCore.QSize(8, 8))
        self.ui2.tbtn_min.setIcon(QtGui.QIcon(QApplication.style().standardIcon(QStyle.SP_TitleBarMinButton)))
        self.ui2.tbtn_min.setIconSize(QtCore.QSize(8, 8))
        self.ui2.tbtn_max.setIcon(QtGui.QIcon(QApplication.style().standardIcon(QStyle.SP_TitleBarMaxButton)))
        self.ui2.tbtn_max.setIconSize(QtCore.QSize(8, 8))
        self.ui2.btn_calc.setIcon(QtGui.QIcon(QApplication.style().standardIcon(QStyle.SP_DialogOkButton)))
        self.ui2.btn_calc.setIconSize(QtCore.QSize(32, 32))
        self.ui2.btn_ok.setIcon(QtGui.QIcon(QApplication.style().standardIcon(QStyle.SP_DialogApplyButton)))
        self.ui2.btn_ok.setIconSize(QtCore.QSize(32, 32))
        self.ui2.btn_reload_glob_vars.setIcon(QtGui.QIcon(QApplication.style().standardIcon(QStyle.SP_BrowserReload)))
        self.ui2.btn_reload_glob_vars.setIconSize(QtCore.QSize(32, 32))
        self.ui2.btn_load_csv_gl_var.setIcon(QtGui.QIcon(QApplication.style().standardIcon(QStyle.SP_DriveHDIcon)))
        self.ui2.btn_load_csv_gl_var.setIconSize(QtCore.QSize(32, 32))


def update_parametrs(self,spis_tk:list,j:int,nn:str):
    nk_mat_tk = 10
    nk_op_tst = 7
    list_vars_vo = deepcopy(self.list_vars_vo)
    for i in range(1,len(list_vars_vo)):
        list_vars_vo[i][3]= "$".join([list_vars_vo[i][0],list_vars_vo[i][3]])
    dict_vo = F.list_to_dict(list_vars_vo,'Код')
    current_row = "$".join([nn,spis_tk[j][3]])
    if current_row not in dict_vo:
        CQT.msgbox(f'{current_row} не найдена в шаблоне под ТКП')
        return spis_tk[j]
    list_vars = [[],[]]
    for key in dict_vo[current_row].keys():
        if key not in ['ДСЕ','Операция','Переход','Код']:
            if dict_vo[current_row][key] != '_':
                list_vars[0].append(key)
                list_vars[1].append(str(dict_vo[current_row][key]))

    time = operacii.vremya_tsht(dict_vo[current_row]['Операция'], list_vars)
    list_mat = operacii.materiali(self, dict_vo[current_row]['Операция'], list_vars)
    #==============ADD OSN MATS±+++++++++++++
    if current_row in self.dict_mat_vo:
        list_vsp_mat_tmp = list_mat.split('{')
        list_osn_mat = self.dict_mat_vo[current_row].split('{')
        for mat in list_vsp_mat_tmp:
            if mat != '':
                list_osn_mat.append(mat)
        list_mat= '{'.join(list_osn_mat)
    #+++++++++++++++++++++++++++++

    if j < len(spis_tk)-1:
        for i in range(j + 1,len(spis_tk)):
            current_row = "$".join([nn, spis_tk[i][3]])
            if spis_tk[i][20] == '1' or spis_tk[i][20] == '0':
                break
            if spis_tk[i][20] == '2':
                list_vars_pereh = [[], []]
                if current_row in dict_vo:
                    for key in dict_vo[current_row].keys():
                        if key not in ['ДСЕ', 'Операция', 'Переход', 'Код']:
                            if dict_vo[current_row][key] != '_':
                                list_vars_pereh[0].append(key)
                                list_vars_pereh[1].append(str(dict_vo[current_row][key]))

                    vrema = operacii.vremya_tsht_perehodi(dict_vo[current_row]['Операция'],spis_tk[i][0], list_vars_pereh, list_vars)
                    #materials = operacii.materiali(self, dict_vo[current_row]['Операция'], arr_tmp)
                    time += vrema
                #for mat in materials:
                #    list_mat.append(mat)
    spis_tk[j][nk_mat_tk] = list_mat
    spis_tk[j][nk_op_tst] = time
    return spis_tk[j]


def DICT_VAR_OPER(self):
    self.DICT_VAR_OPER = dict()
    if self.SPIS_OP == None:
        quit()
    dict_var_oper = F.list_to_dict(self.SPIS_OP,'name')
    #dict_var_oper = list_var_from_txt(F.otkr_f(F.tcfg('oper'), separ='|'))
    for oper in dict_var_oper.keys():
        dict_var_pereh =[]
        if F.nalich_file(F.scfg('oper') + F.sep() + f'{oper}.txt'):
            dict_var_pereh = list_var_from_txt(F.otkr_f(F.scfg('oper') + F.sep() + f'{oper}.txt', separ='|'),2)
        self.DICT_VAR_OPER[oper] = [dict_var_oper[oper],dict_var_pereh]

def list_var_from_txt(list,shag=3):
    dict_ = dict()
    for oper in list:
        if len(oper) > shag:
            list_var = oper[shag].split(';')
            tmp_var = []
            for var in list_var:
                tmp_var.append(var.split(':')[0])
            dict_[oper[0]] = tmp_var
    return dict_