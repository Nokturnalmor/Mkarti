# coding=cp1251
import pythoncom
import win32com.client
import project_cust_38.Cust_SQLite as CSQ
import project_cust_38.Cust_Functions as F


def query_run_unify(V83, querytxt):
    query = V83.NewObject("Query", querytxt)
    return query.Execute().Choose()

def connect():
    print('Try connect to ERP')
    print(F.user_name())
    #put_f = F.scfg('cash') + F.sep() + 'users_erp.txt'
    #if F.nalich_file(put_f) == False:
    #    return 'Не найен файл с ключами пользователей ерп'
    #print("Файл найден")
    #spis_users = F.otkr_f(put_f,True,separ='|')
    spis_users = [["a.belyakov","Беляков Антон Геннадьевич","25012022"],["a.stepanova","Степанова Алёна Сергеевна","25012022"]]
    login = ''
    password = ''
    for i in range(len(spis_users)):
        if spis_users[i][0] == F.user_name():
            print("Пользователь найден")
            login = spis_users[i][1]
            password = spis_users[i][2]
            break
    if login == '' or password == '':
        return 'Не найден логин/пароль'
    #V83_CONN_STRING = 'Srvr="novgorod";Ref="ERP";Usr="Беляков Антон Геннадьевич";Pwd="25012022";'
    V83_CONN_STRING = f'Srvr="novgorod";Ref="ERP";Usr="{login}";Pwd="{password}";'
    print(f'Ввод {V83_CONN_STRING}')
    pythoncom.CoInitialize()
    V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    print('         .... OK')
    print('')
    return

def connect_test():
    print('Try connect to ERP')
    print(F.user_name())
    #put_f = F.scfg('cash') + F.sep() + 'users_erp.txt'
    #if F.nalich_file(put_f) == False:
    #    return 'Не найен файл с ключами пользователей ерп'
    #print("Файл найден")
    #spis_users = F.otkr_f(put_f,True,separ='|')
    spis_users = [["a.belyakov","Беляков Антон Геннадьевич","25012022"],["a.stepanova","Степанова Алёна Сергеевна","25012022"]]
    login = ''
    password = ''
    for i in range(len(spis_users)):
        if spis_users[i][0] == F.user_name():
            print("Пользователь найден")
            login = spis_users[i][1]
            password = spis_users[i][2]
            break
    if login == '' or password == '':
        return 'Не найден логин/пароль'
    #V83_CONN_STRING = 'Srvr="novgorod";Ref="ERP";Usr="Беляков Антон Геннадьевич";Pwd="25012022";'
    V83_CONN_STRING = f'Srvr="KE-IT02";Ref="ERP_Copy";Usr="Test";Pwd="Df90cv";'
    print(f'Ввод {V83_CONN_STRING}')
    pythoncom.CoInitialize()
    try:
        V83_2 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
    except:
        print('err')
    print('         .... OK')
    print('')
    return

def query_mat(mat, V83):
    # get = lambda obj,attr: getattr(obj, str(attr.encode('cp1251', 'ignore')))
    # catalog = getattr(V83.Catalogs, "Документы.ЗаданиеНаРезку")
    spis = [["rez.Код", "rez.Артикул", "rez.Наименование", "rez.ЕдиницаИзмерения", "rez.ПометкаУдаления"]]
    query_mat = f'''ВЫБРАТЬ
        Номенклатура.ПометкаУдаления КАК ПометкаУдаления,
        Номенклатура.Наименование КАК Наименование,
        Номенклатура.Артикул КАК Артикул,
        Номенклатура.ЕдиницаИзмерения.Наименование КАК ЕдиницаИзмерения,
        Номенклатура.Код КАК Код
    ИЗ
        Справочник.Номенклатура КАК Номенклатура
    ГДЕ
        Номенклатура.ВидНоменклатуры.Наименование = "{mat}"'''

    rez = query_run_unify(V83, query_mat)
    while rez.next():
        print(rez.Код, rez.Артикул, rez.Наименование, rez.ЕдиницаИзмерения, rez.ПометкаУдаления)
        spis.append([rez.Код, rez.Артикул, rez.Наименование, rez.ЕдиницаИзмерения, rez.ПометкаУдаления])
    return spis

def export_date_plan(self):
    conn = connect_test()
    return
    vid = 'Болты высокопрочные'
    table = query_mat(vid, conn)

if __name__ == '__main__':
    export_date_plan('')
    a=3