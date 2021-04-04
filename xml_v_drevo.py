import lxml2dict as L
from lxml import etree


def spisok_iz_xml(putt):
    doc = etree.parse(putt)
    rr = doc.getroot()
    tree = L.convert(rr)
    s = list_tree(tree)
    return s
#doc = etree.parse('P:\\Python\\Mkarti\\КТ.1408182.11 (меньше метизов).xml')


def base_sp_names(tree):
    s = []
    sp = tree['Root']['Elements']['Element']['Parameters']['Parameter']
    for i in range(0, len(sp)):
        s.append(sp[i]['@Name'])
    return s

def base_ob(tree,name):
    sp = tree['Root']['Elements']['Element']['Parameters']['Parameter']
    for i in range(0, len(sp)):
        if sp[i]['@Name'] == name:
            otv = sp[i]['@Value']
            return otv
    return None



def sp_imen_child_poputi(putt):
    s=[]
    for i in range(0, len(putt['Children']['Element'])):
        s.append(putt['Children']['Element'][i]['@ObjectId'])
    return s

def spis_child(tree,obj):
    s = []
    if obj == '':
        if type(tree['Root']['Elements']['Element']) == type([]):
            for i in range(0, len(tree['Root']['Elements']['Element'])):
                s.append(tree['Root']['Elements']['Element'][i]['@ObjectId'])
        else:
            s.append(tree['Root']['Elements']['Element']['@ObjectId'])
        return s
    if obj == spis_child(tree,"")[0]:
        return sp_imen_child_poputi(tree['Root']['Elements']['Element'])

    putt = tree['Root']['Elements']['Element']
    putt = naity_put_obj(putt,obj)
    if putt == None:
        return None
    return sp_imen_child_poputi(putt)

def naity_put_obj(putt,obj):
    if putt.get('Children') != None:
        for i in range(0, len(putt['Children']['Element'])):
            #s.append(putt['Children']['Element'][i]['@ObjectId'])
            if putt['Children']['Element'][i]['@ObjectId'] == obj:
                return putt['Children']['Element'][i]
            else:
                naity_put_obj(putt['Children']['Element'][i], obj)
    return None

def znach_param(putt,param):
    for i in range(0, len(putt)):
        if putt[i]['@Name'] == param:
            return putt[i]['@Value']


def oform_strok(putt, ur):
    ID = putt['@ObjectId']
    p1 = znach_param(putt['Parameters']['Parameter'], 'Наименование')
    p2 = znach_param(putt['Parameters']['Parameter'], 'Обозначение полное')
    p3 = znach_param(putt['Parameters']['Parameter'], 'Количество')
    p4 = znach_param(putt['Parameters']['Parameter'], 'Материал')
    p5 = znach_param(putt['Parameters']['Parameter'], 'Материал2')
    p6 = znach_param(putt['Parameters']['Parameter'], 'Материал3')
    p7 = znach_param(putt['Parameters']['Parameter'], 'Количество на изделие')
    p8 = znach_param(putt['Parameters']['Parameter'], 'Масса')
    p9 = znach_param(putt['Parameters']['Parameter'], 'Покупное изделие')

    return [p1, p2,p3, p4, p5, p6, ID, p7, p8, p9, '', '', '', '', '', '', '', '', '', '', ur]

def naidi_child(putt,s,ur):
    if putt.get('Children') != None:
        ur+=1
        for i in range(0, len(putt['Children']['Element'])):
            s.append(oform_strok(putt['Children']['Element'][i],ur))
            s = naidi_child(putt['Children']['Element'][i], s, ur)
    return s


def list_tree(tree):
    ur = 0
    s = []
    if type(tree['Root']['Elements']['Element']) == type([]):
        for i in range(0, len(tree['Root']['Elements']['Element'])):
            s.append(oform_strok(tree['Root']['Elements']['Element'][i],ur))
            s = naidi_child(tree['Root']['Elements']['Element'][i],s,ur)
    else:
        s.append(oform_strok(tree['Root']['Elements']['Element'],ur))
        s = naidi_child(tree['Root']['Elements']['Element'], s, ur)

    return s


#s = list_tree(tree)
#for i in s:
#    pr = ''
#    for j in range(0,6):
#        pr +=  i[j] +"|"
#    print("    " * i[20]+ pr + ' - ' + str(i[20]))


