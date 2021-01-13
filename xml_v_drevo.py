import Cust_Functions as F

import lxml2dict as L
from lxml import etree

doc = etree.parse('P:\\Python\\Mkarti\\КТ.1408182.11 (меньше метизов).xml')

rr = doc.getroot()

tree = L.convert(rr)

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

def naity_put_obj(putt,obj):
    if putt.get('Children') != None:
        for i in range(0, len(putt['Children']['Element'])):
            #s.append(putt['Children']['Element'][i]['@ObjectId'])
            if putt['Children']['Element'][i]['@ObjectId'] == obj:
                return putt['Children']['Element'][i]
            else:
                naity_put_obj(putt['Children']['Element'][i], obj)
    return None

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


obj2 = '0x108000408'

s = spis_child(tree,obj2)
for i in s:
    print(i)

sp = base_sp_names(tree)
#for i in sp:
#    print(i, end=" = ")
#    print(base_ob(tree, i))

#print(base_ob(tree,'Наименование'))

