# -*- coding: utf-8 -*-
"""
Created on Mon Jul 28 14:45:58 2014

@author: russinow
"""

import xml.etree.ElementTree as ET # читаем XML
import xlrd # чтение файлов Excel
import xlwt # пишем в Excel
# файл с настройками
import paths 

# функция обратная к colname
def colindex(colname):
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    k = alphabet.find(colname[len(colname)-1].upper())
    for i in range(len(colname)):
            k = k + (alphabet.find(colname[i].upper())+1)*(26*(len(colname)-i-1))
    return k

#получаем индекс файлов
indexXML = ET.parse(paths.BNCpath + 'Etc/file_index.xml') #(file_in)
index_root = indexXML.getroot()
index = {ir.text[5:8]:ir.text  for ir in index_root.iter('file')}

#читаем базу источников
rb = xlrd.open_workbook(paths.BNCpath + 'BNC_WORLD_INDEX.XLS',formatting_info=True)
sheet = rb.sheet_by_index(0) 

#формируем индекс описания источников
SourceRowIndex = {}
for rownum in range(sheet.nrows):
    SourceRowIndex[sheet.cell(rownum,0).value.encode('ascii','ignore')] =  rownum+1

# создаем выходной файл и листы в нем
#('UU','DE','C1','AB','C2')
wb = xlwt.Workbook()
UU = wb.add_sheet('UU')
DE = wb.add_sheet('DE')
C1 = wb.add_sheet('C1')
AB = wb.add_sheet('AB')
C2 = wb.add_sheet('C2')
meta = wb.add_sheet('meta')

    
# функция возвращающая значение по паре в нотации самого Excel (заодно сразу перевожу в str)
def GetItem(stri,intg):
    return sheet.cell(intg-1,colindex(stri)).value#.encode('ascii','ignore')

#инкапсуляция проверки получения ноды
def GetTextFind(host_el, key, akey):
    try:
        return host_el[key].find(akey).text
    except KeyError:
        return ''
    except AttributeError:
        return ''

def GetAttr(host_el, key, akey):
    try:
        return host_el[key].attrib[akey]
    except KeyError:
        return ''
        
def writeRow(sheet, row, vec):
    for i in range(len(vec)):
        sheet.write(row, i, vec[i])

def integerate(s):
    try:
        return int(s)
    except ValueError:
        return s
#        try:
#           return float(s)
#        except ValueError:
#            return s

#открываем файл аутпута
#output = open(paths.BNCpath + 'output.csv', 'w')
#пишем в него хедеры столбцов
#output.write("leftContext[] ; AN[]       ; ; rightContext[] ; c5 ; pos ; n ; PersonDict[who].attrib['sex'] ;  PersonDict[who].find('age').text ; PersonDict[who].find('persName').text ; PersonDict[who].find('occupation').text ; PersonDict[who].find('dialect').text ; f ; title")
#output.write("left Context  ; Abstr Noun ; ; right Context  ; c5 ; pos ; n ; sex                           ;                        age        ;                       persName        ;                       occupation        ;                       dialect        ; f ; title")

#переменные хода прогресса
progress = 0
sp = ' ; '
num = 0
# счетчик строк
RowNumIndex = {'UU':1,'DE':1,'C1':1,'AB':1,'C2':1}
vec = ['left Context',\
        'Abstr Noun',\
        'right Context',\
        'c5', \
        'pos',\
        'sex',\
        'role',\
        'social class',\
        'dialect',\
        'occupation',\
        'age',\
        'name',\
        'dialect (code)',\
        'file name',\
        'GENRE',\
        'Notes & Alternative Genres',\
        'Interaction Type',\
        'Time Period (Alltim)',\
        'sentence num',\
        'Domain',\
        'Word Total (in source)',\
        'source title'\
        ]
#vec = ['left Context',\#vec = [leftContext[k],\
#        'Abstr Noun',\# AN[k],\
#        'right Context',\# rightContext[k], \
#        'c5', \# c5[k], \
#        'pos',\# pos[k], \
#        'sex',\# GetAttr(PersonDict,who,'sex'), \
#        'role',\# GetAttr(PersonDict,who,'role'), \
#        'social class',\# GetAttr(PersonDict,who,'soc'), \
#        'dialect',\# GetTextFind(PersonDict,who,'dialect'), \
#        'occupation'# GetTextFind(PersonDict,who,'occupation'), \
#        'age',\# GetTextFind(PersonDict,who,'age'), \
#        'name',\# GetTextFind(PersonDict,who,'persName'),\
#        'dialect (code)',\# GetAttr(PersonDict,who,'dialect'), \
#        'file name',\# f, \
#        'sentence num',\# n, \
#        'cource title'\# title ]
        


writeRow(UU, 0, vec)
writeRow(DE, 0, vec)
writeRow(C1, 0, vec)
writeRow(AB, 0, vec)
writeRow(C2, 0, vec)

# прочесываем все файлы на предмет соответствия заданным характеристикам
for f in index:
    progress = progress + 1
    RowNum = SourceRowIndex[f]
    print (str ((len(index) - progress))+'-->'+str(RowNum)+' |=| '+str(f) + ' ЗАПИСЕЙ: '+ str(num))
    # Условия отбора
    if (GetItem('Q',RowNum) == "S") and GetItem('C',RowNum) in ("S_Demog_AB", "S_Demog_C1", "S_Demog_C2", "S_Demog_DE", "S_Demog_Unclassified"):
        tree = ET.parse(paths.BNCpath + 'Texts/' + index[f]) # нужный файл найден, начинаем его анализировать, и расчленять
        root = tree.getroot()       
        #вычленяем хедер и данные из него
        title = tree.find('.//title').text
        #анализатор персон
        PersonDict = {}
        for pers in tree.findall('.//person'):
            PersonDict[pers.attrib['{http://www.w3.org/XML/1998/namespace}id']] = pers 
        # основное, говорящий атрибуты, говорящий ноды, имя источника, заголовок книгиб
        for u in root.findall('.//u'):
            who = u.attrib['who']    
            if GetAttr(PersonDict,who,'dialect') in ('XNC','XNO','XSL','XLO','XMW','XME','XMC','XNE','XLC','XSS'):
                for s in u.findall('s'): #s = root.find('.//s')
                    n = s.attrib['n']
                    leftContext = ['']
                    rightContext = ['']
                    b = [True]# идентификатор до или после. После того, как находим пишем и до и после.
                    flag = [1]                
                    AN = []
                    c5 = []
                    hw = []
                    pos = []
                    i = 0 # идентификатор количества найденых существительных
                    for w in s.findall('.//'): # все содержимое блока.
                        flag[i-1] = 1  
                        word = ''                        
                        if w.tag == 'c':
                            word = w.text               
                        if w.tag == 'w':
                            word = w.text                         
                            if w.attrib['c5'] in ('NN0', 'NN1', 'NN1-AJ0', 'NN1-NP0', 'NN1-WB', 'NN1-WG', 'NN2', 'NN2-WZ', 'NP0-NN1', 'UNC', 'WB-NN1', 'WG-NN1', 'WZ-NN2'): 
                                AN.append(w.text)
                                c5.append(w.attrib['c5'])
                                hw.append(w.attrib['hw'])
                                pos.append(w.attrib['pos'])
                                leftContext.append(leftContext[i])
                                rightContext.append('')
                                b[i] = False
                                flag[i] = 0
                                b.append(True)
                                flag.append(1)
                                i = i + 1                            
                        for ik in range(i+1):
                            if b[ik]: 
                                leftContext[ik] = leftContext[ik] + word
                            else:                            
                                rightContext[ik] = rightContext[ik] + word*flag[ik]
                        #Аутпутим в цикле по range(i)
                    for k in range(i):
                        num = num + 1
                        soc = GetAttr(PersonDict,who,'soc')
                        vec = [leftContext[k],\
                            AN[k],\
                            rightContext[k], \
                            c5[k], \
                            pos[k], \
                            GetAttr(PersonDict,who,'sex'), \
                            GetAttr(PersonDict,who,'role'), \
                            GetAttr(PersonDict,who,'soc'), \
                            GetTextFind(PersonDict,who,'dialect'), \
                            GetTextFind(PersonDict,who,'occupation'), \
                            integerate(GetTextFind(PersonDict,who,'age')), \
                            GetTextFind(PersonDict,who,'persName'),\
                            GetAttr(PersonDict,who,'dialect'), \
                            f, \
                            GetItem('D',RowNum),\
                            GetItem('E',RowNum),\
                            GetItem('O',RowNum),\
                            GetItem('P',RowNum),\
                            integerate(n), \
                            GetItem('C',RowNum),\
                            integerate(GetItem('K',RowNum)),\
                            title ]
                            
                        #('UU','DE','C1','AB','C2')
                        if   soc == 'UU':
                            writeRow(UU, RowNumIndex['UU'], vec)
                            RowNumIndex['UU'] = RowNumIndex['UU'] + 1
                        elif soc == 'DE':
                            writeRow(DE, RowNumIndex['DE'], vec)
                            RowNumIndex['DE'] = RowNumIndex['DE'] + 1
                        elif soc == 'C1':
                            writeRow(C1, RowNumIndex['C1'], vec)
                            RowNumIndex['C1'] = RowNumIndex['C1'] + 1
                        elif soc == 'AB':
                            writeRow(AB, RowNumIndex['AB'], vec)
                            RowNumIndex['AB'] = RowNumIndex['AB'] + 1
                        elif soc == 'C2':
                            writeRow(C2, RowNumIndex['C2'], vec)
                            RowNumIndex['C2'] = RowNumIndex['C2'] + 1
                        else:
                            pass
                      
                        
                       # ('UU','DE','C1','AB','C2')

                   #     ws.write(num, 0, leftContext[k])
                   #     ws.write(num, 1, AN[k])
                   #     ws.write(num, 2, rightContext[k])
                   #     ws.write(num, 3, c5[k])
                   #     ws.write(num, 4, pos[k])
                   #     ws.write(num, 5, n)
                   #     ws.write(num, 6, GetAttr(PersonDict,who,'sex'))
                   #     ws.write(num, 7, GetAttr(PersonDict,who,'role'))
                   #     ws.write(num, 8, GetAttr(PersonDict,who,'soc'))
                   #     ws.write(num, 9, GetAttr(PersonDict,who,'dialect'))
                   #     ws.write(num, 10, GetTextFind(PersonDict,who,'age'))
                   #     ws.write(num, 11, GetTextFind(PersonDict,who,'persName'))
                   #     ws.write(num, 12, GetTextFind(PersonDict,who,'occupation'))
                   #     ws.write(num, 13, GetTextFind(PersonDict,who,'dialect'))
                   #     ws.write(num, 14, f)
                        #ws.write(num, 15, )
                        #ws.write(num, 1, )
                        
                        
                        
                        #a = str(num) + leftContext[k]+sp+ \
                        #        AN[k]+sp+ \
                        #        rightContext[k]+sp+sp+ \
                        #        c5[k] +sp+ \
                        #        pos[k] +sp+ \
                        #        n +sp+ \
                        #        GetAttr(PersonDict,who,'sex') +sp+ \
                        #        GetAttr(PersonDict,who,'role') +sp+ \
                        #        GetAttr(PersonDict,who,'soc') +sp+ \
                        #        GetAttr(PersonDict,who,'dialect') +sp+ \
                        #        GetTextFind(PersonDict,who,'age') +sp+ \
                        #        GetTextFind(PersonDict,who,'persName') +sp+ \
                        #        GetTextFind(PersonDict,who,'occupation') +sp+ \
                        #        GetTextFind(PersonDict,who,'dialect') +sp+ \
                        #        f
                       # output.write(a.encode('utf-8'))
#output.close()
meta.write(0, 0, 'records')
meta.write(0, 1, num)

i = 1
for key in RowNumIndex:
    meta.write(i, 0, key)
    #print (str(i)+','+key)
    meta.write(i, 1, RowNumIndex[key]-1)    
    #print (str(i)+','+str(RowNumIndex[key]))
    i = i+1

wb.save(paths.BNCpath + 'base.xls')
print(num)