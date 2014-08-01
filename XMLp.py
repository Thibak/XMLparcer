# -*- coding: utf-8 -*-
"""
Created on Mon Jul 28 14:45:58 2014

@author: russinow
"""

import xml.etree.ElementTree as ET
#import argparse
#import sys
import xlrd # чтение файлов Excel
# файл с настройками
import paths

#if sys.argv == [sys.argv[0]]:
#	sys.exit()

#формируем парсер аргументов командной строки
#parser = argparse.ArgumentParser(description='convert .xml to .csv')
#parser.add_argument('file', metavar='FILE', type=str, nargs='+',
#                  help='file to convert')
#args = parser.parse_args()


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
    
# функция возвращающая значение по паре в нотации самого Excel (заодно сразу перевожу в str)
def GetItem(stri,intg):
    return sheet.cell(intg-1,colindex(stri)).value.encode('ascii','ignore')

def GetText(host_el):
    if host_el is not None:
        return host_el.text
    else:
        return ''

#открываем файл аутпута
output = open(paths.BNCpath + 'output.csv', 'w')
output.write("leftContext[] ; AN[]       ; ; rightContext[] ; c5 ; pos ; n ; PersonDict[who].attrib['sex'] ;  PersonDict[who].find('age').text ; PersonDict[who].find('persName').text ; PersonDict[who].find('occupation').text ; PersonDict[who].find('dialect').text ; f ; title")
output.write("left Context  ; Abstr Noun ; ; right Context  ; c5 ; pos ; n ; sex                           ;                        age        ;                       persName        ;                       occupation        ;                       dialect        ; f ; title")


#переменные хода прогресса
progress = 0
gut = 0

#for f in args.file:#импортируем файл для конвертации

for f in index:
    progress = progress + 1
    RowNum = SourceRowIndex[f]
    print (str ((len(index) - progress))+'-->'+str(RowNum)+' |=| '+str(f))
    # Условия отбора
    if (GetItem('Q',RowNum) == "S") and GetItem('C',RowNum) in ("S_Demog_AB", "S_Demog_C1", "S_Demog_C2", "S_Demog_DE", "S_Demog_Unclassified"):
        print("CATCH IT: "+str(RowNum)+' |=| '+str(f))  
        gut = gut+1
        # нужный файл найден, начинаем его анализировать, и расчленять
#        tree = ET.parse(paths.BNCpath + 'Texts/' + index[f]) #(file_in)
#        root = tree.getroot()       
        
print('catched: '+ str(gut))
#        
        #for hit in root.iter('hit'):#w = open(f[:len(f)-4] + '.csv', 'w')
            #print BNCpath + index[hit.get('text')]
            #импортируем файл исходник, для получения метаинформации
         #   cource = ET.parse(paths.BNCpath + 'Texts/' + index[hit.get('text')]) #(file_in)
         #   crc_r = cource.getroot()
#			for n in crc_r.findall(\\)
            #print(str(hit.get('text')) + ';' + str(hit.text) + ';' + str(hit.find('kw').text) + ';' + str(hit.find('kw').tail))
            #w.write( "%s; %s; %s; %s \n"  % (hit.get('text'), hit.text, hit.find('kw').text, hit.find('kw').tail))


    #def ff(file_in):  C:/Users/russinow/Desktop/Query1.xml
#tree = ET.parse(args.file) #(file_in)
#root = tree.getroot()
#for hit in root.iter('hit'):
#        print(hit.get('text')+'; ' + hit.text + '; ' + hit.find('kw').text+ '; ' + hit.find('kw').tail)

tree = ET.parse(paths.BNCpath + 'Texts/' + index['KD9']) #(file_in)
root = tree.getroot() 

#вычленяем хедер и данные из него
title = tree.find('.//title').text
#анализатор персон
PersonDict = {}
for pers in tree.findall('.//person'):
    PersonDict[pers.attrib['{http://www.w3.org/XML/1998/namespace}id']] = pers 

sp = ' ; '
# основное, говорящий атрибуты, говорящий ноды, имя источника, заголовок книгиб

for u in root.findall('.//u'):
            who = u.attrib['who']    
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
                            leftContext.append(str(leftContext[i]))
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
                    print (leftContext[k]+sp+ \
                            AN[k]+sp+ \
                            rightContext[k]+sp+sp+ \
                            c5[k] +sp+ \
                            pos[k] +sp+ \
                            n +sp+ \
                            PersonDict[who].attrib['sex'] +sp+ \
                            PersonDict[who].attrib['role'] +sp+ \
                            PersonDict[who].attrib['sex'] +sp+ \
                            PersonDict[who].attrib['soc'] +sp+ \
                            PersonDict[who].attrib['dialect'] +sp+ \
                            GetText(PersonDict[who].find('age')) +sp+ \
                            GetText(PersonDict[who].find('persName')) +sp+ \
                            GetText(PersonDict[who].find('occupation')) +sp+ \
                            GetText(PersonDict[who].find('dialect')) +sp+ \
                            f +sp+ \
                            title)





output.close()