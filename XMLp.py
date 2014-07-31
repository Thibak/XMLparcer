# -*- coding: utf-8 -*-
"""
Created on Mon Jul 28 14:45:58 2014

@author: russinow
"""

import xml.etree.ElementTree as ET
#import argparse
#import sys
import xlrd
# файл с настройками
import paths

#if sys.argv == [sys.argv[0]]:
#	sys.exit()

#формируем парсер аргументов командной строки
#parser = argparse.ArgumentParser(description='convert .xml to .csv')
#parser.add_argument('file', metavar='FILE', type=str, nargs='+',
#                  help='file to convert')
#args = parser.parse_args()


def colname(colx):
    """ 7 => 'H', 27 => 'AB' """
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if colx <= 25:
        return alphabet[colx]
    else:
        xdiv26, xmod26 = divmod(colx, 26)
        return alphabet[xdiv26 - 1] + alphabet[xmod26]
# функция обратная к colname
def colindex(colname):
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if len(colname) == 1:
        return alphabet.find(colname.upper())
    else:
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
    SourceRowIndex[sheet.cell(rownum,0).value.encode('ascii','ignore')] =  rownum
    
#for f in args.file:#импортируем файл для конвертации
for f in index:
        tree = ET.parse(index[f]) #(file_in)
        root = tree.getroot()
        InfRow = SourceRowIndex[]      
        
        w = open(f[:len(f)-4] + '.csv', 'w')
        for hit in root.iter('hit'):
            #print BNCpath + index[hit.get('text')]
            #импортируем файл исходник, для получения метаинформации
            cource = ET.parse(paths.BNCpath + 'Texts/' + index[hit.get('text')]) #(file_in)
            crc_r = cource.getroot()
#			for n in crc_r.findall(\\)
            #print(str(hit.get('text')) + ';' + str(hit.text) + ';' + str(hit.find('kw').text) + ';' + str(hit.find('kw').tail))
            w.write( "%s; %s; %s; %s \n"  % (hit.get('text'), hit.text, hit.find('kw').text, hit.find('kw').tail))


    #def ff(file_in):  C:/Users/russinow/Desktop/Query1.xml
#tree = ET.parse(args.file) #(file_in)
#root = tree.getroot()
#for hit in root.iter('hit'):
#        print(hit.get('text')+'; ' + hit.text + '; ' + hit.find('kw').text+ '; ' + hit.find('kw').tail)