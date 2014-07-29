# -*- coding: utf-8 -*-
"""
Created on Mon Jul 28 14:45:58 2014

@author: russinow
"""

import xml.etree.ElementTree as ET
import argparse



parser = argparse.ArgumentParser(description='convert .xml to .csv')
parser.add_argument('file', metavar='FILE', type=str, nargs='+',
                  help='file to convert')

BNCpath = 'D:/BNCcorp/'

args = parser.parse_args()
a = 1                 

#if args.file == ['None']:
#    print('Set file to convert')
#    a = 0



#получаем индекс файлов
indexXML = ET.parse(BNCpath + 'Etc/file_index.xml') #(file_in)
index_root = indexXML.getroot()

index = {ir.text[5:8]:ir.text  for ir in index_root.iter('file')}

if a == 1: 
    for f in args.file:
        #импортируем файл для конвертации
        tree = ET.parse(f) #(file_in)
        root = tree.getroot()
        
        w = open(f[:len(f)-4] + '.csv', 'w')
        for hit in root.iter('hit'):
            #print BNCpath + index[hit.get('text')]
            #импортируем файл исходник, для получения метаинформации
            cource = ET.parse(BNCpath + 'Texts/' + index[hit.get('text')]) #(file_in)
            crc_r = cource.getroot()
            #print(str(hit.get('text')) + ';' + str(hit.text) + ';' + str(hit.find('kw').text) + ';' + str(hit.find('kw').tail))
            w.write( "%s; %s; %s; %s \n"  % (hit.get('text'), hit.text, hit.find('kw').text, hit.find('kw').tail))


    #def ff(file_in):  C:/Users/russinow/Desktop/Query1.xml
#tree = ET.parse(args.file) #(file_in)
#root = tree.getroot()
#for hit in root.iter('hit'):
#        print(hit.get('text')+'; ' + hit.text + '; ' + hit.find('kw').text+ '; ' + hit.find('kw').tail)