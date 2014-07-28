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
args = parser.parse_args()

for f in args.file:
    tree = ET.parse(f) #(file_in)
    root = tree.getroot()
    
    w = open(f[:len(f)-4] + '.csv', 'w')
    for hit in root.iter('hit'):
        #print(str(hit.get('text')) + ';' + str(hit.text) + ';' + str(hit.find('kw').text) + ';' + str(hit.find('kw').tail))
        w.write( "%s; %s; %s; %s \n"  % (hit.get('text'), hit.text, hit.find('kw').text, hit.find('kw').tail))


    #def ff(file_in):  C:/Users/russinow/Desktop/Query1.xml
#tree = ET.parse(args.file) #(file_in)
#root = tree.getroot()
#for hit in root.iter('hit'):
#        print(hit.get('text')+'; ' + hit.text + '; ' + hit.find('kw').text+ '; ' + hit.find('kw').tail)