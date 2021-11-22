# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 09:12:00 2019

@author: Michał Nowicki

import os 
import urllib
from pandas import DataFrame


def filesercher(file):
    files=os.listdir(file)
    return files 
    
def string_sort (string,seed):
    return string[seed-5000:seed+10000]

def namemain(string):
    result = string.find( '<td class="left">Rubryka 3. Firma, nazwa albo imię i nazwisko podmiotu leczniczego</td>') 
    start=result+117
    stopstr=string[start:start+300]
    stop=stopstr.find('</td>')
    return (string[start:start+stop])

def subname(string):
    result = string.find( '<td class="left">Rubryka 1. Nazwa komórki organizacyjnej</td>') 
    start=result+91
    stopstr=string[start:start+300]
    stop=stopstr.find('</td>')
    return (string[start:start+stop])

def street(string):
    result = string.find( '<td class="leftThickPadding">1. Ulica</td>') 
    start=result+78
    stopstr=string[start:start+300]
    stop=stopstr.find('</td>')
    return (string[start:start+stop])
    
def streetnumber(string):
    stop=string.find('</adr:Budynek>')
    startstr=string[stop-20:stop]
    start=startstr.find('">')
    number=startstr[start+2:]
    return  number
     
def contact_number(string):
    result = string.find( '<td class="leftThickPadding">6. Numer telefonu</td>') 
    start=result+86
    stopstr=string[start:start+300]
    stop=stopstr.find('</td>')
    return (string[start:start+stop])

def email(string):
    result = string.find( '<td class="left">Rubryka 3. Adres poczty elektronicznej</td>') 
    start=result+90
    stopstr=string[start:start+300]
    stop=stopstr.find('</td>')
    return (string[start:start+stop])

def postcode(string):
    stop2=string.find('</adr:KodPocztowy>')
    postcode=string[stop2-6:stop2]
    return postcode

def city(string):
    stop=string.find('</adr:Miejscowosc>')
    startstr=string[stop-40:stop]
    start=startstr.find('">')
    number=startstr[start+2:]
    return number


def main():
    code= 4300 
    file1="file:///C:/Users/Michał20Nowicki/Desktop/rehab/input_data/"
    file2=("C:/Users/Michał Nowicki/Desktop/rehab/input_data/")
    file_out=("C:/Users/Michał Nowicki/Desktop/rehab/output_data/")
    
    
    N=(len(os.listdir(file2)))#(1)
    
    col1=["" for x in range(0,N)]
    col2=["" for x in range(0,N)]
    col3=["" for x in range(0,N)]
    col4=["" for x in range(0,N)]
    col5=["" for x in range(0,N)]
    col6=["" for x in range(0,N)]
    col7=["" for x in range(0,N)]
    col8=["" for x in range(0,N)]
    for i in range (0,N):
        filepath=file1+filesercher(file2)[i]   
        wp = urllib.request.urlopen(filepath)
        pw = wp.read()
        pw_str=pw.decode('utf8')
        result = pw_str.find('<typ:KodResort>'+str(code)+'</typ:KodResort>') 
        neuro_str=string_sort(pw_str,result)
        col1[i]=namemain(pw_str)
        col2[i]=subname(neuro_str)
        col3[i]=street(neuro_str)
        col4[i]=streetnumber(neuro_str)
        col5[i]=city(neuro_str)
        col6[i]=postcode(neuro_str)
        col7[i]=contact_number(neuro_str)
        col8[i]=email(neuro_str)
        
    df = DataFrame({'Nazwa jednowski': col1, 'Nazwa oddziału': col2,'Ulica': col3,'Nr budynku': col4,'Miasto': col5,'Kod pocztowy': col6,'Nr telefonu': col7,'Email': col8})
    df.to_excel(file_out+'database_technomex_'+str(code)+'.xlsx', sheet_name='sheet1', index=False)
    print('database_technomex_'+str(code)+'.xlsx'+' created')
    
if __name__ == "__main__":
    main()
