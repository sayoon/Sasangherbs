# -*- coding: utf-8 -*-
"""
Created on Tue Nov 21 13:24:18 2017

@author: physiology5
"""

import os
os.chdir(r"C:\Users\physiology5\Dropbox\연구\1709_사상본초\171107_사상본초파일")
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import scipy.stats as ss



#DataFrame으로 엑셀파일 불러오기
TCMID_Ingredient= pd.read_excel(r'C:\Users\physiology5\Dropbox\_연구\1709_사상본초\171107_사상본초파일\TCMID_Ingredient.xlsx', sheetname='Sheet1')
SE01 = pd.read_excel('C:\\Users\\physiology5\\Dropbox\\연구\\1709_사상본초\\171107_사상본초\\171017_본초개별파일\\SE01.xlsx', sheetname='TCMID', header=None )




##코드 짜는 과정         
SE01= SE01.iloc[:,0] 
#SE01.str.split(':')[1]

SE1= SE01.str.extract('([0-9]+)')
SE1= int(SE1)
#SE1= pd.DataFrame(SE1)

SE1= pd.DataFrame(SE1, columns=['compound_id'], dtype='str')
SE1= SE1.astype('int')

#column=['compound_id']
#SE1=SE01.reindex(columns=column)        

merge1=pd.merge(SE1, TCMID_Ingredient)





##자동화하기 위해서 함수 생성
def readexcel_merge(file,i):
    SEno = pd.read_excel(file, sheetname='TCMID', header=None)
    SEno = SEno.iloc[:,0] 
    SEno = SEno.str.extract('([0-9]+)')
    SEno = pd.DataFrame(SEno, columns=['compound_id'], dtype='str')
    SEno = SEno.astype('int')
    SEno = pd.merge(SEno, TCMID_Ingredient)
    num = i+1
    SEno.to_excel('%s.xlsx' %('TY'+str(num)))
    

import glob
glob.glob(r"C:\Users\physiology5\Dropbox\연구\1709_사상본초\171107_사상본초\171017_본초개별파일\*")
filename = glob.glob(r"C:\Users\physiology5\Dropbox\연구\1709_사상본초\171107_사상본초\171017_본초개별파일\*")




for i in range(0,8):
    file= filename[i]
    readexcel_merge(file,i)



readexcel_merge(filename[2],2) #연습

               
               
               
               
               
#########180303 추가

import os
os.chdir(r"C:\Users\physiology5\Desktop\본초개별파일+최종정리엑셀-20180228\본초개별파일-TCMID-20180228")


TCMID_Ingredient= pd.read_excel(r'C:\Users\physiology5\Dropbox\_연구\1709_사상본초\171107_사상본초파일\TCMID_Ingredient.xlsx', sheetname='Sheet1')


def readexcel_merge_2(file,i):
    SEno = pd.read_excel(file, header=None)
    SEno = SEno.iloc[:,0] 
    SEno = SEno.str.extract('([0-9]+)')
    SEno = SEno.astype('int')
    SEno = pd.DataFrame({'compound_id': SEno})
    SEno = pd.merge(SEno, TCMID_Ingredient)
    num = i+1
    SEno.to_excel('%s.xlsx' %('PLUS'+str(num)))
    
    
import glob
filename = glob.glob(r"C:\Users\physiology5\Desktop\본초개별파일+최종정리엑셀-20180228\본초개별파일-TCMID-20180228\1.새로 Coumpound연결후 각체질폴더에 합산할 5개\*")



for i in range(0,4):
    file= filename[i]
    readexcel_merge_2(file,i)






##########180313 TCMID 2.0 이용

TCMID_2= pd.read_excel(r'C:\Users\physiology5\Dropbox\_연구\1709_사상본초\180312_TCMID_개별컴파운드\TCMID2.0.xlsx', sheetname='Sheet1')


import glob
filename = glob.glob(r"C:\Users\physiology5\Dropbox\_연구\1709_사상본초\180312_TCMID_개별컴파운드\태양인\*") ##

import os
os.chdir(r'C:\Users\physiology5\Dropbox\_연구\1709_사상본초\180313_TCMID_2_본초2smiles\태양') ##


def readexcel_merge_3(file,i):
    SEno = pd.read_excel(file, sheetname='Sheet1')
    compound_id = SEno.iloc[:,0] 
    compound_id = pd.DataFrame(compound_id, columns=['compound_id'], dtype='str')
    compound_id = compound_id.astype('int')
    SEno = pd.merge(compound_id, TCMID_2)
    num = i+1
    SEno.to_excel('%s.xlsx' %('TY'+str(num)))  ##


for i in range(len(filename)): 
    file= filename[i]
    readexcel_merge_3(file,i)



















