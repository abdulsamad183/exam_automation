# -*- coding: utf-8 -*-
"""
Created on Tue Aug 23 20:13:19 2022

@author: Samad
"""

import pandas as pd
import os 
import msoffcrypto
import io


def check_marks(k,a):
    a=[str(n).strip().upper() for n in a]
    for an in range(len(a)):
        if a[an]=="Answers":
            a[an]='A'
    ab=''.join(a)
    d={}
    marks=0
    sec1=0
    sec2=0
    sec3=0
    sec4=0
    sec5=0
    for i in range(len(k)):
        if k[i]==a[i]:
            marks+=1
            if i>=0 and i<10:
                sec1+=1
                continue
            elif i>=10 and i<20:
                sec2+=1
                continue
            elif i>=20 and i<30:
                sec3+=1
                continue
            elif i>=30 and i<40:
                sec4+=1
                continue
            else:
                sec5+=1
                continue
    d={"Answers":ab,'Sec1':sec1,'Sec2':sec2,'Sec3':sec3,'Sec4':sec4,'Sec5':sec5,"Total":marks}
    return d



key=['B', 'C', 'C', 'A', 'D', 'C', 'B', 'D', 'A', 'B', 'D', 'A', 'D', 'C', 'A', 'A', 'B', 'D', 'C', 'B', 'B', 'D', 'A', 'B', 'C', 'D', 'A', 'D', 'C', 'B', 'D', 'D', 'C', 'A', 'D', 'B', 'A', 'A', 'C', 'C', 'A', 'B', 'A', 'A', 'D', 'B', 'C', 'A', 'D', 'B']
di={}


files=os.listdir()
file=[i for i in files if i[-5:]==".xlsx"]
for i in file:
    print("Completed -",i)
    temp = io.BytesIO()
    id_num=i[:-5]
    with open(i, 'rb') as f:
        excel = msoffcrypto.OfficeFile(f)
        excel.load_key('RLRKV2023')
        excel.decrypt(temp)
        df = pd.read_excel(temp)
        ans=df.iloc[:,18].tolist()
        #print(ans)
        dic=check_marks(key, ans)
        di[id_num]=dic
        
        
pd.DataFrame(di).T.to_csv('final_results.csv')










