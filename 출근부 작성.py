# -*- coding: utf-8 -*-
"""
Created on Wed Dec 30 09:33:46 2020

@author: LEEKWANGJIN
"""

import pandas as pd
import random
import os

from pandas import DataFrame


###############for문 실행 전 반드시 실행할 것###########################
Date_cnt = 0 # 일주일 당 근무 일자. 일반적으로 6(월 ~ 토 이므로)
Pre_idx = 0 # 이전 인덱스 값.
test_1 = 0

Tic_num = 3 # 열화상 근무 인원 수
Ofc_num = 2 # 사무실 근무 인원 수

Tic_time = 8 #열화상 근무 시간. 일반적으로 8 (9 ~ 18)
Ofc_time = 8 #사무실 근무 시간. 일반적으로 8 (9 ~ 18)

Tic_all_time = Tic_time * Tic_num
Ofc_all_time = Ofc_time * Ofc_num
All_time = Tic_all_time + Ofc_all_time
Work_time = 0

Time = ['09 ~ 10', '10 ~ 11', '11 ~ 12', '13 ~ 14', '14 ~ 15', '15 ~ 16', '16 ~ 17', '17 ~ 18','19 ~ 20','20 ~ 21']
Name = []
List = []
unique = []
Un_unique = []
Go = 1
Ofc_List_temp = []
Tic_List_temp = []
df_Ofc = DataFrame()
df_Tic = DataFrame()
df2 = DataFrame()
i = 0
j = 0
pre_i = 0
b= 0
k = 0
Non_work = []
######################################################################

excel = 'C:/Users/LEEKWANGJIN/Desktop/KW/3-2/algorithm/알고리즘/알고리즘/코드/근로장학생 01.18-1.23.xlsx'

df = pd.read_excel(excel)
df = df.drop(columns=['Number', 'Work', 'Name_1', 'Work_number'])
df = df.fillna(0)

for idx, date in enumerate(df.Date) :
    if date == 0 :
        df1 = df[Pre_idx:idx]
        df1.reset_index(inplace=True)
        del df1['index']
        Popul = df1.index[-1] + 1
        Work_time = All_time / Popul #월요일은 4번 들어감
        Tic_work_time = Tic_all_time / Popul
        Ofc_work_time = Ofc_all_time / Popul
        Name = df1['Name'].tolist()
        random.shuffle(Name)
        Ofc_List = Name[:]
        Tic_List = Name[:]
        random.shuffle(Tic_List)
        
        while b < Tic_time  :  
        
            try:
                Tic_List_temp.append(Tic_List[i])
                i += 1
        
            except:
                Tic_List_temp.append('Tic_random')
                
            if len(Tic_List_temp) == Tic_num :      
                for k in range(0, int(Tic_work_time)):
                    if b < Tic_time :
                        df_Tic[Time[b]] = Tic_List_temp
                    else : 
                        break
                    b += 1
                
                Tic_List_temp = []
        
        for i in range(0, Tic_time) :
            for j in range(0, Tic_num) :
                if df_Tic[Time[i]][j] == 'Tic_random' :
                    df_Tic[Time[i]][j] = Tic_List[k]
                    k += 1
                
        b = 0
        i = 0
        j = 0
        k = 0
        
        while b < Ofc_time  :  
        
            try:
                Ofc_List_temp.append(Ofc_List[i])
                i += 1
        
            except:
                Ofc_List_temp.append('Ofc_random')
            
            if len(Ofc_List_temp) == Ofc_num :
                for k in range(0, int(Ofc_work_time)):
                    if b < Ofc_time :
                        df_Ofc[Time[b]] = Ofc_List_temp
                    else : 
                        break
                    b += 1
                
                Ofc_List_temp = []
        
        for i in range(0, Ofc_time) :
            for j in range(0, Ofc_num) :
                if df_Ofc[Time[i]][j] == 'Ofc_random' :
                    df_Ofc[Time[i]][j] = Ofc_List[k]
                    k += 1
                    
        result1 = pd.concat([df_Tic, df_Ofc])
        
        for b in range(0, Tic_time) :
            unique = []
            for name in result1[Time[b]].tolist() :         # 1st loop
                if name not in unique:   # 2nd loop
                    unique.append(name)
                else : 
                    unique.append('다시 돌려주세요')
                    Un_unique.append(name)
            result1[Time[b]] = unique
            
        
                    
        for b in range(0, Tic_time) :
            result1_list = result1[Time[b]].tolist()
            for index, value in enumerate(result1_list) :
                if value == '다시 돌려주세요':
                    for name2 in Un_unique :
                        if name2 not in result1_list :
                            result1_list[index] = name2
            result1[Time[b]] = result1_list
                
        result1.reset_index(drop = True, inplace = True)
        
        for b in range(0, Tic_time) :
            result1_list = result1[Time[b]].tolist()
            for index, value in enumerate(result1_list) :
                if value == '다시 돌려주세요':
                    test_1 = 1
                    print("다시 돌려주세요!!!!!!!!!!!!!!!!!!!")
                    
        if test_1 == 0 :
            if not os.path.exists('근무표.xlsx'):
                with pd.ExcelWriter('근무표.xlsx', mode='w', engine='openpyxl') as writer:
                    result1.to_excel(writer, index=False, sheet_name = Day)
            else:
                with pd.ExcelWriter('근무표.xlsx', mode='a', engine='openpyxl') as writer:
                    result1.to_excel(writer, index=False, sheet_name = Day)
        else : 
            if os.path.isfile('근무표.xlsx'):
                os.remove('근무표.xlsx')
            
        Tic_List_temp = []
        Ofc_List_temp = []
        result1_list = []
        unique = []
        Un_unique = []
        b = 0
        i = 0
        j = 0
        k = 0
        df2 = df2.append(result1)
        Pre_idx = idx+1
        Date_cnt+=1
    
    else : Day = date


                    
