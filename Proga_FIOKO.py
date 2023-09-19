import openpyxl
import copy
import numpy as np
from openpyxl import load_workbook 
import pandas as pd

def mark(x, y):
    if y == '-':
        return '-'
    elif x / y < 5:
        return 2
    elif x / y < 10:
        return 3
    elif x / y < 14:
        return 4
    else:
        return 5
    

def new_sum(row, columns_list, mark_list):
    j = columns_list[0] + '_x'
    if row[j] != 'X' and np.isnan(row[j]):
        return 0
    elif row['num_not'] > 2 or row['are_marks_not_necessary?']:
        return '-'
    else:
        ans = 0
        for i in range(len(columns_list)):
            j = 'crit ' + columns_list[i]
            f = columns_list[i] + '_x'
            ans += (int(~row[j]) + 2) * mark_list[i]
        return(ans)


def get_mark(row, max_sum):
    if int(row['Итого баллов']) == 0 and row['New_sum'] == 0:
        return np.nan
    else:
        return mark(int(row['Итого баллов']) * max_sum , row['New_sum'])
    

def new_column(row, list):
    if row['Пользователь'] in list:
        return False
    else:
        return True


df = pd.read_excel('ФИ8_тест.xlsx', sheet_name='Протокол')
columns_list = df.columns.to_list()[3:-4]
death_list = df[df[columns_list[0]].isna()].index
df = df.drop(death_list, axis=0)
#print(df)
mark_list = [0 for i in range(len(columns_list))]
max_sum = 0
for i in range(len(columns_list)):
    mark_list[i] = int(columns_list[i][-3])
    max_sum += mark_list[i]
criteria1 = pd.DataFrame()
criteria2 = pd.DataFrame()

for i in columns_list:
    j = i + ' check'
    df[j] = (df[i] == 'не пройд.') 
    criteria2[i] = df.groupby(['Пользователь'])[j].mean()
    criteria2[i] = (criteria2[i] > 0.5)
    criteria1[i] = df.groupby(['Пользователь', 'Наименование класса'])[j].mean()
    criteria1[i] = (criteria1[i] >= 0.8)

#print(df.iloc[113])   
#print(criteria2)
criteria1 = criteria1.merge(criteria2,left_on=['Пользователь'], right_index=True, how='left')
spisok1 = []
for i in columns_list:
    j = i + '_x'
    k = i + '_y' 
    criteria1[i] = criteria1[j] + criteria1[k]
    spisok1.append(j)
    spisok1.append(k)
criteria1 = criteria1.drop(spisok1, axis=1) 
criteria2['num_not'] = criteria2.sum(axis=1)  
criteria3 = criteria1.groupby('Пользователь').value_counts().index.to_frame(index=False)['Пользователь'].value_counts()
criteria3 = criteria3.loc[criteria3 == 1].index.to_list()
criteria2.reset_index(inplace=True)
criteria2['are_marks_not_necessary?'] = criteria2.apply(lambda x : new_column(x, criteria3), axis=1)
criteria2 = criteria2.set_index('Пользователь')

df = df.merge(criteria1, left_on=['Пользователь', 'Наименование класса'], right_on=['Пользователь', 'Наименование класса'], how='left')
df = df.merge(criteria2, left_on=['Пользователь'], right_on=['Пользователь'])
for i in columns_list:
    j = 'crit ' + i
    f = i + '_y'
    df[j] = (df[i] | df[f]) 

df['New_sum'] = df.apply(lambda x : new_sum(x, columns_list, mark_list), axis=1)
df['Оценка'] = df.apply(lambda x : get_mark(x, max_sum), axis=1)
# print(df['New_sum'].value_counts())
# j = columns_list[0] + '_x'
#print(df['num_not'].value_counts())

spisok = list(criteria2.columns)
spisok.append('New_sum')
for i in columns_list:
    j = i + ' check'
    k = i + '_y'
    s = 'crit ' + i
    spisok.append(j)
    spisok.append(i)
    spisok.append(k)
    spisok.append(s)
df = df.drop(spisok, axis=1)
df.columns = df.columns.str.replace('_x', '')
df.to_excel('Ответ.xlsx',index=False)
