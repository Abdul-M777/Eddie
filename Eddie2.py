import numpy as np
import pandas as pd
import itertools as it
from time import time
start_time = time()

matchings = 5
output_name = 'output'

df_own = pd.read_excel('Kontoudtog Dental Suite (eget bogføringssystem).xlsx')
#df_own = pd.read_excel('Egen.xlsx')
date_own = np.asarray(df_own['Dato'])
text_own = np.asarray(df_own['Beskrivelse'])
#text_own = np.asarray(df_own['Tekst'])
amount_own = np.asarray(df_own['Beløb'])
N_own = len(amount_own)
print('length own: %d'%N_own)

df_bank = pd.read_excel('Kontoudtog Ringkjøbing Landbobank.xlsx')
#df_bank = pd.read_excel('Bank.xlsx')
date_bank = np.asarray(df_bank['Dato'])
text_bank = np.asarray(df_bank['Tekst'])
amount_bank = np.asarray(df_bank['Beløb'])
N_bank = len(amount_bank)
print('length bank: %d'%N_bank)

#amount_own = np.array([1,2,3,12,12,25,15,15,16,5,7,40])
sum_own_1 = sum(amount_own)
#amount_bank = np.array([3,1,2,3,5,7,4,12,18,30,17,5,41])
sum_bank_1 = sum(amount_bank)
missIds_own = np.array(range(N_own))
missIds_bank = np.array(range(N_bank))

def remove_id(idx,which):
    '''remove idx from the global missing ids, and update global variables
    which: String; 'own' or 'bank': which account to remove ids from '''
    if which == 'own':
        global missIds_own, N_own
        index = np.nonzero(missIds_own == idx)[0][0]
        if index == len(missIds_own)-1:
            missIds_own = missIds_own[:index]
        else:
            missIds_own = np.concatenate((missIds_own[:index],missIds_own[index+1:]))
    elif which == 'bank':
        global missIds_bank, N_bank
        index = np.nonzero(missIds_bank == idx)[0][0]
        if index == len(missIds_bank)-1:
            missIds_bank = missIds_bank[:index]
        else:
            missIds_bank = np.concatenate((missIds_bank[:index],missIds_bank[index+1:]))
    else:
        raise ValueError('type either "own" or "bank"')    
    return None
    
def match(n):
    '''update global variables, matching both forwards and backwards
    n: number of matches to make'''
    
    global amount_bank, amount_own, missIds_own, missIds_bank
    
    if missIds_own is not None:
        for idx1,num1 in zip(missIds_own,amount_own[missIds_own]):
            if missIds_bank is not None:
                for idx2s,num2 in zip(it.combinations(missIds_bank,n),it.combinations(amount_bank[missIds_bank],n)):
                    #if any(j not in missIds_bank for j in idx2s):
                    #    continue # skip if ids have been removed from missIds2 since combinations were made
                    num2 = sum(num2)
                    if num1 == num2:
                        remove_id(idx1,'own')
                        for idx2 in idx2s:
                            remove_id(idx2,'bank')
                        break
                    
    # doing the same thing with own and bank switched
    if missIds_bank is not None:
        for idx1,num1 in zip(missIds_bank,amount_bank[missIds_bank]):
            if missIds_own is not None:
                for idx2s,num2 in zip(it.combinations(missIds_own,n),it.combinations(amount_own[missIds_own],n)):
                    #if any(j not in missIds_own for j in idx2s):
                    #    continue # skip if ids have been removed from missIds2 since combinations were made
                    num2 = sum(num2)
                    if num1 == num2:
                        remove_id(idx1,'bank')
                        for idx2 in idx2s:
                            remove_id(idx2,'own')
                        break
    return None

for n in range(1, matchings+1):
    match(n) # update missing ids

if missIds_own is None:
    missIds_own = []
if missIds_bank is None:
    missIds_bank = []

print('no. of matchings: %d'%matchings)
date_own = date_own[missIds_own]
text_own = text_own[missIds_own]
amount_own = amount_own[missIds_own]
print('length own: %d'%len(amount_own))

date_bank = date_bank[missIds_bank]
text_bank = text_bank[missIds_bank]
amount_bank = amount_bank[missIds_bank]
print('length bank: %d'%len(amount_bank))

sum_own_2 = sum(amount_own)
print('sum own subtracted: %.1f'%(sum_own_1-sum_own_2))
sum_bank_2 = sum(amount_bank)
print('sum bank subtracted: %.1f'%(sum_bank_1-sum_bank_2))

d1 = {'Dato': date_bank, 'Tekst': text_bank, 'Beløb': amount_bank}
d2 = {'Dato': date_own, 'Tekst': text_own, 'Beløb': amount_own}
df1 = pd.DataFrame(data=d1,columns=['Dato', 'Tekst', 'Beløb'])
df2 = pd.DataFrame(data=d2,columns=['Dato', 'Tekst', 'Beløb'])
writer = pd.ExcelWriter('%s.xlsx'%output_name)
df1.to_excel(writer,'Bank')
df2.to_excel(writer,'Eget')
writer.save()
print('%s.xlsx successfully saved!'%output_name)
stop_time = time()
dur = stop_time-start_time
if dur < 60:
    print('time: %.0f s'%dur)
else:
    print('time: %.0f min'%(dur/60.))