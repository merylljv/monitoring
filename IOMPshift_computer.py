import pandas as pd
import random
import numpy as np
from calendar import monthrange

def shift_division(df, with_excess_shift, equal_load):
    name = df['name'].values[0]
    field = df['field_load'].values[0]
    shift = equal_load - field + (name in with_excess_shift)*1
    df['MTshift'] = int(shift/2)
    df['CTshift'] = shift - int(shift/2)
    return df
    
def load_comp(df, with_excess_shift, equal_load):
    name = df['name'].values[0]
    df['load'] = equal_load + (name in with_excess_shift)*1
    return df

names_CTD = ['anne', 'camille', 'eunice', 'lem', 'momay', 'pau', 'harry', 'pati']
names_CTS = ['biboy']
names_MT = ['carlo', 'leo', 'marvin', 'meryll']
names_Data = ['anj', 'cath', 'mart', 'morgan', 'nathan', 'roy']
names_Data += ['brain', 'kate', 'kennex', 'sky']
names_Data2 = ['arnel', 'earl']
names_D = ['junsat']
names_S = ['ivy', 'john', 'kevin', 'prado']
names_A = ['amy', 'ardeth', 'daisy', 'tinc']
names_S1 = ['jec', 'oscar', 'reyn', 'rodney', 'web']

no_shift = pd.read_csv('Monitoring Shift Schedule - fieldwork.csv')
no_shift['ts'] = pd.to_datetime(no_shift['ts'])
no_shift['name'] = no_shift['name'].apply(lambda x: x.lower())

year = int(raw_input('year (e.g. 2017): '))
month = int(raw_input('month (1 to 12): '))

date = pd.to_datetime(str(year)+'-'+str(month)).strftime('%b%Y')

totalCT = int(monthrange(year, month)[1])*2
totalMT = totalCT

##### days in field #####
df = pd.DataFrame({'name': names_CTD + names_CTS+ names_MT + names_D \
        + names_Data + names_Data2 + names_S + names_A + names_S1, \
        'team': ['CTD']*len(names_CTD) + ['CTS']*len(names_CTS) \
        + ['MT']*len(names_MT) + ['Data2']*len(names_D) + ['Data']*len(names_Data) \
        + ['Data2']*len(names_Data2) + ['SWAT']*len(names_S) + ['Admin']*len(names_A) + ['S1']*len(names_S1)})
startTS = pd.to_datetime(str(year)+'-'+str(month))
if month != 12:
    endTS = pd.to_datetime(str(year)+'-'+str(month+1))
else:
    endTS = pd.to_datetime(str(year+1)+'-'+str(1))
field = []
for i in df.name:
    field += [len(no_shift[(no_shift.ts >= startTS)&(no_shift.ts < endTS)&(no_shift.name == i)].drop_duplicates())]
df['field'] = field
df['field_load'] = np.round(df['field']/7.)
df['field_load'] = df['field_load'].apply(lambda x: int(x))

##### CT shifts of Admin #####
dfA = df.loc[df.team == 'Admin']
dfA['CTshift'] = 1
totalCT -= sum(dfA.CTshift)
dfA['MTshift'] = 0

##### CT shifts of SRS1 #####
dfS1 = df.loc[(df.team == 'S1')&(df.name != 'web')]
dfS1['CTshift'] = 1
totalCT -= sum(dfS1.CTshift)
dfS1['MTshift'] = 1
totalMT -= sum(dfS1.MTshift)
new_dfS1 = df.loc[df.name == 'web']
new_dfS1['CTshift'] = 2
totalCT -= sum(new_dfS1.CTshift)
new_dfS1['MTshift'] = 0
totalMT -= sum(new_dfS1.MTshift)
dfS1 = dfS1.append(new_dfS1)

df = df[(df.team != 'Admin')&(df.team != 'S1')]

##### equal load: total CT shift, MT shift, field #####
equal_load = int(totalCT + totalMT + sum(df.field_load))/len(df)

##### personnel with additional shifts #####
excess_shift = int(totalCT + totalMT + sum(df.field_load)) % len(df)

#CT personnel
CT = names_CTD + names_CTS
#non-CT personnel
SD = names_MT + names_D + names_Data + names_Data2 + names_S
#MT, Data, Data2 -> for excess MT shifts
MT = names_MT
Data = names_Data
Data2 = names_Data2
#SWAT and Sir Jun
SD2 = names_D + names_S

with_excess_shift = []
# excess shifts for MT and CT
if excess_shift >= len(CT + MT):
    with_excess_shift = CT + MT
    excess_shift -= len(CT + MT)
    # excess shifts for Data
    if excess_shift >= len(Data):
        with_excess_shift += Data
        excess_shift -= len(Data)
        # excess shifts for Data2
        if excess_shift >= len(Data2):
            with_excess_shift += Data2
            excess_shift -= len(Data2)
            # excess shifts for Data
            if excess_shift >= 0:
                for count in range(excess_shift):
                    while True:
                        random_personnel = random.choice(SD2)
                        if random_personnel in with_excess_shift:
                            continue
                        else:
                            break
                    with_excess_shift += [random_personnel]
        # excess shifts for chosen Data2 team
        else:
            for count in range(excess_shift):
                while True:
                    random_personnel = random.choice(Data2)
                    if random_personnel in with_excess_shift:
                        continue
                    else:
                        break
                with_excess_shift += [random_personnel]
    # excess shifts for chosen Data team
    else:
        for count in range(excess_shift):
            while True:
                random_personnel = random.choice(Data)
                if random_personnel in with_excess_shift:
                    continue
                else:
                    break
            with_excess_shift += [random_personnel]
# excess shifts for chosen MT and/or CT    
else:
    for count in range(excess_shift):
        while True:
            random_personnel = random.choice(CT + MT)
            if random_personnel in with_excess_shift:
                continue
            else:
                break
        with_excess_shift += [random_personnel]

################################## CT SHIFTS ###################################

##### Community team shifts #####
dfCT = df.loc[(df.team == 'CTD')]
dfCT['CTshift'] = equal_load - dfCT['field_load']
dfCT['MTshift'] = [0]*len(dfCT)
excess_dfCT = dfCT[dfCT.name.isin(with_excess_shift)]
excess_dfCT['CTshift'] = excess_dfCT['CTshift'] + 1
dfCT = dfCT[~dfCT.name.isin(with_excess_shift)]
dfCT = dfCT.append(excess_dfCT)
dfCTS = df.loc[(df.team == 'CTS')]
dfCTS_grp = dfCTS.groupby('name', as_index=False)
dfCTS = dfCTS_grp.apply(shift_division, with_excess_shift=with_excess_shift, equal_load=equal_load)
dfCT = dfCT.append(dfCTS)
dfCT['load'] = dfCT['field_load'] + dfCT['CTshift'] + dfCT['MTshift']

##### CT shifts for non-CT personnel #####
total_nonCT = totalCT - dfCT['CTshift'].sum()
nonCT_CTshift = []
for count in range(total_nonCT):
    while True:
        random_personnel = random.choice(SD)
        if random_personnel in nonCT_CTshift:
            continue
        else:
            break
    nonCT_CTshift += [random_personnel]

################################## MT SHIFTS ###################################
dfMT = df = df[(df.team != 'CTD')&(df.team != 'CTS')]
dfMT_grp = dfMT.groupby('name', as_index=False)
dfMT = dfMT_grp.apply(load_comp, with_excess_shift=with_excess_shift, equal_load=equal_load)
with_CTshift = dfMT[dfMT.name.isin(nonCT_CTshift)]
with_CTshift['CTshift'] = 1
dfMT = dfMT[~dfMT.name.isin(nonCT_CTshift)]
dfMT['CTshift'] = 0
dfMT = dfMT.append(with_CTshift)
dfMT['MTshift'] = dfMT['load'] - dfMT['field_load'] - dfMT['CTshift'] 

################################################################################

dfIOMP = dfCT.append(dfMT).append(dfA).append(dfS1)
dfIOMP['MTshift'] = dfIOMP['MTshift'].apply(lambda x: int(x))
dfIOMP['CTshift'] = dfIOMP['CTshift'].apply(lambda x: int(x))
dfIOMP = dfIOMP.fillna(0)
dfIOMP['load'] = dfIOMP['MTshift'] + dfIOMP['field_load'] + dfIOMP['CTshift']
shiftdf = dfIOMP.set_index('name')
shiftdf = shiftdf[['team', 'field', 'MTshift', 'CTshift', 'load']].sort_index()