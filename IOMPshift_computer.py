import pandas as pd
import random
import numpy as np
from calendar import monthrange

names_CTD = ['anne', 'camille', 'claud', 'edch', 'eunice', 'lem', 'mikee', 'momay', 'pau']
names_CTSD = ['harry', 'pati']
names_CTSS = ['biboy']
names_MT = ['carlo', 'leo', 'marvin', 'meryll']
names_Data = ['anj', 'cath', 'marj', 'mart', 'morgan', 'nathan', 'roy']
names_Data += ['brain', 'kate', 'kennex', 'sky']
names_Data2 = ['arnel', 'earl', 'zhey', 'tinb']
names_D = ['junsat']
names_S = ['ivy', 'john', 'kevin', 'prado']
names_A = ['amy', 'ardeth', 'daisy', 'tinc']
names_S1 = ['jec', 'oscar', 'reyn', 'rodney', 'web']

no_shift = pd.read_csv('Monitoring Shift Schedule - fieldwork.csv')
no_shift['ts'] = pd.to_datetime(no_shift['ts'])
no_shift['name'] = no_shift['name'].apply(lambda x: x.lower())
no_shift = no_shift.set_index('name')

year = int(raw_input('year (e.g. 2017): '))
month = int(raw_input('month (1 to 12): '))

date = pd.to_datetime(str(year)+'-'+str(month)).strftime('%b%Y')

totalCT = int(monthrange(year, month)[1])*2
totalMT = totalCT

#days in field
df = pd.DataFrame({'name': names_CTD + names_CTSD + names_CTSS+ names_MT + names_D \
        + names_Data + names_Data2 + names_S + names_A + names_S1, \
        'team': ['CTD']*len(names_CTD) + ['CTSD']*len(names_CTSD) + ['CTSS']*len(names_CTSS) \
        + ['MT']*len(names_MT) + ['Data2']*len(names_D) + ['Data']*len(names_Data) \
        + ['Data2']*len(names_Data2) + ['SWAT']*len(names_S) + ['Admin']*len(names_A) + ['S1']*len(names_S1)})
df = df.set_index('name')
startTS = pd.to_datetime(str(year)+'-'+str(month))
if month != 12:
    endTS = pd.to_datetime(str(year)+'-'+str(month+1))
else:
    endTS = pd.to_datetime(str(year+1)+'-'+str(1))
field = []
for i in df.index:
    field += [len(no_shift[(no_shift.ts >= startTS)&(no_shift.ts < endTS)&(no_shift.index == i)].drop_duplicates())]
df['field'] = field
df['field_load'] = np.round(df['field']/7.)
df['field_load'] = df['field_load'].apply(lambda x: int(x))

#CT shifts of Admin
dfA = df.loc[df.team == 'Admin']
dfA['CTshift'] = 1
totalCT -= len(dfA)
dfA['MTshift'] = 0

#CT shifts of SRS1
dfS1 = df.loc[(df.team == 'S1')&(df.index != 'web')]
dfS1['CTshift'] = 1
totalCT -= sum(dfS1.CTshift)
dfS1['MTshift'] = 1
totalMT -= sum(dfS1.MTshift)
new_dfS1 = df.loc[df.index == 'web']
new_dfS1['CTshift'] = 2
totalCT -= sum(new_dfS1.CTshift)
new_dfS1['MTshift'] = 0
totalMT -= sum(new_dfS1.MTshift)
dfS1 = dfS1.append(new_dfS1)

df = df[(df.team != 'Admin')&(df.team != 'S1')]

EqualLoad = int(totalCT + totalMT + sum(df.field_load))/len(df)

#CT personnel
CT = names_CTD + names_CTSD + names_CTSS
#non-CT personnel
SD = names_MT + names_D + names_Data + names_Data2 + names_S
#MT, Data, Data2 -> for excess MT shifts
MT = names_MT
Data = names_Data
Data2 = names_Data2
#SWAT and Sir Jun
SD2 = names_D + names_S

################################## CT SHIFTS ###################################
#CT shifts (CT, Pati and Harry)
dfCT = df.loc[(df.team == 'CTD')|(df.team == 'CTSD')]
dfCT['CTshift'] = EqualLoad - dfCT['field_load']
dfCT['MTshift'] = 0

#CT shift(s) of non-CT + Biboy
CTexcess = totalCT - sum(dfCT['CTshift'])
if CTexcess <= len(SD):
    CTSDshifts = CTexcess
else:
    for i in [2,3]:
        if CTexcess <= i * (len(SD) + 1):
            break
    CTSDshifts = CTexcess - (i-1) * (len(SD) + 1)
#random non-CT with CT shift
CTSDlst = [df.loc[df.team == 'CTSS'].index[0]]
for i in range(1, int(CTSDshifts)):
    while True:
        randomCT = random.choice(SD)
        if randomCT in CTSDlst:
            continue
        else:
            break
    CTSDlst += [randomCT]
#CT shift(s)
# 1 or 0 shifts (random; except biboy:1)
if CTexcess <= len(SD):
    dfCT_SD = df.loc[df.index.isin(sorted(set(SD) - set(CTSDlst)))]
    dfCT_SD['CTshift'] = 0
    dfCT = dfCT.append(dfCT_SD)
    dfCT_SDe = df.loc[df.index.isin(CTSDlst)]
    dfCT_SDe['CTshift'] = 1
    dfCT = dfCT.append(dfCT_SDe)
# 2 or 1 shifts (random; except biboy:2)
else:
    for i in [2,3]:
        if CTexcess <= i * (len(SD) + 1):
            break
    dfCT_SD = df.loc[df.index.isin(sorted(set(SD) - set(CTSDlst)))]
    dfCT_SD['CTshift'] = i - 1
    dfCT = dfCT.append(dfCT_SD)
    dfCT_SDe = df.loc[df.index.isin(CTSDlst)]
    dfCT_SDe['CTshift'] = i
    dfCT = dfCT.append(dfCT_SDe)
#total load = CTshift and field load
dfCT['load'] = dfCT['field_load'] + dfCT['CTshift']

################################ CT & MT SHIFTS ################################
#equal MT shifts
dfMT = dfCT.loc[(dfCT.team != 'CTD')&(dfCT.team != 'CTSD')]
dfCT = dfCT.loc[(dfCT.team == 'CTD')|(dfCT.team == 'CTSD')]

################################## MT SHIFTS ###################################
#equal shifts
dfMT['load'] = dfMT['field_load'] + dfMT['CTshift']
dfMT['MTshift'] = EqualLoad - dfMT['load']
dfMT['load'] = dfMT['MTshift'] + dfMT['field_load'] + dfMT['CTshift']

#load based on field days and CT shifts
MTexcess = int(totalMT - sum(dfMT['MTshift']))
minload = min(set(dfMT.load))
dfMTmin = dfMT.loc[dfMT.load == minload]
#additional shifts to personnel with load < EqualLoad
while (minload < EqualLoad) and (MTexcess >= len(dfMTmin)):
    dfMTmin = dfMT.loc[dfMT.load == minload]
    dfMTmin['MTshift'] = dfMTmin['MTshift'] + 1
    dfMT = dfMT.loc[dfMT.load != minload]
    dfMT = dfMT.append(dfMTmin)
    dfMT['load'] = dfMT['MTshift'] + dfMT['field_load'] + dfMT['CTshift']
    MTexcess -= len(dfMTmin)
    minload = min(set(dfMT.load))
    dfMTmin = dfMT.loc[dfMT.load == minload]

while minload < EqualLoad:
    MTlst = []
    #additional shift to MT
    if MTexcess >= len(sorted(dfMTmin[dfMTmin.index.isin(MT)].index)):
        MTlst += sorted(dfMTmin[dfMTmin.index.isin(MT)].index)
        MTexcess -= len(MTlst)
        #additional shift to Data
        if MTexcess >= len(sorted(dfMTmin[dfMTmin.index.isin(Data)].index)):
            MTlst += sorted(dfMTmin[dfMTmin.index.isin(Data)].index)
            datalst = set(MTlst) - set(Data)
            MTexcess -= MTexcess - (len(MTlst) - len(datalst))
            #additional shift to Data2
            if MTexcess >= len(sorted(dfMTmin[dfMTmin.index.isin(Data2)].index)):
                MTlst += sorted(dfMTmin[dfMTmin.index.isin(Data2)].index)
                data2lst = set(MTlst) - set(Data) - set(Data2)
                MTexcess -= MTexcess - (len(MTlst) - len(data2lst))
                #additional shift to SD2
                if MTexcess >= len(sorted(dfMTmin[dfMTmin.index.isin(SD2)].index)):
                    dfMTmin['MTshift'] = dfMTmin['MTshift'] + 1
                    dfMTmin['load'] = dfMTmin['MTshift'] + dfMTmin['field_load'] + dfMTmin['CTshift']
                    dfMT = dfMT[~dfMT.index.isin(dfMTmin.index)]
                    dfMT = dfMT.append(dfMTmin)
                    MTexcess = int(totalMT - sum(dfMT['MTshift']))
                    minload = min(set(dfMT.load))
                    dfMTmin = dfMT.loc[dfMT.load == minload]
                #additional shift to random SD2
                elif MTexcess > 0:
                    for i in range(MTexcess):
                        while True:
                            randomMT = random.choice(sorted(dfMTmin[dfMTmin.index.isin(SD2)].index))
                            if randomMT in MTlst:
                                continue
                            else:
                                break
                        MTlst += [randomMT]
                    dfMTmin = dfMT.loc[dfMT.index.isin(MTlst)]
                    dfMTmin['MTshift'] = dfMTmin['MTshift'] + 1
                    dfMTmin['load'] = dfMTmin['MTshift'] + dfMTmin['field_load'] + dfMTmin['CTshift']
                    dfMT = dfMT[~dfMT.index.isin(MTlst)]
                    dfMT = dfMT.append(dfMTmin)
                    break
                else:
                    break
            #additional shift to random Data2
            elif MTexcess > 0:
                for i in range(MTexcess):
                    while True:
                        randomMT = random.choice(sorted(dfMTmin[dfMTmin.index.isin(Data2)].index))
                        if randomMT in MTlst:
                            continue
                        else:
                            break
                    MTlst += [randomMT]
                dfMTmin = dfMT.loc[dfMT.index.isin(MTlst)]
                dfMTmin['MTshift'] = dfMTmin['MTshift'] + 1
                dfMTmin['load'] = dfMTmin['MTshift'] + dfMTmin['field_load'] + dfMTmin['CTshift']
                dfMT = dfMT[~dfMT.index.isin(MTlst)]
                dfMT = dfMT.append(dfMTmin)
                break
            else:
                break
        #additional shift to random Data
        elif MTexcess > 0:
            for i in range(MTexcess):
                while True:
                    randomMT = random.choice(sorted(dfMTmin[dfMTmin.index.isin(Data)].index))
                    if randomMT in MTlst:
                        continue
                    else:
                        break
                MTlst += [randomMT]
            dfMTmin = dfMT.loc[dfMT.index.isin(MTlst)]
            dfMTmin['MTshift'] = dfMTmin['MTshift'] + 1
            dfMTmin['load'] = dfMTmin['MTshift'] + dfMTmin['field_load'] + dfMTmin['CTshift']
            dfMT = dfMT[~dfMT.index.isin(MTlst)]
            dfMT = dfMT.append(dfMTmin)
            break
        else:
            break
    #additional shift to random MT
    elif MTexcess > 0:
        for i in range(MTexcess):
            while True:
                randomMT = random.choice(sorted(dfMTmin[dfMTmin.index.isin(MT)].index))
                if randomMT in MTlst:
                    continue
                else:
                    break
            MTlst += [randomMT]
        dfMTmin = dfMT.loc[dfMT.index.isin(MTlst)]
        dfMTmin['MTshift'] = dfMTmin['MTshift'] + 1
        dfMTmin['load'] = dfMTmin['MTshift'] + dfMTmin['field_load'] + dfMTmin['CTshift']
        dfMT = dfMT[~dfMT.index.isin(MTlst)]
        dfMT = dfMT.append(dfMTmin)
        break
    else:
        break

dfIOMP = dfCT.append(dfMT)
dfIOMP['MTshift'] = dfIOMP['MTshift'].apply(lambda x: int(x))
nonCT = sorted(dfMT[dfMT.CTshift != 0].index)
MTexcess = int(totalMT - sum(dfIOMP['MTshift']))

while MTexcess > 0:
    #additional shift to MT
    if MTexcess >= len(MT):
        dfIOMPmt = dfIOMP[dfIOMP.index.isin(MT)]
        dfIOMPmt['MTshift'] = dfIOMPmt['MTshift'] + 1
        dfIOMP = dfIOMP[~dfIOMP.index.isin(MT)]
        dfIOMP = dfIOMP.append(dfIOMPmt)
        MTexcess -= len(dfIOMPmt)
        #additional shift to CT
        if MTexcess >= len(CT):
            #exchange shift to CTSD
            if len(CT) >= len(CTSDlst):
                #exchange CT to MT shift for all CTSD
                dfIOMPctsd = dfIOMP[dfIOMP.index.isin(CTSDlst)]
                dfIOMPctsd['CTshift'] = dfIOMPctsd['CTshift'] - 1
                dfIOMPctsd['MTshift'] = dfIOMPctsd['MTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(CTSDlst)]
                dfIOMP = dfIOMP.append(dfIOMPctsd)
                #exchange CT to MT shift for random nonCT
                MTlst = []
                for i in range(len(CTSDlst), len(CT)):
                    while True:
                        randomMT = random.choice(nonCT)
                        if randomMT in MTlst:
                            continue
                        else:
                            break
                    MTlst += [randomMT]
                dfIOMPmt = dfIOMP[dfIOMP.index.isin(MTlst)]
                dfIOMPmt['CTshift'] = dfIOMPmt['CTshift'] - 1
                dfIOMPmt['MTshift'] = dfIOMPmt['MTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                dfIOMP = dfIOMP.append(dfIOMPmt)
                #additional CT shift to all CT
                dfIOMPct = dfIOMP[dfIOMP.index.isin(CT)]
                dfIOMPct['CTshift'] = dfIOMPct['CTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(CT)]
                dfIOMP = dfIOMPct.append(dfIOMP)
                MTexcess -= len(CT)
                #additional shift to Data
                if MTexcess >= len(Data):
                    dfIOMPd = dfIOMP[dfIOMP.index.isin(Data)]
                    dfIOMPd['MTshift'] = dfIOMPd['MTshift'] + 1
                    dfIOMP = dfIOMP[~dfIOMP.index.isin(Data)]
                    dfIOMP = dfIOMP.append(dfIOMPd)
                    MTexcess -= len(Data)
                    #additional shift to Data2
                    if MTexcess >= len(Data2):
                        dfIOMPd2 = dfIOMP[dfIOMP.index.isin(Data2)]
                        dfIOMPd2['MTshift'] = dfIOMPd2['MTshift'] + 1
                        dfIOMP = dfIOMP[~dfIOMP.index.isin(Data2)]
                        dfIOMP = dfIOMP.append(dfIOMPd2)
                        MTexcess -= len(Data2)
                        #additional shift to random SWAT or Sir Jun
                        if MTexcess > 0:
                            MTlst = []
                            for i in range(MTexcess):
                                while True:
                                    randomMT = random.choice(SD2)
                                    if randomMT in MTlst:
                                        continue
                                    else:
                                        break
                                MTlst += [randomMT]
                            dfIOMPsd2 = dfIOMP[dfIOMP.index.isin(MTlst)]
                            dfIOMPsd2['MTshift'] = dfIOMPsd2['MTshift'] + 1
                            dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                            dfIOMP = dfIOMP.append(dfIOMPsd2)
                        break
                    #additional shift to random Data2
                    else:
                        MTlst = []
                        for i in range(MTexcess):
                            while True:
                                randomMT = random.choice(Data2)
                                if randomMT in MTlst:
                                    continue
                                else:
                                    break
                            MTlst += [randomMT]
                        dfIOMPd2 = dfIOMP[dfIOMP.index.isin(MTlst)]
                        dfIOMPd2['MTshift'] = dfIOMPd2['MTshift'] + 1
                        dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                        dfIOMP = dfIOMP.append(dfIOMPd2)
                        break
                #additional shift to random Data
                else:
                    MTlst = []
                    for i in range(MTexcess):
                        while True:
                            randomMT = random.choice(Data)
                            if randomMT in MTlst:
                                continue
                            else:
                                break
                        MTlst += [randomMT]
                    dfIOMPd = dfIOMP[dfIOMP.index.isin(MTlst)]
                    dfIOMPd['MTshift'] = dfIOMPd['MTshift'] + 1
                    dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                    dfIOMP = dfIOMP.append(dfIOMPd)
                    break
            #exchange shift to random CTSD
            else:
                #exchange CT to MT shift for random CTSD
                MTlst = []
                for i in range(len(CT)):
                    while True:
                        randomMT = random.choice(CTSDlst)
                        if randomMT in MTlst:
                            continue
                        else:
                            break
                    MTlst += [randomMT]
                dfIOMPctsd = dfIOMP.loc[dfIOMP.index.isin(MTlst)]
                dfIOMPctsd['CTshift'] = dfIOMPctsd['CTshift'] - 1
                dfIOMPctsd['MTshift'] = dfIOMPctsd['MTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                dfIOMP = dfIOMP.append(dfIOMPctsd)
                #additional shift to all CT
                dfIOMPct = dfIOMP[dfIOMP.index.isin(CT)]
                dfIOMPct['CTshift'] = dfIOMPct['CTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(CT)]
                dfIOMP = dfIOMPct.append(dfIOMP)
                MTexcess -= len(CT)
                #additional shift to Data
                if MTexcess >= len(Data):
                    dfIOMPd = dfIOMP[dfIOMP.index.isin(Data)]
                    dfIOMPd['MTshift'] = dfIOMPd['MTshift'] + 1
                    dfIOMP = dfIOMP[~dfIOMP.index.isin(Data)]
                    dfIOMP = dfIOMP.append(dfIOMPd)
                    MTexcess -= len(Data)
                    #additional shift to Data2
                    if MTexcess >= len(Data2):
                        dfIOMPd2 = dfIOMP[dfIOMP.index.isin(Data2)]
                        dfIOMPd2['MTshift'] = dfIOMPd2['MTshift'] + 1
                        dfIOMP = dfIOMP[~dfIOMP.index.isin(Data2)]
                        dfIOMP = dfIOMP.append(dfIOMPd2)
                        MTexcess -= len(Data2)
                        #additional shift to random SWAT or Sir Jun
                        if MTexcess > 0:
                            MTlst = []
                            for i in range(MTexcess):
                                while True:
                                    randomMT = random.choice(SD2)
                                    if randomMT in MTlst:
                                        continue
                                    else:
                                        break
                                MTlst += [randomMT]
                            dfIOMPsd2 = dfIOMP[dfIOMP.index.isin(MTlst)]
                            dfIOMPsd2['MTshift'] = dfIOMPsd2['MTshift'] + 1
                            dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                            dfIOMP = dfIOMP.append(dfIOMPsd2)
                        break
                    #additional shift to random Data2
                    else:
                        MTlst = []
                        for i in range(MTexcess):
                            while True:
                                randomMT = random.choice(Data2)
                                if randomMT in MTlst:
                                    continue
                                else:
                                    break
                            MTlst += [randomMT]
                        dfIOMPd2 = dfIOMP[dfIOMP.index.isin(MTlst)]
                        dfIOMPd2['MTshift'] = dfIOMPd2['MTshift'] + 1
                        dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                        dfIOMP = dfIOMP.append(dfIOMPd2)
                        break
                #additional shift to random Data
                else:
                    MTlst = []
                    for i in range(MTexcess):
                        while True:
                            randomMT = random.choice(Data)
                            if randomMT in MTlst:
                                continue
                            else:
                                break
                        MTlst += [randomMT]
                    dfIOMPd = dfIOMP[dfIOMP.index.isin(MTlst)]
                    dfIOMPd['MTshift'] = dfIOMPd['MTshift'] + 1
                    dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                    dfIOMP = dfIOMP.append(dfIOMPd)
                    break
        #all additional shift to random CT
        else:
            #additional shift to random CT (exchange shift to CTSD and random non-CT)
            if MTexcess >= len(CTSDlst):
                #exchange CT to MT shift for all CTSD
                dfIOMPctsd = dfIOMP[dfIOMP.index.isin(CTSDlst)]
                dfIOMPctsd['CTshift'] = dfIOMPctsd['CTshift'] - 1
                dfIOMPctsd['MTshift'] = dfIOMPctsd['MTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(CTSDlst)]
                dfIOMP = dfIOMP.append(dfIOMPctsd)
                #exchange CT to MT shift for random nonCT
                MTlst = []
                for i in range(len(CTSDlst), MTexcess):
                    while True:
                        randomMT = random.choice(nonCT)
                        if randomMT in MTlst:
                            continue
                        else:
                            break
                    MTlst += [randomMT]
                dfIOMPmt = dfIOMP[dfIOMP.index.isin(MTlst)]
                dfIOMPmt['CTshift'] = dfIOMPmt['CTshift'] - 1
                dfIOMPmt['MTshift'] = dfIOMPmt['MTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                dfIOMP = dfIOMP.append(dfIOMPmt)
                #additional CT shift to CT
                CTlst = []
                for i in range(len(CTSDlst)):
                    while True:
                        randomCT = random.choice(CTSDlst)
                        if randomCT in CTlst:
                            continue
                        else:
                            break
                    CTlst += [randomCT]
                dfIOMPct = dfIOMP[dfIOMP.index.isin(CTlst)]
                dfIOMPct['CTshift'] = dfIOMPct['CTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(CTlst)]
                dfIOMP = dfIOMPct.append(dfIOMP)
                break
            #additional shift to random CT (exchange shift to random CTSD)
            else:
                MTlst = []
                for i in range(MTexcess):
                    while True:
                        randomMT = random.choice(CTSDlst)
                        if randomMT in MTlst:
                            continue
                        else:
                            break
                    MTlst += [randomMT]
                dfIOMPctsd = dfIOMP.loc[dfIOMP.index.isin(MTlst)]
                dfIOMPctsd['CTshift'] = dfIOMPctsd['CTshift'] - 1
                dfIOMPctsd['MTshift'] = dfIOMPctsd['MTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
                dfIOMP = dfIOMP.append(dfIOMPctsd)
                CTlst = []
                for i in range(MTexcess):
                    while True:
                        randomCT = random.choice(CT)
                        if randomCT in CTlst:
                            continue
                        else:
                            break
                    CTlst += [randomCT]
                dfIOMPct = dfIOMP[dfIOMP.index.isin(CTlst)]
                dfIOMPct['CTshift'] = dfIOMPct['CTshift'] + 1
                dfIOMP = dfIOMP[~dfIOMP.index.isin(CTlst)]
                dfIOMP = dfIOMPct.append(dfIOMP)
                break
    #additional shift to random MT
    else:
        MTlst = []
        for i in range(MTexcess):
            while True:
                randomMT = random.choice(MT)
                if randomMT in MTlst:
                    continue
                else:
                    break
            MTlst += [randomMT]
        dfIOMPmt = dfIOMP.loc[dfIOMP.index.isin(MTlst)]
        dfIOMPmt['MTshift'] = dfIOMPmt['MTshift'] + 1
        dfIOMP = dfIOMP[~dfIOMP.index.isin(MTlst)]
        dfIOMP = dfIOMP.append(dfIOMPmt)
        break

dfIOMP = dfIOMP.append(dfA).append(dfS1)
dfIOMP = dfIOMP.fillna(0)
dfIOMP['load'] = dfIOMP['MTshift'] + dfIOMP['field_load'] + dfIOMP['CTshift']

neg_MTshift = dfIOMP[dfIOMP.MTshift < 0]
if len(neg_MTshift) != 0:
    neg_MTshift['CTshift'] = neg_MTshift['CTshift'] + neg_MTshift['MTshift']
    neg_MTshift['MTshift'] = 0
    dfIOMP = dfIOMP[~dfIOMP.index.isin(neg_MTshift.index)]
    dfIOMP = dfIOMP.append(neg_MTshift)

neg_CTshift = dfIOMP[dfIOMP.CTshift < 0]
if len(neg_CTshift) != 0:
    neg_CTshift['MTshift'] = neg_CTshift['MTshift'] + neg_CTshift['CTshift']
    neg_CTshift['CTshift'] = 0
    dfIOMP = dfIOMP[~dfIOMP.index.isin(neg_CTshift.index)]
    dfIOMP = dfIOMP.append(neg_CTshift)

total_CTshift = dfIOMP['CTshift'].sum()
total_MTshift = dfIOMP['MTshift'].sum()
if total_CTshift < total_MTshift:
    no_CT = dfIOMP[~dfIOMP.team.isin(['S1', 'CTD', 'CTSD', 'CTSS', 'Admin'])&(dfIOMP.CTshift == 0)]
    chosen_CT_list = []
    while len(chosen_CT_list) != np.abs(total_CTshift - total_MTshift)/2:
        chosen_CT = random.choice(no_CT.index)
        if chosen_CT not in chosen_CT_list:
            chosen_CT_list += [chosen_CT]
    no_CT = no_CT[no_CT.index.isin(chosen_CT_list)]
    no_CT['CTshift'] += 1
    no_CT['MTshift'] -= 1
    dfIOMP = dfIOMP[~dfIOMP.index.isin(chosen_CT_list)].append(no_CT)
if total_MTshift < total_CTshift:
    with_CT = dfIOMP[~dfIOMP.team.isin(['S1', 'CTD', 'CTSD', 'CTSS', 'Admin'])&(dfIOMP.CTshift != 0)]
    chosen_MT_list = []
    while len(chosen_MT_list) != np.abs(total_MTshift - total_CTshift)/2:
        chosen_MT = random.choice(with_CT.index)
        if chosen_MT not in chosen_MT_list:
            chosen_MT_list += [chosen_MT]
    with_CT = with_CT[with_CT.index.isin(chosen_MT_list)]
    with_CT['MTshift'] += 1
    with_CT['CTshift'] -= 1
    dfIOMP = dfIOMP[~dfIOMP.index.isin(chosen_MT_list)].append(with_CT)

shiftdf = dfIOMP[['team', 'field', 'MTshift', 'CTshift', 'load']].sort_index()