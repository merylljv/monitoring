import pandas as pd
import random
from datetime import time, datetime
from calendar import monthrange

import IOMPshift_computer as scomp

#Month and Year of Shift Schedule
year = scomp.year
month = scomp.month

if month == 2:
    WAshift = 3
else:
    WAshift = 2

#Number of Shifts per Personnel
date = scomp.date
shift_count = scomp.shiftdf

#Empty Shift Schedule
last_shift = int(monthrange(year, month)[1])
startTS = pd.to_datetime(str(year) + '-' + str(month) + '-01 07:30')
endTS = pd.to_datetime(str(year) + '-' + str(month) + '-' + str(last_shift) + ' 19:30')
shiftTS = pd.date_range(start=startTS, end=endTS, freq='12H')
shiftdf = pd.DataFrame({'ts': shiftTS, 'IOMP-MT': ['?']*len(shiftTS), 'IOMP-CT': ['?']*len(shiftTS)})
shiftdf = shiftdf.set_index('ts')

#Fieldwork Schedule
no_shift = scomp.no_shift

################################### MT Shift ###################################
#fills from top to bottom

MTlst = []
for i in sorted(set(shift_count[shift_count.MTshift != 0].index)):
    MTlst += [i]*shift_count[shift_count.index == i]['MTshift'].values[0]

MTset = sorted(set(MTlst))

#First set of MT shifts (equal to number of personnel with MT shifts)
for i in range(len(MTset)):
    while True:
        #random personnel to take MT shift
        randomMT = random.choice(MTset)
        #checks if has fieldwork during that day
        no_shift_lst = no_shift[no_shift.ts == shiftdf.index[i].date()].index
        if randomMT not in no_shift_lst:
            break
        else:
            continue
    shiftdf.loc[shiftdf.index[i]]['IOMP-MT'] = randomMT
    MTset.remove(randomMT)
    MTlst.remove(randomMT)

#Remaining shifts
lastshift = len(shiftdf.loc[shiftdf['IOMP-MT'] != '?'])
while lastshift < len(shiftdf):
    MTset = list(shiftdf.loc[(shiftdf['IOMP-MT'] != '?')&(shiftdf['IOMP-MT'].isin(MTlst))].drop_duplicates(['IOMP-MT'])['IOMP-MT'].values)    
    for i in range(lastshift,lastshift+len(MTset)):
        while True:
            #random personnel to take MT shift
            randomMT = random.choice(MTset)
            #checks if has fieldwork during that day
            no_shift_lst = no_shift[no_shift.ts == shiftdf.index[i].date()].index
            #checks if has shift before and after
            prev_shift_lst = [shiftdf.loc[shiftdf.index[i-1]]['IOMP-MT']]
            #if PM shifts, checks if with PM shift before and after
            if shiftdf.index[i].time() == time(19,30):
                prev_shift_lst += [shiftdf.loc[shiftdf.index[i-2]]['IOMP-MT']]
            if randomMT not in no_shift_lst and randomMT not in prev_shift_lst:
                break
            else:
                continue
        shiftdf.loc[shiftdf.index[i]]['IOMP-MT'] = randomMT
        MTlst.remove(randomMT)
        MTset.remove(randomMT)    
    lastshift = len(shiftdf.loc[shiftdf['IOMP-MT'] != '?'])

################################### CT Shift ###################################
#fills from bottom to top

CTlst = []
for i in sorted(set(shift_count[shift_count.CTshift != 0].index)):
    CTlst += [i]*shift_count[shift_count.index == i]['CTshift'].values[0]

CTset = sorted(shift_count[shift_count['CTshift'] > 2].index) * 2 \
    + scomp.names_A + scomp.names_S1

#First set of MT shifts (equal to number of personnel with CT shifts)
for i in range(len(shiftdf) - len(CTset), len(shiftdf)):
    while True:
        #random personnel to take CT shift
        randomCT = random.choice(CTset)
        #checks if has fieldwork during that day
        no_shift_lst = no_shift[no_shift.ts == shiftdf.index[i].date()].index
        #checks if has shift before and after
        prev_shift_lst = [shiftdf.loc[shiftdf.index[i]]['IOMP-MT']]
        prev_shift_lst += [shiftdf.loc[shiftdf.index[i-1]]['IOMP-MT']]
        prev_shift_lst += [shiftdf.loc[shiftdf.index[i-1]]['IOMP-CT']]
        try:
            prev_shift_lst += [shiftdf.loc[shiftdf.index[i+1]]['IOMP-MT']]
            prev_shift_lst += [shiftdf.loc[shiftdf.index[i+1]]['IOMP-CT']]
        except:
            pass
        #if PM shifts, checks if with PM shift before and after
        if shiftdf.index[i].time() == time(19,30):
            prev_shift_lst += [shiftdf.loc[shiftdf.index[i-2]]['IOMP-MT']]
            prev_shift_lst += [shiftdf.loc[shiftdf.index[i-2]]['IOMP-CT']]
            try:
                prev_shift_lst += [shiftdf.loc[shiftdf.index[i+2]]['IOMP-MT']]
                prev_shift_lst += [shiftdf.loc[shiftdf.index[i+2]]['IOMP-CT']]
            except:
                pass
        if randomCT in scomp.names_A:
            isocal = shiftdf.index[i].isocalendar()
            startTSA = datetime.strptime(str(isocal[0])+'-w'+str(isocal[1]-1)+'-0', '%Y-W%W-%w')
            endTSA = datetime.strptime(str(isocal[0])+'-w'+str(isocal[1])+'-6', '%Y-W%W-%w')
            dfA = shiftdf.loc[(shiftdf.index > startTSA)&(shiftdf.index < endTSA)&(shiftdf['IOMP-CT'].isin(scomp.names_A))]
            if len(dfA) == WAshift:
                continue
        if randomCT not in no_shift_lst and randomCT not in prev_shift_lst:
            break
    shiftdf.loc[shiftdf.index[i]]['IOMP-CT'] = randomCT
    CTset.remove(randomCT)
    CTlst.remove(randomCT)

#Remaining shifts
lastshift = len(shiftdf.loc[shiftdf['IOMP-CT'] == '?'])
for i in range(0, lastshift):
    while True:
        #random personnel to take CT shift
        randomCT = random.choice(CTlst)
        #checks if has fieldwork during that day
        no_shift_lst = no_shift[no_shift.ts == shiftdf.index[i].date()].index
        #checks if has shift before and after
        prev_shift_lst = [shiftdf.loc[shiftdf.index[i]]['IOMP-MT']]
        prev_shift_lst += [shiftdf.loc[shiftdf.index[i-1]]['IOMP-CT']]
        prev_shift_lst += [shiftdf.loc[shiftdf.index[i+1]]['IOMP-CT']]
        prev_shift_lst += [shiftdf.loc[shiftdf.index[i-1]]['IOMP-MT']]
        prev_shift_lst += [shiftdf.loc[shiftdf.index[i+1]]['IOMP-MT']]
        #if PM shifts, checks if with PM shift before and after
        if shiftdf.index[i].time() == time(19,30):
            prev_shift_lst += [shiftdf.loc[shiftdf.index[i-2]]['IOMP-CT']]
            prev_shift_lst += [shiftdf.loc[shiftdf.index[i+2]]['IOMP-CT']]
            prev_shift_lst += [shiftdf.loc[shiftdf.index[i-2]]['IOMP-MT']]
            prev_shift_lst += [shiftdf.loc[shiftdf.index[i+2]]['IOMP-MT']]
        if randomCT in scomp.names_A:
            isocal = shiftdf.index[i].isocalendar()
            startTSA = datetime.strptime(str(isocal[0])+'-w'+str(isocal[1]-1)+'-0', '%Y-W%W-%w')
            endTSA = datetime.strptime(str(isocal[0])+'-w'+str(isocal[1])+'-6', '%Y-W%W-%w')
            dfA = shiftdf.loc[(shiftdf.index > startTSA)&(shiftdf.index < endTSA)&(shiftdf['IOMP-CT'].isin(scomp.names_A))]
            if len(dfA) == WAshift:
                continue
        if randomCT not in no_shift_lst and randomCT not in prev_shift_lst:
            break
    shiftdf.loc[shiftdf.index[i]]['IOMP-CT'] = randomCT
    CTlst.remove(randomCT)

shiftdf['IOMP-CT'] = shiftdf['IOMP-CT'].apply(lambda x: x[0].upper()+x[1:len(x)])
shiftdf['IOMP-MT'] = shiftdf['IOMP-MT'].apply(lambda x: x[0].upper()+x[1:len(x)])
shiftdf['IOMP-CT'] = ','.join(shiftdf['IOMP-CT'].values).replace('Tinc', 'TinC').replace('Tinb', 'TinB').split(',')
shiftdf['IOMP-MT'] = ','.join(shiftdf['IOMP-MT'].values).replace('Tinb', 'TinB').split(',')
shiftdf = shiftdf[['IOMP-MT','IOMP-CT']]

print shiftdf

##################################### EXCEL ####################################

writer = pd.ExcelWriter('MonitoringShift.xlsx')
try:
    allsheet = pd.read_excel('MonitoringShift.xlsx', sheetname=None)
    allsheet[date] = shiftdf.reset_index()
except:
    allsheet = {date: shiftdf.reset_index()}
for sheetname, xlsxdf in allsheet.items():
    xlsxdf.to_excel(writer, sheetname, index=False)
    worksheet = writer.sheets[sheetname]
    #adjust column width
    worksheet.set_column('A:A', 20)
writer.save()