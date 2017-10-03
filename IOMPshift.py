import pandas as pd
import random
from datetime import time, timedelta
from calendar import monthrange
from collections import Counter

import IOMPshift_computer as scomp

#Month and Year of Shift Schedule
year = scomp.year
month = scomp.month

def indiv_sched(ind_sched, ts):
    class_sched = pd.DataFrame(columns=['name', 'ts'])
    name = ind_sched['name'].values[0]
    for i in range(len(ind_sched)):
        day = ind_sched[i:i+1]['day'].values[0]
        time = ind_sched[i:i+1]['time'].values[0]
        ts_sched = ts[(ts.strftime('%A') == day)&(ts.time == pd.to_datetime(time).time())]
        sched = pd.DataFrame({'name': [name]*len(ts_sched), 'ts':ts_sched})
        class_sched = class_sched.append(sched, ignore_index = True)
    return class_sched

def class_sched(ts):
    sched = pd.read_csv('class_sched.csv')
    ind_sched = sched.groupby('name', as_index=False)
    class_sched = ind_sched.apply(indiv_sched, ts=ts)
    class_sched = class_sched.drop_duplicates()
    class_sched = class_sched.reset_index(drop=True)
    return class_sched

# Admin Restrictions
def AMshift(df):
    ts = pd.to_datetime(df['ts'].values[0])
    return ts.time() == time(7, 30)

def weekdayAMshift(df):
    ts = pd.to_datetime(df['ts'].values[0])
    return ts.isocalendar()[2] in range(1,6) and AMshift(df)
    
def weekdayshift(df):
    ts = pd.to_datetime(df['ts'].values[0])
    return ts.isocalendar()[2] in range(1,5) or \
        (ts.isocalendar()[2] == 5 and AMshift(ts))

def restricted_shift(shiftdf, ts):
    if ts.time() == time(19, 30):
        tdelta = 1
    else:
        tdelta = 0.5
    restricted = shiftdf[(shiftdf.index >= ts - timedelta(tdelta))&(shiftdf.index <= ts + timedelta(tdelta))]
    personnel = set(list(restricted['IOMP-CT'].values) + list(restricted['IOMP-MT'].values))
    return list(personnel - {'?'})

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

add_restricted = class_sched(shiftdf.index)

# admin shifts
grpts = shiftdf.reset_index().groupby('ts')
weekdayAM = grpts.apply(weekdayAMshift)
weekday = grpts.apply(weekdayAMshift)
AM = grpts.apply(weekdayAMshift)
amy = random.choice(shiftdf[(AM == True)&(shiftdf['IOMP-CT'] == '?')].index)
shiftdf.loc[shiftdf.index == amy, ['IOMP-CT']] = 'amy'
ardeth = random.choice(shiftdf[(weekdayAM == True)&(shiftdf['IOMP-CT'] == '?')].index)
shiftdf.loc[shiftdf.index == ardeth, ['IOMP-CT']] = 'ardeth'
daisy = random.choice(shiftdf[(AM == True)&(shiftdf['IOMP-CT'] == '?')].index)
shiftdf.loc[shiftdf.index == daisy, ['IOMP-CT']] = 'daisy'
tinc = random.choice(shiftdf[(weekday == True)&(shiftdf['IOMP-CT'] == '?')].index)
shiftdf.loc[shiftdf.index == tinc, ['IOMP-CT']] = 'tinc'

#Fieldwork Schedule
no_shift = scomp.no_shift
no_shift = no_shift[(no_shift.ts >= startTS)&(no_shift.ts <= endTS)]

no_shift_count = Counter(no_shift.index)
no_shift_count = pd.DataFrame({'name': no_shift_count.keys(), 'field_days': no_shift_count.values()})

#Start of shift assigning
MT_field1 = len(no_shift[(no_shift.ts < startTS+timedelta(len(scomp.SD)/2.))&(no_shift.index.isin(scomp.SD))])
MT_field2 = len(no_shift[(no_shift.ts > endTS-timedelta(len(scomp.SD)/2.))&(no_shift.index.isin(scomp.SD))])
CT_field1 = len(no_shift[(no_shift.ts < startTS+timedelta(len(scomp.CT)/2.))&(no_shift.index.isin(scomp.CT))])
CT_field2 = len(no_shift[(no_shift.ts > endTS-timedelta(len(scomp.CT)/2.))&(no_shift.index.isin(scomp.CT))])
if MT_field1 >= MT_field2:
    MT_assign = 'start'
else:
    MT_assign = 'end'
if CT_field1 >= CT_field2:
    CT_assign = 'start'
else:
    CT_assign = 'end'

################################### MT Shift ###################################
#if MT_assign = start: fills from top to bottom
if MT_assign == 'start':
    shiftdf = shiftdf.sort_index()
else:
    shiftdf = shiftdf.sort_index(ascending=False)

MTlst = []
for i in sorted(set(shift_count[shift_count.MTshift != 0].index)):
    MTlst += [i]*shift_count[shift_count.index == i]['MTshift'].values[0]

remaining_MTshift = []
for i in shift_count.index:
    remaining_MTshift += [MTlst.count(i)]
shift_count['remaining_MTshift'] = remaining_MTshift
max_MTshift = max(shift_count.remaining_MTshift)

while max_MTshift > 0:
    MTset = list(shift_count[(shift_count.remaining_MTshift >= max_MTshift-1)&(shift_count.remaining_MTshift > 0)].index)
    #First set of MT shifts (equal to number of personnel with max MT shifts)
    while len(MTset) != 0:
        ts = shiftdf.loc[shiftdf['IOMP-MT'] == '?'].index[0]
        while True:
            #random personnel to take MT shift
            randomMT = random.choice(MTset)
            #checks if has fieldwork during that day
            no_shift_lst = no_shift[no_shift.ts == ts.date()].index
            restricted_list = restricted_shift(shiftdf, ts)
            with_class = list(add_restricted[add_restricted.ts == ts].name)
            if randomMT not in no_shift_lst and randomMT not in restricted_list and randomMT not in with_class:
                break
            else:
                MTset = MTset[1::]+[MTset[0]]
        shiftdf.loc[shiftdf.index == ts, ['IOMP-MT']] = randomMT
        MTlst.remove(randomMT)
        MTset.remove(randomMT)
    
    remaining_MTshift = []
    for i in shift_count.index:
        remaining_MTshift += [MTlst.count(i)]
    shift_count['remaining_MTshift'] = remaining_MTshift
    max_MTshift = max(shift_count.remaining_MTshift)

#################################### CT Shift ###################################

# with more than 11 days fieldwork
excess_field = shift_count[(shift_count.field > 11)&(shift_count.CTshift > 0)].index

for i in excess_field:
    field_days = no_shift[no_shift.index == i].index
    avail_shift = shiftdf[(shiftdf['IOMP-CT'] == '?')&~(shiftdf.index.isin(field_days))].index
    chosen_shift = []
    nth_shift = 0
    while len(chosen_shift) != shift_count[shift_count.index == i]['CTshift'].values[0]:
        chosen_shift += [avail_shift[nth_shift]]
        nth_shift += random.choice([3, 4])
    for chosen_ts in chosen_shift:
        shiftdf.loc[shiftdf.index == chosen_ts, ['IOMP-CT']] = i

#if CT_assign = end: fills from top to bottom
if CT_assign != 'start':
    shiftdf = shiftdf.sort_index()
else:
    shiftdf = shiftdf.sort_index(ascending=False)

CTlst = []
for i in sorted(set(shift_count[(shift_count.CTshift != 0)&(shift_count.team != 'Admin')].index)):
    CTlst += [i]*shift_count[shift_count.index == i]['CTshift'].values[0]

remaining_CTshift = []
for i in shift_count.index:
    remaining_CTshift += [CTlst.count(i)]
shift_count['remaining_CTshift'] = remaining_CTshift
shift_count.loc[shift_count.team == 'Admin', ['remaining_CTshift']] = 0
shift_count.loc[shift_count.index.isin(excess_field) , ['remaining_CTshift']] = 0
for i in excess_field:
    while i in CTlst:
        CTlst.remove(i)

max_CTshift = max(shift_count.remaining_CTshift)

while max_CTshift > 0:
    CTset = list(shift_count[(shift_count.remaining_CTshift > 0) \
                #& (shift_count.remaining_CTshift >= max_CTshift-1) \
                ].index)
    #First set of CT shifts (equal to number of personnel with max CT shifts)
    while len(CTset) != 0:
        ts = shiftdf.loc[shiftdf['IOMP-CT'] == '?'].index[0]
        while True:
            #random personnel to take CT shift
            randomCT = random.choice(CTset)
            #checks if has fieldwork during that day
            no_shift_lst = no_shift[no_shift.ts == ts.date()].index
            restricted_list = restricted_shift(shiftdf, ts)
            with_class = list(add_restricted[add_restricted.ts == ts].name)
            if randomCT not in no_shift_lst and randomCT not in restricted_list and randomCT not in with_class:
                break
            else:
                CTset = CTset[1::]+[CTset[0]]
        shiftdf.loc[shiftdf.index == ts, ['IOMP-CT']] = randomCT
        CTlst.remove(randomCT)
        CTset.remove(randomCT)
    
    remaining_CTshift = []
    for i in shift_count.index:
        remaining_CTshift += [CTlst.count(i)]
    shift_count['remaining_CTshift'] = remaining_CTshift
    max_CTshift = max(shift_count.remaining_CTshift)


shiftdf = shiftdf.sort_index()

shiftdf['IOMP-CT'] = shiftdf['IOMP-CT'].apply(lambda x: x[0].upper()+x[1:len(x)])
shiftdf['IOMP-MT'] = shiftdf['IOMP-MT'].apply(lambda x: x[0].upper()+x[1:len(x)])
shiftdf['IOMP-MT'] = ','.join(shiftdf['IOMP-MT'].values).replace('Tinb', 'TinB').split(',')
shiftdf['IOMP-CT'] = ','.join(shiftdf['IOMP-CT'].values).replace('Tinc', 'TinC').replace('Tinb', 'TinB').split(',')
shiftdf = shiftdf[['IOMP-MT','IOMP-CT']]

print shiftdf

############################## shift count (excel) #############################

writer = pd.ExcelWriter('ShiftCount.xlsx')
shift_count['team'] = ','.join(shift_count['team'].values).replace('CTD', 'CT').replace('CTSD', 'CT').replace('CTSS', 'CT').split(',')
shift_count = shift_count.reset_index().sort_values(['team', 'name'])

try:
    allsheet = pd.read_excel('ShiftCount.xlsx', sheetname=None)
    allsheet[date] = shift_count
except:
    allsheet = {date: shift_count}
for sheetname, xlsxdf in allsheet.items():
    xlsxdf.to_excel(writer, sheetname, index=False)
    worksheet = writer.sheets[sheetname]
writer.save()

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