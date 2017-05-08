import pandas as pd
import random
from datetime import time, timedelta
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

#Start of shift assigning
MT_field1 = len(no_shift[(no_shift.ts >= startTS)&(no_shift.ts < startTS+timedelta(len(scomp.SD)/2.))&(no_shift.index.isin(scomp.SD))])
MT_field2 = len(no_shift[(no_shift.ts <= endTS)&(no_shift.ts > endTS-timedelta(len(scomp.SD)/2.))&(no_shift.index.isin(scomp.SD))])
CT_field1 = len(no_shift[(no_shift.ts >= startTS)&(no_shift.ts < startTS+timedelta(len(scomp.CT)/2.))&(no_shift.index.isin(scomp.CT))])
CT_field2 = len(no_shift[(no_shift.ts <= endTS)&(no_shift.ts > endTS-timedelta(len(scomp.CT)/2.))&(no_shift.index.isin(scomp.CT))])
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
    MTset = list(shift_count[(shift_count.remaining_MTshift.isin([max_MTshift, max_MTshift-1]))&(shift_count.remaining_MTshift > 0)].index)
    random.shuffle(MTset)
    #First set of MT shifts (equal to number of personnel with max MT shifts)
    while len(MTset) != 0:
        while True:
            #random personnel to take MT shift
            randomMT = MTset[0]
            #checks if has fieldwork during that day
            no_shift_lst = no_shift[no_shift.ts == shiftdf.loc[shiftdf['IOMP-MT'] == '?'].index[0].date()].index
            try:
                prev_shift_lst = [shiftdf.loc[shiftdf['IOMP-MT'] != '?']['IOMP-MT'].values[-1]]
            except:
                prev_shift_lst = []
                print 'first MT shift'
            try:
                if shiftdf.loc[shiftdf['IOMP-MT'] == '?'].index[0].time() == time(19,30):
                    prev_shift_lst += [shiftdf.loc[shiftdf['IOMP-MT'] != '?']['IOMP-MT'].values[-2]]
            except:
                print 'first MT night shift'
            if randomMT not in no_shift_lst and randomMT not in prev_shift_lst:
                break
            else:
                MTset = MTset[1::]+[MTset[0]]
        shiftdf.loc[shiftdf.loc[shiftdf['IOMP-MT'] == '?'].index[0]]['IOMP-MT'] = randomMT
        MTlst.remove(randomMT)
        MTset.remove(randomMT)
    
    remaining_MTshift = []
    for i in shift_count.index:
        remaining_MTshift += [MTlst.count(i)]
    shift_count['remaining_MTshift'] = remaining_MTshift
    max_MTshift = max(shift_count.remaining_MTshift)

#################################### CT Shift ###################################
#if CT_assign = end: fills from top to bottom
if CT_assign == 'end':
    shiftdf = shiftdf.sort_index()
else:
    shiftdf = shiftdf.sort_index(ascending=False)

CTlst = []
for i in sorted(set(shift_count[shift_count.CTshift != 0].index)):
    CTlst += [i]*shift_count[shift_count.index == i]['CTshift'].values[0]

remaining_CTshift = []
for i in shift_count.index:
    remaining_CTshift += [CTlst.count(i)]
shift_count['remaining_CTshift'] = remaining_CTshift
max_CTshift = max(shift_count.remaining_CTshift)

while max_CTshift > 0:
    CTset = list(shift_count[(shift_count.remaining_CTshift > 0)].index)
    random.shuffle(CTset)
    #First set of CT shifts (equal to number of personnel with max CT shifts)
    while len(CTset) != 0:
        while True:
            #random personnel to take CT shift
            randomCT = CTset[0]
            #checks if has fieldwork during that day
            no_shift_lst = no_shift[no_shift.ts == shiftdf.loc[shiftdf['IOMP-CT'] == '?'].index[0].date()].index
            prev_shift_lst = [shiftdf.loc[shiftdf['IOMP-CT'] == '?']['IOMP-MT'].values[0]]
            try:
                prev_shift_lst += [shiftdf.loc[shiftdf['IOMP-CT'] != '?']['IOMP-CT'].values[-1]]
                prev_shift_lst += [shiftdf.loc[shiftdf['IOMP-CT'] != '?']['IOMP-MT'].values[-1]]
            except:
                print 'first CT shift'
            try:
                prev_shift_lst += [shiftdf.loc[shiftdf['IOMP-CT'] == '?']['IOMP-MT'].values[1]]
            except:
                print 'last CT shift'
            try:
                if shiftdf.loc[shiftdf['IOMP-CT'] == '?'].index[0].time() == time(19,30):
                    prev_shift_lst += [shiftdf.loc[shiftdf['IOMP-CT'] != '?']['IOMP-CT'].values[-2]]
                    prev_shift_lst += [shiftdf.loc[shiftdf['IOMP-CT'] != '?']['IOMP-MT'].values[-2]]
            except:
                print 'first CT night shift'
            try:
                prev_shift_lst += [shiftdf.loc[shiftdf['IOMP-CT'] == '?']['IOMP-MT'].values[2]]
            except:
                print 'last CT night shift'
            if randomCT not in no_shift_lst and randomCT not in prev_shift_lst:
                break
            else:
                CTset = CTset[1::]+[CTset[0]]
        shiftdf.loc[shiftdf.loc[shiftdf['IOMP-CT'] == '?'].index[0]]['IOMP-CT'] = randomCT
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
shiftdf['IOMP-CT'] = ','.join(shiftdf['IOMP-CT'].values).replace('Tinc', 'TinC').replace('Tinb', 'TinB').split(',')
shiftdf = shiftdf[['IOMP-MT','IOMP-CT']]

print shiftdf

############################## shift count (excel) #############################

writer = pd.ExcelWriter('ShiftCount.xlsx')
try:
    allsheet = pd.read_excel('ShiftCount.xlsx', sheetname=None)
    allsheet[date] = shift_count.reset_index().sort('team')
except:
    allsheet = {date: shift_count.reset_index()}
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