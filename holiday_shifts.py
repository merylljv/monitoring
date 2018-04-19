"""
works only if list of number of holiday shifts (2 * 3 * number of holidays)
is greater than total number of monitoring personnel

multiplied by:
2 for CT and MT
3 for PM shift before the holiday, AM and PM shift on holiday 
"""

from datetime import timedelta
import pandas as pd
import random

def shifts(holiday):
    ts = pd.to_datetime(holiday['ts'].values[0]) + timedelta(hours=7.5)
    ts_list = [ts-timedelta(0.5), ts, ts+timedelta(0.5)]
    df = pd.DataFrame({'ts': ts_list})
    return df

# imports ts of holidays and name & team of dynaslope staff
holidays = pd.read_csv('holidays.csv', names=['ts'])
staff = pd.read_csv('dyna_staff.csv', names=['name', 'team'])

# randomize CT shifts
IOMP_CT = sorted(staff[staff.team.isin(['CT', 'S1', 'admin'])]['name'].values)
random.shuffle(IOMP_CT)
# randomize CT shifts
IOMP_MT = sorted(staff[~staff.team.isin(['CT', 'S1', 'admin'])]['name'].values)
random.shuffle(IOMP_MT)

holiday_grp = holidays.groupby('ts', as_index=False)
holiday_shifts = holiday_grp.apply(shifts).reset_index(drop=True)

MT_shift = IOMP_MT[(len(IOMP_MT) - len(IOMP_CT))/2 ::]
CT_shift = IOMP_CT + IOMP_MT[0 : (len(IOMP_MT) - len(IOMP_CT))/2]
IOMP_MT = IOMP_MT[(len(IOMP_MT) - len(IOMP_CT))/2 ::] + IOMP_MT[0 : (len(IOMP_MT) - len(IOMP_CT))/2]

while len(IOMP_MT) + len(IOMP_CT) < (len(holiday_shifts) * 2) - (len(MT_shift) + len(CT_shift)):
    if len(IOMP_CT) != len(IOMP_MT):
        add = 1
    else:
        add = 0
    MT_shift += IOMP_MT[(len(IOMP_MT) - len(IOMP_CT))/2 + add ::]
    CT_shift += IOMP_CT + IOMP_MT[0 : (len(IOMP_MT) - len(IOMP_CT))/2 + add]
    IOMP_MT = IOMP_MT[(len(IOMP_MT) - len(IOMP_CT))/2 + add ::] + IOMP_MT[0 : (len(IOMP_MT) - len(IOMP_CT))/2 + add]

if (len(holiday_shifts) * 2) - (len(MT_shift) + len(CT_shift)) > 0:
    if len(IOMP_CT) > len(holiday_shifts) - len(CT_shift):
        CT_shift += IOMP_CT[0 : len(holiday_shifts) - len(CT_shift)]
        MT_shift += IOMP_MT[0 : len(holiday_shifts) - len(MT_shift)]
    else:
        CT_shift += IOMP_CT + IOMP_MT[0 : len(holiday_shifts) - (len(CT_shift) + len(IOMP_CT))]
        MT_shift += IOMP_MT[len(holiday_shifts) - (len(CT_shift) + len(IOMP_CT)) : (2 * len(holiday_shifts)) - (len(MT_shift) + len(CT_shift) + len(IOMP_CT))]
    
holiday_shifts['IOMP-MT'] = MT_shift
holiday_shifts['IOMP-CT'] = CT_shift

holiday_shifts['IOMP-CT'] = holiday_shifts['IOMP-CT'].apply(lambda x: x[0].upper()+x[1:len(x)])
holiday_shifts['IOMP-MT'] = holiday_shifts['IOMP-MT'].apply(lambda x: x[0].upper()+x[1:len(x)])
holiday_shifts['IOMP-CT'] = ','.join(holiday_shifts['IOMP-CT'].values).replace('Tinc', 'TinC').split(',')

print holiday_shifts

year = pd.to_datetime(holiday_shifts['ts'].values[0]).strftime('%Y')
              
writer = pd.ExcelWriter('HolidayShift.xlsx')
try:
    allsheet = pd.read_excel('HolidayShift.xlsx', sheetname=None)
    allsheet[year] = holiday_shifts
except:
    allsheet = {year: holiday_shifts}
for sheetname, xlsxdf in allsheet.items():
    xlsxdf.to_excel(writer, sheetname, index=False)
    worksheet = writer.sheets[sheetname]
    #adjust column width
    worksheet.set_column('A:A', 20)
writer.save()