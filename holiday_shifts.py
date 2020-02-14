"""
with the assumption that personnel who can take CT shifts only
is less than personnel who can take both MT and CT shifts

works only if list of number of holiday shifts (2 * 3 * number of holidays)
is greater than total number of monitoring personnel

multiplied by:
2 for CT and MT
3 for PM shift before the holiday, AM and PM shift on holiday 
"""

from datetime import time, timedelta
from calendar import monthrange

import math
import pandas as pd
import random


def check_salary_week(df):
    ts = pd.to_datetime(df['ts'].values[0])
    year = ts.strftime('%Y')
    month = ts.strftime('%m')
    
    months15th = pd.to_datetime(year + '-' + month + '-15')
    start_15thweek = months15th - timedelta(months15th.weekday() + 1)
    end_15thweek = start_15thweek + timedelta(7)

    monthsend = pd.to_datetime(year + '-' + month + '-' + str(monthrange(int(year), int(month))[1]))
    start_monthsend_week = monthsend - timedelta(monthsend.weekday() + 1)
    end_monthsend_week = start_monthsend_week + timedelta(7)

    if int(month) != 1:
        prev_year = year
        prev_month = str(int(month) - 1)
    else:
        prev_year = str(int(year) - 1)
        prev_month = '12'
        
    prev_end_week = pd.to_datetime(prev_year + '-' + prev_month + '-' + str(monthrange(int(prev_year), int(prev_month))[1]))
    start_prev_end_week = prev_end_week - timedelta(prev_end_week.weekday() + 1)
    end_prev_end_week = start_prev_end_week + timedelta(7)

    salary_week = start_15thweek <= ts <= end_15thweek or \
        start_monthsend_week <= ts <= end_monthsend_week or \
        prev_end_week <= ts <= end_prev_end_week
    
    df['salary_week'] = salary_week
    
    return df

def check_weekdayAM(df):
    weekdayAM = pd.to_datetime(df['ts'].values[0]).isocalendar()[2] in range(1, 6) and pd.to_datetime(df['ts'].values[0]).time() == time(7, 30)
    df['weekdayAM'] = weekdayAM
    return df

def shifts(holiday):
    ts = pd.to_datetime(holiday['ts'].values[0]) + timedelta(hours=7.5)
    ts_list = [ts-timedelta(0.5), ts, ts+timedelta(0.5)]
    df = pd.DataFrame({'ts': ts_list})
    return df

# imports ts of holidays and name & team of dynaslope staff
holidays = pd.read_csv('holidays.csv', names=['ts'])
staff = pd.read_csv('dyna_staff.csv', names=['name', 'team'])

# admin
admin = sorted(staff[staff.team == 'admin']['name'].values)
random.shuffle(admin)
# randomize CT shifts
IOMP_CT = sorted(staff[staff.team.isin(['CT', 'S1'])]['name'].values)
random.shuffle(IOMP_CT)
# randomize CT shifts
IOMP_MT = sorted(staff[~staff.team.isin(['CT', 'S1', 'admin'])]['name'].values)
random.shuffle(IOMP_MT)

# holiday shifts
holiday_grp = holidays.groupby('ts', as_index=False)
holiday_shifts = holiday_grp.apply(shifts).reset_index(drop=True)
start_date = pd.to_datetime(min(holidays['ts'].values))
if int(start_date.strftime('%d')) == 1:
    holiday_shifts = holiday_shifts[holiday_shifts.ts >= start_date]
# checks if holiday on a weekday
holiday_grp = holiday_shifts.groupby('ts', as_index=False)
holiday_shifts = holiday_grp.apply(check_weekdayAM).reset_index(drop=True)
# checks if holiday during the salary week (week of 15 and last week of the month)
holiday_grp = holiday_shifts.groupby('ts', as_index=False)
holiday_shifts = holiday_grp.apply(check_salary_week).reset_index(drop=True)

holiday_shifts['IOMP-MT'] = '?'
holiday_shifts['IOMP-CT'] = '?'

shift_each = math.floor((len(holiday_shifts) * 2) / len(staff))

# assign shifts for ate amy:
not_salary_weekdayAM = holiday_shifts[holiday_shifts.weekdayAM & ~holiday_shifts.salary_week]['ts'].values
random.shuffle(not_salary_weekdayAM)
amy_shifts = not_salary_weekdayAM[0:shift_each]
for ts in amy_shifts:
    holiday_shifts.loc[holiday_shifts.ts == ts, 'IOMP-CT'] = 'amy'

# assign shifts for the rest of the admin
weekdayAM_shift = sorted(holiday_shifts[holiday_shifts.weekdayAM & (holiday_shifts['IOMP-CT'] == '?')]['ts'].values)
admin.remove('amy')
admin = admin * shift_each
for admin_i in admin:
    ts = random.choice(weekdayAM_shift)
    weekdayAM_shift.remove(ts)
    holiday_shifts.loc[holiday_shifts.ts == ts, 'IOMP-CT'] = admin_i

total_admin = math.floor(len(holiday_shifts[holiday_shifts['IOMP-CT'] != '?']) / shift_each)

total_MTshift = len(holiday_shifts)
total_CTshift = len(holiday_shifts) - total_admin * shift_each

MT_shift = IOMP_MT[math.floor((len(IOMP_MT) - (len(IOMP_CT) + total_admin))/2) ::]
CT_shift = IOMP_CT + IOMP_MT[0 : math.floor((len(IOMP_MT) - (len(IOMP_CT) + total_admin))/2)]
IOMP_MT = IOMP_MT[math.floor((len(IOMP_MT) - (len(IOMP_CT) + total_admin))/2) ::] + IOMP_MT[0 : math.floor((len(IOMP_MT) - (len(IOMP_CT) + total_admin))/2)]

while len(IOMP_MT) + len(IOMP_CT) < (total_MTshift + total_CTshift) - (len(MT_shift) + len(CT_shift)):
    if len(IOMP_CT) + total_admin != len(IOMP_MT):
        add = 1
    else:
        add = 0
    MT_shift += IOMP_MT[math.floor((len(IOMP_MT) - (len(IOMP_CT) + total_admin))/2) + add ::]
    CT_shift += IOMP_CT + IOMP_MT[0 : math.floor((len(IOMP_MT) - (len(IOMP_CT) + total_admin))/2) + add]
    IOMP_MT = IOMP_MT[math.floor((len(IOMP_MT) - (len(IOMP_CT) + total_admin))/2) + add ::] + IOMP_MT[0 : math.floor((len(IOMP_MT) - (len(IOMP_CT) + total_admin))/2) + add]

if (total_MTshift + total_CTshift) - (len(MT_shift) + len(CT_shift)) > 0:
    if len(IOMP_CT) > total_CTshift - len(CT_shift):
        CT_shift += IOMP_CT[0 : total_CTshift - len(CT_shift)]
        MT_shift += IOMP_MT[0 : total_MTshift - len(MT_shift)]
    else:
        CT_shift += IOMP_CT + IOMP_MT[0 : total_CTshift - (len(CT_shift) + len(IOMP_CT))]
        MT_shift += IOMP_MT[total_CTshift - (len(CT_shift) + len(IOMP_CT)) : (total_MTshift + total_CTshift) - (len(MT_shift) + len(CT_shift) + len(IOMP_CT))]
    
holiday_shifts['IOMP-MT'] = MT_shift
admin_shifts = holiday_shifts[holiday_shifts['IOMP-CT'] != '?']
holiday_shifts = holiday_shifts[holiday_shifts['IOMP-CT'] == '?']
holiday_shifts['IOMP-CT'] = CT_shift
holiday_shifts = holiday_shifts.append(admin_shifts)
holiday_shifts = holiday_shifts.sort_values('ts')

holiday_shifts['IOMP-CT'] = holiday_shifts['IOMP-CT'].apply(lambda x: x[0].upper()+x[1:len(x)])
holiday_shifts['IOMP-MT'] = holiday_shifts['IOMP-MT'].apply(lambda x: x[0].upper()+x[1:len(x)])
holiday_shifts['IOMP-CT'] = ','.join(holiday_shifts['IOMP-CT'].values).replace('Tinc', 'TinC').split(',')

print (holiday_shifts)

year = pd.to_datetime(holiday_shifts['ts'].values[0]).strftime('%Y')
              
writer = pd.ExcelWriter('HolidayShift.xlsx')
try:
    allsheet = pd.read_excel('HolidayShift.xlsx', sheet_name=None)
    allsheet[year] = holiday_shifts
except:
    allsheet = {year: holiday_shifts}
for sheetname, xlsxdf in allsheet.items():
    xlsxdf.to_excel(writer, sheetname, index=False)
    worksheet = writer.sheets[sheetname]
writer.save()