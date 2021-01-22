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
from collections import Counter

import math
import pandas as pd
import random

import shift_sched as sched


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

def main():
    # imports ts of holidays and name & team of dynaslope staff
    holidays = pd.read_csv('holidays.csv', names=['ts'])
    
    personnel_sheet = "personnel"
    key = "1UylXLwDv1W1ukT4YNoUGgHCHF-W8e3F8-pIg1E024ho"
    staff = sched.get_sheet(key, personnel_sheet)
    staff = staff.loc[staff.current == 1, :]
        
    # admin
    admin = sorted(staff[staff.team == 'admin']['Nickname'].values)
    random.shuffle(admin)
    # randomize CT shifts
    IOMP_CT = sorted(staff[staff.team.isin(['CT', 'S1'])]['Nickname'].values)
    random.shuffle(IOMP_CT)
    # randomize CT shifts
    IOMP_MT = sorted(staff[~staff.team.isin(['CT', 'S1', 'admin'])]['Nickname'].values)
    random.shuffle(IOMP_MT)
    
    # holiday shifts
    holiday_grp = holidays.groupby('ts', as_index=False)
    holiday_shifts = holiday_grp.apply(shifts).reset_index(drop=True)
    start_date = pd.to_datetime(min(holidays['ts'].values))
    if int(start_date.strftime('%d')) == 1:
        holiday_shifts = holiday_shifts[holiday_shifts.ts >= start_date]
    holiday_shifts = holiday_shifts.drop_duplicates('ts')
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
        holiday_shifts.loc[holiday_shifts.ts == ts, 'IOMP-CT'] = 'Amy'
    
    # assign shifts for the rest of the admin
    weekdayAM_shift = sorted(holiday_shifts[holiday_shifts.weekdayAM & (holiday_shifts['IOMP-CT'] == '?')]['ts'].values)
    admin.remove('Amy')
    admin = admin * shift_each
    for admin_i in admin:
        ts = random.choice(weekdayAM_shift)
        weekdayAM_shift.remove(ts)
        holiday_shifts.loc[holiday_shifts.ts == ts, 'IOMP-CT'] = admin_i
    
    
    except_admin = staff.loc[staff.team != 'admin', 'Nickname'].values
    random.shuffle(except_admin)
    
    MT_ts = holiday_shifts.loc[holiday_shifts['IOMP-MT'] == '?', 'ts'].values
    CT_ts = holiday_shifts.loc[holiday_shifts['IOMP-CT'] == '?', 'ts'].values
    
    while len(MT_ts) > len(except_admin)/2 and len(CT_ts) > len(except_admin)/2:
        MT = except_admin[0:int((len(except_admin) + (len(MT_ts)-len(CT_ts))) / 2)]
        holiday_shifts.loc[holiday_shifts.ts.isin(MT_ts[0:len(MT)]), 'IOMP-MT'] = MT
        
        CT = except_admin[int((len(except_admin) + (len(MT_ts)-len(CT_ts))) / 2):]
        holiday_shifts.loc[holiday_shifts.ts.isin(CT_ts[0:len(CT)]), 'IOMP-CT'] = CT
    
        MT_ts = holiday_shifts.loc[holiday_shifts['IOMP-MT'] == '?', 'ts'].values
        CT_ts = holiday_shifts.loc[holiday_shifts['IOMP-CT'] == '?', 'ts'].values
    
    MT = except_admin[0:len(MT_ts)]
    holiday_shifts.loc[holiday_shifts.ts.isin(MT_ts[0:len(MT)]), 'IOMP-MT'] = MT
    CT = except_admin[len(MT_ts):len(MT_ts)+len(CT_ts)]
    holiday_shifts.loc[holiday_shifts.ts.isin(CT_ts[0:len(CT)]), 'IOMP-CT'] = CT
    
    
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
    
    return holiday_shifts
    
    
if __name__ == "__main__":
    holiday_shifts = main()

    name_list = set(list(holiday_shifts["IOMP-MT"].values) + list(holiday_shifts["IOMP-CT"].values))
    for name in sorted(name_list):
        print(name)
    #    print(shiftdf.loc[(shiftdf['IOMP-MT'] == name) | (shiftdf['IOMP-CT'] == name), :])
        for ts in holiday_shifts.loc[(holiday_shifts['IOMP-MT'] == name) | (holiday_shifts['IOMP-CT'] == name), 'ts'].values:
            ts = pd.to_datetime(ts)
            if ts.time() == time(19,30):
                cur = holiday_shifts.loc[(holiday_shifts.ts >= ts-timedelta(1)) & (holiday_shifts.ts <= ts+timedelta(1)), :]
            else:
                cur = holiday_shifts.loc[(holiday_shifts.ts >= ts-timedelta(0.5)) & (holiday_shifts.ts <= ts+timedelta(0.5)), :]
            if Counter(cur.values.flatten())[name] != 1:
                print(cur)
