"""
with the assumption that the number of holiday shifts is greater than 
the total number of monitoring personnel, and minimum shift count is 2. 
(e.g for year with 54 holiday shifts, should have 37-54 monitoring personnel)
"""

from datetime import time, timedelta

import numpy as np
import pandas as pd
import random

import shift_sched as sched


def get_holiday_shifts(holiday):
    ts = pd.to_datetime(holiday['ts'].values[0]) + timedelta(hours=7.5)
    ts_list = [ts-timedelta(0.5), ts, ts+timedelta(0.5)]
    df = pd.DataFrame({'ts': ts_list})
    return df

def shift_divider(shift_list, IOMP, admin_list):
    IOMP.loc[:, 'IOMP-MT'] = np.nan
    IOMP.loc[:, 'IOMP-CT'] = np.nan
    shift_count = IOMP.loc[:, ['Nickname', 'AM_shifts', 'IOMP-MT', 'IOMP-CT']].rename(columns={'Nickname': 'name'})
    min_shift = int(np.floor(2 * len(shift_list) / len(shift_count)))
    # admin
    shift_count.loc[shift_count.name.isin(admin_list), 'IOMP-CT'] = min_shift
    # non-admin
    non_admin_list = sorted(set(shift_count['name']) - set(admin_list))
    MT_count = len(shift_list)
    CT_count = len(shift_list) - min_shift * len(admin_list)
    # at least 1 MT shift for non-admin
    shift_count.loc[shift_count.name.isin(non_admin_list), 'IOMP-MT'] = 1
    MT_count -= len(non_admin_list)
    random.shuffle(non_admin_list)
    if CT_count > len(non_admin_list):
        shift_count.loc[shift_count.name.isin(non_admin_list), 'IOMP-CT'] = 1
        CT_count -= len(non_admin_list)
        shift_count.loc[shift_count.name.isin(non_admin_list[:CT_count]), 'IOMP-CT'] += 1
        shift_count.loc[shift_count.name.isin(non_admin_list[CT_count:CT_count+MT_count]), 'IOMP-MT'] += 1
    else:
        shift_count.loc[shift_count.name.isin(non_admin_list[len(non_admin_list)-CT_count:]), 'IOMP-CT'] = 1
        shift_count.loc[shift_count.name.isin(non_admin_list[:MT_count]), 'IOMP-MT'] += 1
    shift_count = shift_count.fillna(0)
    
    return shift_count


def assign_remaining_IOMP(shiftdf, shift_count):
    # unassigned IOMPs
    extra_CT = shift_count.loc[(shift_count['IOMP-CT'] == 2), 'name'].values
    extra_MT = shift_count.loc[(shift_count['IOMP-MT'] == 2), 'name'].values
    other_IOMP = sorted(set(shift_count.loc[((shift_count['IOMP-MT'] != 0) & (shift_count['IOMP-CT'] != 0)), 'name'].values) - set(extra_CT) - set(extra_MT))
    # randomize
    random.shuffle(extra_MT)
    random.shuffle(extra_CT)
    random.shuffle(other_IOMP)
    
    # shifts with unassigned IOMP
    MT_shift = shiftdf.loc[shiftdf['IOMP-MT'] == '?', 'ts'].values
    CT_shift = shiftdf.loc[shiftdf['IOMP-CT'] == '?', 'ts'].values
    
    # assign remaining shifts
    remaining_IOMP = list(extra_MT) + list(extra_CT) + list(other_IOMP)
    MT = remaining_IOMP[::2]
    temp_MT_shift = MT_shift[0:len(MT)]
    MT_shift = sorted(set(MT_shift) - set(temp_MT_shift))
    shiftdf.loc[shiftdf.ts.isin(temp_MT_shift), 'IOMP-MT'] = MT
    CT = remaining_IOMP[1::2]
    temp_CT_shift = CT_shift[0:len(CT)]
    CT_shift = sorted(set(CT_shift) - set(temp_CT_shift))
    shiftdf.loc[shiftdf.ts.isin(temp_CT_shift), 'IOMP-CT'] = CT
    
    # assign extra shifts
    temp_MT_shift = MT_shift[0:len(extra_MT)]
    MT_shift = sorted(set(MT_shift) - set(temp_MT_shift))
    shiftdf.loc[shiftdf.ts.isin(temp_MT_shift), 'IOMP-MT'] = extra_MT
    temp_CT_shift = CT_shift[0:len(extra_CT)]
    CT_shift = sorted(set(CT_shift) - set(temp_CT_shift))
    shiftdf.loc[shiftdf.ts.isin(temp_CT_shift), 'IOMP-CT'] = extra_CT    
    
    # assign remaining flipped shifts
    remaining_IOMP = list(other_IOMP) + list(extra_MT) + list(extra_CT)
    MT = remaining_IOMP[1::2]
    shiftdf.loc[shiftdf.ts.isin(MT_shift), 'IOMP-MT'] = MT
    CT = remaining_IOMP[::2]
    shiftdf.loc[shiftdf.ts.isin(CT_shift), 'IOMP-CT'] = CT
    
    return shiftdf

def main():
    holidays = pd.read_csv('holidays.csv', names=['ts'])
    holiday_grp = holidays.groupby('ts', as_index=False)
    shift_list = holiday_grp.apply(get_holiday_shifts).reset_index(drop=True)
    start_date = pd.to_datetime(min(holidays['ts'].values))
    end_date = pd.to_datetime(max(holidays['ts'].values))
    if int(start_date.strftime('%d')) == 1:
        shift_list = shift_list[shift_list.ts >= start_date]
    shift_list = shift_list.drop_duplicates('ts')
    shiftdf = shift_list.reset_index(drop=True)
    shiftdf.loc[:, 'IOMP-MT'] = '?'
    shiftdf.loc[:, 'IOMP-CT'] = '?'

    # IOMPs    
    personnel_sheet = "personnel"
    key = "1UylXLwDv1W1ukT4YNoUGgHCHF-W8e3F8-pIg1E024ho"
    IOMP = sched.get_sheet(key, personnel_sheet)
    IOMP = IOMP.loc[IOMP.current == 1, :]
    admin_list = sorted(IOMP[IOMP.team == 'admin']['Nickname'].values)

    # Number of shifts
    shift_count = shift_divider(shiftdf, IOMP, admin_list)
    fieldwork = sched.get_field(key, start_date, end_date)
    # Assign admin
    for name in admin_list:
        shiftdf, shift_count = sched.assign_shift(name, shift_count, shiftdf, start_date, end_date, admin_list, fieldwork)
    # Assign no CT
    for name in shift_count.loc[(~shift_count.name.isin(admin_list)) & ((shift_count['IOMP-MT'] == 0) | (shift_count['IOMP-CT'] == 0)), 'name']:
        shiftdf, shift_count = sched.assign_shift(name, shift_count, shiftdf, start_date, end_date, admin_list, fieldwork)
    # Assign remaining IOMP
    shiftdf = assign_remaining_IOMP(shiftdf, shift_count)
    
    sched.shift_validity(shiftdf, shift_count, fieldwork)

    year = pd.to_datetime(shiftdf['ts'].values[0]).strftime('%Y')
                  
    writer = pd.ExcelWriter('HolidayShift.xlsx')
    try:
        allsheet = pd.read_excel('HolidayShift.xlsx', sheet_name=None)
        allsheet[year] = shiftdf
    except:
        allsheet = {year: shiftdf}
    for sheetname, xlsxdf in allsheet.items():
        xlsxdf.to_excel(writer, sheetname, index=False)
    writer.save()
    
    return shiftdf
    
if __name__ == "__main__":
    shiftdf = main()
