from collections import Counter
from calendar import monthrange
from datetime import time, timedelta, datetime
import numpy as np
import pandas as pd
import random

import shift_division as div


def restrict_shift(fieldwork):
    ts = pd.to_datetime(fieldwork['ts'].values[0]) + timedelta(hours=7.5)
    ts_list = [ts-timedelta(0.5), ts, ts+timedelta(0.5)]
    df = pd.DataFrame({'ts': ts_list})
    df['name'] = fieldwork['name'].values[0]
    return df

def holiday_shift(holidays, shiftdf):
    IOMP = ['IOMP-MT', 'IOMP-CT']
    ts = holidays['ts'].values[0]
    shiftdf.loc[shiftdf.ts == ts, IOMP] = holidays[div.holidays.ts == ts][IOMP].values

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
        prev_month = 12
        
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

def allowed_shift(ts, name, shiftdf):
    ts = pd.to_datetime(ts)
    tdelta = 0.5
    if pd.to_datetime(ts).time() == time(19, 30):
        tdelta += 0.5
    not_allowed = np.concatenate(shiftdf[(shiftdf.ts >= ts - timedelta(tdelta)) & \
                           (shiftdf.ts <= ts + timedelta(tdelta))][['IOMP-CT', 'IOMP-MT']].values)
    return name not in not_allowed

#Empty Shift Schedule
shiftTS = pd.date_range(start=div.startTS, end=div.endTS, freq='12H')
shiftdf = pd.DataFrame({'ts': shiftTS, 'IOMP-MT': ['?']*len(shiftTS), 'IOMP-CT': ['?']*len(shiftTS)})

shift_count = div.shift_count.reset_index(drop=True)

# holiday shifts
holiday_ts = div.holidays.groupby('ts', as_index=False)
holiday_ts.apply(holiday_shift, shiftdf=shiftdf)

for MT in div.holidays['IOMP-MT'].values:
    shift_count.loc[shift_count.name == MT, 'MT'] -= 1

for CT in div.holidays['IOMP-CT'].values:
    shift_count.loc[shift_count.name == CT, 'CT'] -= 1


# shift of people with fieldwork
fieldwork = pd.read_csv('Monitoring Shift Schedule - fieldwork.csv')
fieldwork['ts'] = pd.to_datetime(fieldwork['ts'])
fieldwork = fieldwork[(fieldwork.ts >= datetime.strptime(div.date, '%b%Y')) & (fieldwork.ts <= div.endTS)]
fieldwork['name'] = fieldwork['name'].apply(lambda x: x.lower())
fieldwork['id'] = range(len(fieldwork))
fieldwork_id = fieldwork.groupby('id', as_index=False)
field_shifts = fieldwork_id.apply(restrict_shift).drop_duplicates(['ts', 'name']).reset_index(drop=True)

field_shift_count = Counter(field_shifts.name)
field_shift_count = pd.DataFrame({'name': field_shift_count.keys(), 'field_shift_count': field_shift_count.values()})
admin_list = shift_count[(shift_count.team == 'admin')]['name'].values
field_shift_count = field_shift_count[~field_shift_count.name.isin(admin_list)]
field_shift_count = field_shift_count.sort_values('field_shift_count', ascending=False)


# check if weekday AM shift and/or salary week (last week and week of 15th)
shiftdf_grp = shiftdf.groupby('ts', as_index=False)
shiftdf = shiftdf_grp.apply(check_salary_week).reset_index(drop=True)
shiftdf_grp = shiftdf.groupby('ts', as_index=False)
shiftdf = shiftdf_grp.apply(check_weekdayAM).reset_index(drop=True)

# shift of ate amy
if shift_count[shift_count.name == 'amy']['CT'].values[0] != 0:
    not_salary_weekdayAM = sorted(set(shiftdf[shiftdf.weekdayAM & ~shiftdf.salary_week & (shiftdf['IOMP-CT'] == '?')]['ts']) - set(field_shifts[field_shifts.name == 'amy']['ts']))
    allowed = False
    while not allowed:
        ts = random.choice(not_salary_weekdayAM)
        allowed = allowed_shift(ts, 'amy', shiftdf)
    shiftdf.loc[shiftdf.ts == ts, 'IOMP-CT'] = 'amy'
    shift_count.loc[shift_count.name == 'amy', 'CT'] -= 1

# shift of remaining admin
for admin in shift_count[(shift_count.team == 'admin') & (shift_count.CT != 0)]['name'].values:
    weekdayAM = shiftdf[shiftdf.weekdayAM & (shiftdf['IOMP-CT'] == '?')]['ts'].values
    allowed = False
    while not allowed:
        ts = random.choice(weekdayAM)
        allowed = allowed_shift(ts, admin, shiftdf)
    shiftdf.loc[shiftdf.ts == ts, 'IOMP-CT'] = admin
    shift_count.loc[shift_count.name == admin, 'CT'] -= 1

for IOMP in field_shift_count['name'].values:
    for MT in range(shift_count[shift_count.name == IOMP]['MT'].values[0]):
        ts_list = sorted(set(shiftdf[shiftdf['IOMP-MT'] == '?']['ts']) - set(field_shifts[field_shifts.name == IOMP]['ts']))
        allowed = False
        while not allowed:
            ts = random.choice(ts_list)
            allowed = allowed_shift(ts, IOMP, shiftdf)
        shiftdf.loc[shiftdf.ts == ts, 'IOMP-MT'] = IOMP
        shift_count.loc[shift_count.name == IOMP, 'MT'] -= 1
    for CT in range(shift_count[shift_count.name == IOMP]['CT'].values[0]):
        ts_list = sorted(set(shiftdf[shiftdf['IOMP-CT'] == '?']['ts']) - set(field_shifts[field_shifts.name == IOMP]['ts']))
        allowed = False
        while not allowed:
            ts = random.choice(ts_list)
            allowed = allowed_shift(ts, IOMP, shiftdf)
        shiftdf.loc[shiftdf.ts == ts, 'IOMP-CT'] = IOMP
        shift_count.loc[shift_count.name == IOMP, 'CT'] -= 1
                       
# shift of remaining IOMP

#remaining MT
remaining_MT = []
for IOMP in sorted(set(shift_count[shift_count.MT != 0].name)):
    remaining_MT += [IOMP]*shift_count[shift_count.name == IOMP]['MT'].values[0]

random.shuffle(remaining_MT)

for IOMP in remaining_MT:
    ts_list = sorted(shiftdf[shiftdf['IOMP-MT'] == '?']['ts'])
    allowed = False
    while not allowed:
        ts = random.choice(ts_list)
        allowed = allowed_shift(ts, IOMP, shiftdf)
    shiftdf.loc[shiftdf.ts == ts, 'IOMP-MT'] = IOMP
    shift_count.loc[shift_count.name == IOMP, 'MT'] -= 1

#remaining CT
remaining_CT = []
for IOMP in sorted(set(shift_count[shift_count.CT != 0].name)):
    remaining_CT += [IOMP]*shift_count[shift_count.name == IOMP]['CT'].values[0]

random.shuffle(remaining_CT)

for IOMP in remaining_CT:
    ts_list = sorted(shiftdf[shiftdf['IOMP-CT'] == '?']['ts'])
    allowed = False
    while not allowed:
        ts = random.choice(ts_list)
        allowed = allowed_shift(ts, IOMP, shiftdf)
    shiftdf.loc[shiftdf.ts == ts, 'IOMP-CT'] = IOMP
    shift_count.loc[shift_count.name == IOMP, 'CT'] -= 1
                   
shiftdf = shiftdf[['ts', 'IOMP-MT', 'IOMP-CT']]
shiftdf['IOMP-CT'] = shiftdf['IOMP-CT'].apply(lambda x: x[0].upper()+x[1:len(x)])
shiftdf['IOMP-MT'] = shiftdf['IOMP-MT'].apply(lambda x: x[0].upper()+x[1:len(x)])
shiftdf['IOMP-CT'] = ','.join(shiftdf['IOMP-CT'].values).replace('Tinc', 'TinC').split(',')

print shiftdf

##################################### EXCEL ####################################

writer = pd.ExcelWriter('MonitoringShift.xlsx')
try:
    allsheet = pd.read_excel('MonitoringShift.xlsx', sheetname=None)
    allsheet[div.date] = shiftdf
except:
    allsheet = {div.date: shiftdf}
for sheetname, xlsxdf in allsheet.items():
    xlsxdf.to_excel(writer, sheetname, index=False)
    worksheet = writer.sheets[sheetname]
    #adjust column width
    worksheet.set_column('A:A', 20)
writer.save()