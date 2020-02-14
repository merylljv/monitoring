"""

Works with the assumptions that total number of non-admin personnel does not exceed
number of admin personnel less than twice the number of minimum days in a month.
Equivalent to:
    1) 1 shift each for admin personnel 
    2) at least 2 shifts for non-admin personnel

e.g. with 4 admin personnel,
if Feb has 28 days in a month, non-admin personnel should not exceed 52

"""

from collections import Counter
from calendar import monthrange
from datetime import date, datetime, time, timedelta
import numpy as np
import pandas as pd
import random


def get_sheet(key, sheet_name):
    url = 'https://docs.google.com/spreadsheets/d/{key}/gviz/tq?tqx=out:csv&sheet={sheet_name}&headers=1'.format(
        key=key, sheet_name=sheet_name.replace(' ', '%20'))
    df = pd.read_csv(url)
    df = df.drop([col for col in df.columns if col.startswith('Unnamed')], axis=1)
    return df


def get_shift(key, sheet_name):
    df = get_sheet(key, sheet_name)
    df = df.drop([col for col in df.columns if col.startswith('Unnamed')], axis=1)
    df.loc[:, 'Date'] = pd.to_datetime(df.loc[:, 'Date'].ffill())
    df.loc[:, 'Shift'] = pd.to_timedelta(df.loc[:, 'Shift'].map({'AM': timedelta(hours=7.5), 'PM': timedelta(hours=19.5)}))
    df.loc[:, 'ts'] = pd.to_datetime(df.loc[:, 'Date'] + df.loc[:, 'Shift'])
    return df.loc[:, ['ts', 'IOMP-MT', 'IOMP-CT']]


def get_field(key, start, end):
    field_sheet = "fieldwork"
    personnel_sheet = "personnel"
    field = get_sheet(key, field_sheet)
    name = get_sheet(key, personnel_sheet)
    df = pd.merge(field, name, left_on='Personnel', right_on='Fullname')
    df.loc[:, 'Date of Arrival'] = pd.to_datetime(df.loc[:, 'Date of Arrival'].ffill())
    df.loc[:, 'Date of Departure'] = pd.to_datetime(df.loc[:, 'Date of Departure'].ffill())
    df.loc[:, 'ts_range'] = df.apply(lambda row: pd.date_range(start=row['Date of Departure']-timedelta(hours=4.5), end=row['Date of Arrival']+timedelta(hours=19.5), freq='12H'), axis=1)
    df = pd.DataFrame({'name':df.Nickname.repeat(df.ts_range.str.len()), 'ts':sum(map(list, df.ts_range), [])})
    df = df.loc[(df.ts >= start) & (df.ts <= end)]
    return df.reset_index(drop=True)


def get_shift_count(year, month):
    df = pd.DataFrame()
    for i in range(1, month):
        ts = pd.to_datetime(date(year, i, 1)).strftime('%b%Y')
        dfi = pd.read_excel('ShiftCount.xlsx', sheet_name=ts, index_col='name')
        dfi = dfi.loc[:,['IOMP-MT', 'IOMP-CT']]
        dfi.index = dfi.index.str.lower()
        df = df.add(dfi, fill_value=0)
    name = pd.read_csv('dyna_staff.csv')
    name.index = name.loc[:, 'name'].str.lower()
    total_shift = pd.merge(df, name, right_index=True, left_index=True, validate='one_to_one', how='right')
    total_shift = total_shift.fillna(0)
    total_shift.loc[total_shift.new == 'new', 'MT'] = total_shift.loc[total_shift.new == 'new', 'IOMP-MT'] + 2
    total_shift.loc[total_shift.new == 'new', 'CT'] = total_shift.loc[total_shift.new == 'new', 'IOMP-CT'] + 2
    total_shift.loc[:, 'total'] = total_shift['IOMP-MT'] + total_shift['IOMP-CT']
    total_shift = total_shift.reset_index(drop=True).set_index('name').loc[:, ['team', 'IOMP-MT', 'IOMP-CT', 'total']]
    return total_shift.sort_index()


def shift_divider(key, next_start, shift_name, recompute=False):
    try:
        allsheet = pd.read_excel('ShiftCount.xlsx', sheet_name=None)
        if shift_name not in allsheet.keys():
            div_shift = True
        else:
            div_shift = False
    except:
        div_shift = True
    
    if div_shift or recompute:
        total_shift = get_shift_count(year, month)
        shift_num = sorted(set(total_shift.loc[total_shift.team != 'admin', 'total']))
    
        num_days = (next_start - timedelta(1)).day
        
        shift_count = total_shift.loc[:, ['team']].reset_index()
        shift_count.loc[shift_count.team.isin(['MT', 'CT']), 'IOMP-MT'] = 1
        shift_count.loc[shift_count.team == 'admin', 'IOMP-MT'] = 0
        shift_count.loc[:, 'IOMP-CT'] = 1
        
        for least_num in shift_num:
            least_shift = total_shift.loc[total_shift.total == least_num, :].reset_index()
            # CT shifts
            CT_least_shift = least_shift[least_shift.team == 'CT']
            if len(CT_least_shift) <= 2*num_days - sum(shift_count['IOMP-CT']):
                shift_count.loc[shift_count.name.isin(CT_least_shift.name), 'IOMP-CT'] += 1
            else:
                num_CT = int(2*num_days - sum(shift_count['IOMP-CT']))
                CT_list = sorted(CT_least_shift.name)
                random.shuffle(CT_list)
                shift_count.loc[shift_count.name.isin(CT_list[0:num_CT]), 'IOMP-CT'] += 1
            # MT shifts
            MT_least_shift = least_shift[least_shift.team == 'MT']
            if len(MT_least_shift) <= 2*num_days - sum(shift_count['IOMP-MT']):
                shift_count.loc[shift_count.name.isin(MT_least_shift.name), 'IOMP-MT'] += 1
            else:
                num_MT = int(2*num_days - sum(shift_count['IOMP-MT']))
                MT_list = sorted(MT_least_shift.name)
                random.shuffle(MT_list)
                shift_count.loc[shift_count.name.isin(MT_list[0:num_MT]), 'IOMP-MT'] += 1
    
        writer = pd.ExcelWriter('ShiftCount.xlsx')
        shift_count['team'] = ','.join(shift_count['team'].values).split(',')
        shift_count = shift_count.sort_values('name')
        
        try:
            allsheet = pd.read_excel('ShiftCount.xlsx', sheet_name=None)
            allsheet[shift_name] = shift_count
        except:
            allsheet = {shift_name: shift_count}
        for sheet_name, xlsxdf in allsheet.items():
            xlsxdf.loc[:, ['name', 'team', 'IOMP-MT', 'IOMP-CT']].to_excel(writer, sheet_name, index=False)
        writer.save()
    else:
        shift_count = allsheet[shift_name]

    return shift_count

###############################################################################

def restrict_shift(fieldwork):
    ts = pd.to_datetime(fieldwork['ts'].values[0]) + timedelta(hours=7.5)
    ts_list = [ts-timedelta(0.5), ts, ts+timedelta(0.5)]
    df = pd.DataFrame({'ts': ts_list})
    df['name'] = fieldwork['name'].values[0]
    return df

def holiday_shift(holidays, shiftdf):
    IOMP = ['IOMP-MT', 'IOMP-CT']
    ts = holidays['ts'].values[0]
    shiftdf.loc[shiftdf.ts == ts, IOMP] = holidays.loc[holidays.ts == ts, IOMP].values[0]

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
        prev_year = str(year)
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
    
    df.loc[:, 'salary_week'] = salary_week
    
    return df

def check_weekdayAM(df):
    weekdayAM = pd.to_datetime(df['ts'].values[0]).isocalendar()[2] in range(1, 6) and pd.to_datetime(df['ts'].values[0]).time() == time(7, 30)
    df.loc[:, 'weekdayAM'] = weekdayAM
    return df

def allowed_shift(ts, name, shiftdf):
    ts = pd.to_datetime(ts)
    tdelta = 0.5
    if pd.to_datetime(ts).time() == time(19, 30):
        tdelta += 0.5
    not_allowed = np.concatenate(shiftdf.loc[(shiftdf.ts >= ts - timedelta(tdelta)) & \
                           (shiftdf.ts <= ts + timedelta(tdelta)), ['IOMP-CT', 'IOMP-MT']].values)
    return name not in not_allowed

########################

key = "1UylXLwDv1W1ukT4YNoUGgHCHF-W8e3F8-pIg1E024ho"

year = int(input('year (e.g. 2017): '))
month = int(input('month (1 to 12): '))

curr_start = pd.to_datetime(date(year, month, 1)) + timedelta(hours=7.5)
shift_name = curr_start.strftime('%b%Y')
if month != 1:
    prev_start = pd.to_datetime(date(year, month-1, 1))
else:
    prev_start = pd.to_datetime(date(year-1, 12, 1))
if month != 12:
    next_start = pd.to_datetime(date(year, month+1, 1))
else:
    next_start = pd.to_datetime(date(year+1, 1, 1))

prev_shift = get_shift(key, prev_start.strftime('%B %Y'))

shift_count = shift_divider(key, next_start, shift_name)
non_admin = shift_count.loc[shift_count.team != 'admin', :]
admin_list = list(shift_count.loc[shift_count.team == 'admin', :].index)

###############################################################################
        
#Empty Shift Schedule
shiftTS = pd.date_range(start=curr_start, end=next_start, freq='12H')
shiftdf = pd.DataFrame({'ts': shiftTS, 'IOMP-MT': ['?']*len(shiftTS), 'IOMP-CT': ['?']*len(shiftTS)})
shiftdf = prev_shift.loc[prev_shift.ts == max(prev_shift.ts), :].append(shiftdf, sort=False, ignore_index=True)

# holiday shifts
try:
    holidays = pd.read_excel('HolidayShift.xlsx', sheet_name=str(year))
    holidays = holidays[['ts', 'IOMP-MT', 'IOMP-CT']]
    holidays['ts'] = pd.to_datetime(holidays['ts'])
    holidays = holidays[(curr_start <= pd.to_datetime(holidays['ts'])) & (pd.to_datetime(holidays['ts']) < next_start)]
except:
    holidays = pd.DataFrame(data=None, columns=['ts', 'IOMP-MT', 'IOMP-CT'])

holiday_ts = holidays.groupby('ts', as_index=False)
holiday_ts.apply(holiday_shift, shiftdf=shiftdf)

for MT in holidays['IOMP-MT'].values:
    shift_count.loc[shift_count.name == MT, 'IOMP-MT'] -= 1

for CT in holidays['IOMP-CT'].values:
    shift_count.loc[shift_count.name == CT, 'IOMP-CT'] -= 1

# shift of people with fieldwork
fieldwork = get_field(key, curr_start, next_start)
if len(fieldwork) != 0:
    fieldwork.loc[:, 'id'] = range(len(fieldwork))
    fieldwork_id = fieldwork.groupby('id', as_index=False)
    field_shifts = fieldwork_id.apply(restrict_shift).drop_duplicates(['ts', 'name']).reset_index(drop=True)
    field_shift_count = Counter(field_shifts.loc[:, 'name'].values)
    field_shift_count = pd.DataFrame(field_shift_count.items(), columns=['name', 'field_shift_count'])
    field_shift_count = field_shift_count.loc[~field_shift_count.name.isin(admin_list), :]
    field_shift_count = field_shift_count.sort_values('field_shift_count', ascending=False)


# check if weekday AM shift and/or salary week (last week and week of 15th)
shiftdf_grp = shiftdf.groupby('ts', as_index=False)
shiftdf = shiftdf_grp.apply(check_salary_week).reset_index(drop=True)
shiftdf_grp = shiftdf.groupby('ts', as_index=False)
shiftdf = shiftdf_grp.apply(check_weekdayAM).reset_index(drop=True)

if len(fieldwork) != 0:
    admin_field = set(field_shifts.loc[field_shifts.name.isin(admin_list), 'ts'])
    amy_field = set(field_shifts.loc[field_shifts.name == 'Amy', 'ts'])
else:
    admin_field = set([])
    amy_field = set([])

# shift of ate amy
if shift_count.loc[shift_count.name == 'Amy', 'IOMP-CT'].values[0] != 0:
    not_salary_weekdayAM = sorted(set(shiftdf.loc[shiftdf.weekdayAM & ~shiftdf.salary_week & (shiftdf['IOMP-CT'] == '?'), 'ts']) - admin_field)
    not_salary_weekdayAM = sorted(set(map(pd.to_datetime, not_salary_weekdayAM)) - amy_field)
    allowed = False
    while not allowed:
        ts = random.choice(not_salary_weekdayAM)
        allowed = allowed_shift(ts, 'Amy', shiftdf)
    shiftdf.loc[shiftdf.ts == ts, 'IOMP-CT'] = 'amy'
    shift_count.loc[shift_count.name == 'Amy', 'IOMP-CT'] -= 1

# shift of remaining admin
for admin in shift_count.loc[(shift_count.team == 'admin') & (shift_count['IOMP-CT'] != 0), :].name:
    weekdayAM = shiftdf.loc[shiftdf.weekdayAM & (shiftdf['IOMP-CT'] == '?'), 'ts'].values        
    weekdayAM = sorted(set(map(pd.to_datetime, weekdayAM)) - admin_field)
    allowed = False
    while not allowed:
        ts = random.choice(weekdayAM)
        allowed = allowed_shift(ts, admin, shiftdf)
    shiftdf.loc[shiftdf.ts == ts, 'IOMP-CT'] = admin
    shift_count.loc[shift_count.name == admin, 'IOMP-CT'] -= 1

# shift of with fieldwork
if len(fieldwork) != 0:
    for IOMP in field_shift_count['name'].values:
        for MT in range(int(shift_count.loc[shift_count.name == IOMP, 'IOMP-MT'].values[0])):
            ts_list = sorted(set(shiftdf.loc[shiftdf['IOMP-MT'] == '?', 'ts']) - set(field_shifts.loc[field_shifts.name == IOMP, 'ts']))
            allowed = False
            while not allowed:
                ts = random.choice(ts_list)
                allowed = allowed_shift(ts, IOMP, shiftdf)
            shiftdf.loc[shiftdf.ts == ts, 'IOMP-MT'] = IOMP
            shift_count.loc[shift_count.name == IOMP, 'IOMP-MT'] -= 1
        for CT in range(int(shift_count.loc[shift_count.name == IOMP, 'IOMP-CT'].values[0])):
            ts_list = sorted(set(shiftdf.loc[shiftdf['IOMP-CT'] == '?', 'ts']) - set(field_shifts.loc[field_shifts.name == IOMP, 'ts']))
            allowed = False
            while not allowed:
                ts = random.choice(ts_list)
                allowed = allowed_shift(ts, IOMP, shiftdf)
            shiftdf.loc[shiftdf.ts == ts, 'IOMP-CT'] = IOMP
            shift_count.loc[shift_count.name == IOMP, 'IOMP-CT'] -= 1
                           
# shift of remaining IOMP

#remaining MT
remaining_MT = []
for IOMP in sorted(set(shift_count.loc[shift_count['IOMP-MT'] != 0, 'name'].values)):
    remaining_MT += [IOMP]*shift_count.loc[shift_count.name == IOMP, 'IOMP-MT'].values[0]

random.shuffle(remaining_MT)

for IOMP in remaining_MT:
    ts_list = sorted(shiftdf.loc[shiftdf['IOMP-MT'] == '?', 'ts'])
    allowed = False
    while not allowed:
        ts = random.choice(ts_list)
        allowed = allowed_shift(ts, IOMP, shiftdf)
    shiftdf.loc[shiftdf.ts == ts, 'IOMP-MT'] = IOMP
    shift_count.loc[shift_count.name == IOMP, 'IOMP-MT'] -= 1

#remaining CT
remaining_CT = []
for IOMP in sorted(set(shift_count.loc[shift_count['IOMP-CT'] != 0, 'name'].values)):
    remaining_CT += [IOMP]*shift_count.loc[shift_count.name == IOMP, 'IOMP-CT'].values[0]

random.shuffle(remaining_CT)

for IOMP in remaining_CT:
    ts_list = sorted(shiftdf.loc[shiftdf['IOMP-CT'] == '?', 'ts'])
    allowed = False
    while not allowed:
        ts = random.choice(ts_list)
        allowed = allowed_shift(ts, IOMP, shiftdf)
    shiftdf.loc[shiftdf.ts == ts, 'IOMP-CT'] = IOMP
    shift_count.loc[shift_count.name == IOMP, 'IOMP-CT'] -= 1
                   
shiftdf = shiftdf.loc[(shiftdf.ts >= curr_start) & (shiftdf.ts < next_start), ['ts', 'IOMP-MT', 'IOMP-CT']]

print(shiftdf)

writer = pd.ExcelWriter('MonitoringShift.xlsx')
try:
    allsheet = pd.read_excel('MonitoringShift.xlsx', sheet_name=None)
    allsheet[shift_name] = shiftdf
except:
    allsheet = {shift_name: shiftdf}
for sheet_name, xlsxdf in allsheet.items():
    xlsxdf.to_excel(writer, sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
writer.save()