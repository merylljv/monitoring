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
from datetime import date, time, datetime, timedelta
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
    df.loc[:, 'ts'] = pd.to_datetime(df.loc[:, 'Date'] + df.loc [:, 'Shift'])
    return df.loc[:, ['ts', 'IOMP-MT', 'IOMP-CT']]


def get_field(key, start, end):
    field_sheet = "fieldwork"
    personnel_sheet = "personnel"
    field = get_sheet(key, field_sheet)
    field.loc[:, 'Date of Arrival'] = pd.to_datetime(field.loc[:, 'Date of Arrival'].ffill())
    field.loc[:, 'Date of Departure'] = pd.to_datetime(field.loc[:, 'Date of Departure'].ffill())
    name = get_sheet(key, personnel_sheet)
    df = pd.merge(field, name, left_on='Personnel', right_on='Fullname')
    df = df.loc[~((df['Date of Departure'].isnull()) | (df['Date of Arrival'].isnull())), :]
    df.loc[:, 'ts_range'] = df.apply(lambda row: pd.date_range(start=row['Date of Departure']-timedelta(hours=4.5), end=row['Date of Arrival']+timedelta(hours=19.5), freq='12H'), axis=1)
    df = pd.DataFrame({'name':df.Nickname.repeat(df.ts_range.str.len()), 'ts':sum(map(list, df.ts_range), [])})
    df = df.loc[(df.ts >= start) & (df.ts <= end)]
    return df.reset_index(drop=True)


def get_shift_count(year, month, key, backup=2):
    df = pd.DataFrame(columns=['IOMP-MT', 'IOMP-CT']).rename_axis('name')
    for i in range(1, month):
        ts = pd.to_datetime(date(year, i, 1)).strftime('%b%Y')
        dfi = pd.read_excel('ShiftCount.xlsx', sheet_name=ts, index_col='name')
        dfi = dfi.loc[:,['IOMP-MT', 'IOMP-CT']]
        df = df.add(dfi, fill_value=0)
    personnel_sheet = "personnel"
    name = get_sheet(key, personnel_sheet)
    name = name.rename(columns={'Nickname': 'name'})
    name = name.set_index('name')
    total_shift = pd.merge(df, name, right_index=True, left_index=True, validate='one_to_one', how='right')
    total_shift = total_shift.loc[total_shift.current == 1, :]
#    equal shift
    admin = len(total_shift.loc[total_shift.team == 'admin', :])
    eq_shift = ((4*pd.to_datetime('{}-12-31'.format(year)).dayofyear-12*admin)/(len(total_shift)-admin)+backup)/24
    total_shift = total_shift.fillna(0)
    # offset new
    total_shift.loc[:, 'IOMP-MT'] = total_shift.loc[:, 'IOMP-MT'] + total_shift['new'].apply(lambda x: np.ceil(eq_shift*x))
    total_shift.loc[:, 'IOMP-CT'] = total_shift.loc[:, 'IOMP-CT'] + total_shift['new'].apply(lambda x: np.ceil(eq_shift*x))
    total_shift.loc[:, 'total'] = total_shift['IOMP-MT'] + total_shift['IOMP-CT']
    total_shift = total_shift.loc[:, ['team', 'AM_shifts', 'IOMP-MT', 'IOMP-CT', 'total']]
    return total_shift.sort_index()


def shift_divider(key, year, month, next_start, shift_name, recompute=False):
    try:
        allsheet = pd.read_excel('ShiftCount.xlsx', sheet_name=None)
        if shift_name not in allsheet.keys():
            div_shift = True
        else:
            div_shift = False
    except:
        div_shift = True
    if div_shift or recompute:
        total_shift = get_shift_count(year, month, key)
        shift_num = sorted(set(total_shift.loc[total_shift.team != 'admin', 'total']))
        num_days = (next_start - timedelta(1)).day
        shift_count = total_shift.loc[:, ['team', 'AM_shifts']].reset_index()
        shift_count.loc[shift_count.team.isin(['MT', 'CT']), 'IOMP-MT'] = 1
        shift_count.loc[shift_count.team == 'admin', 'IOMP-MT'] = 0
        shift_count.loc[:, 'IOMP-CT'] = 1 
        for least_num in shift_num:
            rem_MT = int(2*num_days - sum(shift_count['IOMP-MT']))
            rem_CT = int(2*num_days - sum(shift_count['IOMP-CT']))
            if rem_MT + rem_CT == 0:
                break
            least_shift = total_shift.loc[(total_shift.total == least_num) & (total_shift.team != 'admin'), :].reset_index()
            CT_least_shift = least_shift[least_shift.team == 'CT']
            MT_least_shift = least_shift[least_shift.team == 'MT']
            # if remaining CT and MT shifts is more than the CT and MT personnel with least shifts
            if len(MT_least_shift) <= rem_MT and len(CT_least_shift) <= rem_CT:
                shift_count.loc[shift_count.name.isin(CT_least_shift.name), 'IOMP-CT'] += 1
                shift_count.loc[shift_count.name.isin(MT_least_shift.name), 'IOMP-MT'] += 1
            # if total remaining shifts is more than total personnel with least shifts
            elif len(least_shift) <= rem_MT + rem_CT:
                # if remaining CT shifts is more than the CT personnel with least shifts
                if len(CT_least_shift) <= rem_CT:
                    # assign all CT with least shifts to CT shifts
                    shift_count.loc[shift_count.name.isin(CT_least_shift.name), 'IOMP-CT'] += 1
                    # remaining MT shifts is less that MT personnel with least shift and some should be assigned to CT shift
                    CT_list = sorted(MT_least_shift.name)
                    random.shuffle(CT_list)
                    shift_count.loc[shift_count.name.isin(CT_list[0:len(CT_list)-rem_MT]), 'IOMP-CT'] += 1
                    MT_list = set(CT_list) - set(CT_list[0:len(CT_list)-rem_MT])
                    shift_count.loc[shift_count.name.isin(MT_list), 'IOMP-MT'] += 1
                # else, remaining MT shifts is more than the MT personnel with least shifts
                else:
                    # assign all MT with least shifts to MT shifts
                    shift_count.loc[shift_count.name.isin(MT_least_shift.name), 'IOMP-MT'] += 1
                    # remaining CT shifts is less that CT personnel with least shift and some should be assigned to MT shift
                    MT_list = sorted(CT_least_shift.name)
                    random.shuffle(MT_list)
                    shift_count.loc[shift_count.name.isin(MT_list[0:len(MT_list)-rem_CT]), 'IOMP-MT'] += 1
                    CT_list = set(MT_list) - set(MT_list[0:len(MT_list)-rem_CT])
                    shift_count.loc[shift_count.name.isin(CT_list), 'IOMP-CT'] += 1
            # randomly choose who gets additional shifts
            else:
                IOMP_list = sorted(least_shift.name)
                random.shuffle(IOMP_list)
                least_shift = total_shift.loc[total_shift.index.isin(IOMP_list[0:(rem_CT+rem_MT)]), :].reset_index()
                CT_least_shift = least_shift[least_shift.team == 'CT']
                MT_least_shift = least_shift[least_shift.team == 'MT']
                # if remaining CT shifts is more than the CT personnel with least shifts
                if len(CT_least_shift) <= rem_CT:
                    # assign all CT with least shifts to CT shifts
                    shift_count.loc[shift_count.name.isin(CT_least_shift.name), 'IOMP-CT'] += 1
                    # remaining MT shifts is less that MT personnel with least shift and some should be assigned to CT shift
                    CT_list = sorted(MT_least_shift.name)
                    random.shuffle(CT_list)
                    shift_count.loc[shift_count.name.isin(CT_list[0:len(CT_list)-rem_MT]), 'IOMP-CT'] += 1
                    MT_list = set(CT_list) - set(CT_list[0:len(CT_list)-rem_MT])
                    shift_count.loc[shift_count.name.isin(MT_list), 'IOMP-MT'] += 1
                # else, remaining MT shifts is more than the MT personnel with least shifts
                else:
                    # assign all MT with least shifts to MT shifts
                    shift_count.loc[shift_count.name.isin(MT_least_shift.name), 'IOMP-MT'] += 1
                    # remaining CT shifts is less that CT personnel with least shift and some should be assigned to MT shift
                    MT_list = sorted(CT_least_shift.name)
                    random.shuffle(MT_list)
                    shift_count.loc[shift_count.name.isin(MT_list[0:len(MT_list)-rem_CT]), 'IOMP-MT'] += 1
                    CT_list = set(MT_list) - set(MT_list[0:len(MT_list)-rem_CT])
                    shift_count.loc[shift_count.name.isin(CT_list), 'IOMP-CT'] += 1
        with pd.ExcelWriter('ShiftCount.xlsx', engine='openpyxl', mode='a') as writer:
            shift_count['team'] = ','.join(shift_count['team'].values).split(',')
            shift_count = shift_count.sort_values('name')
            shift_count.to_excel(writer, sheet_name=shift_name, index=False)
    else:
        shift_count = allsheet[shift_name]

    return shift_count

###############################################################################

def allowed_shifts(name, shiftdf, shift_type, curr_start, next_start, admin_list, fieldwork, satPM):
    ts_list = shiftdf.loc[(shiftdf.ts > curr_start) & ((shiftdf['IOMP-MT'] == name) | (shiftdf['IOMP-CT'] == name)), 'ts'].values
    week_list = sorted(set(list(map(lambda x: pd.to_datetime(x).isocalendar()[1]*(np.floor(pd.to_datetime(x).hour/7)-1), ts_list))) - set([0]))
    shift_list = sorted(map(lambda x: pd.to_datetime(x), shiftdf.loc[(shiftdf.ts > curr_start) & (shiftdf[shift_type] == '?') & (~shiftdf.ts.apply(lambda x: pd.to_datetime(x).isocalendar()[1]).isin(week_list)), 'ts'].values))
    not_allowed = []
    for ts in ts_list:
        ts = pd.to_datetime(ts)
        not_allowed += [ts]
        not_allowed += [ts+timedelta(0.5), ts-timedelta(0.5)]
        if ts.time() == time(19, 30):
            not_allowed += [ts+timedelta(1), ts-timedelta(1)]
    if name in admin_list:
        not_allowed += sorted(np.array(shift_list)[np.invert(list(map(lambda x: (x.isocalendar()[2] in range(1,6)) & (x.time() == time(7,30)), shift_list)))])
#    if name == 'Amy':
#        date_list = pd.date_range(curr_start, next_start, freq='SM')
#        salary_week = list(map(lambda x: x.isocalendar()[1], date_list))
#        not_allowed += sorted(np.array(shift_list)[list(map(lambda x: x.isocalendar()[1] in salary_week, shift_list))])
    shift_list = sorted(set(shift_list) - set(not_allowed) - set(fieldwork.loc[fieldwork.name == name, 'ts']) - satPM)
    return shift_list


def assign_shift(name, shift_count, shiftdf, curr_start, next_start, admin_list, fieldwork, satPM=set()):
    print(name)
    if name in shift_count.loc[shift_count.AM_shifts == 1, 'name'].values:
        PM = set(shiftdf.ts[(shiftdf.ts.dt.time == time(19, 30))])
        weekend = shiftdf.ts[shiftdf.ts.apply(lambda x: pd.to_datetime(x).isocalendar()[2]).isin([6,7])]
        satPM = set(sorted(satPM) + sorted(PM) + sorted(weekend))
    unassigned_shift = shift_count.loc[shift_count.name == name, ['IOMP-MT', 'IOMP-CT']].to_dict(orient='records')[0]
    for shift_type in unassigned_shift.keys():
        count = unassigned_shift.get(shift_type)
        while count > 0:
            shift_list = allowed_shifts(name, shiftdf, shift_type, curr_start, next_start, admin_list, fieldwork, satPM)
            ts = random.choice(shift_list)
            shiftdf.loc[shiftdf.ts == ts, shift_type] = name
            count -= 1
            shift_count.loc[shift_count.name == name, shift_type] -= 1
    return shiftdf, shift_count


def assign_holiday_shifts(holidays, shiftdf):
    IOMP = ['IOMP-MT', 'IOMP-CT']
    ts = holidays['ts'].values[0]
    shiftdf.loc[shiftdf.ts == ts, IOMP] = holidays.loc[holidays.ts == ts, IOMP].values[0]


def get_holidays(curr_start, next_start):
    try:
        holidays = pd.read_excel('HolidayShift.xlsx', sheet_name=str(curr_start.year))
        holidays = holidays[['ts', 'IOMP-MT', 'IOMP-CT']]
        holidays['ts'] = pd.to_datetime(holidays['ts'])
        holidays = holidays[(curr_start <= pd.to_datetime(holidays['ts'])) & (pd.to_datetime(holidays['ts']) < next_start)]
    except:
        holidays = pd.DataFrame(data=None, columns=['ts', 'IOMP-MT', 'IOMP-CT'])
    return holidays


def assign_with_holiday_shifts(holidays, shiftdf, shift_count, year, curr_start, next_start, admin_list, fieldwork, satPM):
    print('### holiday ###')
    holidays = holidays.fillna('?')
    holiday_ts = holidays.groupby('ts', as_index=False)
    holiday_ts.apply(assign_holiday_shifts, shiftdf=shiftdf)
    
    for MT in holidays['IOMP-MT'].values:
        shift_count.loc[shift_count.name == MT, 'IOMP-MT'] -= 1
    
    for CT in holidays['IOMP-CT'].values:
        shift_count.loc[shift_count.name == CT, 'IOMP-CT'] -= 1
    
    with_holiday_shifts = sorted(set(np.ndarray.flatten(shiftdf.loc[(shiftdf.ts >= curr_start), ['IOMP-MT', 'IOMP-CT']].values)) - set(['?']))
    
    for name in with_holiday_shifts:
        shiftdf, shift_count = assign_shift(name, shift_count, shiftdf, curr_start, next_start, admin_list, fieldwork, satPM)
    
    return shiftdf, shift_count


def assign_admin_shifts(shiftdf, shift_count, admin_list, curr_start, next_start, fieldwork, satPM):
    print('### admin ###')
    for name in shift_count.loc[((shift_count.team == 'admin') & (shift_count['IOMP-CT'] != 0)), 'name'].values:
        shiftdf, shift_count = assign_shift(name, shift_count, shiftdf, curr_start, next_start, admin_list, fieldwork, satPM)
    return shiftdf, shift_count


def assign_with_fieldwork(fieldwork, shiftdf, shift_count, admin_list, curr_start, next_start, satPM):
    print('### fieldwork ###')
    field_shift_count = Counter(fieldwork.loc[:, 'name'].values)
    field_shift_count = pd.DataFrame(field_shift_count.items(), columns=['name', 'field_shift_count'])
    field_shift_count = field_shift_count.loc[~field_shift_count.name.isin(admin_list), :]
    field_shift_count = field_shift_count.sort_values('field_shift_count', ascending=False)
    with_field = set(field_shift_count.name) - set(shift_count.loc[shift_count.apply(lambda row: (row['IOMP-MT']+row['IOMP-CT'])==0, axis=1), 'name'])
    for name in with_field - set(np.ndarray.flatten(shiftdf.loc[shiftdf.ts > curr_start, ['IOMP-MT', 'IOMP-CT']].values)):
        shiftdf, shift_count = assign_shift(name, shift_count, shiftdf, curr_start, next_start, admin_list, fieldwork, satPM)
    return shiftdf, shift_count


def assign_satPM_shifts(shiftdf, shift_count, curr_start, next_start, admin_list, fieldwork, satPM):
    print('### saturday PM ###')
    PM = set(shiftdf.ts[(shiftdf.ts.dt.time == time(19, 30))])
    MT_least_shift = list(shift_count.loc[(shift_count['IOMP-MT'] != 0) & (shift_count['IOMP-MT'] + shift_count['IOMP-CT'] == 2), 'name'].values)
    random.shuffle(MT_least_shift)
    MT_least_shift = MT_least_shift[0:len(satPM)]
    if len(MT_least_shift) < len(satPM):
        temp_list = list(shift_count.loc[(shift_count['IOMP-MT'] != 0) & (shift_count['IOMP-MT'] + shift_count['IOMP-CT'] != 2), 'name'].values)
        random.shuffle(temp_list)
        MT_least_shift += temp_list[0:len(satPM)-len(MT_least_shift)]
    shiftdf.loc[shiftdf.ts.isin(satPM), 'IOMP-MT'] = MT_least_shift
    shift_count.loc[shift_count.name.isin(MT_least_shift), 'IOMP-MT'] -= 1
    for name in MT_least_shift:
        shiftdf, shift_count = assign_shift(name, shift_count, shiftdf, curr_start, next_start, admin_list, fieldwork, satPM=PM)
    CT_least_shift = shift_count.loc[(shift_count['IOMP-CT'] != 0) & (shift_count['IOMP-MT'] + shift_count['IOMP-CT'] == 2), 'name'].values
    CT_least_shift = sorted(set(CT_least_shift) - set(MT_least_shift))
    random.shuffle(CT_least_shift)
    CT_least_shift = CT_least_shift[0:len(satPM)]
    if len(CT_least_shift) < len(satPM):
        temp_list = sorted(set(shift_count.loc[(shift_count['IOMP-CT'] != 0) & (shift_count['IOMP-MT'] + shift_count['IOMP-CT'] != 2), 'name'].values) - set(MT_least_shift))
        random.shuffle(temp_list)
        CT_least_shift += temp_list[0:len(satPM)-len(CT_least_shift)]
    shiftdf.loc[shiftdf.ts.isin(satPM), 'IOMP-CT'] = CT_least_shift
    shift_count.loc[shift_count.name.isin(CT_least_shift), 'IOMP-CT'] -= 1
    for name in CT_least_shift:
        shiftdf, shift_count = assign_shift(name, shift_count, shiftdf, curr_start, next_start, admin_list, fieldwork, satPM=PM)
    return shiftdf, shift_count


def assign_remaining_IOMP(shiftdf, shift_count):
    print('### remaining IOMP ###')
    # unassigned IOMPs
    extra_CT = shift_count.loc[(shift_count['IOMP-CT'] == 2), 'name'].values
    extra_MT = shift_count.loc[(shift_count['IOMP-MT'] == 2), 'name'].values
    remaining_IOMP = sorted(set(shift_count.loc[((shift_count['IOMP-MT'] != 0) & (shift_count['IOMP-CT'] != 0)), 'name'].values) - set(extra_CT) - set(extra_MT))
    # randomize
    random.shuffle(extra_MT)
    random.shuffle(extra_CT)
    random.shuffle(remaining_IOMP)
    
    # shifts with unassigned IOMP
    MT_shift = shiftdf.loc[shiftdf['IOMP-MT'] == '?', 'ts'].values
    CT_shift = shiftdf.loc[shiftdf['IOMP-CT'] == '?', 'ts'].values
    
    # assign extra shifts
    temp_MT_shift = MT_shift[0:len(extra_MT)]
    MT_shift = sorted(set(MT_shift) - set(temp_MT_shift))
    shiftdf.loc[shiftdf.ts.isin(temp_MT_shift), 'IOMP-MT'] = extra_MT
    temp_CT_shift = CT_shift[0:len(extra_CT)]
    CT_shift = sorted(set(CT_shift) - set(temp_CT_shift))
    shiftdf.loc[shiftdf.ts.isin(temp_CT_shift), 'IOMP-CT'] = extra_CT

    # assign remaining shifts
    remaining_IOMP = list(remaining_IOMP) + list(extra_MT) + list(extra_CT)
    MT = remaining_IOMP[1::2]   
    temp_MT_shift = MT_shift[0:len(MT)]
    MT_shift = sorted(set(MT_shift) - set(temp_MT_shift))
    shiftdf.loc[shiftdf.ts.isin(temp_MT_shift), 'IOMP-MT'] = MT
    CT = remaining_IOMP[::2]
    temp_CT_shift = CT_shift[0:len(CT)]
    CT_shift = sorted(set(CT_shift) - set(temp_CT_shift))
    shiftdf.loc[shiftdf.ts.isin(temp_CT_shift), 'IOMP-CT'] = CT
    # flip
    MT = remaining_IOMP[::2]
    shiftdf.loc[shiftdf.ts.isin(MT_shift), 'IOMP-MT'] = MT
    CT = remaining_IOMP[1::2]
    shiftdf.loc[shiftdf.ts.isin(CT_shift), 'IOMP-CT'] = CT
    
    return shiftdf


def assign_schedule(key, previous_vpl, recompute=False):
    now = datetime.now()
    year = now.year
    month = now.month
    if month != 12:
       month += 1
    else:
      month = 1
      year += 1
    
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
    
    shift_count = shift_divider(key, year, month, next_start, shift_name,
                                recompute=recompute)
    admin_list = list(shift_count.loc[shift_count.team == 'admin', :].name)
    fieldwork = get_field(key, curr_start, next_start)
            
    # Empty shift schedule
    shiftTS = pd.date_range(start=curr_start, end=next_start, freq='12H')
    shiftdf = pd.DataFrame({'ts': shiftTS, 'IOMP-MT': ['?']*len(shiftTS), 'IOMP-CT': ['?']*len(shiftTS)})
    shiftdf = pd.concat([prev_shift.loc[prev_shift.ts == max(prev_shift.ts), :],shiftdf], sort=False, ignore_index=True)
    # shiftdf = prev_shift.loc[prev_shift.ts == max(prev_shift.ts), :].append(shiftdf, sort=False, ignore_index=True)
    holidays = get_holidays(curr_start, next_start)
#    satPM = set(shiftdf.ts[(shiftdf.ts.apply(lambda x: pd.to_datetime(x).isocalendar()[2]) == 6) & (shiftdf.ts.dt.month == month) & (shiftdf.ts.dt.time == time(19, 30))])
#    satPM -= set(holidays.ts)
    satPM = set()
    
    # Holiday shifts
    shiftdf, shift_count = assign_with_holiday_shifts(holidays, shiftdf, shift_count, year, curr_start, next_start, admin_list, fieldwork, satPM)
    
    # Shift of admin
    shiftdf, shift_count = assign_admin_shifts(shiftdf, shift_count, admin_list, curr_start, next_start, fieldwork, satPM)
    
    # Shift of personnel with fieldwork
    shiftdf, shift_count = assign_with_fieldwork(fieldwork, shiftdf, shift_count, admin_list, curr_start, next_start, satPM)

#    # Shift with saturday PM
#    shiftdf, shift_count = assign_satPM_shifts(shiftdf, shift_count, curr_start, next_start, admin_list, fieldwork, satPM)
    
    # Shift of remaining personnel
    shiftdf = assign_remaining_IOMP(shiftdf, shift_count)
    
    total_shift = get_shift_count(year, month+1, key)    
    vpl = total_shift.loc[total_shift.total == sorted(set(total_shift.total))[1]].index
    vpl = vpl[~vpl.isin(previous_vpl)]
    print('### VPL: ### \n', '\n'.join(sorted(vpl)))   
    writer = pd.ExcelWriter('MonitoringShift.xlsx')
    try:
        allsheet = pd.read_excel('MonitoringShift.xlsx', sheet_name=None)
        allsheet[shift_name] = shiftdf.loc[(shiftdf.ts >= curr_start), :]
    except:
        allsheet = {shift_name: shiftdf.loc[(shiftdf.ts >= curr_start), :]}
    for sheet_name, xlsxdf in allsheet.items():
        xlsxdf.to_excel(writer, sheet_name, index=False)
    writer.close()    
    
    return shiftdf, shift_count, fieldwork


def shift_validity(shiftdf, shift_count, fieldwork):
    print("### processed ###")
    for name in shift_count.name:
        print(name)
        ts_list = sorted(shiftdf.loc[((shiftdf['IOMP-MT'] == name) | (shiftdf['IOMP-CT'] == name)), 'ts'].values)
        week_list = list(map(lambda x: pd.to_datetime(x).isocalendar()[1], ts_list))
        if len(week_list) != len(set(week_list)):
            df = pd.DataFrame({'ts': ts_list, 'week': week_list})
            df.loc[df.ts.dt.time == time(7,30), 'shift_count'] = 1
            df.loc[df.ts.dt.time == time(19,30), 'shift_count'] = 2
            if any(df.groupby('week')['shift_count'].agg('sum') > 4):
                print('same week!', ts_list)
        ts_index = np.arange(len(ts_list))[pd.Series(ts_list).diff() < timedelta(1.5)]
        if len(ts_index) > 0:
            for i in ts_index:
                ts = pd.to_datetime(ts_list[i])
                prev_ts = pd.to_datetime(ts_list[i-1])
                allowed_ts = ts - timedelta(1)
                if ts.time() == time(19,30):
                    allowed_ts -= timedelta(0.5)
                if prev_ts > allowed_ts:
                    print('consecutive shift!', prev_ts, ts)
        if len(set(ts_list) - set(fieldwork.loc[fieldwork.name == name, 'ts'].values)) != len(ts_list):
            print('fieldwork!', ts_list)
        
            
            
###############################################################################
    
if __name__ == "__main__":
    start_time = datetime.now()
    
    previous_vpl = ['Jel', 'Marj', 'Leb', 'Kennex', 'Edch']
    key = "1UylXLwDv1W1ukT4YNoUGgHCHF-W8e3F8-pIg1E024ho"
    recompute = False
        
    shiftdf, shift_count, fieldwork = assign_schedule(key, previous_vpl, recompute=recompute)
    shift_validity(shiftdf, shift_count, fieldwork)
 
    ########## check shift validity with end of previous shift
    
    print('runtime =', datetime.now()-start_time)

