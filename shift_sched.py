"""

Works with the assumptions that total number of non-admin personnel does not exceed
number of admin personnel less than twice the number of minimum days in a month.
Equivalent to:
    1) 2 shift each for admin personnel 
    2) 3 shift each for lrap personnel
    3) at least 2 shifts for non-admin personnel

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
    personnel_sheet = "personnel1"
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
    """get required shift count for a month"""
    
    df = pd.DataFrame(columns=['IOMP-MT', 'IOMP-CT']).rename_axis('name')
    for i in range(1, month):
        ts = pd.to_datetime(date(year, i, 1)).strftime('%b%Y')
        dfi = pd.read_excel('ShiftCount.xlsx', sheet_name=ts, index_col='name')
        dfi = dfi.loc[:,['IOMP-MT', 'IOMP-CT']]
        df = df.add(dfi, fill_value=0)
    personnel_sheet = "personnel1"
    name = get_sheet(key, personnel_sheet)
    name = name.rename(columns={'Nickname': 'name'})
    name = name.set_index('name')
    total_shift = pd.merge(df, name, right_index=True, left_index=True, validate='one_to_one', how='right')
    total_shift = total_shift.loc[total_shift.current == 1, :]
    admin = len(total_shift.loc[total_shift.team == 'admin', :])
    lrap = len(total_shift.loc[total_shift.team == 'lrap', :])
    eq_shift = ((4*pd.to_datetime('{}-12-31'.format(year)).dayofyear-((24*admin)+(36*lrap)))/(len(total_shift)-(admin+lrap))+backup)/24
    total_shift = total_shift.fillna(0)
    # offset new
    total_shift.loc[:, 'IOMP-MT'] = total_shift.loc[:, 'IOMP-MT'] + total_shift['new'].apply(lambda x: np.ceil(eq_shift*x))
    total_shift.loc[:, 'IOMP-CT'] = total_shift.loc[:, 'IOMP-CT'] + total_shift['new'].apply(lambda x: np.ceil(eq_shift*x))
    total_shift.loc[:, 'total'] = total_shift['IOMP-MT'] + total_shift['IOMP-CT']
    total_shift = total_shift.loc[:, ['team', 'AM_shifts', 'IOMP-MT', 'IOMP-CT', 'total']]
    return total_shift.sort_index()

def shift_divider(key, year, month, next_start, shift_name, recompute=False):
    """assign number of shifts for each personnel for month to be generated
    
        args:
            recompute (bool) - false if shift count is already generated
    
    """
    
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
        shift_num = sorted(set(total_shift.loc[~total_shift.team.isin(['admin', 'lrap']), 'total']))
        num_days = (next_start - timedelta(1)).day
        shift_count = total_shift.loc[:, ['team', 'AM_shifts']].reset_index()
        
        #Add 2 CT shift to LRAP and admin, add 1 CT and 1 MT shift to remaining staff
        shift_count.loc[shift_count.team.isin(['MT', 'CT', 'lrap']), 'IOMP-MT'] = 1
        shift_count.loc[shift_count.team == 'admin', 'IOMP-MT'] = 0
        shift_count.loc[:, 'IOMP-CT'] = 1 
        shift_count.loc[shift_count.team.isin(['admin', 'lrap']), 'IOMP-CT'] += 1        
        
        for least_num in shift_num:
            rem_MT = int(2 * num_days - sum(shift_count['IOMP-MT']))  # Remaining MT shifts
            rem_CT = int(2 * num_days - sum(shift_count['IOMP-CT']))  # Remaining CT shifts
            
            if rem_MT + rem_CT == 0:  # If no remaining shifts
                print("no remaining shift to be assigned")
                break
            
            least_shift = total_shift.loc[(total_shift.total == least_num) & (~total_shift.team.isin(['admin', 'lrap'])), :].reset_index()
        
            
            CT_least_shift = least_shift[least_shift.team == 'CT']
            MT_least_shift = least_shift[least_shift.team == 'MT']
        
            # Assign shifts to CT personnel first
            ct_assign = min(len(CT_least_shift), rem_CT)
            shift_count.loc[shift_count.name.isin(CT_least_shift.name[:ct_assign]), 'IOMP-CT'] += 1
            rem_CT -= ct_assign  # Update remaining CT shifts
        
            # Assign shifts to MT personnel second
            mt_assign = min(len(MT_least_shift), rem_MT)
            shift_count.loc[shift_count.name.isin(MT_least_shift.name[:mt_assign]), 'IOMP-MT'] += 1
            rem_MT -= mt_assign  # Update remaining MT shifts
            
            # If still remaining shifts, randomly assign
            while rem_CT + rem_MT > 0:
                # Sort the least_shift dataframe based on the total shifts (CT and MT) in ascending order
                least_shift['total_shifts'] = least_shift.apply(
                    lambda row: (
                        shift_count.loc[shift_count.name == row.name, 'IOMP-CT'].values[0] 
                        if not shift_count.loc[shift_count.name == row.name].empty else 0
                    ) + (
                        shift_count.loc[shift_count.name == row.name, 'IOMP-MT'].values[0]
                        if not shift_count.loc[shift_count.name == row.name].empty else 0
                    ), axis=1
                )
                least_shift = least_shift.sort_values(by='total_shifts')  # Sort by the total shifts in ascending order
                
                CT_least_shift = least_shift[least_shift.team == 'CT']
                MT_least_shift = least_shift[least_shift.team == 'MT']
            
                # Re-assign remaining CT shifts to those with the least total shifts
                ct_assign = min(len(CT_least_shift), rem_CT)
                shift_count.loc[shift_count.name.isin(CT_least_shift.name[:ct_assign]), 'IOMP-CT'] += 1
                rem_CT -= ct_assign  # Update remaining CT shifts
                
                # Re-assign remaining T shifts to those with the least total shifts
                mt_assign = min(len(MT_least_shift), rem_MT)
                shift_count.loc[shift_count.name.isin(MT_least_shift.name[:mt_assign]), 'IOMP-MT'] += 1
                rem_MT -= mt_assign  # Update remaining MT shifts
                    
        count_file = 'ShiftCount.xlsx'
    
        shift_count['team'] = ','.join(shift_count['team'].values).split(',')
        shift_count = shift_count.sort_values('name')
        
        try:
            with pd.ExcelWriter(count_file, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
                shift_count.loc[:, ['name', 'team', 'AM_shifts', 'IOMP-MT', 'IOMP-CT']].to_excel(
                    writer, sheet_name=shift_name, index=False
                )
        except FileNotFoundError:
            # If the file doesn't exist, create it and add the first sheet
            with pd.ExcelWriter(count_file, engine='openpyxl', mode='w') as writer:
                shift_count.loc[:, ['name', 'team', 'AM_shifts', 'IOMP-MT', 'IOMP-CT']].to_excel(
                    writer, sheet_name=shift_name, index=False)

    return shift_count

##########################################################################################################################

def allowed_shifts(name, shiftdf, shift_type, curr_start, next_start, admin_list, fieldwork, satPM):
    """returns a sorted list of timestamps representing the allowed shifts for a personnel"""
    
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
    shift_list = sorted(set(shift_list) - set(not_allowed) - set(fieldwork.loc[fieldwork.name == name, 'ts']) - satPM)
    return shift_list


def assign_shift(name, shift_count, shiftdf, curr_start, next_start, admin_list, fieldwork, satPM=set()):
    """assign a name on each shift ts"""
    
    if name in shift_count.loc[shift_count.AM_shifts == 1, 'name'].values:
        PM = set(shiftdf.ts[(shiftdf.ts.dt.time == time(19, 30))])
        weekend = shiftdf.ts[shiftdf.ts.apply(lambda x: pd.to_datetime(x).isocalendar()[2]).isin([6,7])]
        satPM = set(sorted(satPM) + sorted(PM) + sorted(weekend))
        
    shift_data = shift_count.loc[shift_count.name == name, ['IOMP-MT', 'IOMP-CT']]
    if not shift_data.empty:
        unassigned_shift = shift_data.to_dict(orient='records')[0]
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
    print('### holiday shifts assigned ###')
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
    print('### admin shifts assigned ###')
    for name in shift_count.loc[(shift_count.team.isin(['admin', 'lrap']) & (shift_count['IOMP-CT'] != 0)), 'name'].values:
        shiftdf, shift_count = assign_shift(name, shift_count, shiftdf, curr_start, next_start, admin_list, fieldwork, satPM)
    return shiftdf, shift_count

def assign_with_fieldwork(fieldwork, shiftdf, shift_count, admin_list, curr_start, next_start, satPM):
    print('### fieldwork shifts assigned ###')
    field_shift_count = Counter(fieldwork.loc[:, 'name'].values)
    field_shift_count = pd.DataFrame(field_shift_count.items(), columns=['name', 'field_shift_count'])
    field_shift_count = field_shift_count.loc[~field_shift_count.name.isin(admin_list), :]
    field_shift_count = field_shift_count.sort_values('field_shift_count', ascending=False)
    with_field = set(field_shift_count.name) - set(shift_count.loc[shift_count.apply(lambda row: (row['IOMP-MT']+row['IOMP-CT'])==0, axis=1), 'name'])
    for name in with_field - set(np.ndarray.flatten(shiftdf.loc[shiftdf.ts > curr_start, ['IOMP-MT', 'IOMP-CT']].values)):
        shiftdf, shift_count = assign_shift(name, shift_count, shiftdf, curr_start, next_start, admin_list, fieldwork, satPM)
    return shiftdf, shift_count

def assign_remaining_IOMP(shiftdf, shift_count):
    print('### remaining IOMP shifts assigned ###')
    extra_CT = shift_count.loc[shift_count.index.repeat(shift_count['IOMP-CT'])]['name'].tolist()
    extra_MT = shift_count.loc[shift_count.index.repeat(shift_count['IOMP-MT'])]['name'].tolist()
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
    
    return shiftdf

def main(key, previous_vpl, recompute=False):
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
    admin_list = list(shift_count.loc[shift_count.team.isin(['admin', 'lrap']), :].name)
    fieldwork = get_field(key, curr_start, next_start)
            
    # Empty shift schedule
    shiftTS = pd.date_range(start=curr_start, end=next_start, freq='12H')
    shiftdf = pd.DataFrame({'ts': shiftTS, 'IOMP-MT': ['?']*len(shiftTS), 'IOMP-CT': ['?']*len(shiftTS)})
    shiftdf = pd.concat([prev_shift.loc[prev_shift.ts == max(prev_shift.ts), :], shiftdf], sort=False, ignore_index=True)
    holidays = get_holidays(curr_start, next_start)
    satPM = set()
    
    # Holiday shifts
    shiftdf, shift_count = assign_with_holiday_shifts(holidays, shiftdf, shift_count, year, curr_start, next_start, admin_list, fieldwork, satPM)
    
    # Shift of admin
    shiftdf, shift_count = assign_admin_shifts(shiftdf, shift_count, admin_list, curr_start, next_start, fieldwork, satPM)
    
    # Shift of personnel with fieldwork
    shiftdf, shift_count = assign_with_fieldwork(fieldwork, shiftdf, shift_count, admin_list, curr_start, next_start, satPM)

    # Shift of remaining personnel
    shiftdf = assign_remaining_IOMP(shiftdf, shift_count)
    
    total_shift = get_shift_count(year, month+1, key)
    vpl = total_shift.loc[total_shift.total == sorted(set(total_shift.total))[1]].index
    vpl = vpl[~vpl.isin(previous_vpl)]
    print('### VPL: ### \n', '\n'.join(sorted(vpl)))    
    
    # Write in xlsx
    shift_file = 'MonitoringShift.xlsx'
    new_sheet = shiftdf.loc[(shiftdf.ts >= curr_start), :]
    try:
        with pd.ExcelWriter(shift_file, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
            new_sheet.to_excel(writer, sheet_name=shift_name, index=False)
    except FileNotFoundError:
        with pd.ExcelWriter(shift_file, engine='openpyxl', mode='w') as writer:
            new_sheet.to_excel(writer, sheet_name=shift_name, index=False)
    
    return shiftdf, shift_count, fieldwork


def shift_validity(shiftdf, shift_count, fieldwork):
    for name in shift_count.name:
        ts_list = sorted(shiftdf.loc[((shiftdf['IOMP-MT'] == name) | (shiftdf['IOMP-CT'] == name)), 'ts'].values)
        week_list = list(map(lambda x: pd.to_datetime(x).isocalendar()[1], ts_list))
        if len(week_list) != len(set(week_list)):
            df = pd.DataFrame({'ts': ts_list, 'week': week_list})
            df.loc[df.ts.dt.time == time(7,30), 'shift_count'] = 1
            df.loc[df.ts.dt.time == time(19,30), 'shift_count'] = 2
            if any(df.groupby('week')['shift_count'].agg('sum') > 4):
                print(name, ': same week!', ts_list)
        ts_index = np.arange(len(ts_list))[pd.Series(ts_list).diff() < timedelta(1.5)]
        if len(ts_index) > 0:
            for i in ts_index:
                ts = pd.to_datetime(ts_list[i])
                prev_ts = pd.to_datetime(ts_list[i-1])
                allowed_ts = ts - timedelta(1)
                if ts.time() == time(19,30):
                    allowed_ts -= timedelta(0.5)
                if prev_ts > allowed_ts:
                    print(name, ': consecutive shift!', prev_ts, ts)
        if len(set(ts_list) - set(fieldwork.loc[fieldwork.name == name, 'ts'].values)) != len(ts_list):
            print(name,': fieldwork!', ts_list)      
            
###############################################################################
    
if __name__ == "__main__":
    start_time = datetime.now()
    
    previous_vpl = ['Jel', 'Marj', 'Leb', 'Kennex', 'Edch']
    key = "1TGXlVe10LuRzAsIFqOjgucn_HI3mQz9t4g8xQdupQpk" #shift sched sheet
    recompute = True
        
    shiftdf, shift_count, fieldwork = main(key, previous_vpl, recompute=recompute)
    
    ########## check shift validity with end of previous shift ##########
    shift_validity(shiftdf, shift_count, fieldwork)
    
    print('runtime =', datetime.now()-start_time)