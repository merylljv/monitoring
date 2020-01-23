"""
works with the following assumptions:
1. scheduled fieldwork for at least 1 IOMP
2. twice the number of IOMP that takes CT shifts only > 
    number of IOMP that takes both MT and CT shifts > 
        number of IOMP that takes CT shifts only
"""

import math
import pandas as pd
import random

from calendar import monthrange


year = int(input('year (e.g. 2017): '))
month = int(input('month (1 to 12): '))

total_days = int(monthrange(year, month)[1])
start_ts = pd.to_datetime(str(year) + '-' + str(month) + '-01 07:30')
end_ts = pd.to_datetime(str(year) + '-' + str(month) + '-' + str(total_days) + ' 19:30')
date = start_ts.strftime('%b%Y')

try:
    holidays = pd.read_excel('HolidayShift.xlsx', sheetname=str(year))
    holidays = holidays[['ts', 'IOMP-MT', 'IOMP-CT']]
    holidays['ts'] = pd.to_datetime(holidays['ts'])
    holidays = holidays[(start_ts <= pd.to_datetime(holidays['ts'])) & (pd.to_datetime(holidays['ts']) <= end_ts)]
except:
    holidays = pd.DataFrame(data=None, columns=['ts', 'IOMP-MT', 'IOMP-CT'])
holidays['IOMP-CT'] = holidays['IOMP-CT'].apply(lambda x: x.lower())
holidays['IOMP-MT'] = holidays['IOMP-MT'].apply(lambda x: x.lower())


try:
    allsheet = pd.read_excel('ShiftCount.xlsx', sheetname=None)
    if date not in allsheet.keys():
        div_shift = True
    else:
        div_shift = False
except:
    div_shift = True

if div_shift:    
    shift_count = pd.read_csv('dyna_staff.csv', names=['name', 'team'])
    shift_count['MT'] = 0
    shift_count['CT'] = 0
               
    for CT in holidays['IOMP-CT'].values:
        shift_count.loc[shift_count.name == CT, 'CT'] += 1
    
    # 1 shift for admin, min_shift for others
    admin_shift = shift_count[shift_count.team == 'admin']
    admin_shift['CT'] = 1
               
    # 3 shifts for S1
    S1_shift = shift_count[shift_count.team == 'S1']
    S1_shift['CT'] = 3
               
    min_shift = math.floor(int(total_days * 4 - len(admin_shift) - (3 * len(S1_shift))) / int(len(shift_count) - len(admin_shift) - len(S1_shift)))
    
    pureCT_shift = shift_count[shift_count.team == 'CT']
    pureCT_shift['CT'] = min_shift
    both_shift = shift_count[~shift_count.team.isin(['S1', 'CT', 'admin'])]
    MT_min = math.floor(min([min_shift, int(total_days * 2) / len(both_shift)]))
    if MT_min == min_shift:
        both_shift['MT'] = MT_min - both_shift['CT']
    else:
        both_shift['MT'] = MT_min - both_shift['CT']
        remaining_MT = total_days * 2 - (len(both_shift) * MT_min)
        remaining_MT_to_CT = len(both_shift) - remaining_MT
        MT_list = sorted(set(both_shift['name']) - set(both_shift[both_shift.CT != 0]['name']))
        random.shuffle(MT_list)
        to_CTshift = MT_list[0:remaining_MT_to_CT]
        both_shift.loc[both_shift.name.isin(to_CTshift), 'CT'] += 1
        both_shift.loc[~both_shift.name.isin(to_CTshift), 'MT'] += 1
    shift_count = admin_shift.append(S1_shift).append(pureCT_shift).append(both_shift)
    
    # random personnel to not take excess shift
    num_no_excess = (len(shift_count) - len(admin_shift) - len(S1_shift)) - (total_days * 4 - (sum(shift_count['MT']) + sum(shift_count['CT'])))
    
    
    if num_no_excess != len(shift_count) - (len(admin_shift) + len(S1_shift)):
        try:
            no_excess = pd.read_csv('no_excess_shift.csv', names=['date', 'name'])
        except:
            no_excess = pd.DataFrame(data=None, columns=['name'])
            
        if len(no_excess) == len(both_shift) + len(pureCT_shift) or len(no_excess) == 0:
            IOMP = shift_count[~shift_count.team.isin(['admin', 'S1'])]['name'].values
            random.shuffle(IOMP)
            curr_no_excess = pd.DataFrame({'name': IOMP[0:num_no_excess]})
            curr_no_excess['date'] = date
            curr_no_excess = curr_no_excess[['date', 'name']]
            curr_no_excess.to_csv('no_excess_shift.csv', header=False, index=False)
        elif len(no_excess) <= len(both_shift) + len(pureCT_shift) - num_no_excess:
            IOMP = sorted(set(shift_count[~shift_count.team.isin(['admin', 'S1'])]['name']) - set(no_excess['name']))
            random.shuffle(IOMP)
            curr_no_excess = pd.DataFrame({'name': IOMP[0:num_no_excess]})
            curr_no_excess['date'] = date
            curr_no_excess = curr_no_excess[['date', 'name']]
            curr_no_excess.to_csv('no_excess_shift.csv', header=False, mode='a', index=False)
        else:
            curr_no_excess = pd.DataFrame({'name': sorted(set(shift_count[~shift_count.team.isin(['admin', 'S1'])]['name']) - set(no_excess['name']))})
            IOMP = shift_count[~shift_count.team.isin(['admin', 'S1'])]['name'].values
            random.shuffle(IOMP)
            curr_no_excess2 = pd.DataFrame({'name': IOMP[0:num_no_excess-len(curr_no_excess)]})
            curr_no_excess2['date'] = date
            curr_no_excess2 = curr_no_excess2[['date', 'name']]
            curr_no_excess2.to_csv('no_excess_shift.csv', header=False, index=False)
            curr_no_excess = curr_no_excess.append(curr_no_excess2)
        CTexcess_pool = sorted(set(pureCT_shift['name']) - set(curr_no_excess['name']))
        numCTexcess = total_days*2 - sum(shift_count.CT)
        try:
            shift_count.loc[shift_count.name.isin(random.sample(CTexcess_pool, numCTexcess)), 'CT'] = min_shift + 1
        except:
            shift_count.loc[shift_count.name.isin(random.sample(CTexcess_pool, len(CTexcess_pool))), 'CT'] = min_shift + 1
        CTexcess = total_days * 2 - sum(shift_count['CT'])
        MT = sorted(set(both_shift['name']) - set(curr_no_excess['name']))
        random.shuffle(MT)
        MT_to_CT = MT[0:CTexcess]
        shift_count.loc[shift_count.name.isin(MT_to_CT), 'CT'] += 1
        MT = MT[CTexcess:]
        numMTexcess = total_days*2 - sum(shift_count.MT)
        shift_count.loc[shift_count.name.isin(random.sample(MT, numMTexcess)), 'MT'] += 1
    
    ############################## shift count (excel) #############################                                           
    writer = pd.ExcelWriter('ShiftCount.xlsx')
    shift_count['team'] = ','.join(shift_count['team'].values).split(',')
    shift_count = shift_count.sort_values(['team', 'name'])
    
    try:
        allsheet = pd.read_excel('ShiftCount.xlsx', sheetname=None)
        allsheet[date] = shift_count
    except:
        allsheet = {date: shift_count}
    for sheetname, xlsxdf in allsheet.items():
        xlsxdf.to_excel(writer, sheetname, index=False)
        worksheet = writer.sheets[sheetname]
    writer.save()