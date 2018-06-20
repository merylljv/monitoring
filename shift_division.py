import pandas as pd
import random
from calendar import monthrange


year = int(raw_input('year (e.g. 2017): '))
month = int(raw_input('month (1 to 12): '))

total_days = int(monthrange(year, month)[1])
startTS = pd.to_datetime(str(year) + '-' + str(month) + '-01 07:30')
endTS = pd.to_datetime(str(year) + '-' + str(month) + '-' + str(total_days) + ' 19:30')
date = startTS.strftime('%b%Y')

try:
    holidays = pd.read_excel('HolidayShift.xlsx', sheetname=str(year))
    holidays = holidays[['ts', 'IOMP-MT', 'IOMP-CT']]
    holidays['ts'] = pd.to_datetime(holidays['ts'])
    holidays = holidays[(startTS <= pd.to_datetime(holidays['ts'])) & (pd.to_datetime(holidays['ts']) <= endTS)]
except:
    holidays = pd.DataFrame(data=None, columns=['ts', 'IOMP-MT', 'IOMP-CT'])
holidays['IOMP-CT'] = holidays['IOMP-CT'].apply(lambda x: x.lower())
holidays['IOMP-MT'] = holidays['IOMP-MT'].apply(lambda x: x.lower())

shift_count = pd.read_csv('dyna_staff.csv', names=['name', 'team'])
shift_count['MT'] = 0
shift_count['CT'] = 0
           
for CT in holidays['IOMP-CT'].values:
    shift_count.loc[shift_count.name == CT, 'CT'] += 1

# 1 shift for admin, min_shift for others
admin_shift = shift_count[shift_count.team == 'admin']
admin_shift['CT'] = 1

min_shift = int(total_days * 4 - len(admin_shift)) / int(len(shift_count) - len(admin_shift))

pureCT_shift = shift_count[shift_count.team.isin(['S1', 'CT'])]
pureCT_shift['CT'] = min_shift
both_shift = shift_count[~shift_count.team.isin(['S1', 'CT', 'admin'])]
both_shift['MT'] = min_shift - both_shift['CT']
shift_count = admin_shift.append(pureCT_shift).append(both_shift)

# random personnel to not take excess shift
num_no_excess = (len(shift_count) - len(admin_shift)) - (total_days * 4 - (sum(shift_count['MT']) + sum(shift_count['CT'])))

try:
    no_excess = pd.read_csv('no_excess_shift.csv', names=['date', 'name'])
except:
    no_excess = pd.DataFrame(data=None, columns=['name'])
    
if len(no_excess) == len(both_shift) + len(pureCT_shift) or len(no_excess) == 0:
    IOMP = shift_count[shift_count.team != 'admin']['name'].values
    random.shuffle(IOMP)
    curr_no_excess = pd.DataFrame({'name': IOMP[0:num_no_excess]})
    curr_no_excess['date'] = date
    curr_no_excess = curr_no_excess[['date', 'name']]
    curr_no_excess.to_csv('no_excess_shift.csv', header=False, index=False)
elif len(no_excess) <= len(both_shift) + len(pureCT_shift) - num_no_excess:
    IOMP = sorted(set(shift_count[shift_count.team != 'admin']['name']) - set(no_excess['name']))
    random.shuffle(IOMP)
    curr_no_excess = pd.DataFrame({'name': IOMP[0:num_no_excess]})
    curr_no_excess['date'] = date
    curr_no_excess = curr_no_excess[['date', 'name']]
    curr_no_excess.to_csv('no_excess_shift.csv', header=False, mode='a', index=False)
else:
    curr_no_excess = pd.DataFrame({'name': sorted(set(shift_count[shift_count.team != 'admin']['name']) - set(no_excess['name']))})
    IOMP = shift_count[shift_count.team != 'admin']['name'].values
    random.shuffle(IOMP)
    curr_no_excess2 = pd.DataFrame({'name': IOMP[0:num_no_excess-len(curr_no_excess)]})
    curr_no_excess2['date'] = date
    curr_no_excess2 = curr_no_excess2[['date', 'name']]
    curr_no_excess2.to_csv('no_excess_shift.csv', header=False, index=False)
    curr_no_excess = curr_no_excess.append(curr_no_excess2)

shift_count.loc[shift_count.name.isin(set(pureCT_shift['name']) - set(curr_no_excess['name'])), 'CT'] = min_shift + 1
shift_count.loc[shift_count.name.isin(set('biboy') - set(curr_no_excess['name'])), 'CT'] += 1
CTexcess = total_days * 2 - sum(shift_count['CT'])
MT = sorted(set(both_shift['name']) - set('biboy') - set(curr_no_excess['name']))
random.shuffle(MT)
MT_to_CT = MT[0:CTexcess]
shift_count.loc[shift_count.name.isin(MT_to_CT), 'CT'] += 1
MT = MT[CTexcess::]
shift_count.loc[shift_count.name.isin(MT), 'MT'] += 1

############################## shift count (excel) #############################                                           
writer = pd.ExcelWriter('ShiftCount.xlsx')
shift_count['team'] = ','.join(shift_count['team'].values).replace('CTS', 'CT').split(',')
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