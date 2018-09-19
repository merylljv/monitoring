from collections import Counter
from calendar import monthrange
from datetime import time, timedelta, datetime
import numpy as np
import pandas as pd
import random

def restrict_shift(fieldwork):
    ts = pd.to_datetime(fieldwork['ts'].values[0]) + timedelta(hours=7.5)
    ts_list = [ts-timedelta(0.5), ts, ts+timedelta(0.5)]
    df = pd.DataFrame({'ts': ts_list})
    df['name'] = fieldwork['name'].values[0]
    return df


def allowed_shift(ts, name, shiftdf):
    ts = pd.to_datetime(ts)
    tdelta = 0.5
    if pd.to_datetime(ts).time() == time(19, 30):
        tdelta += 0.5
    not_allowed = np.concatenate(shiftdf[(shiftdf.ts >= ts - timedelta(tdelta)) & \
                           (shiftdf.ts <= ts + timedelta(tdelta))][['IOMP-CT', 'IOMP-MT']].values)
    return name not in not_allowed


shiftdf = pd.read_csv('for_backup.csv')
shiftdf['ts'] = pd.to_datetime(shiftdf['ts'])
shiftdf['IOMP-CT'] = shiftdf['IOMP-CT'].apply(lambda x: x.lower())
shiftdf['IOMP-MT'] = shiftdf['IOMP-MT'].apply(lambda x: x.lower())
shiftdf['backup'] = '?'

try:
    prev_backup = pd.read_csv('backup.csv')
except:
    prev_backup = pd.DataFrame(data=None, columns=['ts', 'backup'])
dyna_backup = pd.read_csv('dyna_staff.csv', names=['name', 'team'])

num_backup = len(shiftdf)

if len(prev_backup) == 0:
    IOMP = dyna_backup['name'].values
    random.shuffle(IOMP)
    IOMP_backup = IOMP[0:num_backup]
elif len(dyna_backup) >= len(prev_backup) + num_backup:
    IOMP = sorted(set(dyna_backup['name']) - set(prev_backup['name']))
    random.shuffle(IOMP)
    IOMP_backup = IOMP[0:num_backup]
else:
    curr_no_backup = sorted(set(dyna_backup['name']) - set(prev_backup['name']))
    IOMP = dyna_backup['name'].values
    random.shuffle(IOMP)
    IOMP_backup = list(IOMP[0:num_backup-len(curr_no_backup)]) + curr_no_backup
    IOMP_backup.shuffle(IOMP_backup)
    prev_backup = pd.DataFrame(data=None, columns=['ts', 'backup'])


# shifts of with fieldwork
fieldwork = pd.read_csv('Monitoring Shift Schedule - fieldwork.csv')
fieldwork['ts'] = pd.to_datetime(fieldwork['ts'])
start_ts = pd.to_datetime(pd.to_datetime(min(shiftdf.ts)).date())
fieldwork = fieldwork[(fieldwork.ts >= start_ts) & (fieldwork.ts <= max(shiftdf.ts))]

if len(fieldwork) != 0:
    fieldwork['name'] = fieldwork['name'].apply(lambda x: x.lower())
    fieldwork['id'] = range(len(fieldwork))
    fieldwork_id = fieldwork.groupby('id', as_index=False)
    field_shifts = fieldwork_id.apply(restrict_shift).drop_duplicates(['ts', 'name']).reset_index(drop=True)
    
    field_shift_count = Counter(field_shifts.name)
    field_shift_count = pd.DataFrame({'name': field_shift_count.keys(), 'field_shift_count': field_shift_count.values()})
    field_shift_count = field_shift_count[field_shift_count.name.isin(IOMP_backup)]
    field_shift_count = field_shift_count.sort_values('field_shift_count', ascending=False)

    for IOMP in field_shift_count['name'].values:
        for backup in range(IOMP_backup.count(IOMP)):
            ts_list = sorted(set(shiftdf[shiftdf['backup'] == '?']['ts']) - set(field_shifts[field_shifts.name == IOMP]['ts']))
            allowed = False
            while not allowed:
                ts = random.choice(ts_list)
                allowed = allowed_shift(ts, IOMP, shiftdf)
            shiftdf.loc[shiftdf.ts == ts, 'backup'] = IOMP
            IOMP_backup.remove(IOMP)
      
                 
# shifts of remaining IOMP
for IOMP in IOMP_backup:
    ts_list = sorted(shiftdf[shiftdf['backup'] == '?']['ts'])
    allowed = False
    while not allowed:
        ts = random.choice(ts_list)
        allowed = allowed_shift(ts, IOMP, shiftdf)
    shiftdf.loc[shiftdf.ts == ts, 'backup'] = IOMP

shiftdf = shiftdf[['ts', 'backup']]
shiftdf['backup'] = shiftdf['backup'].apply(lambda x: x[0].upper()+x[1:len(x)])
print shiftdf

##################################### EXCEL ####################################

prev_backup = prev_backup.append(shiftdf)
prev_backup.to_csv('backup.csv')