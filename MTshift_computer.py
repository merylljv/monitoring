import pandas as pd
import numpy as np
import CTshift_computer as CT

# NOTE: names in shifts_Mo.csv should be arranged such that Monitoring team comes first

df  = pd.read_csv('field_MT.csv',index_col=0,names=['name', 'field_days'])

total = CT.total
senslope = 48 - CT.S_Shifts
MT = 4
total = total - senslope

excess = np.round((total % len(df))/MT, 0)

df['missed_shifts'] = [np.nan] * len(df)

#priority is equal to the num
for Mo in df.index:
    df.ix[df.index == Mo, 'missed_shifts'] = int(df.loc[df.index == Mo]['field_days'])*2

for Mo in list(df.index):
    if Mo in list(df.index[0:MT]):
        df.ix[df.index == Mo, 'priority'] = (total + excess - df['missed_shifts']) / total
    else:
        df.ix[df.index == Mo, 'priority'] = (total - df['missed_shifts']) / total

equal_shift=total/len(df)
df['take_shifts']=  equal_shift * ((total-df['missed_shifts'])/total)
df['take_shifts'] = df.take_shifts.round()
total = total - sum(df.take_shifts)

df=df.sort('priority', ascending=False)
cnt=0

while total > 0:
    df.ix[cnt, 'take_shifts'] += 1
    total -= 1
#    print 'shift added to ' + df.index[cnt]
    cnt +=1
    if cnt == len(df):
        cnt=0
    
df = df[['field_days', 'take_shifts']]
print df