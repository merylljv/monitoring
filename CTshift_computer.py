import pandas as pd
import numpy as np

# NOTE: names in shifts_Mo.csv should be arranged such that Monitoring team comes first

total = int(raw_input('enter total number of days this month: '))*2
CT_shifts = 0

df  = pd.read_csv('field_CT.csv',index_col=0,names=['name', 'field_days'])

for CT in df.index:
    df.ix[df.index == CT, 'take_shifts'] = 4 - round(int(df.loc[df.index == CT]['field_days'])/5., 0)
    CT_shifts += 4 - round(int(df.loc[df.index == CT]['field_days'])/5., 0)
    
print df

CT_excess = total - CT_shifts

MT = 11

if CT_excess > MT * 2:
    MT_Shifts = int(CT_excess/MT)
    S_Shifts = int(CT_excess - int(CT_excess/MT) * MT)
elif CT_excess < 11:
    S_Shifts = CT_excess
else:
    MT_Shifts = 1
    S_Shifts = int(CT_excess - MT)

print 'CT shifts from MT = ', MT_Shifts, 'each'
print 'CT shifts from Senslope = ', S_Shifts