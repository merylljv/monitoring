# -*- coding: utf-8 -*-
"""
Created on Wed Jan 22 13:41:37 2020

@author: Data Scientist 1
"""

from datetime import timedelta
import pandas as pd

def get_sheet(key, sheet_name):
    url = 'https://docs.google.com/spreadsheets/d/{key}/gviz/tq?tqx=out:csv&sheet={sheet_name}&headers=1'.format(
        key=key, sheet_name=sheet_name.replace(' ', '%20'))
    df = pd.read_csv(url)
    df = df.drop([col for col in df.columns if col.startswith('Unnamed')], axis=1)
    return df

def get_shift(sheet_name):
    key = "1UylXLwDv1W1ukT4YNoUGgHCHF-W8e3F8-pIg1E024ho"
    df = get_sheet(key, sheet_name)
    df = df.drop([col for col in df.columns if col.startswith('Unnamed')], axis=1)
    df.loc[:, 'Date'] = pd.to_datetime(df.loc[:, 'Date'].ffill())
    df.loc[:, 'Shift'] = pd.to_timedelta(df.loc[:, 'Shift'].map({'AM': timedelta(hours=7.5), 'PM': timedelta(hours=19.5)}))
    df.loc[:, 'ts'] = pd.to_datetime(df.loc[:, 'Date'] + df.loc[:, 'Shift'])
    df = df.rename(columns={'IOMP-MT': 'MT', 'IOMP-CT': 'CT'})
    return df.loc[:, ['ts', 'MT', 'CT']]

def get_field():
    key = "1cXUaikP9ZjIpraPl4MrG5YWHmp_a6VbcNlxmmD7_ivs"
    sheet_name = "Sched"
    field = get_sheet(key, sheet_name)
    sheet_name = "Personnel"
    name = get_sheet(key, sheet_name)
    df = pd.merge(field, name, left_on='Personnel', right_on='Fullname')
    df.loc[:, 'Date of Arrival'] = pd.to_datetime(df.loc[:, 'Date of Arrival'])
    df.loc[:, 'Date of Departure'] = pd.to_datetime(df.loc[:, 'Date of Departure'])
    df.loc[:, 'ts_range'] = df.apply(lambda row: pd.date_range(start=row['Date of Departure']-timedelta(hours=4.5), end=row['Date of Arrival']+timedelta(hours=19.5), freq='12H'), axis=1)
    return df.loc[:, ['Nickname', 'Date of Departure', 'Date of Arrival', 'ts_range']]

shift = get_shift('January 2020')
field = get_field()