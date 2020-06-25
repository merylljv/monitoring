# -*- coding: utf-8 -*-
"""
Created on Wed May 27 23:13:33 2020

@author: Meryll
"""

import pandas as pd

sched = pd.read_csv("covid_sched.csv")

fieldwork = pd.DataFrame()

for col in (set(sched.columns) - set(["name"])):
    names = sched.loc[sched[col].isin(["K", "L"]), "name"]
    ts = "7/{}/2020".format(col)
    fieldwork = fieldwork.append(pd.DataFrame({"name": names,
                                               "ts": [ts]*len(names)}),
                                ignore_index=True)
    
fieldwork = fieldwork.sort_values(["name", "ts"])

fieldwork.to_csv("fieldwork.csv", index=False)