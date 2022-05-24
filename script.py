# -*- coding: UTF-8 -*-
"""
Script.py
---------
Author: Emil Bjørn, 2022-05-22
https://github.com/EmilBjorn/MicrosoftTeams-Shifts-to-ICS

The purpose of this script is to convert shifts from 
the exportable format of Microsoft Teams shift module 
into the standardised iCalendar format used by many common 
calendar apps.
"""
#%%
import pandas as pd

usermail = 'embn@dksund.dk'
df = pd.read_excel('test\ScheduleReport.xlsx', sheet_name=0)

with open('output.ics', mode='a', encoding='utf-8') as f:

    pass

# Inspiration til hvordan man kan loope over rækkerne
# Det er ikke særligt hurtigt, men performance burde ikke være et problem(?)
# for i in df[:10].iterrows():
#   print(i[1]['Gruppe'])
