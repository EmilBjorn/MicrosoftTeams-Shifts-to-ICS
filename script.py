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

TODO:
- Forstå alle fields i ics
- Find min/max dato og slet vagter i det interval - muligt?
    - eller må man bare nøjes med at gøre det en gang og så manuelt holde styr på ændringer?
- ordne timezone i dtstart og dtend
"""
#%%
import pandas as pd
import pytz
import socket
from datetime import datetime as dt, time

usermail = 'embn@dksund.dk'
usermail = usermail.lower()
df = pd.read_excel('test\ScheduleReport.xlsx', sheet_name=0)

with open('output.ics', mode='r+', encoding='utf-8') as f:
    idcount = 0

    for i in df.iterrows():
        if i[1]['Mail (arbejde)'].lower() == usermail:
            dtstart_date = dt.strptime(i[1]['Startdato for vagten'],
                                       "%m/%d/%Y")
            dtstart_time = time.fromisoformat(
                i[1]['Starttidspunkt for vagten'])
            dtstart = dt.combine(dtstart_date,
                                 dtstart_time).strftime("%Y%m%dT%H%M%SZ")

            dtslut_date = dt.strptime(i[1]['Slutdato for vagten'], "%m/%d/%Y")
            dtslut_time = time.fromisoformat(i[1]['Sluttidspunkt for vagten'])
            dtslut = dt.combine(dtslut_date,
                                dtslut_time).strftime("%Y%m%dT%H%M%SZ")

            summary = i[1]['Brugerdefineret mærkat']

            uid = dt.strftime(
                dt.now(),
                "%Y%m%dT%H%M%S") + str(idcount) + '@' + socket.gethostname()
            idcount += 1

            f.write(f"""BEGIN:VEVENT
DTSTART:{dtstart}
DTEND:{dtslut}
DTSTAMP:20220524T215156Z
UID:{uid}
CREATED:20210524T091707Z
DESCRIPTION:
LAST-MODIFIED:20210524T091707Z
LOCATION:
SEQUENCE:0
STATUS:CONFIRMED
SUMMARY:{summary}
TRANSP:OPAQUE
END:VEVENT
""")

# Inspiration til hvordan man kan loope over rækkerne
# Det er ikke særligt hurtigt, men performance burde ikke være et problem(?)
# for i in df[:10].iterrows():
#   print(i[1]['Gruppe'])
