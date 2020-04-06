#!/usr/bin/env python

import sched, time
import os
import subprocess
from datetime import date, timedelta
from shutil import copyfile
from zipfile import ZipFile
import datetime
import holidays

s = sched.scheduler(time.time, time.sleep)


ONE_DAY = datetime.timedelta(days=1)
HOLIDAYS_US = holidays.US()

def next_business_day(todaydate):
    next_day = todaydate + ONE_DAY
    while next_day.weekday() in holidays.WEEKEND or next_day in HOLIDAYS_US:
        next_day += ONE_DAY
    return next_day


def import_newTG(sc, todaydate, skipProcess):

    ############################################ CreateFiles ###########################################################
    
    os.chdir(r'C:\Users\bailey.lin\Documents\Python Tools\Python-Scheduler') ## Directory where code is stored
    
    TGDir = r'C:\Users\bailey.lin\Documents\Python Tools\TestSFTP' ## Directory where File will be Created
    currfile = 'Trading_Grid_' + today.strftime("%Y%m%d") + '.xlsx' #FileName to lookup
    if os.path.exists(TGDir + "\\" + currfile):
        if not skipProcess:
            f1 = open('run.bat', 'w+')
            f1.write('cscript script.vbs "C:\\Users\\bailey.lin\\Documents\\Python Tools\\Python-Scheduler\\Scheduler.xlsm" ' + today.strftime("%m/%d/%Y")) ### Macro you want to run, plus any parameters
            f1.close()

            subprocess.call('run.bat') ### run.bat calls script.vbs, which can manipulate the workbook then call a Macro

    else:
        print("File not Found, retrying")
        s.enter(60, 1, import_newTG, (sc,todaydate,skipProcess,))
        

if __name__ == "__main__":
    today = date(2020,4,6) # start Date

    skipProcess = [False, False, False, False]

    while True:

        s.enter(0, 1, import_newTG, (s,today,skipProcess[0], ))
        s.run()

        today = next_business_day(today)

        skipProcess = [False, False, False, False]
    