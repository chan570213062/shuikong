#-*- coding: utf-8 -*-
import datetime
import time

def set():
    startdate = time.mktime(datetime.datetime.now().timetuple())
    deadline = time.mktime(datetime.date(2018,12,31).timetuple())
    if startdate <= deadline:
        return True
    else:
        return False