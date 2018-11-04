#-*- coding: utf-8 -*-
import datetime
import time

def set():
    startdate = time.mktime(datetime.datetime.now().timetuple())
    deadline = time.mktime(datetime.date(2018,11,30).timetuple())
    if startdate <= deadline:
        return True
    else:
        return False