# -*- coding: utf-8 -*-
"""
Created on Fri Oct  6 16:55:56 2023

@author: Leonidas Zouloumis
"""

import datetime as dat

def GetMonthtxt(month):
    if month == 1:
        return 'January'
    elif month == 2:
        return 'February'
    elif month == 3:
        return 'March'
    elif month == 4:
        return 'April'
    elif month == 5:
        return 'May'
    elif month == 6:
        return 'June'
    elif month == 7:
        return 'July'
    elif month == 8:
        return 'August'
    elif month == 9:
        return 'September'
    elif month == 10:
        return 'October'
    elif month == 11:
        return 'November'
    elif month == 12:
        return 'December'

def TimeList(dt, time_cnt):
    cntList = [None] * time_cnt
    for i in range(len(cntList)):
        cntList[i] = dat.datetime(year=2021, month=1, day=6, hour=i*dt//3600, minute=i*dt%3600//60, second=i*dt%3600%60)
    return cntList

def CustomTimeList(dt, acstart, acend):
    cntList = [] #* ((acend-acstart)*dt//3600)
    for i in range(acstart, acend):
        cntList.append(dat.datetime(year=2021, month=1, day=6, hour=i*dt//3600, minute=i*dt%3600//60, second=i*dt%3600%60))
    return cntList

def DatetoNum(day_number=1):
    if day_number <= 31:
        month = 1
        day = day_number
    elif day_number <= 59:
        month = 2
        day = day_number - 31
    elif day_number <= 90:
        month = 3
        day = day_number - 59
    elif day_number <= 120:
        month = 4
        day = day_number - 90
    elif day_number <= 151:
        month = 5
        day = day_number - 120
    elif day_number <= 181:
        month = 6
        day = day_number - 151
    elif day_number <= 212:
        month = 7
        day = day_number - 181
    elif day_number <= 243:
        month = 8
        day = day_number - 212
    elif day_number <= 273:
        month = 9
        day = day_number - 243
    elif day_number <= 304:
        month = 10
        day = day_number - 273
    elif day_number <= 334:
        month = 11
        day = day_number - 304
    else:
        month = 12
        day = day_number - 334
    return month, day