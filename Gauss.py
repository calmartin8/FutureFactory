# -*- coding: utf-8 -*-
"""
Created on Wed Dec 28 15:10:01 2022

@author: clawi
"""

import serial.tools.list_ports

import os
    
import ctypes

from tkinter import *

path = os.path.join("C:\\", "Program Files (x86)\Hirst Magnetic Instruments Ltd\gm0\windows64","gm0.dll")

from enum import Enum

gm0 = ctypes.WinDLL(path)

class gm_time(ctypes.Structure):
    _fields_ = [("sec", ctypes.c_byte), ("min", ctypes.c_byte), ("hour", ctypes.c_byte), ("day", ctypes.c_byte), ("month", ctypes.c_byte), ("year", ctypes.c_byte)]

class gm_store(ctypes.Structure):
    _fields_ = [("time", gm_time), ("range", ctypes.c_byte), ("mode", ctypes.c_byte), ("units", ctypes.c_byte), ("value", ctypes.c_float)]

handle = ctypes.c_long()
data = gm_store()
sampleindex = ctypes.c_int()
samplerate = ctypes.c_double()
initstatus = ctypes.c_bool()

Comport = ctypes.c_int()

unitsrange = [[ctypes.c_float(1.0), ctypes.c_float(1.0), "nT"], [ctypes.c_float(10.0), ctypes.c_float(100.0), "uT"], [ctypes.c_float(100.0), ctypes.c_float(1000.0), "mT"], [ctypes.c_float(1000.0), ctypes.c_float(10000.0), "T"], [ctypes.c_float(0.001), ctypes.c_float(0.01), "G"]]
unitsrangefmt = ["{:.1f}", "{:.1f}", "{:.1f}", "{:.2f}", "{:.2f}"]
unitsrangesclar = ["{:.1f}", "{:.1f}", "{:.2f}", "{:.3f}", "{:.3f}"]
baseunits = ["nT", "uT", "mT", "T", "G"]
modestr = ["NULL", "NORMAL", "MIN", "MAX", "AVER"]

mode = int()

dont_schedule = ctypes.c_bool()

extratimerproc = ctypes.c_bool()

class ConnectState(Enum):
    Idle = 1
    Connecting = 2
    Connected = 3
    Disconnecting = 4
    Failed = 5

State = ConnectState

# start stop functions
gm0_newgm = ctypes.windll.gm0.gm0_newgm
gm0_newgm.argtypes = [ctypes.c_long, ctypes.c_long]
gm0_newgm.restype = ctypes.c_long

gm0_startconnect = ctypes.windll.gm0.gm0_startconnect
gm0_startconnect.argtypes = [ctypes.c_long]
gm0_startconnect.restype = ctypes.c_long

gm0_killgm = ctypes.windll.gm0.gm0_killgm
gm0_killgm.argtypes = [ctypes.c_long]
gm0_killgm.restype = ctypes.c_long

gm0_getconnect = ctypes.windll.gm0.gm0_getconnect
gm0_getconnect.argtypes = [ctypes.c_long]
gm0_getconnect.restype = ctypes.c_long

# set functions
gm0_setrange = ctypes.windll.gm0.gm0_setrange
gm0_setrange.argtypes = [ctypes.c_long, ctypes.c_byte]
gm0_setrange.restype = ctypes.c_long

gm0_setunits = ctypes.windll.gm0.gm0_setunits
gm0_setunits.argtypes = [ctypes.c_long, ctypes.c_byte]
gm0_setunits.restype = ctypes.c_long

gm0_setmode = ctypes.windll.gm0.gm0_setmode
gm0_setmode.argtypes = [ctypes.c_long, ctypes.c_byte]
gm0_setmode.restype = ctypes.c_long

gm0_isnewdata = ctypes.windll.gm0.gm0_isnewdata
gm0_isnewdata.argtypes = [ctypes.c_long]
gm0_isnewdata.restype = ctypes.c_long

# get functions
gm0_getrange = ctypes.windll.gm0.gm0_getrange
gm0_getrange.argtypes = [ctypes.c_long]
gm0_getrange.restype = ctypes.c_long

gm0_getunits = ctypes.windll.gm0.gm0_getunits
gm0_getunits.argtypes = [ctypes.c_long, ctypes.c_byte]
gm0_getunits.restype = ctypes.c_longlong

gm0_getmode = ctypes.windll.gm0.gm0_getmode
gm0_getmode.argtypes = [ctypes.c_long]
gm0_getmode.restype = ctypes.c_long

gm0_getvalue = ctypes.windll.gm0.gm0_getvalue
gm0_getvalue.argtypes = [ctypes.c_long]
gm0_getvalue.restype = ctypes.c_double

# null and peak detect
gm0_donull = ctypes.windll.gm0.gm0_donull
gm0_donull.argtypes = [ctypes.c_long]
gm0_donull.restype = ctypes.c_long

gm0_doaz = ctypes.windll.gm0.gm0_doaz
gm0_doaz.argtypes = [ctypes.c_long]
gm0_doaz.restype = ctypes.c_long

gm0_resetnull = ctypes.windll.gm0.gm0_resetnull
gm0_resetnull.argtypes = [ctypes.c_long]
gm0_resetnull.restype = ctypes.c_long

gm0_resetpeak = ctypes.windll.gm0.gm0_resetpeak
gm0_resetpeak.argtypes = [ctypes.c_long]
gm0_resetpeak.restype = ctypes.c_long

# time functions
gm0_sendtime = ctypes.windll.gm0.gm0_sendtime
gm0_sendtime.argtypes = [ctypes.c_long, ctypes.c_bool]
gm0_sendtime.restype = ctypes.c_long

gm0_settime2 = ctypes.windll.gm0.gm0_settime2
gm0_settime2.argtypes = [ctypes.c_long, gm_time]
gm0_settime2.restype = ctypes.c_long

gm0_gettime = ctypes.windll.gm0.gm0_gettime
gm0_gettime.argtypes = [ctypes.c_long]
gm0_gettime.restype = gm_time

gm0_getstore = ctypes.windll.gm0.gm0_getstore
gm0_getstore.argtypes = [ctypes.c_long, ctypes.c_long]
gm0_getstore.restype = gm_store

# call back functions

gm0_setcallback2 = ctypes.windll.gm0.gm0_setcallback2
gm0_setcallback2.argtypes = [ctypes.c_long, ctypes.c_long]
gm0_setcallback2.restype = ctypes.c_long

gm0_setconnectcallback = ctypes.windll.gm0.gm0_setconnectcallback
gm0_setconnectcallback.argtypes = [ctypes.c_long, ctypes.c_long]
gm0_setconnectcallback.restype = ctypes.c_long

# disable datamode functions

gm0_startcmd = ctypes.windll.gm0.gm0_startcmd
gm0_startcmd.argtypes = [ctypes.c_long]
gm0_startcmd.restypes = ctypes.c_long

gm0_endcmd = ctypes.windll.gm0.gm0_endcmd
gm0_endcmd.argtypes = [ctypes.c_long]
gm0_endcmd.restypes = ctypes.c_long


gm0_enabledebug = ctypes.windll.gm0.gm0_enabledebug
gm0_enabledebug.restypes = ctypes.c_long


KillTimer = ctypes.windll.user32.KillTimer
KillTimer.argtypes = [ctypes.c_long, ctypes.c_long]
KillTimer.restypes = ctypes.c_long

SetTimer = ctypes.windll.user32.SetTimer
SetTimer.argtypes = [ctypes.c_long, ctypes.c_long]
SetTimer.restypes = ctypes.c_long

lngTimer = ctypes.c_long()

## Gaussmeter Command Wrappers

def setrange(handle, range):
    if handle < 0:
        return
    ctypes.windll.gm0.gm0_setrange(handle,range)
    
def setmode(handle, mode):
    if handle < 0:
        return
    ctypes.windll.gm0.gm0_setmode(handle, mode)
    
def setunits(handle,units):
    if handle < 0:
        return
    ctypes.windll.gm0.gm0_setunits(handle, units)
    
def donull(handle):
    if handle < 0:
        return
    input("Shield probe and press enter")
    ctypes.windll.gm0.gm0_donull(handle)
    print("Null finished")

def doaz(handle):
    if handle < 0:
        return
    input("Shield probe and press enter")
    ctypes.windll.gm0.gm0_doaz(handle)
    print("Auto zero finished")

def startcmdseq(handle):
    if handle < 0:
        return
    ctypes.windll.gm0.gm0_startcmd(handle)

def endcmdseq(handle):
    if handle < 0:
        return
    ctypes.windll.gm0.gm0_endcmd(handle)

def getregister(handle, reg):
    if handle < 0:
        return
    if reg >= 0 and reg < 100:
        return ctypes.windll.gm0.gm0_.getstore(handle, reg)
    else:
        print("Invalid register")

def setsystime(handle):
    import datetime
    if handle < 0:
        return
    now = datetime.datetime.now()
    thetime = {'year': now.year - 2000, 'month': now.month, 'day': now.day,
               'hour': now.hour, 'min': now.minute, 'sec': now.second}
    ctypes.windll.gm0.gm0_settime2(handle, thetime)
    print("GM08 time set to system time")

def getgmtime(handle):
    if handle < 0:
        return
    return ctypes.windll.gm0.gm_0.gettime(handle)


def connectcallback():
    UserForm1.config(text="Connected")

def datacallback(gm_handle):
    temprange = gm0_getrange(gm_handle)
    data.value = gm0_getvalue(gm_handle)
    data.mode = gm0_getmode(gm_handle)
    data.units = gm0_getunits(gm_handle)
    if temprange > 3:
        temprange -= 4
    data.range = temprange
    UserForm1.config(text="New data received")

root = Tk()
UserForm1 = Label(root, text="Not connected")
UserForm1.pack()
root.mainloop()

## Timer Functions

import time

def sampletosheetproc():
    while True:
        UserForm1.writedatatosheet(data, False)
        if dont_schedule:
            break
        time.sleep(gm0.samplerate)


## Init and kill    
        
def startconnect(mode):
    gm0.handle = gm0_newgm(Comport, mode)
    if gm0.handle < 0:
        print("An error has occured opening the requested port")
        return
    State = Connecting
    Start_Timer()
    gm0.gm0_startconnect(handle)

def switchcomspeed():
    mode = int()
    if State == Failed:
        gm0_killgm(gm0.handle)
        mode += 1
        if mode == 2:
            mode = 0
        startconnect(mode)

def init():
    cleanup()
    
    gm0_enabledebug()
    
    unitsrange[0][0] = "T"
    unitsrange[0][1] = "mT"
    unitsrange[0][2] = "mT"
    unitsrange[0][3] = "mT"
    
    unitsrangesclar[0][0] = 1
    unitsrangesclar[0][1] = 1000
    unitsrangesclar[0][2] = 1000
    unitsrangesclar[0][3] = 1000
    
    unitsrangefmt[0][0] = " 0.000;-0.000"
    unitsrangefmt[0][1] = " 000.0;-000.0"
    unitsrangefmt[0][2] = " 00.00;-00.00"
    unitsrangefmt[0][3] = " 0.000;-0.000"
    
    unitsrange[1][0] = "kG"
    unitsrange[1][1] = "kG"
    unitsrange[1][2] = "G"
    unitsrange[1][3] = "G"
    
    unitsrangefmt[1][0] = " 00.00;-00.00"
    unitsrangefmt[1][1] = " 0.000;-0.000"
    unitsrangefmt[1][2] = " 000.0;-000.0"
    unitsrangefmt[1][3] = " 00.00;-00.00"
    
    unitsrangesclar[1][0] = 0.001
    unitsrangesclar[1][1] = 0.001
    unitsrangesclar[1][2] = 1
    unitsrangesclar[1][3] = 1
    
    unitsrange[2][0] = "kA/m"
    unitsrange[2][1] = "kA/m"
    unitsrange[2][2] = "kA/m"
    unitsrange[2][3] = "kA/m"
    
    unitsrangefmt[2][0] = " 0000;-0000"
    unitsrangefmt[2][1] = " 000.0;-000.0"
    unitsrangefmt[2][2] = " 00.00;-00.00"
    unitsrangefmt[2][3] = " 0.000;-0.000"
    
    unitsrangesclar[2][0] = 0.001
    unitsrangesclar[2][1] = 0.001
    unitsrangesclar[2][2] = 0.001
    unitsrangesclar[2][3] = 0.001
    
    unitsrange[3][0] = "kOe"
    unitsrange[3][1] = "kOe"
    unitsrange[3][2] = "Oe"
    unitsrange[3][3] = "Oe"
    
    unitsrangefmt[3][0] = " 00.00;-00.00"
    unitsrangefmt[3][1] = " 000.0;-000.0"
    unitsrangefmt[3][2] = " 0.000;-0.000"
    unitsrangefmt[3][3] = " 00.00;-00.00"
  
    unitsrangesclar[3][0] = 0.001
    unitsrangesclar[3][1] = 0.001
    unitsrangesclar[3][2] = 1
    unitsrangesclar[3][3] = 1
  
    baseunits[0] = "T"
    baseunits[1] = "G"
    baseunits[2] = "A/m"
    baseunits[3] = "Oe"
            
    modestr[0] = "DC"
    modestr[1] = "DC Pk"
    modestr[2] = "AC"
    modestr[3] = "AC Mx"
    modestr[4] = "AX Pk"
    
def cleanup():
    handle = ctypes.c_long()
    State = Disconnecting
    
    Kill_Timer()
    
    if handle >= 0:
        ctypes.windll.gm0.gm0_killgm(handle)
        return
    State = Idle
    handle -= 1
    
def run_gaussmeter():
    UserForm1.Show(False)
    
## Helper functions

    
    
def makeactualvalue(ldata):
    tempvalue = ldata.value * unitsrangesclar[ldata.units][ldata.range]
    return tempvalue


#Win32 API Callbacks

def Start_Timer():
    lngTimer = SetTimer(0, 0, 100, V_B_E)

def Kill_Timer():
    KillTimer(0, lngTimer)

def V_B_E(hWnd, uMsg, wParam, lParam):
    try:
        State = ConnectState
        if State == Disconnecting:
            Kill_Timer()
        if State == ConnectState.Idle:
            return
        if State == ConnectState.connected:
            if gm0_isnewdata(handle):
                datacallback(gm0.handle)
        if State == ConnectState.Connecting:
            x = gm0_getconnect(handle)
            if x == 1:
                State = ConnectState.connected
                connectcallback()
    except:
        Kill_Timer()

    
    
    
        
        


