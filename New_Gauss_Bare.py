
import time
from new_gm0 import *
from py import *
 
# Connect to the first USB GM08
#  -1 is the first USB device, -2 is the 2nd

dataset = []

handle = gm0_newgm(-1,0)
gm0_startconnect(handle)

while gm0_getconnect(handle) == 0:
    print("Connecting ....")
    time.sleep(0.5)

if gm0_getconnect(handle) == 1:
    print("Connected")

if gm0_getconnect(handle) <0:
    print("Error connecting")
    time.sleep(5)
    exit()


print("Setting Mode DC")
gm0_setmode(handle,0)
time.sleep(0.5)

print("Setting Units Tesla")
gm0_setunits(handle,0)
time.sleep(0.5)

for x in range(4):
    print("Nulling range :",x)
    gm0_setrange(handle,x) 
    time.sleep(0.5)
    gm0_donull(handle)
    time.sleep(0.5)
    gm0_resetnull(handle)
    time.sleep(0.5)

print("Setting Range Auto")
gm0_setrange(handle,4) #4 is auto range
time.sleep(0.5)

for x in range(100):

    val = gm0_getvalue(handle)
    range = gm0_getrange(handle)
    autorange=0;
    if range>0x0:
        autorange=1
        
    range &= 0x03
    
    mode = gm0_getmode(handle)
    units = gm0_getunits(handle)

    
    print("{:.6f}".format(val),gm0_Units[units],gm0_Mode[mode],"Range :",range," Auto :",autorange)
    dataset.append(float("{:.6f}".format(val)))
    time.sleep(0.1)

def find(dataset):
    print(str(max(dataset)) + " is the maximum value")
    print(str(min(dataset)) + " is the minimum value")
    if (abs(min(dataset))) > max(dataset):
        return("This magnet is a south pole")
    elif (abs(min(dataset))) < max(dataset):
        return("This magnet is a north pole")
    else:
        return("Error during magnetic characterisation, please retry")


    
