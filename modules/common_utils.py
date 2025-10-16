from py_stealth import *

WAIT_TIME = 500
WAIT_LAG_TIME = 10000
DEBUG = True

char = Self()
CHAR = Self()

def wait_lag(wait_time=WAIT_TIME, lag_time=WAIT_LAG_TIME):
    Wait(wait_time)
    CheckLag(lag_time)
    return

def debug(message, message_color=66):
    if DEBUG:
        AddToSystemJournal(message)
        ClientPrintEx(char, message_color, 1, message)
