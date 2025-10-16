from py_stealth import *
from modules.common_utils import wait_lag

def connect():
    if not Connected():
        while not Connected():
            print("Character is not connected. Connecting...")
            Connect()
            Wait(15000)
    if Connected():
        wait_lag(1500)

def reconnect():
    print("Reconnecting char...")
    Disconnect()
    Wait(8000)
    connect()

def check_connection():
    if not Connected():
        AddToSystemJournal("No connection.")
        while not Connected():
            AddToSystemJournal(">>> Trying to connect...")
            connect()
            Wait(5000)
        AddToSystemJournal("*** Server connection restored.")
    return

def print_uptime():
    if Connected():
        uptime = ConnectedTime()
        print("Char is connected since: %s" % (uptime))
    else:
        print("Char is not connected yet.")

def get_char_name():
    if not CharName() or (CharName() == "Unknown Name"):
        print("Waiting for stealth to detect char name")
        while not CharName() or (CharName() == "Unknown Name"):
            Wait(500)
    else:
        char_name = CharName()
        return char_name
