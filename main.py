from re import T
from win32clipboard import OpenClipboard, GetClipboardData, CloseClipboard
from PIL import ImageGrab
from PIL.PngImagePlugin import PngImageFile
from time import sleep
from keyboard import is_pressed, 
from win32process import GetProcessWindowStation, GetWindowThreadProcessId
from win32gui import GetForegroundWindow, GetWindowText
from psutil import Process
import string
from ctypes import windll
import os
def getClipboardData():
    try:
        OpenClipboard()
        data = GetClipboardData()
        CloseClipboard()
    except TypeError:
        data = ImageGrab.grabclipboard()
    return data


def getDrives():
    drives = []
    bitmask = windll.kernel32.GetLogicalDrives()
    for letter in string.ascii_uppercase:
        if bitmask & 1:
            drives.append(letter)
        bitmask >>= 1

    return drives




def getActiveWindowName():
    pid = GetWindowThreadProcessId(GetForegroundWindow())
    return Process(pid[-1]).name()

def getActiveWindowTitle():
    return GetWindowText(GetForegroundWindow())



        




def pressCtrlV():
    press_and_release('ctrl+v')
    print("press ctrl v ")


if __name__ == "__main__":
    while True:
        try:
            if is_pressed('shift+v'):

                returnValue = getClipboardData()
                activeWindow = getActiveWindowName()
                openWindowIsFolder = False
                if activeWindow == "explorer.exe":
                    openWindowIsFolder = True
                    openWindowsFolderName = getActiveWindowTitle() 

                if  type(returnValue) == str:
                    if openWindowIsFolder:
                        # if open windows name == exploler.exe
                        # String to text file.

                        # ????????\{openWindowsFolderName}\saveHere.txt
                        #                    ^
                        #                   ||
                        #how can i get previous from this location
                        print('str to file')
                    else:
                        pressCtrlV()


                elif  type(returnValue) == PngImageFile:
                    if openWindowIsFolder:
                        # if open windows name == exploler.exe
                        # Image to img file
                        print('img to file')

                    else:
                        pressCtrlV()
                elif type(returnValue) == list:
                    print("file to folder")
                    pressCtrlV()
                else:
                    print(type(returnValue))
        except Exception as e:
            print(e)
        sleep(0.3)

