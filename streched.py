# true streched valorant script by Akeush // pasted from github and shit

import pywintypes
import win32api
import win32con

devmode = pywintypes.DEVMODEType()
devmode.PelsWidth = int(input("Enter desired width (e.g., 1480): "))
devmode.PelsHeight = int(input("Enter desired Height (e.g., 1080): "))

devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT

win32api.ChangeDisplaySettings(devmode, 0)

import time
import ctypes


#time.sleep(1)

window_title = "VALORANT  "

# Search window
window_handle = ctypes.windll.user32.FindWindowW(None, window_title)

if window_handle == 0:
    print("Valorant not found")
else:
    # Change window properties
    style = ctypes.windll.user32.GetWindowLongW(window_handle, ctypes.c_int(-16))
    style = style & ~0x00800000  # WS_DLGFRAME removal
    style = style & ~0x00040000  # WS_BORDER removal
    ctypes.windll.user32.SetWindowLongW(window_handle, ctypes.c_int(-16), style)

    # Maximize window
    ctypes.windll.user32.ShowWindow(window_handle, ctypes.c_int(3))  # SW_MAXIMIZE

    print("True stretched applied")
