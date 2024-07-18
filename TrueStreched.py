import pywintypes
import win32api
import win32con
import time
import ctypes
import os
import psutil
import subprocess
import platform

# Define the ASCII art banner
ASCII_ART = """
  _______                 _____ _                 _              _ 
 |__   __|               / ____| |               | |            | |
    | |_ __ _   _  ___  | (___ | |_ _ __ ___  ___| |__   ___  __| |
    | | '__| | | |/ _ \  \___ \| __| '__/ _ \/ __| '_ \ / _ \/ _` |
    | | |  | |_| |  __/  ____) | |_| | |  __| (__| | | |  __| (_| |
    |_|_|   \__,_|\___| |_____/ \__|_|  \___|\___|_| |_|\___|\__,_|
                                                                   
                                                                   
"""

# Windows API constants for console color
STD_OUTPUT_HANDLE = -11
FOREGROUND_RED = 0x0004

def clear_console():
    """Clears the console screen based on the platform."""
    if platform.system() == 'Windows':
        os.system('cls')  # Clear screen for Windows
        print(ASCII_ART)
    else:
        os.system('clear')  # Clear screen for Unix/Linux/MacOS
        print(ASCII_ART)

def change_console_title():
    """Changes the console title to include PC name and additional information."""
    pc_name = platform.node()
    title = f"Welcome ({pc_name}) - True Streched by Akeush / Early1488"
    ctypes.windll.kernel32.SetConsoleTitleW(title)

def set_console_color(color):
    """Sets the console text color."""
    handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
    ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)

def change_resolution(width, height):
    """Changes the display resolution."""
    devmode = pywintypes.DEVMODEType()
    devmode.PelsWidth = width
    devmode.PelsHeight = height
    devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT
    win32api.ChangeDisplaySettings(devmode, 0)

def modify_window_properties(window_handle):
    """Modifies window properties to remove dialog frame and border."""
    style = ctypes.windll.user32.GetWindowLongW(window_handle, ctypes.c_int(-16))
    style &= ~0x00800000  # WS_DLGFRAME removal
    style &= ~0x00040000  # WS_BORDER removal
    ctypes.windll.user32.SetWindowLongW(window_handle, ctypes.c_int(-16), style)

def maximize_window(window_handle):
    """Maximizes the specified window."""
    ctypes.windll.user32.ShowWindow(window_handle, ctypes.c_int(3))  # SW_MAXIMIZE

def set_process_priority(process_name, new_priority):
    """Sets the priority of a process."""
    priority_classes = {
        1: psutil.IDLE_PRIORITY_CLASS,
        2: psutil.BELOW_NORMAL_PRIORITY_CLASS,
        3: psutil.NORMAL_PRIORITY_CLASS,
        4: psutil.ABOVE_NORMAL_PRIORITY_CLASS,
        5: psutil.HIGH_PRIORITY_CLASS,
        6: psutil.REALTIME_PRIORITY_CLASS
    }
    
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == process_name:
            try:
                p = psutil.Process(proc.info['pid'])
                p.nice(priority_classes[new_priority])
                print(f"Changed priority of {process_name} (PID: {p.pid}) to {new_priority}")
            except psutil.AccessDenied:
                print(f"Access denied to change priority of {process_name} (PID: {proc.info['pid']})")
            except psutil.NoSuchProcess:
                print(f"No such process: {process_name} (PID: {proc.info['pid']})")
            except Exception as e:
                print(f"An error occurred: {e}")

def choose_priority_class():
    """Prompts user to choose a priority class for the process."""
    priority_classes = {
        1: "Idle",
        2: "Below Normal",
        3: "Normal",
        4: "Above Normal",
        5: "High",
        6: "Realtime"
    }
    
    print("\nChoose a priority class:")
    for key, value in priority_classes.items():
        print(f"{key}. {value}")
    
    while True:
        try:
            choice = int(input("Enter the number corresponding to the desired priority class: "))
            if choice in priority_classes:
                return choice
            else:
                print("Invalid choice. Please enter a number between 1 and 6.")
        except ValueError:
            print("Invalid input. Please enter a numeric value.")

def launch_valorant():
    """Launches Valorant using subprocess."""
    valorant_launcher_path = r"C:\Riot Games\Riot Client\RiotClientServices.exe"
    launch_arguments = ["--launch-product=valorant", "--launch-patchline=live"]
    
    try:
        subprocess.Popen([valorant_launcher_path] + launch_arguments, shell=True)
        print("Launching Valorant...")
        return True
    except Exception as e:
        print(f"Failed to launch Valorant: {e}")
        return False

def monitor_valorant_exit(window_title, native_width, native_height):
    """Monitors the Valorant process and restores native resolution upon exit."""
    while True:
        time.sleep(5)  # Check every 5 seconds
        if not any(proc.name() == "VALORANT-Win64-Shipping.exe" for proc in psutil.process_iter()):
            print("Valorant process exited. Restoring native resolution...")
            change_resolution(native_width, native_height)
            break

def main():
    change_console_title()  # Change console title
    set_console_color(FOREGROUND_RED)  # Set console color to red
    clear_console()  # Clear console screen
    #print(ASCII_ART)  # Print the ASCII art banner
    
    window_title = "VALORANT  "
    process_name = "VALORANT-Win64-Shipping.exe"

    # Attempt to find the Valorant window
    window_handle = ctypes.windll.user32.FindWindowW(None, window_title)

    if window_handle == 0:
        print("Valorant not found. Attempting to launch Valorant...")

        if not launch_valorant():
            input("Press Enter to exit...")
            return
        
        print("Waiting for Valorant to start...")
        time.sleep(30)  # Wait for Valorant to launch

        # Retry finding the Valorant window
        window_handle = ctypes.windll.user32.FindWindowW(None, window_title)

        if window_handle == 0:
            print("Valorant still not found. Exiting...")
            input("Press Enter to exit...")
            return

    clear_console()  # Clear console screen after Valorant has launched
    
    # Save native resolution
    native_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    native_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    
    # Prompt user for resolution
    while True:
        try:
            width = int(input("Enter desired width (e.g., 1480): "))
            height = int(input("Enter desired height (e.g., 1080): "))
            break
        except ValueError:
            print("Invalid input. Please enter numeric values for width and height.")
    
    # Change resolution
    change_resolution(width, height)
    
    # Modify window properties
    modify_window_properties(window_handle)

    # Maximize window
    maximize_window(window_handle)

    # Choose priority class
    new_priority = choose_priority_class()

    # Set process priority
    set_process_priority(process_name, new_priority)
    clear_console()
    
    print("\nTrue stretched applied")
    print("\nDon't close the program; it will revert to native when you close Valorant")
    
    # Monitor Valorant process exit
    monitor_valorant_exit(window_title, native_width, native_height)

    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
