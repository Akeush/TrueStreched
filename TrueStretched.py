import pywintypes
import win32api
import win32con
import win32com.client
import time
import ctypes
import os
import psutil
import subprocess
import platform
import stat


# Define the ASCII art banner
ASCII_ART = """
  _______                 _____ _                 _              _ 
 |__   __|               / ____| |               | |            | |
    | |_ __ _   _  ___  | (___ | |_ _ __ ___  ___| |__   ___  __| |
    | | '__| | | |/ _ \  \___ \| __| '__/ _ \/ __| '_ \ / _ \/ _` |
    | | |  | |_| |  __/  ____) | |_| | |  __| (__| | | |  __| (_| |
    |_|_|   \__,_|\___| |_____/ \__|_|  \___|\___|_| |_|\___|\__,_|
                                                                   
                                                                   
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
    pc_name = os.getlogin()
    title = f"Welcome ({pc_name}) - True Streched by Kaliert / Sayk"
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


def find_shortcut(base_dir, keyword="Riot"):
    """Search for a shortcut file in the base directory containing the keyword."""
    for root, _, files in os.walk(base_dir):
        for file in files:
            if keyword.lower() in file.lower() and file.endswith(".lnk"):
                return os.path.join(root, file)
    return None

def launch_valorant():
    """Launches Valorant using subprocess."""
    valorant_launcher_path = r"C:\\Riot Games\\Riot Client\\RiotClientServices.exe"

    # Check if the direct path exists
    if os.path.exists(valorant_launcher_path):
        launch_arguments = ["--launch-product=valorant", "--launch-patchline=live"]
        try:
            subprocess.Popen([valorant_launcher_path] + launch_arguments, shell=True)
            print("Launching Valorant using the direct path...")
            return True
        except Exception as e:
            print(f"Failed to launch Valorant using the direct path: {e}")
    
    # If the direct path doesn't exist, try finding the shortcut dynamically
    base_dir = r"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Riot Games"
    shortcut_path = find_shortcut(base_dir)

    if shortcut_path:
        try:
            valorant_launcher_path = get_shortcut_target(shortcut_path)
            launch_arguments = ["--launch-product=valorant", "--launch-patchline=live"]
            subprocess.Popen([valorant_launcher_path] + launch_arguments, shell=True)
            print("Launching Valorant using the shortcut path...")
            return True
        except Exception as e:
            print(f"Error launching Valorant from shortcut: {e}")
            return False
    else:
        print("Shortcut not found. Valorant could not be launched.")
        return False

def close_valorant(process_name):
    """Closes the Valorant process if it is running."""
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == process_name:
            try:
                proc.terminate()
                print(f"Valorant process (PID: {proc.info['pid']}) terminated successfully.")
                return True
            except psutil.AccessDenied:
                print(f"Access denied to terminate Valorant process (PID: {proc.info['pid']}).")
                return False
            except psutil.NoSuchProcess:
                print("Valorant process no longer exists.")
                return False
            except Exception as e:
                print(f"An error occurred while terminating Valorant: {e}")
                return False
    print("Valorant process not found.")
    return False


def get_shortcut_target(shortcut_path):
    """Get the target path of a shortcut (.lnk)."""
    shell = win32com.client.Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    return shortcut.TargetPath

def monitor_valorant_exit(window_title, native_width, native_height, game_user_settings_path):
    """Monitors the Valorant process and restores native resolution upon exit."""
    while True:
        time.sleep(5)  # Check every 5 seconds
        if not any(proc.name() == "VALORANT-Win64-Shipping.exe" for proc in psutil.process_iter()):
            print("Valorant process exited. Restoring native resolution...")
            change_resolution(native_width, native_height)
            modify_game_user_settings(game_user_settings_path, native_width, native_height, True)
            break

def remove_read_only(file_path):
    """Removes the read-only attribute from the file."""
    try:
        os.chmod(file_path, stat.S_IWRITE)  # Grant write permissions
        print(f"Read-only attribute removed from file: {file_path}")
    except Exception as e:
        print(f"Error removing read-only attribute: {e}")

def set_read_only(file_path):
    """Sets the read-only attribute on the file."""
    try:
        os.chmod(file_path, stat.S_IREAD)  # Grant only read permissions
        print(f"File {file_path} is now read-only.")
    except Exception as e:
        print(f"Error changing file permissions: {e}")

def modify_game_user_settings(file_path, width, height, reset=False):
    try:
        # Remove the read-only attribute before modifying the file
        remove_read_only(file_path)
        print(f"Modifying file: {file_path}\n")
        print(f"New resolution: {width}x{height}\n")
        with open(file_path, 'r') as file:
            lines = file.readlines()

        fullscreen_mode_found = False
        for i, line in enumerate(lines):
            # Check if the line contains exactly "FullscreenMode="
            if line.strip().startswith("FullscreenMode="):
                if not line.strip() == "FullscreenMode=2":
                    lines[i] = "FullscreenMode=2\n"
                elif line.strip() == "FullscreenMode=2" and reset:
                    lines[i] = ""
                fullscreen_mode_found = True
                break

        # If the line "FullscreenMode=" was not found, add it to the first empty line
        if not fullscreen_mode_found:
            for i, line in enumerate(lines):
                if line.strip() == "":
                    lines.insert(i, "FullscreenMode=2\n")
                    break

        # Modify all resolution lines
        for i, line in enumerate(lines):
            if line.strip().startswith("ResolutionSizeX="):
                lines[i] = f"ResolutionSizeX={width}\n"
            elif line.strip().startswith("ResolutionSizeY="):
                lines[i] = f"ResolutionSizeY={height}\n"
            elif line.strip().startswith("LastUserConfirmedResolutionSizeX="):
                lines[i] = f"LastUserConfirmedResolutionSizeX={width}\n"
            elif line.strip().startswith("LastUserConfirmedResolutionSizeY="):
                lines[i] = f"LastUserConfirmedResolutionSizeY={height}\n"
            elif line.strip().startswith("LastUserConfirmedDesiredScreenWidth="):
                lines[i] = f"LastUserConfirmedDesiredScreenWidth={width}\n"
            elif line.strip().startswith("LastUserConfirmedDesiredScreenHeight="):
                lines[i] = f"LastUserConfirmedDesiredScreenHeight={height}\n"

        # Check and modify the bShouldLetterbox line
        for i, line in enumerate(lines):
            if line.strip().startswith("bShouldLetterbox="):
                    lines[i] = "bShouldLetterbox=False\n"

        # Write the modifications to the file
        with open(file_path, 'w') as file:
            file.writelines(lines)
        print("File modification completed.")

        # Set the file to read-only after modification
        set_read_only(file_path)

    except FileNotFoundError:
        print(f"Error: The file {file_path} was not found.")
    except Exception as e:
        print(f"Error: {str(e)}")



                    
def get_current_resolution():
    """Gets the current resolution applied to the main screen."""
    user32 = ctypes.windll.user32
    user32.SetProcessDPIAware()  # To get the actual resolution if DPI scaling is enabled
    screen_width = user32.GetSystemMetrics(0)  # Screen width
    screen_height = user32.GetSystemMetrics(1)  # Screen height
    return screen_width, screen_height

def display_folder_contents(width, height):
    # Get the AppData\Local\VALORANT\Saved\Config path
    appdata = os.getenv('LOCALAPPDATA')  # Get the path to AppData\Local
    if not appdata:
        print("Unable to find the LOCALAPPDATA folder.")
        return None

    game_path = os.path.join(appdata, 'VALORANT', 'Saved', 'Config')

    # Check if the folder exists
    if not os.path.exists(game_path):
        print(f"The folder does not exist: {game_path}")
        return None

    # Access the Windows folder and look for the configuration file
    windows_path = os.path.join(game_path, 'Windows')
    ini_file = os.path.join(windows_path, 'RiotLocalMachine.ini')

    # Check if the configuration file exists and display the last opened account
    if os.path.exists(ini_file):
        try:
            with open(ini_file, 'r', encoding='utf-8') as f:
                content = f.read()

                # Search for the last opened account
                for line in content.splitlines():
                    if "LastKnownUser=" in line:
                        last_account = line.split("=", 1)[1].strip()
                        print(f"\nID of the last use account: {last_account}\n")
                        break
        except Exception as e:
            print(f"Unable to read the configuration file: {e}")
    else:
        print(f"The configuration file does not exist.\n")

    print(f"Account List: {game_path}\n")

    # List the folders and assign a value to each folder
    valid_folders = []
    try:
        folders = [d for d in os.listdir(game_path) if os.path.isdir(os.path.join(game_path, d))]
        for folder in folders:
            folder_path = os.path.join(game_path, folder)

            # Check if the folder contains a 'Windows' subfolder
            windows_path = os.path.join(folder_path, 'Windows')
            if os.path.isdir(windows_path):
                # Check if the file 'GameUserSettings.ini' exists in this Windows folder
                game_ini_file = os.path.join(windows_path, 'GameUserSettings.ini')
                if os.path.exists(game_ini_file):
                    valid_folders.append(folder)

        # Display valid folders
        if valid_folders:
            for index, folder in enumerate(valid_folders, start=1):
                print(f"[{index}] {folder}")
        else:
            print("No folder contains a 'Windows' subfolder with 'GameUserSettings.ini'.")
    except Exception as e:
        print(f"Error reading folders: {e}")
        return None

    # Ask the user to choose a folder
    try:
        choice = int(input("\nChoose the account you want to play with: "))
        if 1 <= choice <= len(valid_folders):
            chosen_folder = valid_folders[choice - 1]
            chosen_folder_path = os.path.join(game_path, chosen_folder)
            game_user_settings_path = os.path.join(chosen_folder_path, 'Windows', 'GameUserSettings.ini')

            clear_console()

            print(f"\nYou have chosen : {chosen_folder}\n")

            # Return the path of the chosen folder
            return game_user_settings_path
        else:
            print("Invalid value. No folder matches.")
    except ValueError:
        print("Invalid entry. Please enter a number.")
        return None


def main():
    change_console_title()  # Change console title
    set_console_color(FOREGROUND_RED)  # Set console color to red
    clear_console()  # Clear console screen

    window_title = "VALORANT  "
    process_name = "VALORANT-Win64-Shipping.exe"

    # Save native resolution
    native_width, native_height = get_current_resolution()
    print(f"Native Resolution: {native_width}x{native_height}")

    try:
        width = int(input("Enter the width of the resolution (e.g., 1920): "))
        height = int(input("Enter the height of the resolution (e.g., 1080): "))
    except ValueError:
        print("Invalid resolution input. Please enter numbers only.")
        input("Press Enter to exit...")
        return
    clear_console()
    new_priority = choose_priority_class()
    clear_console()
    # Save the selected user folder for later use
    game_user_settings_path = display_folder_contents(width, height)
    if game_user_settings_path:
        # Use the game_user_settings_path variable to modify the settings
        modify_game_user_settings(game_user_settings_path, width, height)

    # Continue with other operations like resolution change, launching Valorant, etc.
    change_resolution(width, height)
    time.sleep(5) 
    window_handle = ctypes.windll.user32.FindWindowW(None, window_title)
    if window_handle == 0:
        clear_console()

        if not launch_valorant():
            input("Press Enter to exit...")
            return
        
        print("Waiting for Valorant to start...")
        time.sleep(30)  # Wait for Valorant to launch

        # Retry finding the Valorant window
        window_handle = ctypes.windll.user32.FindWindowW(None, window_title)

        if window_handle == 0:
            print("Valorant not found.. Exiting...")
            input("Press Enter to exit...")
            return


    # Set process priority
    set_process_priority(process_name, new_priority)
    clear_console()

    print("\nTrue stretched applied")
    print("\nDon't close the program; it will revert to native when you close Valorant")

    # Monitor Valorant process exit
    monitor_valorant_exit(window_title, native_width, native_height, game_user_settings_path)
    remove_read_only(game_user_settings_path)
    clear_console()

    print("True stretched reverted to native resolution.")
    print ("Settings file is now back to normal")
    print("Thank you for using True Stretched by Kaliert / Sayk")
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()

