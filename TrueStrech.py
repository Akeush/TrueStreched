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
import configparser
from typing import List, Tuple, Optional

FOREGROUND_RED = 0x0004

ASCII_ART = """
  _______                 _____ _                 _      
 |__   __|               / ____| |               | |     
    | |_ __ _   _  ___  | (___ | |_ _ __ ___  ___| |__ 
    | | '__| | | |/ _ \  \___ \| __| '__/ _ \/ __| '_ \ 
    | | |  | |_| |  __/  ____) | |_| | |  __| (__| | | |
    |_|_|   \__,_|\___| |_____/ \__|_|  \___|\___|_| |_|
                                                        
                                                        
"""
STD_OUTPUT_HANDLE = -11
VALORANT_PROCESS_NAME = "VALORANT-Win64-Shipping.exe"


def get_riot_client_path() -> Optional[str]:
    """
    Attempts to locate the Riot Client Services executable using registry keys.
    Returns the full path if found, or None otherwise.
    """
    import winreg
    possible_keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Riot Games\Riot Client"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Riot Games\Riot Client")
    ]
    for hive, key in possible_keys:
        try:
            with winreg.OpenKey(hive, key) as reg_key:
                install_folder, _ = winreg.QueryValueEx(reg_key, "InstallFolder")
            candidate = os.path.join(install_folder, "RiotClientServices.exe")
            if os.path.exists(candidate):
                return candidate
        except Exception:
            continue
    return None


class ConsoleManager:
    @staticmethod
    def clear() -> None:
        """Clears the console and displays the ASCII banner."""
        os.system('cls' if platform.system() == 'Windows' else 'clear')
        print(ASCII_ART)

    @staticmethod
    def set_title(pc_name: str) -> None:
        """Sets the console window title."""
        ctypes.windll.kernel32.SetConsoleTitleW(f"True Stretch - {pc_name}")

    @staticmethod
    def set_color(color: int) -> None:
        """Sets the console text color."""
        handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
        ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)


class ResolutionManager:
    @staticmethod
    def get_current() -> Tuple[int, int]:
        """Returns the current screen resolution."""
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

    @staticmethod
    def list_available() -> List[Tuple[int, int]]:
        """Lists all available display resolutions."""
        resolutions = []
        i = 0
        while True:
            try:
                mode = win32api.EnumDisplaySettings(None, i)
                res = (mode.PelsWidth, mode.PelsHeight)
                if res not in resolutions:
                    resolutions.append(res)
                    print(f"[{len(resolutions)}]  {res[0]}x{res[1]}")
                i += 1
            except pywintypes.error:
                break
        return resolutions

    @staticmethod
    def set_resolution(width: int, height: int) -> None:
        """Changes the display resolution."""
        devmode = pywintypes.DEVMODEType()
        devmode.PelsWidth = width
        devmode.PelsHeight = height
        devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT
        win32api.ChangeDisplaySettings(devmode, 0)


class ValorantConfigManager:
    def __init__(self):
        self.appdata_path = os.getenv('LOCALAPPDATA')
        self.game_path = os.path.join(self.appdata_path, 'VALORANT', 'Saved', 'Config')

    def find_config_files(self) -> List[Tuple[str, str]]:
        """Searches for all GameUserSettings.ini files for different accounts."""
        config_files = []
        if not os.path.exists(self.game_path):
            return config_files

        for folder in os.listdir(self.game_path):
            folder_path = os.path.join(self.game_path, folder)
            ini_path = os.path.join(folder_path, 'Windows', 'GameUserSettings.ini')
            if os.path.exists(ini_path):
                config_files.append((folder, ini_path))
        return config_files

    def modify_config(self, file_path: str, width: int, height: int, reset: bool = False) -> None:
        """Modifies (or resets) the configuration file."""
        try:
            self._set_file_permissions(file_path, read_only=False)

            with open(file_path, 'r') as f:
                lines = f.readlines()

            lines = self._update_config_lines(lines, width, height, reset)

            with open(file_path, 'w') as f:
                f.writelines(lines)

            if not reset:
                self._set_file_permissions(file_path, read_only=True)

        except Exception as e:
            print(f"Error modifying config: {str(e)}")

    def _update_config_lines(self, lines: List[str], width: int, height: int, reset: bool) -> List[str]:
        """Updates the configuration lines with the new settings."""
        new_lines = []
        fullscreen_added = False
        first_empty_line_found = False

        for line in lines:
            line_stripped = line.strip()

            if line_stripped.startswith("FullscreenMode="):
                continue

            if not reset and not fullscreen_added and line_stripped == "" and not first_empty_line_found:
                new_lines.append("FullscreenMode=2\n\n")
                fullscreen_added = True
                first_empty_line_found = True
                continue

            updated = False
            for prefix, value in [
                ("ResolutionSizeX=", width),
                ("ResolutionSizeY=", height),
                ("LastUserConfirmedResolutionSizeX=", width),
                ("LastUserConfirmedResolutionSizeY=", height),
                ("LastUserConfirmedDesiredScreenWidth=", width),
                ("LastUserConfirmedDesiredScreenHeight=", height)
            ]:
                if line_stripped.startswith(prefix):
                    new_lines.append(f"{prefix}{value}\n")
                    updated = True
                    break

            if not updated:
                if line_stripped.startswith("bShouldLetterbox=") and not reset:
                    new_lines.append("bShouldLetterbox=False\n")
                else:
                    new_lines.append(line)

        if not reset and not fullscreen_added:
            new_lines.append("FullscreenMode=2\n")

        return new_lines

    @staticmethod
    def _set_file_permissions(file_path: str, read_only: bool) -> None:
        """Sets the read-only attribute on the file."""
        try:
            if read_only:
                os.chmod(file_path, stat.S_IREAD)
            else:
                os.chmod(file_path, stat.S_IWRITE)
        except Exception as e:
            print(f"Error setting file permissions: {str(e)}")


class ProcessManager:
    PRIORITY_MAP = {
        1: psutil.IDLE_PRIORITY_CLASS,
        2: psutil.BELOW_NORMAL_PRIORITY_CLASS,
        3: psutil.NORMAL_PRIORITY_CLASS,
        4: psutil.ABOVE_NORMAL_PRIORITY_CLASS,
        5: psutil.HIGH_PRIORITY_CLASS,
        6: psutil.REALTIME_PRIORITY_CLASS
    }

    @staticmethod
    def set_priority(process_name: str, priority_level: int) -> None:
        """Sets the priority of a process for all matching instances."""
        if priority_level not in ProcessManager.PRIORITY_MAP:
            raise ValueError("Invalid priority level")

        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == process_name:
                try:
                    p = psutil.Process(proc.info['pid'])
                    p.nice(ProcessManager.PRIORITY_MAP[priority_level])
                    print(f"Set {process_name} (PID: {p.pid}) to priority {priority_level}")
                except (psutil.AccessDenied, psutil.NoSuchProcess) as e:
                    print(f"Error setting priority: {str(e)}")


class ValorantLauncher:
    @staticmethod
    def launch() -> bool:
        """
        Attempts to launch Valorant using different methods.
        First, it tries to locate the Riot Client executable via the registry.
        If that fails, it attempts to locate a shortcut.
        """
        riot_client_path = get_riot_client_path()
        if riot_client_path:
            try:
                subprocess.Popen([riot_client_path, "--launch-product=valorant", "--launch-patchline=live"])
                return True
            except Exception as e:
                print(f"Failed to launch via Riot Client: {str(e)}")

        shortcut = ValorantLauncher._find_shortcut()
        if shortcut:
            try:
                target = ValorantLauncher._get_shortcut_target(shortcut)
                subprocess.Popen([target, "--launch-product=valorant", "--launch-patchline=live"])
                return True
            except Exception as e:
                print(f"Failed to launch via shortcut: {str(e)}")

        print("Could not find Valorant installation")
        return False

    @staticmethod
    def _find_shortcut() -> Optional[str]:
        """Searches for the Valorant shortcut in common locations."""
        locations = [
            os.path.join(os.getenv('PROGRAMDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Riot Games'),
            os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Riot Games')
        ]
        for loc in locations:
            for root, _, files in os.walk(loc):
                for file in files:
                    if file.lower().endswith('.lnk') and 'valorant' in file.lower():
                        return os.path.join(root, file)
        return None

    @staticmethod
    def _get_shortcut_target(shortcut_path: str) -> str:
        """Extracts the target path of a Windows shortcut."""
        shell = win32com.client.Dispatch("WScript.Shell")
        return shell.CreateShortCut(shortcut_path).TargetPath


class PresetManager:
    @staticmethod
    def get_preset_dir() -> str:
        """Returns the path to the presets folder (now TrueStrech)."""
        appdata_path = os.getenv('APPDATA')
        preset_dir = os.path.join(appdata_path, 'TrueStrech', 'presets')
        if not os.path.exists(preset_dir):
            os.makedirs(preset_dir)
        return preset_dir

    @staticmethod
    def list_presets() -> List[str]:
        """Lists all available presets."""
        preset_dir = PresetManager.get_preset_dir()
        return [f for f in os.listdir(preset_dir) if f.endswith('.ini')]

    @staticmethod
    def create_preset() -> None:
        """Creates a new preset using the same menu as the normal mode."""
        ConsoleManager.clear()
        print("=== Create New Preset ===\n")
        
        # Display the current resolution and list available resolutions
        native_res = ResolutionManager.get_current()
        print(f"Current resolution: {native_res[0]}x{native_res[1]}\n")
        print("Available resolutions:")
        resolutions = ResolutionManager.list_available()
        res_choice = int(input("\nSelect the resolution number: ")) - 1
        new_res = resolutions[res_choice]
        ConsoleManager.clear()  # Clear after resolution selection.
        
        # Priority selection menu
        print("Priority levels:")
        print("1: Idle\n2: Below Normal\n3: Normal\n4: Above Normal\n5: High\n6: Realtime")
        priority = int(input("\nSelect the priority level (1-6): "))
        ConsoleManager.clear() 
        
        print("Configuration modification:")
        print("1: Apply to all accounts")
        print("2: Choose a specific account")
        account_choice = input("\nSelect an option (1/2): ")
        ConsoleManager.clear()
        apply_to = "all" if account_choice.strip() == '1' else "specific"
        
        preset_name = input("Preset name: ").strip() + ".ini"
        preset_path = os.path.join(PresetManager.get_preset_dir(), preset_name)
        
        config = configparser.ConfigParser()
        config['Settings'] = {
            'ResolutionX': str(new_res[0]),
            'ResolutionY': str(new_res[1]),
            'Priority': str(priority),
            'ApplyTo': apply_to
        }
        
        with open(preset_path, 'w') as configfile:
            config.write(configfile)
        
        print(f"\nPreset '{preset_name}' created successfully!")
        time.sleep(2)

    @staticmethod
    def load_preset(preset_name: str) -> dict:
        """Loads the configuration from a preset."""
        preset_path = os.path.join(PresetManager.get_preset_dir(), preset_name)
        config = configparser.ConfigParser()
        config.read(preset_path)
        return {
            'resolution': (
                int(config['Settings']['ResolutionX']),
                int(config['Settings']['ResolutionY'])
            ),
            'priority': int(config['Settings']['Priority']),
            'apply_to': config['Settings'].get('ApplyTo', 'all')
        }


def main_menu() -> str:
    """Displays the main menu."""
    ConsoleManager.clear()
    print("=== Main Menu ===")
    print("1 - Use Normal Mode")
    print("2 - Load/Create a Preset")
    print("3 - Quit")
    return input("\nSelect an option: ")


def preset_menu() -> Optional[dict]:
    """Displays the preset management menu."""
    while True:
        ConsoleManager.clear()
        print("=== Preset Management ===")
        presets = PresetManager.list_presets()

        for idx, preset in enumerate(presets, 1):
            print(f"{idx} - {preset[:-4]}")

        print(f"\n{len(presets)+1} - Create a new preset")
        print(f"{len(presets)+2} - Return to main menu")

        choice = input("\nSelect an option: ")

        try:
            choice = int(choice)
            if 1 <= choice <= len(presets):
                return PresetManager.load_preset(presets[choice-1])
            elif choice == len(presets) + 1:
                PresetManager.create_preset()
            elif choice == len(presets) + 2:
                return None
        except (ValueError, IndexError):
            print("Invalid selection!")
            time.sleep(1)


def run_normal_mode():
    """Runs the program in normal mode."""
    native_res = ResolutionManager.get_current()
    print(f"\nCurrent resolution: {native_res[0]}x{native_res[1]}")
    
    # Resolution selection
    print("\nAvailable resolutions:")
    resolutions = ResolutionManager.list_available()
    choice = int(input("\nSelect the resolution number: ")) - 1
    new_res = resolutions[choice]
    ConsoleManager.clear() 
    
    # Apply the selected system resolution
    ResolutionManager.set_resolution(*new_res)
    print(f"System resolution set to: {new_res[0]}x{new_res[1]}")
    time.sleep(1)
    ConsoleManager.clear()
    
    # Priority selection
    print("Priority levels:")
    print("1: Idle\n2: Below Normal\n3: Normal\n4: Above Normal\n5: High\n6: Realtime")
    priority = int(input("\nSelect the priority level (1-6): "))
    ConsoleManager.clear() 
    
    # Modify configuration files
    config_manager = ValorantConfigManager()
    config_files = config_manager.find_config_files()
    
    print("Configuration modification:")
    print("1: Apply to all accounts")
    print("2: Choose a specific account")
    acc_option = input("Select an option (1/2): ")
    ConsoleManager.clear()  
    
    if acc_option.strip() == '1':
        for _, path in config_files:
            config_manager.modify_config(path, *new_res)
    elif acc_option.strip() == '2':
        print("Available accounts:")
        for i, (account, _) in enumerate(config_files, 1):
            print(f"{i}: {account}")
        acc_choice = int(input("\nSelect the account number: ")) - 1
        config_manager.modify_config(config_files[acc_choice][1], *new_res)
    ConsoleManager.clear()  
    
    print("Launching Valorant...")
    ValorantLauncher.launch()
    
    print("\nWaiting for Valorant to start...")
    while True:
        if any(p.name() == VALORANT_PROCESS_NAME for p in psutil.process_iter()):
            ProcessManager.set_priority(VALORANT_PROCESS_NAME, priority)
            break
        time.sleep(5)
    
    print("\nValorant is running... (Close the window to exit)")
    while any(p.name() == VALORANT_PROCESS_NAME for p in psutil.process_iter()):
        time.sleep(5)
    
    return native_res


def main():
    ConsoleManager.set_color(FOREGROUND_RED)
    ConsoleManager.clear()
    ConsoleManager.set_title(os.getlogin())

    native_res = None
    config_files = []

    try:
        while True:
            choice = main_menu()

            if choice == '1':
                native_res = run_normal_mode()
                break
            elif choice == '2':
                preset = preset_menu()
                ConsoleManager.clear()
                if preset:
                    native_res = ResolutionManager.get_current()
                    new_res = preset['resolution']
                    priority = preset['priority']
                    ResolutionManager.set_resolution(*new_res)
                    config_manager = ValorantConfigManager()
                    config_files = config_manager.find_config_files()

                    if preset.get('apply_to', 'all') == 'specific':
                        print("Available accounts:")
                        for i, (account, _) in enumerate(config_files, 1):
                            print(f"{i}: {account}")
                        acc_choice = int(input("\nSelect the account number: ")) - 1
                        config_manager.modify_config(config_files[acc_choice][1], *new_res)
                    else:
                        for _, path in config_files:
                            config_manager.modify_config(path, *new_res)

                    print("Launching Valorant...")
                    ValorantLauncher.launch()

                    print("\nWaiting for Valorant to start...")
                    while True:
                        if any(p.name() == VALORANT_PROCESS_NAME for p in psutil.process_iter()):
                            ProcessManager.set_priority(VALORANT_PROCESS_NAME, priority)
                            break
                        time.sleep(5)

                    print("\nValorant is running... (Close the window to exit)")
                    while any(p.name() == VALORANT_PROCESS_NAME for p in psutil.process_iter()):
                        time.sleep(5)

                    break
            elif choice == '3':
                print("\nClosing...")
                return
            else:
                print("Invalid option!")
                time.sleep(1)

    except Exception as e:
        print(f"\nError : {str(e)}")
    finally:
        if native_res:
            ResolutionManager.set_resolution(*native_res)
            if config_files:
                config_manager = ValorantConfigManager()
                for _, path in config_files:
                    config_manager.modify_config(path, native_res[0], native_res[1], reset=True)
        print("\n(Settings restored. Closing...)")


if __name__ == "__main__":
    main()