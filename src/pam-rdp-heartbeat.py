"""
MIT License

Copyright (c) 2024-2026 PAM-Exchange

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import sys
import os
import re
import time
import math
import ctypes
import ctypes.wintypes
import threading
import configparser
import winreg
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

# Program Metadata
gVersion = "1.2.0.13"
gProgramTitle = "PAM RDP Heartbeat"
gSelfPid = os.getpid()
gSelfPidHex = f"{gSelfPid:04x}"

# Log Levels
LOG_ALWAYS = 0
LOG_ERROR = 1
LOG_WARNING = 2
LOG_INFO = 3
LOG_DEBUG = 4
LOG_TRACE = 5

temp_dir = os.environ.get("TEMP", os.environ.get("TMP", os.path.expanduser("~")))
gLogFile = os.path.join(temp_dir, "pam-rdp-heartbeat.log")
gLogLevel = LOG_DEBUG

# Win32 API Constants
WM_ACTIVATE = 0x0006
WA_ACTIVE = 1
WM_SYSCOMMAND = 0x0112
SC_RESTORE = 0xF120
SC_MAXIMIZE = 0xF030
WM_HOTKEY = 0x0312
WM_USER = 0x0400
WM_TRAYICON = WM_USER + 1
WM_CLOSE = 0x0010
WM_COMMAND = 0x0111
WM_NULL = 0x0000
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101

WS_EX_LAYERED = 0x00080000
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000
GWL_EXSTYLE = -20
GWL_STYLE = -16
LWA_ALPHA = 0x00000002

HWND_BOTTOM = 1
HWND_TOP = 0
HWND_TOPMOST = -1
HWND_NOTOPMOST = -2

SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOACTIVATE = 0x0010

SW_HIDE = 0
SW_SHOWNORMAL = 1
SW_SHOWMINIMIZED = 2
SW_MAXIMIZE = 3
SW_SHOWNOACTIVATE = 4
SW_SHOW = 5
SW_MINIMIZE = 6
SW_RESTORE = 9

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
VK_SCROLL = 0x91
VK_NUMLOCK = 0x90

EVENT_SYSTEM_MINIMIZESTART = 0x0016
WINEVENT_OUTOFCONTEXT = 0x0001
GW_OWNER = 4

NIM_ADD = 0x00000000
NIM_MODIFY = 0x00000001
NIM_DELETE = 0x00000002
NIF_ICON = 0x00000001
NIF_MESSAGE = 0x00000002
NIF_TIP = 0x00000004

WM_LBUTTONDBLCLK = 0x0203
WM_RBUTTONUP = 0x0205
WM_LBUTTONUP = 0x0202

HSHELL_WINDOWACTIVATED = 4
HSHELL_RUDEAPPACTIVATED = 0x8004

# Win32 Structures
class SYSTEMTIME(ctypes.Structure):
    _fields_ = [
        ("wYear", ctypes.wintypes.WORD),
        ("wMonth", ctypes.wintypes.WORD),
        ("wDayOfWeek", ctypes.wintypes.WORD),
        ("wDay", ctypes.wintypes.WORD),
        ("wHour", ctypes.wintypes.WORD),
        ("wMinute", ctypes.wintypes.WORD),
        ("wSecond", ctypes.wintypes.WORD),
        ("wMilliseconds", ctypes.wintypes.WORD),
    ]

class FILETIME(ctypes.Structure):
    _fields_ = [
        ("dwLowDateTime", ctypes.wintypes.DWORD),
        ("dwHighDateTime", ctypes.wintypes.DWORD),
    ]

class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]

class WINDOWPLACEMENT(ctypes.Structure):
    _fields_ = [
        ("length", ctypes.wintypes.UINT),
        ("flags", ctypes.wintypes.UINT),
        ("showCmd", ctypes.wintypes.UINT),
        ("ptMinPosition", ctypes.wintypes.POINT),
        ("ptMaxPosition", ctypes.wintypes.POINT),
        ("rcNormalPosition", RECT),
    ]

class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.wintypes.DWORD),
        ("hWnd", ctypes.wintypes.HWND),
        ("uID", ctypes.wintypes.UINT),
        ("uFlags", ctypes.wintypes.UINT),
        ("uCallbackMessage", ctypes.wintypes.UINT),
        ("hIcon", ctypes.wintypes.HICON),
        ("szTip", ctypes.c_wchar * 128),
        ("dwState", ctypes.wintypes.DWORD),
        ("dwStateMask", ctypes.wintypes.DWORD),
        ("szInfo", ctypes.c_wchar * 256),
        ("uTimeoutOrVersion", ctypes.wintypes.UINT),
        ("szInfoTitle", ctypes.c_wchar * 64),
        ("dwInfoFlags", ctypes.wintypes.DWORD),
    ]

# Win32 Functions
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
gdi32 = ctypes.windll.gdi32
shell32 = ctypes.windll.shell32
dwmapi = ctypes.windll.dwmapi

WINEVENTPROC = ctypes.WINFUNCTYPE(
    None,
    ctypes.wintypes.HANDLE,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.HWND,
    ctypes.c_long,
    ctypes.c_long,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD,
)

# Set restype/argtypes for 64-bit safety
user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.GetWindowLongPtrW.argtypes = [ctypes.wintypes.HWND, ctypes.c_int]

user32.SetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.SetWindowLongPtrW.argtypes = [ctypes.wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]

user32.CallWindowProcW.restype = ctypes.c_ssize_t
user32.CallWindowProcW.argtypes = [ctypes.c_ssize_t, ctypes.wintypes.HWND, ctypes.wintypes.UINT, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM]

user32.GetWindowLongW.restype = ctypes.c_long
user32.GetWindowLongW.argtypes = [ctypes.wintypes.HWND, ctypes.c_int]

user32.SetWindowLongW.restype = ctypes.c_long
user32.SetWindowLongW.argtypes = [ctypes.wintypes.HWND, ctypes.c_int, ctypes.c_long]

kernel32.OpenProcess.restype = ctypes.wintypes.HANDLE
kernel32.OpenProcess.argtypes = [ctypes.wintypes.DWORD, ctypes.wintypes.BOOL, ctypes.wintypes.DWORD]

user32.SetWinEventHook.restype = ctypes.wintypes.HANDLE
user32.SetWinEventHook.argtypes = [
    ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.HMODULE,
    WINEVENTPROC,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD
]

user32.UnhookWinEvent.restype = ctypes.wintypes.BOOL
user32.UnhookWinEvent.argtypes = [ctypes.wintypes.HANDLE]

user32.RegisterShellHookWindow.restype = ctypes.wintypes.BOOL
user32.RegisterShellHookWindow.argtypes = [ctypes.wintypes.HWND]

user32.DeregisterShellHookWindow.restype = ctypes.wintypes.BOOL
user32.DeregisterShellHookWindow.argtypes = [ctypes.wintypes.HWND]

user32.RegisterWindowMessageW.restype = ctypes.wintypes.UINT
user32.RegisterWindowMessageW.argtypes = [ctypes.wintypes.LPCWSTR]

# Logging Utilities
def log(level, line, msg):
    global gLogLevel, gLogFile, gSelfPidHex
    if level <= gLogLevel:
        level_map = {
            LOG_ALWAYS: "ALW",
            LOG_ERROR: "ERR",
            LOG_WARNING: "WRN",
            LOG_INFO: "INF",
            LOG_DEBUG: "DBG",
            LOG_TRACE: "TRC",
        }
        level_str = level_map.get(level, "DBG")
        timestamp = datetime.now().strftime("%y/%m/%d %H:%M:%S")
        txt = f"{timestamp} [{gSelfPidHex}] {level_str} {msg} [{line}]\n"
        try:
            with open(gLogFile, "a", encoding="utf-8") as f:
                f.write(txt)
        except Exception:
            pass

def logAlways(line, msg): log(LOG_ALWAYS, line, msg)
def logError(line, msg): log(LOG_ERROR, line, msg)
def logWarning(line, msg): log(LOG_WARNING, line, msg)
def logInfo(line, msg): log(LOG_INFO, line, msg)
def logDebug(line, msg): log(LOG_DEBUG, line, msg)
def logTrace(line, msg): log(LOG_TRACE, line, msg)

def logRoll(filename, max_size_mb=5, keep=5):
    if not os.path.exists(filename):
        return
    try:
        size = os.path.getsize(filename)
        if size >= max_size_mb * 1024 * 1024:
            logDebug(sys._getframe().f_lineno, f"LogRoll: roll files - filename= '{filename}', size= {size} bytes")
            for k in range(keep - 1, 0, -1):
                fn1 = f"{filename}.{k + 1}"
                fn2 = f"{filename}.{k}"
                if os.path.exists(fn2):
                    if os.path.exists(fn1):
                        os.remove(fn1)
                    os.rename(fn2, fn1)
            fn1 = f"{filename}.1"
            if os.path.exists(fn1):
                os.remove(fn1)
            os.rename(filename, fn1)
    except Exception as e:
        logError(sys._getframe().f_lineno, f"LogRoll error: {e}")

# Registry Helper
def readRegistry(root_key, subkey, value_name, default_val=0):
    try:
        key = winreg.OpenKey(root_key, subkey, 0, winreg.KEY_READ)
        val, _ = winreg.QueryValueEx(key, value_name)
        winreg.CloseKey(key)
        return val
    except Exception:
        return default_val

# Process Utilities
def getProcessCreationTime(pid):
    try:
        hProcess = kernel32.OpenProcess(0x0410, False, pid)  # PROCESS_QUERY_INFORMATION | PROCESS_VM_READ
        if not hProcess:
            return None
        ftCreate = FILETIME()
        ftExit = FILETIME()
        ftKernel = FILETIME()
        ftUser = FILETIME()
        success = kernel32.GetProcessTimes(
            hProcess,
            ctypes.byref(ftCreate),
            ctypes.byref(ftExit),
            ctypes.byref(ftKernel),
            ctypes.byref(ftUser)
        )
        kernel32.CloseHandle(hProcess)
        if not success:
            return None

        ftLocal = FILETIME()
        kernel32.FileTimeToLocalFileTime(ctypes.byref(ftCreate), ctypes.byref(ftLocal))
        st = SYSTEMTIME()
        kernel32.FileTimeToSystemTime(ctypes.byref(ftLocal), ctypes.byref(st))

        dt = datetime(st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond)
        return dt
    except Exception as e:
        logError(sys._getframe().f_lineno, f"getProcessCreationTime error: {e}")
        return None

def formatSeconds(seconds):
    seconds = int(seconds)
    hrs = seconds // 3600
    seconds %= 3600
    mins = seconds // 60
    secs = seconds % 60
    return f"{hrs:02d}:{mins:02d}:{secs:02d}"

def isWindowFullScreen(hwnd):
    if not hwnd or not user32.IsWindow(hwnd):
        return False
    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w = rect.right - rect.left
    h = rect.bottom - rect.top
    screen_w = user32.GetSystemMetrics(0)
    screen_h = user32.GetSystemMetrics(1)
    if (style & 0x20800000) or h < screen_h or w < screen_w:
        return False
    return True

def isInvisibleWin10BackgroundAppWindow(hwnd):
    cloaked = ctypes.c_int(0)
    hr = dwmapi.DwmGetWindowAttribute(hwnd, 14, ctypes.byref(cloaked), ctypes.sizeof(cloaked))
    if hr == 0:
        return bool(cloaked.value)
    return False

# Core Heartbeat Application Class
class PAMRDPHeartbeatApp:
    def __init__(self, root):
        self.root = root
        self.root.title(gProgramTitle)

        # Internal State
        self.window_list = {}  # hwnd: state (0: normal, 1: fullscreen hidden, 2: windowed hidden)
        self.click_active = True
        self.gui_hidden = False
        self.win_event_hook = None
        self.event_hook_proc_c = None
        self.shell_msg_id = 0
        self.tray_nid = None

        # Configuration defaults
        self.set_variable_defaults()
        self.load_variables_from_property()
        self.load_variables_from_registry()

        # Validate screensaver
        if not self.test_screensaver():
            logError(sys._getframe().f_lineno, "ScreenSaver test failed. Exiting.")
            sys.exit(1)

        # Setup GUI & Timers
        self.setup_ui()
        self.setup_system_tray()
        self.setup_hooks_and_hotkeys()

        self.refresh_process()
        self.start_timers()

        self.root.protocol("WM_DELETE_WINDOW", self.on_gui_close)

    def set_variable_defaults(self):
        self.margin_x = 10
        self.margin_y = 10
        self.btn_width = 80
        self.btn_height = 26
        self.start_x = 10
        self.start_y = 10
        self.language = "en-US"

        self.heartbeat_frequency = 60
        self.heartbeat_frequency_txt = "&Heartbeat frequency (sec)"
        self.heartbeat_minimum = 10
        self.refresh_frequency = 10
        self.refresh_frequency_txt = "&Refresh frequency (sec)"
        self.refresh_minimum = 5
        self.start_minimized = 0
        self.start_minimized_txt = "Start &Minimized"
        self.stay_on_top = 1
        self.stay_on_top_txt = "Stay on &Top"
        self.use_transparent = 1
        self.use_transparent_txt = "Use Transparent"
        self.screensaver_required = 1
        self.screensaver_idle_maximum = 1800  # 30 minutes

        self.btn_exit_txt = "E&xit"
        self.btn_close_txt = "&Close"
        self.btn_update_txt = "&Update"
        self.btn_cancel_txt = "&Cancel"

        self.list_rows = 10
        self.list_col_title_txt = "Title"
        self.list_col_title_wdt = 170
        self.list_col_pid_txt = "PID"
        self.list_col_pid_wdt = 50
        self.list_col_start_txt = "Start"
        self.list_col_start_wdt = 55
        self.list_col_duration_txt = "Duration"
        self.list_col_duration_wdt = 55
        self.list_col_state_txt = "State"
        self.list_col_state_wdt = 50
        self.list_col_hwnd_txt = "HWND"
        self.list_col_hwnd_wdt = 70

        self.mnu_file_txt = "&File"
        self.mnu_file_exit_txt = "E&xit"
        self.mnu_settings_txt = "&Settings"
        self.mnu_settings_stay_on_top_txt = "Stay on &Top"
        self.mnu_settings_start_minimized_txt = "Start &Minimized"
        self.mnu_settings_preferences_txt = "&Preferences"
        self.mnu_action_txt = "&Action"
        self.mnu_action_refresh_txt = "&Refresh"
        self.mnu_action_show_all_txt = "&Show All"
        self.mnu_action_minimize_all_txt = "&Minimize All"
        self.mnu_help_txt = "&Help"
        self.mnu_help_quick_guide_txt = "Quick &guide"
        self.mnu_help_about_txt = "&About"

        self.about_txt = "The program sends heartbeat signals to RDP sessions. When RDP sessions are minimized, the RDP GUI is hidden and heartbeat messages are still being accepted by the session."
        self.quick_guide_txt = (
            "In the menu item 'Settings > Preferences' timers for background actions can be defined. "
            "The heartbeat timer will control how often the heartbeat is send to both active and hidden RDP sessions. "
            "The value should depend of the timeout values defined on the servers connected to from the Desktop. "
            "Typically, sending a heartbeat signal every 30'th second is sufficient. "
            "The refresh timer will control how often the list of RDP sessions is updated. Typical values are every 10 to 30 seconds. "
            "It is also possible to refresh the list manually by pressing F5. "
            "The flags 'Stay on Top' and 'Start Minimized' are controlling the PAM RDP Heartbeat window behaviour.\n\n"
            "Alt-Ctrl-ScrollLock\nWill show/activate or hide the GUI window. This will work even when showing an RDP fullscreen session.\n\n"
            "When the PAM RDP Heartbeat GUI is active, additional functions are available.\n\n"
            "Alt-Ctrl-Down\nMinimize/hide all RDP windows.\n\n"
            "Alt-Ctrl-Up\nRestore all hidden RDP windows.\n\n"
            "F5\nManually refresh the list of RDP sessions."
        )

        appdata = os.environ.get("APPDATA", "")
        prop_dir = os.path.join(appdata, "PAM-Exchange", "PAM-RDP-Heartbeat")
        os.makedirs(prop_dir, exist_ok=True)
        self.user_property_filename = os.path.join(prop_dir, "pam-rdp-heartbeat.properties")

    def load_variables_from_property(self):
        global gLogLevel
        if not os.path.exists(self.user_property_filename):
            logAlways(sys._getframe().f_lineno, f"Properties file not found '{self.user_property_filename}' - using defaults")
            return

        config = configparser.ConfigParser()
        try:
            config.read(self.user_property_filename, encoding="utf-8")
            if "main" in config:
                main_sec = config["main"]
                log_level_str = main_sec.get("LogLevel", "INFO").upper()
                level_map = {"ERROR": LOG_ERROR, "WARNING": LOG_WARNING, "INFO": LOG_INFO, "DEBUG": LOG_DEBUG, "TRACE": LOG_TRACE}
                gLogLevel = level_map.get(log_level_str, LOG_DEBUG)

                self.heartbeat_frequency = main_sec.getint("HeartbeatFrequency", self.heartbeat_frequency)
                self.refresh_frequency = main_sec.getint("RefreshFrequency", self.refresh_frequency)
                self.start_minimized = 1 if main_sec.getboolean("StartMinimized", bool(self.start_minimized)) else 0
                self.start_x = main_sec.getint("StartX", self.start_x)
                self.start_y = main_sec.getint("StartY", self.start_y)
                self.stay_on_top = 1 if main_sec.getboolean("StayOnTop", bool(self.stay_on_top)) else 0
                self.language = main_sec.get("Language", self.language)
        except Exception as e:
            logError(sys._getframe().f_lineno, f"Error reading properties file: {e}")

    def load_variables_from_registry(self):
        key = r"Software\PAM-Exchange\PAM-RDP-Heartbeat"
        self.heartbeat_minimum = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "HeartbeatMinimum", self.heartbeat_minimum))
        self.list_col_duration_wdt = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "listColDurationWdt", self.list_col_duration_wdt))
        self.list_col_hwnd_wdt = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "listColHwndWdt", self.list_col_hwnd_wdt))
        self.list_col_pid_wdt = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "listColPIDWdt", self.list_col_pid_wdt))
        self.list_col_start_wdt = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "listColStartWdt", self.list_col_start_wdt))
        self.list_col_state_wdt = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "listColStateWdt", self.list_col_state_wdt))
        self.list_col_title_wdt = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "listColTitleWdt", self.list_col_title_wdt))
        self.list_rows = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "listRows", self.list_rows))
        self.refresh_minimum = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "RefreshMinimum", self.refresh_minimum))
        self.use_transparent = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "UseTransparent", self.use_transparent))
        self.screensaver_required = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "ScreenSaverRequired", self.screensaver_required))
        self.screensaver_idle_maximum = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, key, "ScreenSaverIdleMaximum", self.screensaver_idle_maximum))

        if self.heartbeat_frequency < self.heartbeat_minimum:
            self.heartbeat_frequency = self.heartbeat_minimum
        if self.refresh_frequency < self.refresh_minimum:
            self.refresh_frequency = self.refresh_minimum

    def save_variables_to_property(self):
        try:
            x = self.root.winfo_x()
            y = self.root.winfo_y()
            config = configparser.ConfigParser()
            level_map = {LOG_ERROR: "ERROR", LOG_WARNING: "WARNING", LOG_INFO: "INFO", LOG_DEBUG: "DEBUG", LOG_TRACE: "TRACE"}
            config["main"] = {
                "LogLevel": level_map.get(gLogLevel, "DEBUG"),
                "HeartbeatFrequency": str(self.heartbeat_frequency),
                "RefreshFrequency": str(self.refresh_frequency),
                "StartX": str(x),
                "StartY": str(y),
                "StartMinimized": "true" if self.start_minimized else "false",
                "StayOnTop": "true" if self.stay_on_top else "false",
                "Language": self.language,
            }
            with open(self.user_property_filename, "w", encoding="utf-8") as f:
                config.write(f)
            logDebug(sys._getframe().f_lineno, f"Saved properties to '{self.user_property_filename}'")
        except Exception as e:
            logError(sys._getframe().f_lineno, f"save_variables_to_property error: {e}")

    def test_screensaver(self):
        if not self.screensaver_required:
            logInfo(sys._getframe().f_lineno, "TestScreenSaver: ScreenSaver on user's desktop is not required")
            return True

        ss_machine_ok = False
        ss_inactivity_timeout = int(readRegistry(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "InactivityTimeoutSecs", 0))
        if ss_inactivity_timeout != 0:
            ss_machine_ok = True

        ss_local_ok = False
        ss_local_active = int(readRegistry(winreg.HKEY_CURRENT_USER, r"Control Panel\Desktop", "ScreenSaveActive", 0))
        ss_local_is_secure = int(readRegistry(winreg.HKEY_CURRENT_USER, r"Control Panel\Desktop", "ScreenSaverIsSecure", 0))
        ss_local_timeout = int(readRegistry(winreg.HKEY_CURRENT_USER, r"Control Panel\Desktop", "ScreenSaveTimeOut", 0))
        if ss_local_active == 1 and ss_local_is_secure == 1 and ss_local_timeout <= self.screensaver_idle_maximum:
            ss_local_ok = True

        ss_gpo_ok = False
        ss_gpo_active = int(readRegistry(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "ScreenSaveActive", 0))
        ss_gpo_is_secure = int(readRegistry(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "ScreenSaverIsSecure", 0))
        ss_gpo_timeout = int(readRegistry(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "ScreenSaveTimeOut", 0))
        if ss_gpo_active == 1 and ss_gpo_is_secure == 1 and ss_gpo_timeout <= self.screensaver_idle_maximum:
            ss_gpo_ok = True

        if not (ss_local_ok or ss_gpo_ok or ss_machine_ok):
            logError(sys._getframe().f_lineno, "TestScreenSaver: Active screensaver not found or invalid timeout. PAM-RDP-Heartbeat is not started")
            messagebox.showerror(
                gProgramTitle,
                "An active password protected ScreenSaver is not identified on the desktop. "
                "If defined, the timeout may be set beyond the allowed maximum. "
                "This is a mandatory requirement when using the PAM RDP Heartbeat program.\n\n"
                "PAM RDP Heartbeat is not started."
            )
            return False
        return True

    def setup_ui(self):
        if self.stay_on_top:
            self.root.attributes("-topmost", True)

        # Menu bar
        menubar = tk.Menu(self.root)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label=self.mnu_file_exit_txt, command=self.on_exit)
        menubar.add_cascade(label=self.mnu_file_txt.replace("&", ""), menu=file_menu)

        # Action Menu
        action_menu = tk.Menu(menubar, tearoff=0)
        action_menu.add_command(label=f"{self.mnu_action_refresh_txt}\tF5".replace("&", ""), command=self.refresh_process)
        action_menu.add_command(label=self.mnu_action_show_all_txt.replace("&", ""), command=self.show_all_mstsc_window)
        action_menu.add_command(label=self.mnu_action_minimize_all_txt.replace("&", ""), command=self.minimize_all_rdp)
        menubar.add_cascade(label=self.mnu_action_txt.replace("&", ""), menu=action_menu)

        # Settings Menu
        self.stay_on_top_var = tk.BooleanVar(value=bool(self.stay_on_top))
        self.start_minimized_var = tk.BooleanVar(value=bool(self.start_minimized))

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_checkbutton(label=self.mnu_settings_stay_on_top_txt.replace("&", ""), variable=self.stay_on_top_var, command=self.toggle_stay_on_top)
        settings_menu.add_checkbutton(label=self.mnu_settings_start_minimized_txt.replace("&", ""), variable=self.start_minimized_var, command=self.toggle_start_minimized)
        settings_menu.add_separator()
        settings_menu.add_command(label=self.mnu_settings_preferences_txt.replace("&", ""), command=self.open_preferences)
        menubar.add_cascade(label=self.mnu_settings_txt.replace("&", ""), menu=settings_menu)

        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label=self.mnu_help_quick_guide_txt.replace("&", ""), command=self.show_quick_guide)
        help_menu.add_separator()
        help_menu.add_command(label=self.mnu_help_about_txt.replace("&", ""), command=self.show_about)
        menubar.add_cascade(label=self.mnu_help_txt.replace("&", ""), menu=help_menu)

        self.root.config(menu=menubar)

        # Session Table (Treeview)
        columns = ("Title", "PID", "Start", "Duration", "State", "HWND")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", height=self.list_rows)

        self.tree.heading("Title", text=self.list_col_title_txt, command=lambda: self.sort_column("Title", False))
        self.tree.heading("PID", text=self.list_col_pid_txt, command=lambda: self.sort_column("PID", False))
        self.tree.heading("Start", text=self.list_col_start_txt, command=lambda: self.sort_column("Start", False))
        self.tree.heading("Duration", text=self.list_col_duration_txt, command=lambda: self.sort_column("Duration", False))
        self.tree.heading("State", text=self.list_col_state_txt, command=lambda: self.sort_column("State", False))
        self.tree.heading("HWND", text=self.list_col_hwnd_txt, command=lambda: self.sort_column("HWND", False))

        self.tree.column("Title", width=self.list_col_title_wdt)
        self.tree.column("PID", width=self.list_col_pid_wdt, anchor="center")
        self.tree.column("Start", width=self.list_col_start_wdt, anchor="center")
        self.tree.column("Duration", width=self.list_col_duration_wdt, anchor="center")
        self.tree.column("State", width=self.list_col_state_wdt, anchor="center")
        self.tree.column("HWND", width=self.list_col_hwnd_wdt, anchor="center")

        self.tree.pack(padx=self.margin_x, pady=self.margin_y, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self.on_double_click_session)
        self.tree.bind("<Button-3>", self.on_right_click_session)

        # Buttons
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(padx=self.margin_x, pady=(0, self.margin_y), fill=tk.X)

        self.btn_exit = ttk.Button(btn_frame, text=self.btn_exit_txt.replace("&", ""), command=self.on_exit)
        self.btn_exit.pack(side=tk.LEFT)

        self.btn_close = ttk.Button(btn_frame, text=self.btn_close_txt.replace("&", ""), command=self.btn_close_event)
        self.btn_close.pack(side=tk.RIGHT)

        # Keyboard Bindings
        self.root.bind("<F5>", lambda e: self.refresh_process())
        self.root.bind("<Control-Alt-Down>", lambda e: self.minimize_all_rdp())
        self.root.bind("<Control-Alt-Up>", lambda e: self.show_all_mstsc_window())
        self.root.bind("<Escape>", lambda e: self.btn_close_event())

        # Set Window Position and Initial Visibility
        self.root.geometry(f"+{self.start_x}+{self.start_y}")
        if self.start_minimized:
            self.root.withdraw()
            self.gui_hidden = True

    def sort_column(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.getchildren('')]
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)
        self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))

    def setup_system_tray(self):
        hwnd = self.root.winfo_id()

        # Register Shell Hook
        self.shell_msg_id = user32.RegisterWindowMessageW("SHELLHOOK")
        user32.RegisterShellHookWindow(hwnd)

        self.tray_nid = NOTIFYICONDATAW()
        self.tray_nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        self.tray_nid.hWnd = hwnd
        self.tray_nid.uID = 1001
        self.tray_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        self.tray_nid.uCallbackMessage = WM_TRAYICON
        self.tray_nid.hIcon = user32.LoadIconW(0, ctypes.cast(32512, ctypes.wintypes.LPCWSTR))  # IDI_APPLICATION
        self.tray_nid.szTip = gProgramTitle

        shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(self.tray_nid))

        # Subclass Tkinter WndProc to handle Win32 messages (WM_TRAYICON, WM_HOTKEY, ShellHook)
        self.old_wnd_proc = user32.GetWindowLongPtrW(hwnd, -4)  # GWLP_WNDPROC = -4

        WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.wintypes.HWND, ctypes.wintypes.UINT, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM)

        def py_wnd_proc(hWnd, msg, wParam, lParam):
            if msg == WM_TRAYICON:
                if lParam in (WM_LBUTTONDBLCLK, WM_LBUTTONUP):
                    self.program_unhide()
                elif lParam == WM_RBUTTONUP:
                    self.show_tray_menu()
                return 0
            elif msg == WM_HOTKEY:
                if wParam == 1:  # Ctrl-Alt-ScrollLock
                    self.toggle_program_visibility()
                return 0
            elif self.shell_msg_id and msg == self.shell_msg_id:
                if wParam in (HSHELL_WINDOWACTIVATED, HSHELL_RUDEAPPACTIVATED):
                    buf = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(lParam, buf, 256)
                    if buf.value == "TscShellContainerClass":
                        if self.click_active:
                            logDebug(sys._getframe().f_lineno, f"ShellHook: HSHELL_WINDOWACTIVATED lParam={lParam}")
                            self.activate_window(lParam)
                return 0

            return user32.CallWindowProcW(self.old_wnd_proc, hWnd, msg, wParam, lParam)

        self.wnd_proc_c = WNDPROC(py_wnd_proc)
        wnd_proc_ptr = ctypes.cast(self.wnd_proc_c, ctypes.c_void_p).value
        user32.SetWindowLongPtrW(hwnd, -4, wnd_proc_ptr)

    def show_tray_menu(self):
        tray_menu = tk.Menu(self.root, tearoff=0)
        tray_menu.add_command(label="Show", command=self.program_unhide)
        tray_menu.add_separator()
        tray_menu.add_command(label="Exit", command=self.on_exit)
        try:
            pt = ctypes.wintypes.POINT()
            user32.GetCursorPos(ctypes.byref(pt))
            tray_menu.tk_popup(pt.x, pt.y)
        finally:
            tray_menu.grab_release()

    def setup_hooks_and_hotkeys(self):
        hwnd = self.root.winfo_id()
        # Register HotKey: Ctrl+Alt+ScrollLock
        success = user32.RegisterHotKey(hwnd, 1, MOD_CONTROL | MOD_ALT, VK_SCROLL)
        logInfo(sys._getframe().f_lineno, f"RegisterHotKey Ctrl+Alt+ScrollLock: {success}")

        # WinEventHook for EVENT_SYSTEM_MINIMIZESTART (0x0016)
        def event_hook_proc(hWinEventHook, event, hwnd_evt, idObject, idChild, dwEventThread, dwmsEventTime):
            if event == EVENT_SYSTEM_MINIMIZESTART:
                buf = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd_evt, buf, 256)
                class_name = buf.value
                logDebug(sys._getframe().f_lineno, f"EventHookProc: EVENT_SYSTEM_MINIMIZESTART class= {class_name}, hwnd= {hwnd_evt}")
                if class_name == "TscShellContainerClass":
                    self.hide_window(hwnd_evt)

        self.event_hook_proc_c = WINEVENTPROC(event_hook_proc)
        self.win_event_hook = user32.SetWinEventHook(
            EVENT_SYSTEM_MINIMIZESTART,
            EVENT_SYSTEM_MINIMIZESTART,
            0,
            self.event_hook_proc_c,
            0,
            0,
            WINEVENT_OUTOFCONTEXT
        )

    def start_timers(self):
        self.heartbeat_timer_callback()
        self.refresh_timer_callback()

    def heartbeat_timer_callback(self):
        self.heartbeat_process()
        self.root.after(self.heartbeat_frequency * 1000, self.heartbeat_timer_callback)

    def refresh_timer_callback(self):
        self.refresh_process()
        self.root.after(self.refresh_frequency * 1000, self.refresh_timer_callback)

    def refresh_process(self):
        for item in self.tree.getchildren():
            self.tree.delete(item)

        # 1. Enumerate mstsc windows (TscShellContainerClass)
        def enum_mstsc(hwnd, lparam):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buf, 256)
            if buf.value == "TscShellContainerClass":
                title_buf = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(hwnd, title_buf, 256)
                pid = ctypes.wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

                ct = getProcessCreationTime(pid.value)
                if ct:
                    now = datetime.now()
                    delta = (now - ct).total_seconds()
                    duration = formatSeconds(delta)
                    start_time = ct.strftime("%H:%M:%S")
                else:
                    duration = "00:00:00"
                    start_time = "00:00:00"

                if hwnd not in self.window_list:
                    self.window_list[hwnd] = 0
                state = self.window_list[hwnd]

                self.tree.insert("", tk.END, values=(title_buf.value, pid.value, start_time, duration, state, f"0x{hwnd:08x}"))
                logTrace(sys._getframe().f_lineno, f"RefreshProcess mstsc: title='{title_buf.value}', pid={pid.value}, hwnd={hwnd}")
            return True

        ENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        user32.EnumWindows(ENUMPROC(enum_mstsc), 0)

        # 2. Enumerate Symantec PAM applets (CAPAMClient.exe / SunAwtFrame)
        def enum_pam(hwnd, lparam):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buf, 256)
            if buf.value == "SunAwtFrame":
                pid = ctypes.wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                title_buf = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(hwnd, title_buf, 256)
                title = title_buf.value

                # Filter out PAM non-RDP applets
                exclude = ["CA PAM LDAP Browser", "CA Privileged Access Manager", "Symantec PAM LDAP Browser", "Symantec Privileged Access Manager"]
                if not any(ex in title for ex in exclude):
                    ct = getProcessCreationTime(pid.value)
                    if ct:
                        delta = (datetime.now() - ct).total_seconds()
                        duration = formatSeconds(delta)
                        start_time = ct.strftime("%H:%M:%S")
                    else:
                        duration = "00:00:00"
                        start_time = "00:00:00"

                    if hwnd not in self.window_list:
                        self.window_list[hwnd] = 0
                    state = self.window_list[hwnd]

                    self.tree.insert("", tk.END, values=(title, pid.value, start_time, duration, state, f"0x{hwnd:08x}"))
                    logTrace(sys._getframe().f_lineno, f"RefreshProcess SymantecPAM: title='{title}', pid={pid.value}, hwnd={hwnd}")
            return True

        user32.EnumWindows(ENUMPROC(enum_pam), 0)

    def heartbeat_process(self):
        # Heartbeat to mstsc windows
        def enum_mstsc_hb(hwnd, lparam):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buf, 256)
            if buf.value == "TscShellContainerClass":
                if user32.IsIconic(hwnd):
                    self.hide_window(hwnd)
                # WM_ACTIVATE(0x0006) WA_ACTIVE(1)
                user32.PostMessageW(hwnd, WM_ACTIVATE, WA_ACTIVE, 0)
                logDebug(sys._getframe().f_lineno, f"HeartbeatProcess: (Mstsc) PostMessage WM_ACTIVATE to hwnd= {hwnd}")
            return True

        ENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        user32.EnumWindows(ENUMPROC(enum_mstsc_hb), 0)

        # Heartbeat to Symantec PAM windows (Post WM_KEYDOWN/WM_KEYUP to window without hardware toggle)
        def enum_pam_hb(hwnd, lparam):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buf, 256)
            if buf.value == "SunAwtFrame":
                if user32.GetForegroundWindow() == hwnd:
                    logDebug(sys._getframe().f_lineno, f"HeartbeatProcess: (SymantecPAM Applet) PostMessage key to hwnd= {hwnd}")
                    user32.PostMessageW(hwnd, WM_KEYDOWN, VK_NUMLOCK, 0)
                    user32.PostMessageW(hwnd, WM_KEYUP, VK_NUMLOCK, 0)
            return True

        user32.EnumWindows(ENUMPROC(enum_pam_hb), 0)

    def hide_window(self, hwnd):
        logDebug(sys._getframe().f_lineno, f"HideWindow: hwnd= {hwnd}")
        self.click_active = False

        # Determine fullscreen state BEFORE restoring
        is_fs = isWindowFullScreen(hwnd)

        # Set transparency to 0
        style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style | WS_EX_LAYERED)
        user32.SetLayeredWindowAttributes(hwnd, 0, 0, LWA_ALPHA)

        # Restore window
        user32.ShowWindow(hwnd, SW_RESTORE)

        if is_fs:
            user32.PostMessageW(hwnd, WM_SYSCOMMAND, SC_RESTORE, 0)
            self.window_list[hwnd] = 1  # 1 - fullscreen
        else:
            self.window_list[hwnd] = 2  # 2 - windowed

        # Send to bottom
        user32.SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE)

        next_hwnd = self.next_window()
        logDebug(sys._getframe().f_lineno, f"HideWindow: nextHwnd= {next_hwnd}")

        # Restore alpha and activate next window
        next_style = user32.GetWindowLongW(next_hwnd, GWL_EXSTYLE)
        if next_style & WS_EX_LAYERED:
            user32.SetLayeredWindowAttributes(next_hwnd, 0, 255, LWA_ALPHA)
        user32.SetForegroundWindow(next_hwnd)
        self.click_active = True

    def activate_window(self, hwnd):
        self.click_active = False
        logDebug(sys._getframe().f_lineno, f"ActivateWindow: hwnd= {hwnd}")

        style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if style & WS_EX_LAYERED:
            user32.SetLayeredWindowAttributes(hwnd, 0, 255, LWA_ALPHA)

        user32.SetForegroundWindow(hwnd)
        if self.window_list.get(hwnd, 0) & 1:
            user32.PostMessageW(hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
        self.window_list[hwnd] = 0
        self.click_active = True

    def show_all_mstsc_window(self):
        self.click_active = False
        def enum_restore(hwnd, lparam):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buf, 256)
            if buf.value in ("TscShellContainerClass", "SunAwtFrame"):
                style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                if style & WS_EX_LAYERED:
                    user32.SetLayeredWindowAttributes(hwnd, 0, 255, LWA_ALPHA)
                user32.SetForegroundWindow(hwnd)
            return True

        ENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        user32.EnumWindows(ENUMPROC(enum_restore), 0)
        self.click_active = True

    def minimize_all_rdp(self):
        def enum_min(hwnd, lparam):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buf, 256)
            if buf.value in ("TscShellContainerClass", "SunAwtFrame"):
                user32.ShowWindow(hwnd, SW_MINIMIZE)
            return True

        ENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        user32.EnumWindows(ENUMPROC(enum_min), 0)
        self.program_unhide()

    def next_window(self):
        alt_tab_windows = []

        def enum_next(hwnd, lparam):
            if user32.IsWindowVisible(hwnd):
                owner = user32.GetWindow(hwnd, GW_OWNER)
                popup = user32.GetLastActivePopup(owner) if owner else hwnd
                if popup == hwnd:
                    es = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                    if not ((es & WS_EX_TOOLWINDOW) and not (es & WS_EX_APPWINDOW)) and not isInvisibleWin10BackgroundAppWindow(hwnd):
                        alt_tab_windows.append(hwnd)
            return True

        ENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        user32.EnumWindows(ENUMPROC(enum_next), 0)

        for hwnd in alt_tab_windows:
            state = self.window_list.get(hwnd, 0)
            if state != 1 and state != 2:
                return hwnd

        progman = user32.FindWindowW("Progman", "Program Manager")
        return progman if progman else self.root.winfo_id()

    def program_unhide(self):
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        if not self.stay_on_top:
            self.root.after(10, lambda: self.root.attributes("-topmost", False))
        self.gui_hidden = False

    def toggle_program_visibility(self):
        if self.gui_hidden:
            self.program_unhide()
        else:
            self.btn_close_event()

    def btn_close_event(self):
        self.click_active = False
        self.root.withdraw()
        self.gui_hidden = True
        next_hwnd = self.next_window()
        user32.SetForegroundWindow(next_hwnd)

    def on_double_click_session(self, event):
        item = self.tree.selection()
        if item:
            values = self.tree.item(item[0], "values")
            if values and len(values) >= 6:
                hwnd = int(values[5], 16)
                self.activate_window(hwnd)

    def on_right_click_session(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Show", command=lambda: self.on_double_click_session(None))
            menu.tk_popup(event.x_root, event.y_root)

    def toggle_stay_on_top(self):
        self.stay_on_top = 1 if self.stay_on_top_var.get() else 0
        self.root.attributes("-topmost", bool(self.stay_on_top))

    def toggle_start_minimized(self):
        self.start_minimized = 1 if self.start_minimized_var.get() else 0

    def open_preferences(self):
        dialog = tk.Toplevel(self.root)
        dialog.title(self.mnu_settings_preferences_txt.replace("&", ""))
        dialog.transient(self.root)
        dialog.grab_set()

        dialog.geometry(f"+{self.root.winfo_x() + 40}+{self.root.winfo_y() + 40}")

        ttk.Label(dialog, text=self.heartbeat_frequency_txt.replace("&", "")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        hb_entry = ttk.Entry(dialog, width=10)
        hb_entry.insert(0, str(self.heartbeat_frequency))
        hb_entry.grid(row=0, column=1, padx=10, pady=5)

        ttk.Label(dialog, text=self.refresh_frequency_txt.replace("&", "")).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ref_entry = ttk.Entry(dialog, width=10)
        ref_entry.insert(0, str(self.refresh_frequency))
        ref_entry.grid(row=1, column=1, padx=10, pady=5)

        top_var = tk.BooleanVar(value=bool(self.stay_on_top))
        min_var = tk.BooleanVar(value=bool(self.start_minimized))

        ttk.Checkbutton(dialog, text=self.stay_on_top_txt.replace("&", ""), variable=top_var).grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        ttk.Checkbutton(dialog, text=self.start_minimized_txt.replace("&", ""), variable=min_var).grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=10)

        def on_update():
            try:
                hb = int(hb_entry.get())
                ref = int(ref_entry.get())
                self.heartbeat_frequency = max(self.heartbeat_minimum, hb)
                self.refresh_frequency = max(self.refresh_minimum, ref)
                self.stay_on_top = 1 if top_var.get() else 0
                self.start_minimized = 1 if min_var.get() else 0
                self.stay_on_top_var.set(bool(self.stay_on_top))
                self.start_minimized_var.set(bool(self.start_minimized))
                self.root.attributes("-topmost", bool(self.stay_on_top))
                dialog.destroy()
            except ValueError:
                messagebox.showerror(gProgramTitle, "Please enter valid integer frequencies.")

        ttk.Button(btn_frame, text=self.btn_cancel_txt.replace("&", ""), command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text=self.btn_update_txt.replace("&", ""), command=on_update).pack(side=tk.RIGHT, padx=5)

    def show_quick_guide(self):
        messagebox.showinfo(f"{gProgramTitle} - Quick guide", self.quick_guide_txt)

    def show_about(self):
        msg = f"{gProgramTitle}\n-----------------------\n{self.about_txt}\n\nVersion {gVersion}\nCopyright (c) 2024-2026 PAM-Exchange"
        messagebox.showinfo(f"{gProgramTitle} - About", msg)

    def on_gui_close(self):
        self.btn_close_event()

    def on_exit(self):
        self.save_variables_to_property()
        if self.win_event_hook:
            user32.UnhookWinEvent(self.win_event_hook)
        if self.tray_nid:
            shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(self.tray_nid))
        user32.DeregisterShellHookWindow(self.root.winfo_id())
        user32.UnregisterHotKey(self.root.winfo_id(), 1)
        self.show_all_mstsc_window()
        self.root.destroy()
        sys.exit(0)

def main():
    logRoll(gLogFile, 5, 5)
    logInfo(sys._getframe().f_lineno, f"Start ----- gVersion= {gVersion}")
    root = tk.Tk()
    app = PAMRDPHeartbeatApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
