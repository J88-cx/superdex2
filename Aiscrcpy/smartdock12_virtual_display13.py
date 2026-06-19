import ctypes
import json
import os
import queue
import re
from enum import IntFlag
import subprocess
import sys
import threading
import time
from ctypes import wintypes

import psutil
import win32com.client
import win32gui
import win32process
from PyQt5 import QtCore, QtGui, QtWidgets

from gesture_engine import FixedGestureRecognizer, TouchStrokeCollector

for stream in (sys.stdout, sys.stderr):
    if hasattr(stream, "reconfigure"):
        stream.reconfigure(encoding="utf-8", errors="replace")

def get_bundled_path(*paths):
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *paths)

BUNDLED_SCRCPY_DIR = get_bundled_path("scrcpy")
os.environ["PATH"] = BUNDLED_SCRCPY_DIR + os.pathsep + os.environ.get("PATH", "")

ADB_PATH = os.path.join(BUNDLED_SCRCPY_DIR, "adb.exe")
SCRCPY_PATH = os.path.join(BUNDLED_SCRCPY_DIR, "scrcpy.exe")
APP_ICON_PATH = get_bundled_path("icon.ico")
SUBPROCESS_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)

SUPERDEX_PACKAGE = "com.xiaoxishu.superdex"
SUPERDEX_ACTIVITY = f"{SUPERDEX_PACKAGE}/.activities.LauncherActivity"
SUPERDEX_DOCK_SERVICE = f"{SUPERDEX_PACKAGE}/.services.DockService"
ACTION_SHOW_DOCK = "SHOW_DOCK"
STATE_FILE = "initial_state.json"
APP_SETTINGS_FILE = "superdex_settings.json"
STATE_LOCK = threading.Lock()
SETTINGS_LOCK = threading.Lock()
HID_KEYBOARD_MIN_SDK = 30
HID_IME_SETTING = "show_ime_with_hard_keyboard"
VIRTUAL_DISPLAY_CTRL_SHORTCUT_KEYS = {"a", "c", "v"}
SUPERDEX_INPUT_SETTLE_DELAY = 0.8
SCRCPY_SHORTCUT_MOD = "rctrl"
SCRCPY_MODE_VIRTUAL = "virtual"
SCRCPY_MODE_MIRROR = "mirror"
SCRCPY_MODE_DEFAULT = SCRCPY_MODE_VIRTUAL
USE_SCRCPY_VIRTUAL_DISPLAY = True
VIRTUAL_DISPLAY_TARGET_SHORT_DP = 960
VIRTUAL_DISPLAY_WAIT_TIMEOUT = 8.0
VIRTUAL_DISPLAY_WAIT_INTERVAL = 0.25
VIRTUAL_DISPLAY_SCREEN_CHECK_INTERVAL = 5
VIRTUAL_DISPLAY_DISABLE_SYSTEM_DECORATIONS = True
VIRTUAL_DISPLAY_DISABLE_SCREENSAVER = True
VIRTUAL_DISPLAY_START_APP_PACKAGE = SUPERDEX_PACKAGE
VIRTUAL_DISPLAY_FORCE_LANDSCAPE = True
VIRTUAL_DISPLAY_IME_POLICY = "local"
VIRTUAL_DISPLAY_WAKE_RECOVER_DELAY = 0.35
DEFAULT_SCREEN_OFF_TIMEOUT_MS = "30000"
SUPERDEX_FORCED_SCREEN_OFF_TIMEOUTS = {"2147483647", "86400000"}
KEYBOARD_MAPPING_MODE = "mapped"
KEYBOARD_LIKE_MODES = {KEYBOARD_MAPPING_MODE, "uhid"}
MAPPED_MODE_SCRCPY_KEYBOARD = "uhid"
GLOBAL_STATE_KEYS = [
    "policy_control",
    "display_cutout_force_fullscreen",
    "force_fullscreen",
    "desktop_mode",
    "desktop_mode_force_resizable",
    "stay_on_while_plugged_in",
]
SYSTEM_STATE_KEYS = [
    "user_rotation",
    "accelerometer_rotation",
]
SUPERDEX_FORCED_GLOBALS = {
    "policy_control": "immersive.full=*",
    "display_cutout_force_fullscreen": "1",
    "force_fullscreen": "1",
}
SUPERDEX_FORCED_SYSTEM = {
    "user_rotation": "1",
    "accelerometer_rotation": "0",
}
SUPERDEX_PRELAUNCH_CLEAN_RESTART = True
SUPERDEX_PRELAUNCH_CLEAN_DELAY = 0.35
RESTART_THIRD_PARTY_APPS_ON_RESTORE = True
# Apps that were already running before Superdex can keep stale density/bitmap
# caches after wm size/density is restored, so restart all running third-party
# packages on restore.
RESTART_ONLY_APPS_STARTED_DURING_SUPERDEX = False
THIRD_PARTY_RESTART_EXCLUDE_PACKAGES = {
    SUPERDEX_PACKAGE,
}

def sanitize_scrcpy_mode(mode):
    mode = str(mode or "").strip().lower()
    if mode in {SCRCPY_MODE_VIRTUAL, SCRCPY_MODE_MIRROR}:
        return mode
    return SCRCPY_MODE_DEFAULT


def load_app_settings():
    with SETTINGS_LOCK:
        if not os.path.exists(APP_SETTINGS_FILE):
            return {}
        try:
            with open(APP_SETTINGS_FILE, "r", encoding="utf-8") as f:
                settings = json.load(f)
        except Exception:
            return {}
    return settings if isinstance(settings, dict) else {}


def save_app_settings(settings):
    with SETTINGS_LOCK:
        tmp_file = APP_SETTINGS_FILE + ".tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
        os.replace(tmp_file, APP_SETTINGS_FILE)


def get_scrcpy_display_mode():
    return sanitize_scrcpy_mode(load_app_settings().get("scrcpy_display_mode"))


def use_scrcpy_virtual_display():
    return get_scrcpy_display_mode() == SCRCPY_MODE_VIRTUAL


def set_scrcpy_display_mode(mode):
    global USE_SCRCPY_VIRTUAL_DISPLAY
    mode = sanitize_scrcpy_mode(mode)
    settings = load_app_settings()
    settings["scrcpy_display_mode"] = mode
    save_app_settings(settings)
    USE_SCRCPY_VIRTUAL_DISPLAY = mode == SCRCPY_MODE_VIRTUAL
    return mode


USE_SCRCPY_VIRTUAL_DISPLAY = use_scrcpy_virtual_display()


ANDROID_KEY_MAP = {
    "space": "KEYCODE_SPACE",
    "enter": "KEYCODE_ENTER",
    "backspace": "KEYCODE_DEL",
    "delete": "KEYCODE_FORWARD_DEL",
    "tab": "KEYCODE_TAB",
    "left": "KEYCODE_DPAD_LEFT",
    "right": "KEYCODE_DPAD_RIGHT",
    "up": "KEYCODE_DPAD_UP",
    "down": "KEYCODE_DPAD_DOWN",
    "home": "KEYCODE_MOVE_HOME",
    "end": "KEYCODE_MOVE_END",
    "page up": "KEYCODE_PAGE_UP",
    "page down": "KEYCODE_PAGE_DOWN",
    ",": "KEYCODE_COMMA",
    ".": "KEYCODE_PERIOD",
    "-": "KEYCODE_MINUS",
    "=": "KEYCODE_EQUALS",
    "[": "KEYCODE_LEFT_BRACKET",
    "]": "KEYCODE_RIGHT_BRACKET",
    "\\": "KEYCODE_BACKSLASH",
    ";": "KEYCODE_SEMICOLON",
    "'": "KEYCODE_APOSTROPHE",
    "/": "KEYCODE_SLASH",
    "`": "KEYCODE_GRAVE",
}

SHIFTED_KEY_MAP = {
    "1": "KEYCODE_1",
    "2": "KEYCODE_2",
    "3": "KEYCODE_3",
    "4": "KEYCODE_4",
    "5": "KEYCODE_5",
    "6": "KEYCODE_6",
    "7": "KEYCODE_7",
    "8": "KEYCODE_8",
    "9": "KEYCODE_9",
    "0": "KEYCODE_0",
    "-": "KEYCODE_MINUS",
    "=": "KEYCODE_EQUALS",
    "[": "KEYCODE_LEFT_BRACKET",
    "]": "KEYCODE_RIGHT_BRACKET",
    "\\": "KEYCODE_BACKSLASH",
    ";": "KEYCODE_SEMICOLON",
    "'": "KEYCODE_APOSTROPHE",
    ",": "KEYCODE_COMMA",
    ".": "KEYCODE_PERIOD",
    "/": "KEYCODE_SLASH",
    "`": "KEYCODE_GRAVE",
}

SHIFTED_SYMBOL_KEY_MAP = {
    "!": "KEYCODE_1",
    "@": "KEYCODE_2",
    "#": "KEYCODE_3",
    "$": "KEYCODE_4",
    "%": "KEYCODE_5",
    "^": "KEYCODE_6",
    "&": "KEYCODE_7",
    "*": "KEYCODE_8",
    "(": "KEYCODE_9",
    ")": "KEYCODE_0",
    "_": "KEYCODE_MINUS",
    "+": "KEYCODE_EQUALS",
    "{": "KEYCODE_LEFT_BRACKET",
    "}": "KEYCODE_RIGHT_BRACKET",
    "|": "KEYCODE_BACKSLASH",
    ":": "KEYCODE_SEMICOLON",
    '"': "KEYCODE_APOSTROPHE",
    "<": "KEYCODE_COMMA",
    ">": "KEYCODE_PERIOD",
    "?": "KEYCODE_SLASH",
    "~": "KEYCODE_GRAVE",
}

KEY_NAME_ALIASES = {
    "question mark": "?",
    "slash": "/",
    "forward slash": "/",
    "backslash": "\\",
    "back slash": "\\",
    "comma": ",",
    "oem comma": ",",
    "period": ".",
    "oem period": ".",
    "dot": ".",
    "minus": "-",
    "dash": "-",
    "equals": "=",
    "equal": "=",
    "semicolon": ";",
    "apostrophe": "'",
    "quote": "'",
    "grave": "`",
    "backtick": "`",
    "left bracket": "[",
    "right bracket": "]",
    "less": "<",
    "less than": "<",
    "left angle bracket": "<",
    "greater": ">",
    "greater than": ">",
    "right angle bracket": ">",
}

CTRL_COMBO_KEYS = {
    "a": "KEYCODE_A",
    "c": "KEYCODE_C",
    "v": "KEYCODE_V",
    "x": "KEYCODE_X",
    "z": "KEYCODE_Z",
    "y": "KEYCODE_Y",
    "f": "KEYCODE_F",
    "s": "KEYCODE_S",
    "left": "KEYCODE_DPAD_LEFT",
    "right": "KEYCODE_DPAD_RIGHT",
    "up": "KEYCODE_DPAD_UP",
    "down": "KEYCODE_DPAD_DOWN",
}

SHIFT_COMBO_KEYS = {
    "tab",
    "left",
    "right",
    "up",
    "down",
    "home",
    "end",
    "page up",
    "page down",
    "delete",
}

MODIFIER_NAME_MAP = {
    "shift": "shift",
    "left shift": "shift",
    "right shift": "shift",
    "ctrl": "ctrl",
    "left ctrl": "ctrl",
    "right ctrl": "ctrl",
    "alt": "alt",
    "left alt": "alt",
    "right alt": "alt",
    "alt gr": "alt",
    "windows": "windows",
    "left windows": "windows",
    "right windows": "windows",
}

COMMON_LAUNCHERS = [
    "com.android.launcher/.Launcher",
    "com.google.android.apps.nexuslauncher/.NexusLauncherActivity",
    "com.miui.home/.launcher.Launcher",
    "com.oppo.launcher/.Launcher",
    "com.huawei.android.launcher/.Launcher",
    "com.zui.launcher/.drawer.NormalLauncher",
    "com.sec.android.app.launcher/.activities.LauncherActivity",
]


class InputDevices(IntFlag):
    NONE = 0
    TOUCH_SCREEN = 1 << 0
    TOUCH_PAD = 1 << 1
    MOUSE = 1 << 2
    PEN = 1 << 3
    TOUCH_DEVICE = TOUCH_SCREEN | TOUCH_PAD

WH_MOUSE_LL = 14
HC_ACTION = 0
WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_MBUTTONDOWN = 0x0207
WM_XBUTTONDOWN = 0x020B
WM_TOUCH = 0x0240
WM_POINTERUPDATE = 0x0245
WM_POINTERDOWN = 0x0246
WM_POINTERUP = 0x0247
TOUCH_COORD_TO_PIXEL = 100
TOUCHEVENTF_MOVE = 0x0001
TOUCHEVENTF_DOWN = 0x0002
TOUCHEVENTF_UP = 0x0004
PT_TOUCHPAD = 5
XBUTTON1 = 0x0001
XBUTTON2 = 0x0002
VIRTUAL_EDGE_SIDE_HOT_RATIO = 0.08
VIRTUAL_EDGE_BOTTOM_HOT_RATIO = 0.14
VIRTUAL_EDGE_SIDE_HOT_MIN = 48
VIRTUAL_EDGE_SIDE_HOT_MAX = 160
VIRTUAL_EDGE_BOTTOM_HOT_MIN = 64
VIRTUAL_EDGE_BOTTOM_HOT_MAX = 220
VIRTUAL_EDGE_SWIPE_RATIO = 0.045
VIRTUAL_EDGE_SWIPE_MIN = 24
VIRTUAL_EDGE_SWIPE_MAX = 96
VIRTUAL_EDGE_HORIZONTAL_RATIO = 0.06
VIRTUAL_EDGE_HOLD_DELAY = 0.35
VIRTUAL_EDGE_ENABLE_LEFT_BUTTON_SIDE_GESTURES = False
SW_SHOW = 5
SW_RESTORE = 9
HWND_TOPMOST = -1
HWND_NOTOPMOST = -2
ASFW_ANY = -1
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_SHOWWINDOW = 0x0040
VK_D = 0x44
VK_O = 0x4F
VK_TAB = 0x09
VK_SHIFT = 0x10
VK_MENU = 0x12
VK_LWIN = 0x5B
VK_RCONTROL = 0xA3
VK_MEDIA_PLAY_PAUSE = 0xB3
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP = 0xAF
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP = 0x0040
MOUSEEVENTF_XDOWN = 0x0080
MOUSEEVENTF_XUP = 0x0100
EXTENDED_HOST_KEYS = {VK_LWIN}
VOLUME_STEP_PERCENT = 2
SYNTHETIC_INPUT_GUARD_SECONDS = 0.45
HOST_INPUT_STEP_DELAY = 0.008
HOST_HOTKEY_HOLD_DELAY = 0.024
HOST_MOUSE_CLICK_DELAY = 0.012
HOST_MOUSE_REPEAT_DELAY = 0.024
HOST_WINDOW_ACTIVATE_CONFIRM_SECONDS = 0.55
HOST_WINDOW_ACTIVATE_RETRY_DELAY = 0.045
HOST_WINDOW_ACTIVATE_STABLE_SECONDS = 0.12
HOST_WINDOW_ACTIVATE_POST_RETRY_DELAYS = (0.08, 0.2, 0.45, 0.9, 1.5)
RAW_FINGER_COUNT_SETTLE_SECONDS = 0.035
RAW_FINGER_COUNT_SETTLE_FRAMES = 2
RAW_LIVE_ACTION_RELEASE_TIMEOUT_SECONDS = 0.75
GESTURE_DEBUG = os.environ.get("SUPERDEX_GESTURE_DEBUG") == "1"
GESTURE_BUILD_TAG = "gesture-global-raw-20260613-02"
BLACK_WINDOW_MIN_SAMPLES = 16
BLACK_WINDOW_BRIGHTNESS_LIMIT = 10
BLACK_WINDOW_RATIO = 0.92
SCREEN_OFF_GUARD_ITERATIONS = 18
SCREEN_OFF_GUARD_INTERVAL = "0.04"
SCREEN_OFF_FALLBACK_DELAY = 0.08
SCREEN_ON_STABILIZE_SECONDS = 3.0
SCREEN_ON_STABILIZE_INTERVAL = 0.35

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class INPUT_UNION(ctypes.Union):
    _fields_ = [
        ("mi", MOUSEINPUT),
        ("ki", KEYBDINPUT),
        ("hi", HARDWAREINPUT),
    ]


class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [
        ("type", wintypes.DWORD),
        ("u", INPUT_UNION),
    ]


class POINTER_INFO(ctypes.Structure):
    _fields_ = [
        ("pointerType", ctypes.c_int),
        ("pointerId", ctypes.c_uint32),
        ("frameId", ctypes.c_uint32),
        ("pointerFlags", wintypes.DWORD),
        ("sourceDevice", wintypes.HANDLE),
        ("hwndTarget", wintypes.HWND),
        ("ptPixelLocation", wintypes.POINT),
        ("ptHimetricLocation", wintypes.POINT),
        ("ptPixelLocationRaw", wintypes.POINT),
        ("ptHimetricLocationRaw", wintypes.POINT),
        ("dwTime", wintypes.DWORD),
        ("historyCount", ctypes.c_uint32),
        ("InputData", ctypes.c_int32),
        ("dwKeyStates", wintypes.DWORD),
        ("PerformanceCount", ctypes.c_uint64),
        ("ButtonChangeType", ctypes.c_int),
    ]


class TOUCHINPUT(ctypes.Structure):
    _fields_ = [
        ("x", wintypes.LONG),
        ("y", wintypes.LONG),
        ("hSource", wintypes.HANDLE),
        ("dwID", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("dwMask", wintypes.DWORD),
        ("dwTime", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
        ("cxContact", wintypes.DWORD),
        ("cyContact", wintypes.DWORD),
    ]


class RAWINPUTDEVICE(ctypes.Structure):
    _fields_ = [
        ("usUsagePage", wintypes.USHORT),
        ("usUsage", wintypes.USHORT),
        ("dwFlags", wintypes.DWORD),
        ("hwndTarget", wintypes.HWND),
    ]


class RAWINPUTHEADER(ctypes.Structure):
    _fields_ = [
        ("dwType", wintypes.DWORD),
        ("dwSize", wintypes.DWORD),
        ("hDevice", wintypes.HANDLE),
        ("wParam", wintypes.WPARAM),
    ]


class RAWHID(ctypes.Structure):
    _fields_ = [
        ("dwSizeHid", wintypes.DWORD),
        ("dwCount", wintypes.DWORD),
    ]


class RID_DEVICE_INFO_MOUSE(ctypes.Structure):
    _fields_ = [
        ("dwId", wintypes.DWORD),
        ("dwNumberOfButtons", wintypes.DWORD),
        ("dwSampleRate", wintypes.DWORD),
        ("fHasHorizontalWheel", wintypes.BOOL),
    ]


class RID_DEVICE_INFO_KEYBOARD(ctypes.Structure):
    _fields_ = [
        ("dwType", wintypes.DWORD),
        ("dwSubType", wintypes.DWORD),
        ("dwKeyboardMode", wintypes.DWORD),
        ("dwNumberOfFunctionKeys", wintypes.DWORD),
        ("dwNumberOfIndicators", wintypes.DWORD),
        ("dwNumberOfKeysTotal", wintypes.DWORD),
    ]


class RID_DEVICE_INFO_HID(ctypes.Structure):
    _fields_ = [
        ("dwVendorId", wintypes.DWORD),
        ("dwProductId", wintypes.DWORD),
        ("dwVersionNumber", wintypes.DWORD),
        ("usUsagePage", wintypes.USHORT),
        ("usUsage", wintypes.USHORT),
    ]


class RID_DEVICE_INFO_UNION(ctypes.Union):
    _fields_ = [
        ("mouse", RID_DEVICE_INFO_MOUSE),
        ("keyboard", RID_DEVICE_INFO_KEYBOARD),
        ("hid", RID_DEVICE_INFO_HID),
    ]


class RID_DEVICE_INFO(ctypes.Structure):
    _anonymous_ = ("_u",)
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("dwType", wintypes.DWORD),
        ("_u", RID_DEVICE_INFO_UNION),
    ]


class HIDP_LINK_COLLECTION_NODE(ctypes.Structure):
    _fields_ = [
        ("LinkUsage", ctypes.c_short),
        ("LinkUsagePage", ctypes.c_short),
        ("Parent", ctypes.c_short),
        ("NumberOfChildren", ctypes.c_short),
        ("NextSibling", ctypes.c_short),
        ("FirstChild", ctypes.c_short),
        ("CollectionType", ctypes.c_byte),
        ("_reserved0", ctypes.c_byte),
        ("_reserved1", ctypes.c_byte),
        ("_reserved2", ctypes.c_byte),
        ("UserContext", wintypes.HANDLE),
    ]


RAWINPUTHEADER_SIZE = ctypes.sizeof(RAWINPUTHEADER)
RAWHID_SIZE = ctypes.sizeof(RAWHID)
RAWINPUT_BUFFER_HEADER_SIZE = RAWINPUTHEADER_SIZE + RAWHID_SIZE

RIDEV_INPUTSINK = 0x00000100
RIDEV_EXINPUTSINK = 0x00001000
RIDEV_DEVNOTIFY = 0x00002000
RIDEV_PAGEONLY = 0x00000020
RID_INPUT = 0x10000003
RIM_TYPEHID = 2
RIDI_DEVICENAME = 0x20000007
RIDI_DEVICEINFO = 0x2000000B
RIDI_PREPARSEDDATA = 0x20000005
WM_INPUT = 0x00FF
WM_INPUT_DEVICE_CHANGE = 0x00FE
DIGITIZER_USAGE_PAGE = 0x0D
TOUCH_PAD_USAGE = 0x05
TOUCH_SCREEN_USAGE = 0x04
PEN_USAGE = 0x02
GENERIC_DESKTOP_PAGE = 0x01
HIDP_STATUS_SUCCESS = (0x0 << 28) | (0x11 << 16) | 0
HIDP_REPORT_TYPE_INPUT = 0
CONTACT_IDENTIFIER_ID = 0x51
CONTACT_COUNT_ID = 0x54
FINGER_ID = 0x22
TIP_ID = 0x42
X_COORDINATE_ID = 0x30
Y_COORDINATE_ID = 0x31
IN_RANGE_ID = 0x32
BARREL_BUTTON_ID = 0x44
INVERT_ID = 0x3C
ERASER_ID = 0x45

def touch_state_from_flags(flags):
    if flags & TOUCHEVENTF_DOWN:
        return "down"
    if flags & TOUCHEVENTF_UP:
        return "up"
    if flags & TOUCHEVENTF_MOVE:
        return "move"
    return "move"


def touch_point_to_tuple(x, y):
    return (x / TOUCH_COORD_TO_PIXEL, y / TOUCH_COORD_TO_PIXEL)


class HID_RAW_INPUT_HELPERS:
    def __init__(self):
        self.user32 = ctypes.WinDLL("user32", use_last_error=True)
        self.hid = ctypes.WinDLL("hid", use_last_error=True)

        self.user32.RegisterRawInputDevices.argtypes = [
            ctypes.POINTER(RAWINPUTDEVICE),
            wintypes.UINT,
            wintypes.UINT,
        ]
        self.user32.RegisterRawInputDevices.restype = wintypes.BOOL
        self.user32.GetRawInputData.argtypes = [
            wintypes.HANDLE,
            wintypes.UINT,
            ctypes.c_void_p,
            ctypes.POINTER(wintypes.UINT),
            wintypes.UINT,
        ]
        self.user32.GetRawInputData.restype = wintypes.UINT
        self.user32.GetRawInputDeviceInfoW.argtypes = [
            wintypes.HANDLE,
            wintypes.UINT,
            ctypes.c_void_p,
            ctypes.POINTER(wintypes.UINT),
        ]
        self.user32.GetRawInputDeviceInfoW.restype = wintypes.UINT
        self.hid.HidP_GetLinkCollectionNodes.argtypes = [
            ctypes.POINTER(HIDP_LINK_COLLECTION_NODE),
            ctypes.POINTER(ctypes.c_int),
            ctypes.c_void_p,
        ]
        self.hid.HidP_GetLinkCollectionNodes.restype = ctypes.c_int
        self.hid.HidP_GetUsages.argtypes = [
            ctypes.c_int,
            ctypes.c_ushort,
            ctypes.c_short,
            ctypes.POINTER(ctypes.c_ushort),
            ctypes.POINTER(ctypes.c_int),
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_int,
        ]
        self.hid.HidP_GetUsages.restype = ctypes.c_int
        self.hid.HidP_MaxUsageListLength.argtypes = [
            ctypes.c_int,
            ctypes.c_ushort,
            ctypes.c_void_p,
        ]
        self.hid.HidP_MaxUsageListLength.restype = ctypes.c_ulong
        self.hid.HidP_GetUsageValue.argtypes = [
            ctypes.c_int,
            ctypes.c_ushort,
            ctypes.c_short,
            ctypes.c_ushort,
            ctypes.POINTER(ctypes.c_int),
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_int,
        ]
        self.hid.HidP_GetUsageValue.restype = ctypes.c_int
        self.hid.HidP_GetScaledUsageValue.argtypes = [
            ctypes.c_int,
            ctypes.c_ushort,
            ctypes.c_short,
            ctypes.c_ushort,
            ctypes.POINTER(ctypes.c_int),
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_int,
        ]
        self.hid.HidP_GetScaledUsageValue.restype = ctypes.c_int


RAW_HELPERS = None


def get_raw_helpers():
    global RAW_HELPERS
    if RAW_HELPERS is None:
        RAW_HELPERS = HID_RAW_INPUT_HELPERS()
    return RAW_HELPERS


def run_host(command):
    result = subprocess.run(
        ["powershell", "-Command", command],
        text=True,
        capture_output=True,
        creationflags=SUBPROCESS_FLAGS,
        encoding="utf-8",
        errors="ignore",
    )
    return (result.stdout or "").strip()


def adb(args, capture_output=True, check=False, timeout=None):
    kwargs = {
        "text": True,
        "check": check,
        "timeout": timeout,
        "creationflags": SUBPROCESS_FLAGS,
        "encoding": "utf-8",
        "errors": "ignore",
    }
    if capture_output:
        kwargs["capture_output"] = True
    else:
        kwargs["stdout"] = subprocess.DEVNULL
        kwargs["stderr"] = subprocess.DEVNULL

    command = [ADB_PATH] + args
    try:
        result = subprocess.run(command, **kwargs)
    except subprocess.TimeoutExpired:
        print(f"⚠️ adb 命令超时，已跳过：{' '.join(map(str, args))}")
        return ""
    if not capture_output:
        return ""
    return (result.stdout or "").strip()


def is_process_running(proc):
    return proc is not None and proc.poll() is None


def stop_process(proc, timeout=3):
    if not is_process_running(proc):
        return
    try:
        proc.terminate()
        proc.wait(timeout=timeout)
    except Exception:
        try:
            proc.kill()
            proc.wait(timeout=timeout)
        except Exception:
            pass


def send_home_key(device_id):
    adb(["-s", device_id, "shell", "input", "keyevent", "KEYCODE_HOME"], capture_output=False)


def start_scrcpy_screen_off_helper(device_id):
    args = [
        SCRCPY_PATH,
        "-s",
        device_id,
        "--no-window",
        "--no-audio",
        "--display-id=0",
        "--turn-screen-off",
        "--keep-active",
        "--time-limit=1",
        "--no-cleanup",
    ]
    try:
        subprocess.Popen(
            args,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=SUBPROCESS_FLAGS,
        )
    except Exception:
        pass


def start_scrcpy_screen_on_helper(device_id):
    # scrcpy powers the device screen on by default when it starts unless
    # --no-power-on is set. Use a short hidden helper to ask the scrcpy server
    # to restore the physical screen power mode instead of relying only on adb
    # keyevents, which do not undo scrcpy's own --turn-screen-off state on some
    # devices.
    # To force the server to initialize the display stream (which actually turns
    # the screen on), we must not disable video completely. Instead, we record
    # to the OS null device.
    null_device = "NUL" if os.name == "nt" else "/dev/null"
    args = [
        SCRCPY_PATH,
        "-s",
        device_id,
        "--no-window",
        "--no-audio",
        "--display-id=0",
        f"--record={null_device}",
        "--record-format=mp4",
        "--time-limit=2",
        "--no-cleanup",
    ]
    try:
        subprocess.Popen(
            args,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=SUBPROCESS_FLAGS,
        )
    except Exception:
        pass


def start_physical_screen_off_guard(device_id):
    guard_script = (
        "i=0; "
        f"while [ $i -lt {SCREEN_OFF_GUARD_ITERATIONS} ]; do "
        "cmd display power-off 0 >/dev/null 2>&1; "
        f"sleep {SCREEN_OFF_GUARD_INTERVAL} >/dev/null 2>&1; "
        "i=$((i+1)); "
        "done"
    )
    adb(
        ["-s", device_id, "shell", "sh", "-c", f"({guard_script}) &"],
        capture_output=False,
    )


def send_home_with_screen_off_guard(device_id):
    guard_script = (
        "i=0; "
        f"while [ $i -lt {SCREEN_OFF_GUARD_ITERATIONS} ]; do "
        "cmd display power-off 0 >/dev/null 2>&1; "
        f"sleep {SCREEN_OFF_GUARD_INTERVAL} >/dev/null 2>&1; "
        "i=$((i+1)); "
        "done"
    )
    adb(
        [
            "-s",
            device_id,
            "shell",
            "sh",
            "-c",
            f"({guard_script}) & input keyevent KEYCODE_HOME",
        ],
        capture_output=False,
    )


def force_physical_screen_off(device_id, use_scrcpy_helper=True):
    adb(
        ["-s", device_id, "shell", "cmd", "display", "power-off", "0"],
        capture_output=False,
        timeout=2,
    )
    if use_scrcpy_helper:
        start_scrcpy_screen_off_helper(device_id)


def physical_screen_policy_reports_off(display_output, window_policy):
    text = f"{display_output or ''}\n{window_policy or ''}"
    if re.search(
        r"(?is)DisplayViewport\{type=INTERNAL\b[^}]*\bisActive=false\b",
        text,
    ):
        return True
    if re.search(r"(?i)\bscreenState\s*=\s*SCREEN_STATE_OFF\b", text):
        return True
    if re.search(r"(?i)mScreenOn(?:Fully|Early)?=false", text):
        return True
    return False


def is_physical_screen_on(device_id):
    output = adb(["-s", device_id, "shell", "dumpsys", "display"], timeout=3)
    normalized = output or ""
    window_policy = None

    # Physical screen check: look at DisplayDeviceInfo blocks with the address
    # field (type=LOCAL or address={port=…}) which always describe the physical
    # hardware panel. In virtual-display mode scrcpy creates a logical display
    # whose state can be ON even though the real built-in screen backlight is
    # OFF, so we must never match logical Display 0 blocks that lack an address
    # field or a type=INTERNAL marker.
    PHYSICAL_DISPLAY_LINE_RE = re.compile(
        r"(?is)"
        r"DisplayDeviceInfo\{"
        r"(?:[^}]*?"
        r"(?:type\s*=\s*INTERNAL|"
        r"address\s*=\s*\{port=\d+\}|"
        r"内置屏幕|built-in\s+screen)"
        r")"
    )

    lines = normalized.splitlines()
    for index, line in enumerate(lines):
        if not PHYSICAL_DISPLAY_LINE_RE.search(line):
            continue
        block = "\n".join(lines[index : index + 10])
        if re.search(r"(?i)\b(?:state|mState)\s*[=: ]\s*(ON|ON_SUSPEND|VR)\b", block):
            window_policy = adb(
                ["-s", device_id, "shell", "dumpsys", "window", "policy"],
                timeout=3,
            )
            if physical_screen_policy_reports_off(normalized, window_policy):
                return False
            return True
        if re.search(r"(?i)\b(?:state|mState)\s*[=: ]\s*(OFF|DOZE|DOZE_SUSPEND)\b", block):
            return False

    # Fallback: search the whole dumpsys output for any physical-display block
    display_zero_patterns = (
        r"(?is)DisplayDeviceInfo\{[^}]*?\bbuilt-in\s+screen\b[^}]*?\bstate\s+([A-Z_]+)",
        r"(?is)DisplayDeviceInfo\{[^}]*?(?:内置屏幕|built-in\s+screen)[^}]*?\bmState\s*=\s*([A-Z_]+)",
        r"(?is)DisplayDeviceInfo\{[^}]*?(?:内置屏幕|built-in\s+screen)[^}]*?\bstate\s+([A-Z_]+)",
        r"(?is)DisplayDeviceInfo\{[^}]*?\bFLAG_SECURE\b[^}]*?\bstate\s+([A-Z_]+)[^}]*?\baddress\s+{port=\d+}",
        r"(?is)DisplayDeviceInfo\{[^}]*?\bFLAG_SECURE\b[^}]*?\bmState\s*=\s*([A-Z_]+)[^}]*?\baddress\s+{port=\d+}",
    )
    for pattern in display_zero_patterns:
        match = re.search(pattern, normalized)
        if not match:
            continue
        state = match.group(1).upper()
        if state in {"ON", "ON_SUSPEND", "VR"}:
            window_policy = adb(
                ["-s", device_id, "shell", "dumpsys", "window", "policy"],
                timeout=3,
            )
            if physical_screen_policy_reports_off(normalized, window_policy):
                return False
            return True
        if state in {"OFF", "DOZE", "DOZE_SUSPEND"}:
            return False

    if window_policy is None:
        window_policy = adb(["-s", device_id, "shell", "dumpsys", "window", "policy"], timeout=3)
    if physical_screen_policy_reports_off(normalized, window_policy):
        return False

    # Do not treat WindowPolicy/PowerManager "screen on" or "awake" as physical
    # panel proof here. In virtual-display mode Android can be interactive while
    # the built-in display backlight remains off. Only an explicit display 0 /
    # built-in-display ON state is accepted as success.
    return None


def start_physical_screen_on_stabilizer(device_id, duration=SCREEN_ON_STABILIZE_SECONDS):
    # Some devices briefly light the built-in panel and then turn it off again
    # while a virtual display remains active. Keep reasserting physical display 0
    # power-on for a short stabilization window so stale scrcpy/screen-off state
    # or an overlapping screen-off guard cannot immediately win after the button
    # reports success.
    iterations = max(1, int(duration / SCREEN_ON_STABILIZE_INTERVAL))
    stabilize_script = (
        "settings put system screen_off_timeout 2147483647 >/dev/null 2>&1; "
        "settings put global stay_on_while_plugged_in 7 >/dev/null 2>&1; "
        "svc power stayon true >/dev/null 2>&1; "
        f"i=0; while [ $i -lt {iterations} ]; do "
        "cmd display power-on 0 >/dev/null 2>&1; "
        "cmd power wakeup >/dev/null 2>&1; "
        "input keyevent KEYCODE_WAKEUP >/dev/null 2>&1; "
        "input keyevent 224 >/dev/null 2>&1; "
        f"sleep {SCREEN_ON_STABILIZE_INTERVAL} >/dev/null 2>&1; "
        "i=$((i+1)); "
        "done"
    )
    adb(
        ["-s", device_id, "shell", "sh", "-c", f"({stabilize_script}) &"],
        capture_output=False,
        timeout=2,
    )


def get_physical_screen_state_debug(device_id):
    display = adb(["-s", device_id, "shell", "dumpsys", "display"], timeout=3)
    policy = adb(["-s", device_id, "shell", "dumpsys", "window", "policy"], timeout=3)
    display_states = []
    for line in (display or "").splitlines():
        if re.search(r"(?i)(displayid\s*=\s*0|display\s+0|built-in|mState=| state )", line):
            stripped = line.strip()
            if stripped:
                display_states.append(stripped[:180])
        if len(display_states) >= 8:
            break
    policy_states = []
    for line in (policy or "").splitlines():
        if re.search(r"(?i)(mScreenOn|mAwake|mWakefulness|screen)", line):
            stripped = line.strip()
            if stripped:
                policy_states.append(stripped[:180])
        if len(policy_states) >= 6:
            break
    return "; ".join(display_states + policy_states) or "未能读取到物理屏状态"


def run_physical_screen_wake_sequence(device_id, include_power_key=False):
    # scrcpy's screen-off feature uses Android display power mode. On Android 15+
    # `cmd display power-on 0` is the documented shell path, while older builds
    # often need PowerManager.wakeUp (`cmd power wakeup`) or injected wake keys.
    # Keep all paths idempotent and only include KEYCODE_POWER as a later fallback
    # because POWER toggles when the device is already on.
    power_key_command = "input keyevent KEYCODE_POWER >/dev/null 2>&1; " if include_power_key else ""
    wake_script = (
        "cmd display power-on 0 >/dev/null 2>&1; "
        "cmd power wakeup >/dev/null 2>&1; "
        "input keyevent KEYCODE_WAKEUP >/dev/null 2>&1; "
        "input keyevent 224 >/dev/null 2>&1; "
        "input keyevent WAKEUP >/dev/null 2>&1; "
        f"{power_key_command}"
        "wm dismiss-keyguard >/dev/null 2>&1"
    )
    adb(
        ["-s", device_id, "shell", "sh", "-c", wake_script],
        capture_output=False,
        timeout=3,
    )


def run_physical_screen_power_button_wake(device_id):
    # When scrcpy/Android has left the physical display in a logically ON but
    # visually black state, normal WAKEUP/display-power commands may be ignored.
    # Reset physical display 0 to a known OFF state first, then inject Android's
    # POWER key over adb (not the phone hardware button) to perform the same wake
    # path the user sees when manually pressing power.
    wake_script = (
        "cmd display power-off 0 >/dev/null 2>&1; "
        "sleep 0.15; "
        "input keyevent KEYCODE_POWER >/dev/null 2>&1; "
        "sleep 0.25; "
        "cmd power wakeup >/dev/null 2>&1; "
        "cmd display power-on 0 >/dev/null 2>&1; "
        "input keyevent KEYCODE_WAKEUP >/dev/null 2>&1; "
        "input keyevent 224 >/dev/null 2>&1; "
        "wm dismiss-keyguard >/dev/null 2>&1"
    )
    adb(
        ["-s", device_id, "shell", "sh", "-c", wake_script],
        capture_output=False,
        timeout=4,
    )


def wait_for_stable_physical_screen_on(device_id, duration=SCREEN_ON_STABILIZE_SECONDS):
    started = time.time()
    deadline = time.time() + max(0.1, duration)
    saw_on = False
    while time.time() < deadline:
        state = is_physical_screen_on(device_id)
        if state is False and time.time() - started >= SCREEN_ON_STABILIZE_INTERVAL * 2:
            return False
        if state is True:
            saw_on = True
        time.sleep(SCREEN_ON_STABILIZE_INTERVAL)
    return saw_on and is_physical_screen_on(device_id) is True


def turn_physical_screen_on(device_id):
    # If the screen is logically ON (but physically OFF due to scrcpy --turn-screen-off),
    # force it logically OFF first using KEYCODE_SLEEP (223). This guarantees that
    # the subsequent wake sequence (using WAKEUP keys) triggers a real transition from
    # OFF to ON, forcing the hardware display panel to power back on.
    if is_physical_screen_on(device_id) is True:
        adb(["-s", device_id, "shell", "input", "keyevent", "223"], timeout=2)
        time.sleep(0.25)

    enable_never_sleep(device_id)
    start_scrcpy_screen_on_helper(device_id)
    start_physical_screen_on_stabilizer(device_id)
    time.sleep(0.35)

    for _ in range(3):
        run_physical_screen_wake_sequence(device_id, include_power_key=False)
        start_physical_screen_on_stabilizer(device_id, duration=1.5)
        time.sleep(0.18)
        if wait_for_stable_physical_screen_on(device_id, duration=1.2):
            return True

    # If safe wake/display-power commands did not prove that the built-in panel
    # is ON, send Android POWER as the final fallback. This previously did not
    # run when logical display/window states were misread as success, so the real
    # phone panel stayed dark even though the UI reported success.
    for _ in range(2):
        run_physical_screen_power_button_wake(device_id)
        start_physical_screen_on_stabilizer(device_id, duration=1.5)
        time.sleep(0.3)
        if wait_for_stable_physical_screen_on(device_id, duration=1.2):
            return True

    start_scrcpy_screen_on_helper(device_id)
    start_physical_screen_on_stabilizer(device_id)
    time.sleep(0.6)
    if wait_for_stable_physical_screen_on(device_id):
        return True

    run_physical_screen_power_button_wake(device_id)
    start_physical_screen_on_stabilizer(device_id)
    time.sleep(0.3)
    return wait_for_stable_physical_screen_on(device_id)


def turn_physical_screen_on_force(device_id):
    """Brute-force physical screen wake by cycling through off→POWER→on.

    In virtual-display mode scrcpy's --turn-screen-off leaves the physical
    panel in a logically-ON-but-visually-dark state that normal wake commands
    cannot escape.  This helper deliberately cycles the hardware power state
    via Android's POWER key, which is the same path the user takes when
    manually pressing the phone's power button.
    """
    enable_never_sleep(device_id)
    start_scrcpy_screen_on_helper(device_id)

    # Cycle 1: power-off → POWER → wake
    adb(
        ["-s", device_id, "shell", "cmd", "display", "power-off", "0"],
        capture_output=False,
        timeout=2,
    )
    time.sleep(0.3)
    adb(
        ["-s", device_id, "shell", "input", "keyevent", "KEYCODE_POWER"],
        capture_output=False,
        timeout=2,
    )
    time.sleep(0.5)
    start_physical_screen_on_stabilizer(device_id, duration=2.0)
    if wait_for_stable_physical_screen_on(device_id, duration=2.0):
        return True

    # Cycle 2: run the full power-button-wake sequence
    run_physical_screen_power_button_wake(device_id)
    start_physical_screen_on_stabilizer(device_id, duration=2.0)
    time.sleep(0.5)
    if wait_for_stable_physical_screen_on(device_id, duration=2.0):
        return True

    # Cycle 3: last-resort — start another scrcpy helper, then POWER again
    start_scrcpy_screen_on_helper(device_id)
    time.sleep(0.5)
    run_physical_screen_power_button_wake(device_id)
    start_physical_screen_on_stabilizer(device_id, duration=2.0)
    time.sleep(0.8)
    return wait_for_stable_physical_screen_on(device_id, duration=2.0)


def turn_physical_screen_off_keep_awake(device_id):
    keep_virtual_display_device_awake(device_id)
    start_physical_screen_off_guard(device_id)
    force_physical_screen_off(device_id, use_scrcpy_helper=True)


def recover_mirror_home_keep_screen_off(device_id):
    start_scrcpy_screen_off_helper(device_id)
    time.sleep(0.03)
    send_home_with_screen_off_guard(device_id)
    time.sleep(SCREEN_OFF_FALLBACK_DELAY)
    force_physical_screen_off(device_id)


def get_foreground_window_pid():
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return None
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return pid or None


def get_foreground_scrcpy_context(manager):
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return None, None
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    if not pid:
        return None, None
    device_id = manager.get_device_id_by_pid(pid)
    if not device_id:
        return None, None
    return hwnd, device_id


def is_window_mostly_black(hwnd):
    try:
        left, top, right, bottom = win32gui.GetClientRect(hwnd)
        width = right - left
        height = bottom - top
        if width <= 0 or height <= 0:
            return False
        origin_x, origin_y = win32gui.ClientToScreen(hwnd, (0, 0))
        hdc = win32gui.GetDC(0)
    except Exception:
        return False

    samples = 0
    black_samples = 0
    try:
        for fx in (0.18, 0.32, 0.5, 0.68, 0.82):
            for fy in (0.22, 0.38, 0.54, 0.7):
                color = win32gui.GetPixel(
                    hdc,
                    origin_x + int(width * fx),
                    origin_y + int(height * fy),
                )
                if color < 0:
                    continue
                red = color & 0xFF
                green = (color >> 8) & 0xFF
                blue = (color >> 16) & 0xFF
                brightness = (red * 299 + green * 587 + blue * 114) / 1000
                samples += 1
                if brightness <= BLACK_WINDOW_BRIGHTNESS_LIMIT:
                    black_samples += 1
    finally:
        try:
            win32gui.ReleaseDC(0, hdc)
        except Exception:
            pass

    return samples >= BLACK_WINDOW_MIN_SAMPLES and black_samples / samples >= BLACK_WINDOW_RATIO


def get_client_point(hwnd, screen_x, screen_y):
    try:
        left, top, right, bottom = win32gui.GetClientRect(hwnd)
        width = right - left
        height = bottom - top
        if width <= 0 or height <= 0:
            return None
        client_x, client_y = win32gui.ScreenToClient(hwnd, (screen_x, screen_y))
        if client_x < 0 or client_y < 0 or client_x >= width or client_y >= height:
            return None
        return client_x, client_y, width, height
    except Exception:
        return None


def screen_to_client_point(hwnd, screen_x, screen_y):
    try:
        return win32gui.ScreenToClient(hwnd, (screen_x, screen_y))
    except Exception:
        return screen_x, screen_y


def adb_input_keyevent_on_display(device_id, display_id, keycode):
    if display_id is None:
        return False

    output = adb(
        [
            "-s",
            device_id,
            "shell",
            "input",
            "-d",
            str(display_id),
            "keyevent",
            keycode,
        ],
        capture_output=True,
        timeout=2,
    )
    lowered = (output or "").lower()
    if "unknown" not in lowered and "invalid" not in lowered and "error" not in lowered:
        return True

    output = adb(
        ["-s", device_id, "shell", "input", "keyevent", keycode],
        capture_output=True,
        timeout=2,
    )
    lowered = (output or "").lower()
    return "unknown" not in lowered and "invalid" not in lowered and "error" not in lowered


def adb_output_ok(output):
    lowered = (output or "").lower()
    return all(
        token not in lowered
        for token in ("unknown", "invalid", "error", "exception", "not found", "can't find", "usage:")
    )


def adb_input_keycombination_on_display(device_id, display_id, keycodes):
    if display_id is None:
        return False

    output = adb(
        [
            "-s",
            device_id,
            "shell",
            "input",
            "-d",
            str(display_id),
            "keycombination",
        ] + list(keycodes),
        capture_output=True,
        timeout=2,
    )
    if adb_output_ok(output):
        return True

    output = adb(
        ["-s", device_id, "shell", "input", "keycombination"] + list(keycodes),
        capture_output=True,
        timeout=2,
    )
    return adb_output_ok(output)


def send_virtual_display_ctrl_shortcut(device_id, display_id, key_name):
    if key_name == "a":
        return adb_input_keycombination_on_display(
            device_id,
            display_id,
            ["KEYCODE_CTRL_LEFT", "KEYCODE_A"],
        )

    return False


def show_superdex_task_view_on_display(device_id, display_id):
    if display_id is None:
        return False

    output = adb(
        [
            "-s",
            device_id,
            "shell",
            "am",
            "start-foreground-service",
            "-n",
            SUPERDEX_DOCK_SERVICE,
            "-a",
            ACTION_SHOW_DOCK,
            "--ei",
            "launch_display_id",
            str(display_id),
        ],
        capture_output=True,
        timeout=3,
    )
    lowered = (output or "").lower()
    if "error" not in lowered and "exception" not in lowered and "not found" not in lowered:
        return True

    return adb_input_keyevent_on_display(device_id, display_id, "KEYCODE_APP_SWITCH")


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("pt", POINT),
        ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


def resolve_android_key(name):
    if not name:
        return None
    name = KEY_NAME_ALIASES.get(str(name).lower(), str(name).lower())
    if len(name) == 1 and "a" <= name <= "z":
        return f"KEYCODE_{name.upper()}"
    if len(name) == 1 and name.isdigit():
        return f"KEYCODE_{name}"
    return ANDROID_KEY_MAP.get(name)


def get_modifier_name(name):
    return MODIFIER_NAME_MAP.get(str(name).lower())


def build_android_input_command(event, active_modifiers):
    if not event.name:
        return None

    name = KEY_NAME_ALIASES.get(str(event.name).lower(), str(event.name).lower())
    if get_modifier_name(name):
        return None

    ctrl_pressed = "ctrl" in active_modifiers
    alt_pressed = "alt" in active_modifiers
    windows_pressed = "windows" in active_modifiers
    shift_pressed = "shift" in active_modifiers

    if windows_pressed or alt_pressed:
        return None

    if ctrl_pressed:
        return None

    shifted_symbol_keycode = SHIFTED_SYMBOL_KEY_MAP.get(name)
    if shifted_symbol_keycode:
        return ("keycombination", ["KEYCODE_SHIFT_LEFT", shifted_symbol_keycode])

    base_keycode = resolve_android_key(name)
    if base_keycode:
        if shift_pressed and name in SHIFT_COMBO_KEYS:
            return None
        if shift_pressed and (
            name in SHIFTED_KEY_MAP
            or len(name) == 1 and "a" <= name <= "z"
        ):
            return ("keycombination", ["KEYCODE_SHIFT_LEFT", base_keycode])
        return ("keyevent", [base_keycode])

    return None


class KeyboardMapper:
    def __init__(self, manager):
        self.manager = manager
        self.key_queue = queue.Queue()
        self.pressed_scan_codes = set()
        self.virtual_shortcut_scan_codes = set()
        self.suppressed_windows_scan_codes = set()
        self.windows_shortcut_modifier_replayed = False
        self.active_modifiers = set()
        self.synthetic_shortcut_until = 0.0
        self.suppress_f12_until_up = False
        self.keyboard_lib = None
        self.enabled = False

        try:
            self.user32 = ctypes.WinDLL("user32", use_last_error=True)
            self.user32.keybd_event.argtypes = [
                ctypes.c_ubyte,
                ctypes.c_ubyte,
                wintypes.DWORD,
                ctypes.c_void_p,
            ]
            self.user32.keybd_event.restype = None
            import keyboard

            self.keyboard_lib = keyboard
            self.keyboard_lib.hook(self.handle_event, suppress=True)
            self.enabled = True
            threading.Thread(target=self.worker_loop, daemon=True).start()
            print("ℹ️ 已启用本地键盘映射转发。")
        except Exception as exc:
            print(f"⚠️ 本地键盘映射不可用：{exc}")

    def get_foreground_device_id(self):
        pid = get_foreground_window_pid()
        if pid is None:
            return None
        return self.manager.get_device_id_by_pid(pid)

    def get_active_device_id(self):
        device_id = self.get_foreground_device_id()
        if not device_id:
            return None
        if self.manager.get_keyboard_mode(device_id) != KEYBOARD_MAPPING_MODE:
            return None
        return device_id

    def handle_event(self, event):
        if not self.enabled:
            return True

        if event.event_type not in ("down", "up"):
            return True

        if time.monotonic() < self.synthetic_shortcut_until:
            return True

        key_name = str(event.name).lower() if event.name else ""
        modifier_name = get_modifier_name(key_name) if key_name else None
        foreground_hwnd, foreground_device_id = get_foreground_scrcpy_context(self.manager)

        if modifier_name == "windows" and event.event_type == "up":
            if event.scan_code in self.suppressed_windows_scan_codes:
                self.suppressed_windows_scan_codes.discard(event.scan_code)
                if not self.suppressed_windows_scan_codes:
                    self.release_windows_shortcut_modifier()
                    self.active_modifiers.discard("windows")
                return False

        if self.suppressed_windows_scan_codes:
            self.active_modifiers.add("windows")

        if foreground_device_id and self.should_forward_windows_shortcut_event(modifier_name):
            if event.event_type == "down" and not modifier_name:
                self.ensure_windows_shortcut_modifier_replayed()
            return True

        if foreground_device_id and modifier_name == "windows":
            if event.event_type == "down":
                self.suppressed_windows_scan_codes.add(event.scan_code)
                self.active_modifiers.add(modifier_name)
            return False

        if event.event_type == "up" and event.scan_code in self.virtual_shortcut_scan_codes:
            self.virtual_shortcut_scan_codes.discard(event.scan_code)
            return False

        if key_name == "esc" and foreground_device_id:
            if event.event_type == "down":
                proc = self.manager.get_scrcpy_proc(foreground_device_id)
                if is_process_running(proc):
                    print(f"ℹ️ Esc 退出 scrcpy：{foreground_device_id}")
                    stop_process(proc, timeout=1)
            return False

        if key_name == "f12" and foreground_device_id:
            if event.event_type == "up" and self.suppress_f12_until_up:
                self.suppress_f12_until_up = False
                return False
            if event.event_type == "down":
                virtual_display_id = self.manager.get_virtual_display_id(foreground_device_id)
                if virtual_display_id is not None:
                    self.suppress_f12_until_up = True
                    threading.Thread(
                        target=recover_virtual_display_activity,
                        args=(foreground_device_id, self.manager),
                        daemon=True,
                    ).start()
                    return False
                if is_window_mostly_black(foreground_hwnd):
                    self.suppress_f12_until_up = True
                    threading.Thread(
                        target=recover_mirror_home_keep_screen_off,
                        args=(foreground_device_id,),
                        daemon=True,
                    ).start()
                    return False
            return True

        if modifier_name:
            if event.event_type == "down":
                self.active_modifiers.add(modifier_name)
            else:
                self.active_modifiers.discard(modifier_name)

        if self.handle_virtual_display_ctrl_shortcut(event, key_name, foreground_device_id):
            return False

        device_id = self.get_active_device_id()
        if not device_id:
            if event.event_type == "up":
                self.pressed_scan_codes.discard(event.scan_code)
            return True

        if modifier_name:
            return True

        if event.event_type == "up":
            if event.scan_code in self.pressed_scan_codes:
                self.pressed_scan_codes.discard(event.scan_code)
                return False
            return True

        if event.scan_code in self.pressed_scan_codes:
            return False

        command = build_android_input_command(event, self.active_modifiers)
        if not command:
            return True

        self.pressed_scan_codes.add(event.scan_code)
        self.key_queue.put((device_id, command))
        return False

    def handle_virtual_display_ctrl_shortcut(self, event, key_name, device_id):
        if event.event_type != "down":
            return False
        if not device_id:
            return False
        if key_name not in VIRTUAL_DISPLAY_CTRL_SHORTCUT_KEYS:
            return False
        if event.scan_code in self.virtual_shortcut_scan_codes:
            return True
        if "ctrl" not in self.active_modifiers:
            return False
        if "alt" in self.active_modifiers or "windows" in self.active_modifiers:
            return False
        display_id = self.manager.get_virtual_display_id(device_id)
        if display_id is None:
            return False

        self.virtual_shortcut_scan_codes.add(event.scan_code)
        self.key_queue.put((device_id, ("virtual_ctrl_shortcut", (display_id, key_name))))
        return True

    def should_forward_windows_shortcut_event(self, modifier_name):
        return bool(self.suppressed_windows_scan_codes) and modifier_name != "windows"

    def ensure_windows_shortcut_modifier_replayed(self):
        if self.windows_shortcut_modifier_replayed or not self.keyboard_lib:
            return

        try:
            self.keyboard_lib.press("left windows")
            self.windows_shortcut_modifier_replayed = True
        except Exception as exc:
            print(f"Replay Windows shortcut modifier failed: {exc}")

    def release_windows_shortcut_modifier(self):
        if not self.windows_shortcut_modifier_replayed or not self.keyboard_lib:
            self.windows_shortcut_modifier_replayed = False
            return

        try:
            self.keyboard_lib.release("left windows")
        except Exception as exc:
            print(f"Release Windows shortcut modifier failed: {exc}")
        finally:
            self.windows_shortcut_modifier_replayed = False

    def send_scrcpy_clipboard_shortcut(self, key_name, shift=False):
        self.synthetic_shortcut_until = time.monotonic() + 0.35
        hotkey = f"right ctrl+{'shift+' if shift else ''}{key_name}"
        try:
            if self.keyboard_lib is not None:
                self.keyboard_lib.press_and_release(hotkey)
                return
        except Exception:
            pass

        key_code = ord(str(key_name).upper())
        self.user32.keybd_event(VK_RCONTROL, 0, KEYEVENTF_EXTENDEDKEY, None)
        try:
            if shift:
                self.user32.keybd_event(VK_SHIFT, 0, 0, None)
                time.sleep(0.01)
            time.sleep(0.01)
            self.user32.keybd_event(key_code, 0, 0, None)
            time.sleep(0.01)
            self.user32.keybd_event(key_code, 0, KEYEVENTF_KEYUP, None)
            time.sleep(0.01)
            if shift:
                self.user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, None)
        finally:
            self.user32.keybd_event(
                VK_RCONTROL,
                0,
                KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP,
                None,
            )

    def worker_loop(self):
        while True:
            device_id, command = self.key_queue.get()
            if not is_device_ready(device_id):
                continue

            command_type, payload = command
            if command_type == "keycombination":
                adb(["-s", device_id, "shell", "input", "keycombination"] + payload, capture_output=False)
            elif command_type == "keyevent":
                adb(["-s", device_id, "shell", "input", "keyevent"] + payload, capture_output=False)
            elif command_type == "virtual_ctrl_shortcut":
                display_id, key_name = payload
                if key_name == "a":
                    send_virtual_display_ctrl_shortcut(device_id, display_id, key_name)
                elif key_name == "c":
                    self.send_scrcpy_clipboard_shortcut("c")
                    adb_input_keyevent_on_display(device_id, display_id, "KEYCODE_COPY")
                elif key_name == "v":
                    # MOD+v syncs the PC clipboard to Android before paste. MOD+Shift+v
                    # only types text as key events, which is flaky for IME/Unicode text.
                    self.send_scrcpy_clipboard_shortcut("v")


class TouchGestureBridge:
    def __init__(self, gesture_sink):
        print(f"Touch gesture bridge build: {GESTURE_BUILD_TAG}")
        self.gesture_sink = gesture_sink
        self.collector = TouchStrokeCollector()
        self.live_recognizer = FixedGestureRecognizer()
        self.registered_hwnds = set()
        self.raw_registered_hwnds = set()
        self.user32 = None
        self.raw_helpers = None
        self.raw_window_hwnd = None
        self.raw_window_ready = threading.Event()
        self.active_source = None
        self.active_device = InputDevices.NONE
        self.allowed_devices = (
            InputDevices.TOUCH_SCREEN | InputDevices.TOUCH_PAD | InputDevices.PEN
        )
        self.device_cache = {}
        self.raw_required_contact_count = 0
        self.raw_output_contacts = []
        self.raw_last_touch_count = 0
        self.raw_count_changed_at = 0.0
        self.raw_count_stable_frames = 0
        self.raw_live_action_fired = False
        self.raw_live_action_at = 0.0
        self.raw_suppress_until_idle = False
        self.raw_completion_needs_idle_suppression = False
        self.raw_deferred_strokes = None
        self.raw_deferred_finger_count = 0

        try:
            self.user32 = ctypes.WinDLL("user32", use_last_error=True)
            self.user32.RegisterTouchWindow.argtypes = [wintypes.HWND, wintypes.ULONG]
            self.user32.RegisterTouchWindow.restype = wintypes.BOOL
            self.user32.GetTouchInputInfo.argtypes = [
                wintypes.HANDLE,
                wintypes.UINT,
                ctypes.POINTER(TOUCHINPUT),
                ctypes.c_int,
            ]
            self.user32.GetTouchInputInfo.restype = wintypes.BOOL
            self.user32.CloseTouchInputHandle.argtypes = [wintypes.HANDLE]
            self.user32.CloseTouchInputHandle.restype = wintypes.BOOL
            self.user32.GetPointerInfo.argtypes = [
                ctypes.c_uint32,
                ctypes.POINTER(POINTER_INFO),
            ]
            self.user32.GetPointerInfo.restype = wintypes.BOOL
            self.user32.SkipPointerFrameMessages.argtypes = [ctypes.c_uint32]
            self.user32.SkipPointerFrameMessages.restype = wintypes.BOOL
            if hasattr(self.user32, "RegisterTouchpadCapableWindow"):
                self.user32.RegisterTouchpadCapableWindow.argtypes = [
                    wintypes.HWND,
                    wintypes.BOOL,
                ]
                self.user32.RegisterTouchpadCapableWindow.restype = wintypes.BOOL
            if hasattr(self.user32, "RegisterPointerInputTarget"):
                self.user32.RegisterPointerInputTarget.argtypes = [
                    wintypes.HWND,
                    ctypes.c_int,
                ]
                self.user32.RegisterPointerInputTarget.restype = wintypes.BOOL
            self.raw_helpers = get_raw_helpers()
            threading.Thread(target=self._run_raw_input_window, daemon=True).start()
        except Exception as exc:
            print(f"⚠️ 触控输入桥接初始化失败：{exc}")
            self.user32 = None
            self.raw_helpers = None

    @staticmethod
    def qt_point_to_tuple(point):
        for attr_name in ("screenPos", "scenePos", "pos"):
            getter = getattr(point, attr_name, None)
            if getter is None:
                continue
            try:
                pos = getter()
            except TypeError:
                pos = getter
            if pos is not None:
                return float(pos.x()), float(pos.y())
        return 0.0, 0.0

    @staticmethod
    def qt_touch_state_to_name(state):
        if state & QtCore.Qt.TouchPointPressed:
            return "down"
        if state & QtCore.Qt.TouchPointReleased:
            return "up"
        if state & QtCore.Qt.TouchPointCanceled:
            return "cancel"
        if state & QtCore.Qt.TouchPointMoved:
            return "move"
        if state & QtCore.Qt.TouchPointStationary:
            return "move"
        return None

    @staticmethod
    def source_device_from_usage(usage):
        if usage == TOUCH_PAD_USAGE:
            return InputDevices.TOUCH_PAD
        if usage == TOUCH_SCREEN_USAGE:
            return InputDevices.TOUCH_SCREEN
        if usage == PEN_USAGE:
            return InputDevices.PEN
        return InputDevices.NONE

    def register_window(self, hwnd):
        if not hwnd:
            return False
        self._register_qt_touch_window(hwnd)
        self._register_touchpad_capable_window(hwnd)
        return True

    def _register_touchpad_capable_window(self, hwnd):
        register = getattr(self.user32, "RegisterTouchpadCapableWindow", None) if self.user32 else None
        if not register:
            return False
        hwnd_value = int(hwnd)
        if register(wintypes.HWND(hwnd_value), True):
            if GESTURE_DEBUG:
                print("RegisterTouchpadCapableWindow succeeded")
            return True
        error_code = ctypes.get_last_error()
        if GESTURE_DEBUG:
            print(f"RegisterTouchpadCapableWindow failed: err={error_code}")
        return False

    def _register_touchpad_pointer_target(self, hwnd):
        register = getattr(self.user32, "RegisterPointerInputTarget", None) if self.user32 else None
        if not register:
            return False
        hwnd_value = int(hwnd)
        if register(wintypes.HWND(hwnd_value), PT_TOUCHPAD):
            if GESTURE_DEBUG:
                print("RegisterPointerInputTarget(PT_TOUCHPAD) succeeded")
            return True
        error_code = ctypes.get_last_error()
        if GESTURE_DEBUG:
            print(f"RegisterPointerInputTarget(PT_TOUCHPAD) failed: err={error_code}")
        return False

    def _run_raw_input_window(self):
        if not self.raw_helpers:
            self.raw_window_ready.set()
            return

        class_name = "SuperdexRawInputWindow"

        def wndproc(hwnd, msg, wparam, lparam):
            if msg in (WM_POINTERDOWN, WM_POINTERUPDATE, WM_POINTERUP):
                try:
                    if self.handle_native_pointer_message(hwnd, msg, wparam, lparam):
                        return 0
                except Exception as exc:
                    print(f"⚠️ Precision Touchpad pointer 处理失败：{exc}")
            if msg == WM_INPUT or msg == WM_INPUT_DEVICE_CHANGE:
                try:
                    if self.handle_native_raw_input_message(msg, wparam, lparam):
                        return 0
                except Exception as exc:
                    print(f"⚠️ 原始输入窗口处理失败：{exc}")
            return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

        try:
            hinstance = ctypes.windll.kernel32.GetModuleHandleW(None)
            wc = win32gui.WNDCLASS()
            wc.hInstance = hinstance
            wc.lpszClassName = class_name
            wc.lpfnWndProc = wndproc
            try:
                win32gui.RegisterClass(wc)
            except win32gui.error:
                pass
            hwnd = win32gui.CreateWindow(
                class_name,
                "Superdex Raw Input",
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                hinstance,
                None,
            )
            self.raw_window_hwnd = hwnd
            self.raw_window_ready.set()
            self._register_raw_input_window(hwnd)
            self._register_touchpad_pointer_target(hwnd)
            print("Raw digitizer input window enabled")
            win32gui.PumpMessages()
        except Exception as exc:
            print(f"⚠️ 原始输入窗口启动失败：{exc}")
            self.raw_window_ready.set()

    def _register_qt_touch_window(self, hwnd):
        if not self.user32:
            return False
        hwnd_value = int(hwnd)
        if hwnd_value in self.registered_hwnds:
            return True
        if self.user32.RegisterTouchWindow(wintypes.HWND(hwnd_value), 0):
            self.registered_hwnds.add(hwnd_value)
            return True
        error_code = ctypes.get_last_error()
        print(f"⚠️ RegisterTouchWindow 失败：{error_code}")
        return False

    def _register_raw_input_window(self, hwnd=None):
        if not self.raw_helpers:
            return False
        if hwnd is None:
            hwnd = self.raw_window_hwnd
        if not hwnd:
            return False
        hwnd_value = int(hwnd)
        if hwnd_value in self.raw_registered_hwnds:
            return True
        usages = []
        if self.allowed_devices & InputDevices.TOUCH_PAD:
            usages.append(TOUCH_PAD_USAGE)
        if self.allowed_devices & InputDevices.TOUCH_SCREEN:
            usages.append(TOUCH_SCREEN_USAGE)
        if self.allowed_devices & InputDevices.PEN:
            usages.append(PEN_USAGE)

        ok = True
        raw_input_flags = RIDEV_INPUTSINK | RIDEV_EXINPUTSINK | RIDEV_DEVNOTIFY
        registrations = [(usage, raw_input_flags) for usage in usages]
        if usages:
            registrations.insert(0, (0, raw_input_flags | RIDEV_PAGEONLY))
        registered_usages = []
        for usage, flags in registrations:
            rid = RAWINPUTDEVICE(
                usUsagePage=DIGITIZER_USAGE_PAGE,
                usUsage=usage,
                dwFlags=flags,
                hwndTarget=wintypes.HWND(hwnd_value),
            )
            if not self.raw_helpers.user32.RegisterRawInputDevices(
                ctypes.byref(rid), 1, ctypes.sizeof(rid)
            ):
                ok = False
                print(
                    "RegisterRawInputDevices failed: "
                    f"usage={usage}, flags=0x{flags:04x}, err={ctypes.get_last_error()}"
                )
            else:
                registered_usages.append(f"{usage}:0x{flags:04x}")
                continue
                print(f"⚠️ RegisterRawInputDevices 失败：usage={usage}, err={ctypes.get_last_error()}")
        if ok:
            self.raw_registered_hwnds.add(hwnd_value)
            print(
                "Raw digitizer input registered: "
                f"hwnd={hwnd_value}, usages={','.join(registered_usages)}"
            )
        return ok

    def _set_source(self, source_name, source_device):
        self.active_source = source_name
        self.active_device = source_device

    def _clear_source_if_idle(self):
        if not self.collector.has_active_contacts():
            self.active_source = None
            self.active_device = InputDevices.NONE

    def _reset_raw_touch_state(self):
        self.raw_required_contact_count = 0
        self.raw_output_contacts = []
        self.raw_last_touch_count = 0
        self.raw_count_changed_at = 0.0
        self.raw_count_stable_frames = 0
        self.raw_live_action_fired = False
        self.raw_live_action_at = 0.0
        self.raw_suppress_until_idle = False
        self.raw_completion_needs_idle_suppression = False
        self.raw_deferred_strokes = None
        self.raw_deferred_finger_count = 0

    def _dispatch_completed(self, completed):
        if not completed:
            return
        strokes, finger_count = completed
        if strokes:
            if self.raw_live_action_fired:
                if GESTURE_DEBUG:
                    print("Touch gesture completion suppressed after live trigger")
                self._reset_raw_touch_state()
                self.active_source = None
                self.active_device = InputDevices.NONE
                return
            source_name = self.active_source
            if GESTURE_DEBUG:
                print(
                    "Touch gesture completed: "
                    f"source={source_name}, "
                    f"device={self.active_device.name}, "
                    f"fingers={finger_count}, "
                    f"points={[len(stroke) for stroke in strokes]}"
                )
            suppress_until_idle = self.raw_completion_needs_idle_suppression
            self.gesture_sink(strokes, finger_count)
            if source_name == "raw-input":
                self._reset_raw_touch_state()
                if suppress_until_idle:
                    self.raw_suppress_until_idle = True
            self.active_source = None
            self.active_device = InputDevices.NONE

    def _try_dispatch_live_raw_gesture(self):
        if self.raw_live_action_fired:
            return False
        if self.raw_count_changed_at:
            stable_elapsed = time.monotonic() - self.raw_count_changed_at
            if (
                self.raw_count_stable_frames < RAW_FINGER_COUNT_SETTLE_FRAMES
                and stable_elapsed < RAW_FINGER_COUNT_SETTLE_SECONDS
            ):
                return False
        strokes, finger_count = self.collector.active_strokes_snapshot()
        if finger_count < 2:
            return False
        if any(len(stroke) < 2 for stroke in strokes):
            return False
        gesture_match = self.live_recognizer.recognize(strokes, finger_count)
        if not gesture_match or gesture_match.gesture.direction == "tap":
            return False
        if (
            finger_count >= 3
            and gesture_match.action.action_type != "activate_window_class"
        ):
            self.raw_deferred_strokes = [list(stroke) for stroke in strokes]
            self.raw_deferred_finger_count = finger_count
            if GESTURE_DEBUG:
                print(
                    "Touch gesture deferred trigger: "
                    f"fingers={finger_count}, "
                    f"action={gesture_match.action.action_type}, "
                    f"score={gesture_match.score:.1f}"
                )
            return False
        if GESTURE_DEBUG:
            print(
                "Touch gesture live trigger: "
                f"fingers={finger_count}, "
                f"action={gesture_match.action.action_type}, "
                f"score={gesture_match.score:.1f}"
            )
        if self.gesture_sink(strokes, finger_count):
            self.raw_live_action_fired = True
            self.raw_live_action_at = time.monotonic()
            return True
        return False

    def _dispatch_deferred_raw_gesture(self):
        if not self.raw_deferred_strokes or self.raw_deferred_finger_count <= 0:
            return False
        strokes = [list(stroke) for stroke in self.raw_deferred_strokes]
        finger_count = int(self.raw_deferred_finger_count)
        self.raw_deferred_strokes = None
        self.raw_deferred_finger_count = 0
        if GESTURE_DEBUG:
            print(
                "Touch gesture deferred dispatch: "
                f"fingers={finger_count}, "
                f"points={[len(stroke) for stroke in strokes]}"
            )
        return bool(self.gesture_sink(strokes, finger_count))

    def handle_qt_touch_event(self, event):
        touch_points = event.touchPoints()
        if not touch_points:
            return False
        if self.active_source not in (None, "qt"):
            return False
        self._set_source("qt", InputDevices.TOUCH_SCREEN)

        completed = None
        for point in touch_points:
            state_name = self.qt_touch_state_to_name(point.state())
            if state_name is None:
                continue
            completed = self.collector.process_contact(
                point.id(),
                self.qt_point_to_tuple(point),
                state_name,
            ) or completed

        if completed:
            self._dispatch_completed(completed)
        elif event.type() in (QtCore.QEvent.TouchEnd, QtCore.QEvent.TouchCancel):
            self._clear_source_if_idle()
        return True

    def handle_native_touch_message(self, hwnd, message, w_param, l_param):
        if not self.user32 or message != WM_TOUCH:
            return False
        if self.active_source not in (None, "native-touch"):
            return False
        self._set_source("native-touch", InputDevices.TOUCH_SCREEN)

        touch_count = int(w_param) & 0xFFFF
        if touch_count <= 0:
            self._clear_source_if_idle()
            return False

        touch_inputs = (TOUCHINPUT * touch_count)()
        if not self.user32.GetTouchInputInfo(
            wintypes.HANDLE(int(l_param)),
            touch_count,
            touch_inputs,
            ctypes.sizeof(TOUCHINPUT),
        ):
            print(f"⚠️ GetTouchInputInfo 失败：{ctypes.get_last_error()}")
            self._clear_source_if_idle()
            return False

        try:
            completed = None
            for touch_input in touch_inputs:
                state_name = touch_state_from_flags(touch_input.dwFlags)
                completed = self.collector.process_contact(
                    touch_input.dwID,
                    touch_point_to_tuple(touch_input.x, touch_input.y),
                    state_name,
                ) or completed
            if completed:
                self._dispatch_completed(completed)
            else:
                self._clear_source_if_idle()
            return True
        finally:
            self.user32.CloseTouchInputHandle(wintypes.HANDLE(int(l_param)))

    @staticmethod
    def pointer_id_from_wparam(w_param):
        return int(w_param) & 0xFFFF

    @staticmethod
    def pointer_state_from_message(message):
        if message == WM_POINTERDOWN:
            return "down"
        if message == WM_POINTERUP:
            return "up"
        if message == WM_POINTERUPDATE:
            return "move"
        return None

    def handle_native_pointer_message(self, hwnd, message, w_param, l_param):
        if not self.user32 or message not in (WM_POINTERDOWN, WM_POINTERUPDATE, WM_POINTERUP):
            return False
        if self.active_source not in (None, "pointer-touchpad"):
            return False

        pointer_id = self.pointer_id_from_wparam(w_param)
        pointer_info = POINTER_INFO()
        if not self.user32.GetPointerInfo(pointer_id, ctypes.byref(pointer_info)):
            if GESTURE_DEBUG:
                print(f"⚠️ GetPointerInfo 失败：{ctypes.get_last_error()}")
            self._clear_source_if_idle()
            return False
        if pointer_info.pointerType != PT_TOUCHPAD:
            return False

        state_name = self.pointer_state_from_message(message)
        if not state_name:
            return False

        self._set_source("pointer-touchpad", InputDevices.TOUCH_PAD)
        point = (
            float(pointer_info.ptHimetricLocation.x),
            float(pointer_info.ptHimetricLocation.y),
        )
        completed = self.collector.process_contact(
            pointer_info.pointerId,
            point,
            state_name,
        )
        if completed:
            self._dispatch_completed(completed)
        else:
            self._clear_source_if_idle()
        return True

    def handle_native_raw_input_message(self, message, w_param, l_param):
        if not self.raw_helpers or message not in (WM_INPUT, WM_INPUT_DEVICE_CHANGE):
            return False

        if message == WM_INPUT_DEVICE_CHANGE:
            self.device_cache.clear()
            self._reset_raw_touch_state()
            return False

        if self.active_source not in (None, "raw-input"):
            return False

        raw_input = self._read_raw_input(w_param, l_param)
        if raw_input is None:
            self._clear_source_if_idle()
            return False

        buffer, header, rawhid = raw_input
        source_device = self._get_source_device(header.hDevice)
        if source_device == InputDevices.NONE or not (source_device & self.allowed_devices):
            return False

        self._set_source("raw-input", source_device)
        completed = self._process_raw_digitizer(buffer, header, rawhid, source_device)
        if completed:
            self._dispatch_completed(completed)
        else:
            self._clear_source_if_idle()
        return True

    def _read_raw_input(self, w_param, l_param):
        size = wintypes.UINT(0)
        hrawinput = wintypes.HANDLE(int(l_param))
        result = self.raw_helpers.user32.GetRawInputData(
            hrawinput,
            RID_INPUT,
            None,
            ctypes.byref(size),
            RAWINPUTHEADER_SIZE,
        )
        if result == 0xFFFFFFFF or size.value == 0:
            return None

        buffer = ctypes.create_string_buffer(size.value)
        result = self.raw_helpers.user32.GetRawInputData(
            hrawinput,
            RID_INPUT,
            ctypes.addressof(buffer),
            ctypes.byref(size),
            RAWINPUTHEADER_SIZE,
        )
        if result == 0xFFFFFFFF:
            print(f"⚠️ GetRawInputData 失败：{ctypes.get_last_error()}")
            return None

        header = RAWINPUTHEADER.from_buffer_copy(buffer)
        if header.dwType != RIM_TYPEHID:
            return None
        rawhid = RAWHID.from_buffer_copy(buffer, RAWINPUTHEADER_SIZE)
        return buffer, header, rawhid

    def _get_source_device(self, device_handle):
        key = int(device_handle)
        if key in self.device_cache:
            return self.device_cache[key]

        usage = self._get_device_usage(device_handle)
        source_device = self.source_device_from_usage(usage)
        self.device_cache[key] = source_device
        return source_device

    def _get_device_usage(self, device_handle):
        size = wintypes.UINT(ctypes.sizeof(RID_DEVICE_INFO))
        info = RID_DEVICE_INFO()
        info.cbSize = ctypes.sizeof(RID_DEVICE_INFO)
        if self.raw_helpers.user32.GetRawInputDeviceInfoW(
            device_handle,
            RIDI_DEVICEINFO,
            ctypes.addressof(info),
            ctypes.byref(size),
        ) == 0xFFFFFFFF:
            return None
        if info.dwType != RIM_TYPEHID:
            return None
        if info.hid.usUsagePage != DIGITIZER_USAGE_PAGE:
            return None

        name = self._get_device_name(device_handle)
        if name and ("VIRTUAL_DIGITIZER" in name.upper() or "ROOT" in name.upper()):
            return None
        return int(info.hid.usUsage)

    def _get_device_name(self, device_handle):
        size = wintypes.UINT(0)
        if self.raw_helpers.user32.GetRawInputDeviceInfoW(
            device_handle,
            RIDI_DEVICENAME,
            None,
            ctypes.byref(size),
        ) == 0xFFFFFFFF or size.value == 0:
            return ""
        buffer = ctypes.create_unicode_buffer(size.value)
        if self.raw_helpers.user32.GetRawInputDeviceInfoW(
            device_handle,
            RIDI_DEVICENAME,
            ctypes.addressof(buffer),
            ctypes.byref(size),
        ) == 0xFFFFFFFF:
            return ""
        return buffer.value or ""

    def _get_preparsed_data(self, device_handle):
        size = wintypes.UINT(0)
        if self.raw_helpers.user32.GetRawInputDeviceInfoW(
            device_handle,
            RIDI_PREPARSEDDATA,
            None,
            ctypes.byref(size),
        ) == 0xFFFFFFFF or size.value == 0:
            return None, None
        buffer = ctypes.create_string_buffer(size.value)
        if self.raw_helpers.user32.GetRawInputDeviceInfoW(
            device_handle,
            RIDI_PREPARSEDDATA,
            ctypes.addressof(buffer),
            ctypes.byref(size),
        ) == 0xFFFFFFFF:
            return None, None
        return buffer, ctypes.addressof(buffer)

    def _get_link_collection_nodes(self, preparsed_data):
        count = ctypes.c_int(0)
        self.raw_helpers.hid.HidP_GetLinkCollectionNodes(None, ctypes.byref(count), preparsed_data)
        if count.value <= 0:
            return []
        nodes = (HIDP_LINK_COLLECTION_NODE * count.value)()
        if self.raw_helpers.hid.HidP_GetLinkCollectionNodes(nodes, ctypes.byref(count), preparsed_data) != HIDP_STATUS_SUCCESS:
            return []
        return nodes

    def _hid_get_usage_value(self, preparsed_data, report_ptr, report_len, link_collection, usage):
        value = ctypes.c_int(0)
        if self.raw_helpers.hid.HidP_GetUsageValue(
            HIDP_REPORT_TYPE_INPUT,
            GENERIC_DESKTOP_PAGE if usage in (X_COORDINATE_ID, Y_COORDINATE_ID) else DIGITIZER_USAGE_PAGE,
            link_collection,
            usage,
            ctypes.byref(value),
            preparsed_data,
            report_ptr,
            report_len,
        ) != HIDP_STATUS_SUCCESS:
            return None
        return value.value

    def _hid_get_scaled_usage_value(self, preparsed_data, report_ptr, report_len, link_collection, usage):
        value = ctypes.c_int(0)
        if self.raw_helpers.hid.HidP_GetScaledUsageValue(
            HIDP_REPORT_TYPE_INPUT,
            GENERIC_DESKTOP_PAGE,
            link_collection,
            usage,
            ctypes.byref(value),
            preparsed_data,
            report_ptr,
            report_len,
        ) != HIDP_STATUS_SUCCESS:
            return None
        return value.value

    def _hid_get_coordinate_value(self, preparsed_data, report_ptr, report_len, link_collection, usage):
        value = self._hid_get_scaled_usage_value(
            preparsed_data,
            report_ptr,
            report_len,
            link_collection,
            usage,
        )
        if value is not None:
            return value
        value = self._hid_get_usage_value(
            preparsed_data,
            report_ptr,
            report_len,
            link_collection,
            usage,
        )
        return 0 if value is None else value

    def _hid_get_usages(self, preparsed_data, report_ptr, report_len, link_collection):
        usage_len = ctypes.c_int(0)
        usage_len.value = int(
            self.raw_helpers.hid.HidP_MaxUsageListLength(
                HIDP_REPORT_TYPE_INPUT,
                DIGITIZER_USAGE_PAGE,
                preparsed_data,
            )
        )
        if usage_len.value <= 0:
            return []
        usage_list = (ctypes.c_ushort * usage_len.value)()
        if self.raw_helpers.hid.HidP_GetUsages(
            HIDP_REPORT_TYPE_INPUT,
            DIGITIZER_USAGE_PAGE,
            link_collection,
            usage_list,
            ctypes.byref(usage_len),
            preparsed_data,
            report_ptr,
            report_len,
        ) != HIDP_STATUS_SUCCESS:
            return []
        return list(usage_list[: usage_len.value])

    @staticmethod
    def _finger_link_collections(nodes):
        finger_nodes = [
            index
            for index, node in enumerate(nodes)
            if int(node.LinkUsagePage) == DIGITIZER_USAGE_PAGE
            and int(node.LinkUsage) == FINGER_ID
        ]
        if finger_nodes:
            return finger_nodes

        if not nodes:
            return [0]

        child_count = max(0, int(nodes[0].NumberOfChildren))
        if child_count > 0:
            return list(range(1, min(len(nodes), child_count + 1)))

        return list(range(1, len(nodes))) or [0]

    @staticmethod
    def _touch_link_collections(nodes):
        if not nodes:
            return [0]

        child_count = max(0, int(nodes[0].NumberOfChildren))
        if child_count > 0:
            return list(range(1, min(len(nodes), child_count + 1)))

        return TouchGestureBridge._finger_link_collections(nodes)

    @staticmethod
    def _synthetic_contact_id(device_handle, link_collection):
        return ((int(device_handle) & 0xFFFF) << 16) | int(link_collection)

    def _collect_raw_touch_contacts(self, raw_buffer, header, rawhid, preparsed_data):
        base_address = ctypes.addressof(raw_buffer)
        data_offset = RAWINPUT_BUFFER_HEADER_SIZE
        packet_size = int(rawhid.dwSizeHid)
        packet_count = max(1, int(rawhid.dwCount))
        first_packet_ptr = ctypes.c_void_p(base_address + data_offset)
        contact_count = self._hid_get_usage_value(
            preparsed_data,
            first_packet_ptr,
            packet_size,
            0,
            CONTACT_COUNT_ID,
        )

        nodes = self._get_link_collection_nodes(preparsed_data)
        link_collections = self._touch_link_collections(nodes)
        if not link_collections:
            return []

        candidates = []
        for packet_index in range(packet_count):
            packet_ptr = ctypes.c_void_p(base_address + data_offset + packet_size * packet_index)
            for node_index in link_collections:
                contact_id = self._hid_get_usage_value(
                    preparsed_data,
                    packet_ptr,
                    packet_size,
                    node_index,
                    CONTACT_IDENTIFIER_ID,
                )
                if contact_id is None or int(contact_id) <= 0:
                    contact_id = self._synthetic_contact_id(header.hDevice, node_index)

                x_value = self._hid_get_coordinate_value(
                    preparsed_data,
                    packet_ptr,
                    packet_size,
                    node_index,
                    X_COORDINATE_ID,
                )
                y_value = self._hid_get_coordinate_value(
                    preparsed_data,
                    packet_ptr,
                    packet_size,
                    node_index,
                    Y_COORDINATE_ID,
                )
                usage_list = self._hid_get_usages(
                    preparsed_data,
                    packet_ptr,
                    packet_size,
                    node_index,
                )
                candidates.append(
                    (
                        int(contact_id),
                        (float(x_value), float(y_value)),
                        TIP_ID in usage_list,
                    )
                )

        if contact_count is not None and int(contact_count) > 0:
            self.raw_required_contact_count = int(contact_count)
            self.raw_output_contacts = []

        if self.raw_required_contact_count > 0:
            for contact in candidates:
                self.raw_output_contacts.append(contact)
                self.raw_required_contact_count -= 1
                if self.raw_required_contact_count == 0:
                    contacts = self.raw_output_contacts
                    self.raw_output_contacts = []
                    return contacts
            return []

        return [
            contact
            for contact in candidates
            if contact[2] or self.collector.is_active(contact[0])
        ]

    def _feed_raw_touch_contacts(self, contacts, state):
        completed = None
        for contact_id, point, tip_active in contacts:
            if state != "up" and not tip_active:
                continue
            completed = self.collector.process_contact(
                contact_id,
                point,
                state,
            ) or completed
        return completed

    def _end_raw_touch_contacts(self, contacts):
        contact_points = {int(contact_id): point for contact_id, point, _ in contacts}
        return self.collector.end_all_contacts(contact_points)

    def _dispatch_raw_touch_contacts(self, contacts, source_device):
        if not contacts:
            return None

        active_contacts = [
            (contact_id, point, tip_active)
            for contact_id, point, tip_active in contacts
            if tip_active
        ]
        active_count = len(active_contacts)
        release_count = sum(1 for _, _, tip_active in contacts if not tip_active)

        if GESTURE_DEBUG and (
            release_count or active_count != self.raw_last_touch_count
        ):
            ids = [contact_id for contact_id, _, tip_active in contacts if tip_active]
            print(
                "Raw touch frame: "
                f"device={source_device.name}, "
                f"active={active_count}, "
                f"last={self.raw_last_touch_count}, "
                f"released={release_count}, "
                f"active_ids={ids}"
            )

        if self.raw_suppress_until_idle:
            if active_count <= 0:
                if GESTURE_DEBUG:
                    print("Raw touch release completed after deferred trigger")
                self.collector.reset()
                self._reset_raw_touch_state()
            return None

        if self.raw_live_action_fired:
            if active_count <= 0:
                if GESTURE_DEBUG:
                    print("Raw touch release completed after live trigger")
                self.collector.reset()
                self._reset_raw_touch_state()
                return None
            live_elapsed = (
                time.monotonic() - self.raw_live_action_at
                if self.raw_live_action_at
                else 0.0
            )
            if live_elapsed < RAW_LIVE_ACTION_RELEASE_TIMEOUT_SECONDS:
                return None
            if GESTURE_DEBUG:
                print(
                    "Raw touch live trigger release timeout; "
                    "resetting stale gesture state"
                )
            self.collector.reset()
            self._reset_raw_touch_state()
            # Continue processing the current active frame as the start of a new
            # gesture. Some host actions (notably Show Desktop/Win+D) can steal
            # focus before Windows delivers the release frame, otherwise the
            # stale live-trigger flag suppresses every later gesture.

        if active_count <= 0:
            if self._dispatch_deferred_raw_gesture():
                self.collector.reset()
                self._reset_raw_touch_state()
                return None
            completed = self._end_raw_touch_contacts(contacts)
            self.raw_last_touch_count = 0
            self.raw_count_changed_at = 0.0
            self.raw_count_stable_frames = 0
            return completed

        if (
            release_count
            and active_count < self.raw_last_touch_count
            and self.collector.has_active_contacts()
        ):
            if self._dispatch_deferred_raw_gesture():
                self.collector.reset()
                self._reset_raw_touch_state()
                self.raw_suppress_until_idle = True
                return None
            completed = self._end_raw_touch_contacts(contacts)
            self.raw_last_touch_count = active_count
            self.raw_count_changed_at = 0.0
            self.raw_count_stable_frames = 0
            if active_count > 0:
                self.raw_completion_needs_idle_suppression = True
            return completed

        if active_count != self.raw_last_touch_count:
            if self.collector.has_active_contacts():
                self.collector.reset()
            self.raw_deferred_strokes = None
            self.raw_deferred_finger_count = 0
            self.raw_live_action_fired = False
            completed = self._feed_raw_touch_contacts(active_contacts, "down")
            self.raw_last_touch_count = active_count
            self.raw_count_changed_at = time.monotonic()
            self.raw_count_stable_frames = 0
            return completed

        self.raw_count_stable_frames += 1
        completed = self._feed_raw_touch_contacts(active_contacts, "move")
        if not completed:
            self._try_dispatch_live_raw_gesture()

        if completed:
            self.raw_last_touch_count = 0
        return completed

    def _process_raw_digitizer(self, raw_buffer, header, rawhid, source_device):
        preparsed_buffer, preparsed_data = self._get_preparsed_data(header.hDevice)
        if not preparsed_data:
            return None

        base_address = ctypes.addressof(raw_buffer)
        data_offset = RAWINPUT_BUFFER_HEADER_SIZE
        completed = None

        if source_device in (InputDevices.TOUCH_SCREEN, InputDevices.TOUCH_PAD):
            contacts = self._collect_raw_touch_contacts(
                raw_buffer,
                header,
                rawhid,
                preparsed_data,
            )
            completed = self._dispatch_raw_touch_contacts(contacts, source_device)
        elif source_device == InputDevices.PEN:
            packet_ptr = ctypes.c_void_p(base_address + data_offset)
            usage_list = self._hid_get_usages(preparsed_data, packet_ptr, rawhid.dwSizeHid, 0)
            pen_state = 0
            for usage in usage_list:
                if usage in (TIP_ID, IN_RANGE_ID, BARREL_BUTTON_ID, INVERT_ID, ERASER_ID):
                    pen_state |= 1
            x_value = self._hid_get_coordinate_value(
                preparsed_data,
                packet_ptr,
                rawhid.dwSizeHid,
                0,
                X_COORDINATE_ID,
            )
            y_value = self._hid_get_coordinate_value(
                preparsed_data,
                packet_ptr,
                rawhid.dwSizeHid,
                0,
                Y_COORDINATE_ID,
            )
            contact_id = 0
            state = "down" if pen_state and not self.collector.is_active(contact_id) else "move"
            if not pen_state:
                state = "up"
            completed = self.collector.process_contact(
                contact_id,
                (float(x_value), float(y_value)),
                state,
            ) or completed

        return completed

    def _finalize_if_idle(self):
        if not self.collector.has_active_contacts():
            self.active_source = None
            self.active_device = InputDevices.NONE



class MouseShortcutMapper:
    def __init__(self, manager, keyboard_mapper=None):
        self.manager = manager
        self.keyboard_mapper = keyboard_mapper
        self.command_queue = queue.Queue()
        self.hook_handle = None
        self.callback_ref = None
        self.edge_gesture = None
        self.edge_gesture_lock = threading.Lock()
        self.fixed_gesture_recognizer = FixedGestureRecognizer()
        self.fixed_gesture_action_seq = 0
        self.synthetic_mouse_until = 0.0
        self.switch_to_this_window = None

        try:
            self.user32 = ctypes.WinDLL("user32", use_last_error=True)
            self.kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            self.user32.SetWindowsHookExW.argtypes = [
                ctypes.c_int,
                ctypes.c_void_p,
                ctypes.c_void_p,
                wintypes.DWORD,
            ]
            self.user32.SetWindowsHookExW.restype = ctypes.c_void_p
            self.user32.CallNextHookEx.argtypes = [
                ctypes.c_void_p,
                ctypes.c_int,
                wintypes.WPARAM,
                wintypes.LPARAM,
            ]
            self.user32.CallNextHookEx.restype = wintypes.LPARAM
            self.user32.GetMessageW.argtypes = [
                ctypes.POINTER(wintypes.MSG),
                ctypes.c_void_p,
                wintypes.UINT,
                wintypes.UINT,
            ]
            self.user32.GetMessageW.restype = wintypes.BOOL
            self.user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
            self.user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
            self.user32.keybd_event.argtypes = [
                ctypes.c_ubyte,
                ctypes.c_ubyte,
                wintypes.DWORD,
                ctypes.c_void_p,
            ]
            self.user32.keybd_event.restype = None
            self.user32.mouse_event.argtypes = [
                wintypes.DWORD,
                wintypes.DWORD,
                wintypes.DWORD,
                wintypes.DWORD,
                ctypes.c_void_p,
            ]
            self.user32.mouse_event.restype = None
            self.user32.SendInput.argtypes = [
                wintypes.UINT,
                ctypes.POINTER(INPUT),
                ctypes.c_int,
            ]
            self.user32.SendInput.restype = wintypes.UINT
            self.user32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
            self.user32.SetCursorPos.restype = wintypes.BOOL
            self.user32.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
            self.user32.FindWindowW.restype = wintypes.HWND
            self.user32.GetForegroundWindow.argtypes = []
            self.user32.GetForegroundWindow.restype = wintypes.HWND
            self.user32.GetWindowThreadProcessId.argtypes = [
                wintypes.HWND,
                ctypes.POINTER(wintypes.DWORD),
            ]
            self.user32.GetWindowThreadProcessId.restype = wintypes.DWORD
            self.user32.AttachThreadInput.argtypes = [
                wintypes.DWORD,
                wintypes.DWORD,
                wintypes.BOOL,
            ]
            self.user32.AttachThreadInput.restype = wintypes.BOOL
            self.user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
            self.user32.ShowWindow.restype = wintypes.BOOL
            self.user32.ShowWindowAsync.argtypes = [wintypes.HWND, ctypes.c_int]
            self.user32.ShowWindowAsync.restype = wintypes.BOOL
            try:
                self.switch_to_this_window = self.user32.SwitchToThisWindow
                self.switch_to_this_window.argtypes = [wintypes.HWND, wintypes.BOOL]
                self.switch_to_this_window.restype = None
            except AttributeError:
                self.switch_to_this_window = None
            self.user32.BringWindowToTop.argtypes = [wintypes.HWND]
            self.user32.BringWindowToTop.restype = wintypes.BOOL
            self.user32.SetWindowPos.argtypes = [
                wintypes.HWND,
                wintypes.HWND,
                ctypes.c_int,
                ctypes.c_int,
                ctypes.c_int,
                ctypes.c_int,
                wintypes.UINT,
            ]
            self.user32.SetWindowPos.restype = wintypes.BOOL
            self.user32.SetActiveWindow.argtypes = [wintypes.HWND]
            self.user32.SetActiveWindow.restype = wintypes.HWND
            self.user32.SetFocus.argtypes = [wintypes.HWND]
            self.user32.SetFocus.restype = wintypes.HWND
            self.user32.IsIconic.argtypes = [wintypes.HWND]
            self.user32.IsIconic.restype = wintypes.BOOL
            self.user32.IsWindowVisible.argtypes = [wintypes.HWND]
            self.user32.IsWindowVisible.restype = wintypes.BOOL
            self.user32.SetForegroundWindow.argtypes = [wintypes.HWND]
            self.user32.SetForegroundWindow.restype = wintypes.BOOL
            if hasattr(self.user32, "AllowSetForegroundWindow"):
                self.user32.AllowSetForegroundWindow.argtypes = [wintypes.DWORD]
                self.user32.AllowSetForegroundWindow.restype = wintypes.BOOL
            self.kernel32.GetCurrentThreadId.argtypes = []
            self.kernel32.GetCurrentThreadId.restype = wintypes.DWORD
            threading.Thread(target=self.worker_loop, daemon=True).start()
            threading.Thread(target=self.hook_loop, daemon=True).start()
            print("ℹ️ 已启用鼠标快捷键兜底。")
        except Exception as exc:
            print(f"⚠️ 鼠标快捷键兜底不可用：{exc}")

    def get_foreground_device_id(self):
        pid = get_foreground_window_pid()
        if pid is None:
            return None
        return self.manager.get_device_id_by_pid(pid)

    def enqueue_keyevent(self, device_id, keycode):
        if device_id:
            self.command_queue.put((device_id, "keyevent", keycode))

    def enqueue_notifications(self, device_id):
        if device_id:
            self.command_queue.put((device_id, "notifications", None))

    def enqueue_home_keep_screen_off(self, device_id):
        if device_id:
            self.command_queue.put((device_id, "home_keep_screen_off", None))

    def enqueue_recover_virtual_display(self, device_id):
        if device_id:
            self.command_queue.put((device_id, "recover_virtual_display", None))

    def enqueue_scrcpy_home(self, device_id, keep_screen_off=False):
        if device_id:
            command_type = "home_keep_screen_off" if keep_screen_off else "keyevent"
            payload = None if keep_screen_off else "KEYCODE_HOME"
            self.command_queue.put((device_id, command_type, payload))

    def enqueue_virtual_display_back(self, device_id, display_id):
        if device_id and display_id is not None:
            self.command_queue.put((device_id, "virtual_display_back", display_id))

    def enqueue_virtual_display_task_view(self, device_id, display_id):
        if device_id and display_id is not None:
            self.command_queue.put((device_id, "virtual_display_task_view", display_id))

    def enqueue_virtual_display_horizontal_swipe(self, device_id, display_id, direction):
        if device_id and display_id is not None:
            self.command_queue.put(
                (device_id, "virtual_display_horizontal_swipe", (display_id, direction))
            )

    def enqueue_fixed_gesture_action(self, gesture_match):
        if gesture_match:
            threading.Thread(
                target=self.execute_fixed_gesture_action,
                args=(gesture_match.action,),
                daemon=True,
            ).start()

    def handle_fixed_touchpad_gesture(self, strokes, finger_count=None):
        """Recognize sampled multi-finger strokes and enqueue the mapped host action.

        This is the fixed GestureSign-style entry point for touchpad/touch HID
        capture code: callers provide one stroke per finger, and this method
        handles shape recognition, threshold judgment, and action dispatch.
        """
        gesture_match = self.fixed_gesture_recognizer.recognize(strokes, finger_count)
        if not gesture_match:
            if GESTURE_DEBUG:
                centroid = self.fixed_gesture_recognizer.analyzer.centroid_stroke(strokes)
                direction, score = self.fixed_gesture_recognizer.classify_motion(centroid)
                print(
                    "Touch gesture missed: "
                    f"fingers={finger_count or len(strokes)}, "
                    f"points={[len(stroke) for stroke in strokes]}, "
                    f"direction={direction}, "
                    f"score={score:.1f}"
                )
            return False
        print(
            "Touch gesture matched: "
            f"{gesture_match.action.action_type} "
            f"({gesture_match.action.trigger}, score={gesture_match.score:.1f})"
        )
        self.enqueue_fixed_gesture_action(gesture_match)
        return True

    def allow_synthetic_keyboard_input(self):
        if self.keyboard_mapper is not None:
            self.keyboard_mapper.synthetic_shortcut_until = (
                time.monotonic() + SYNTHETIC_INPUT_GUARD_SECONDS
            )

    def allow_synthetic_mouse_input(self):
        self.synthetic_mouse_until = time.monotonic() + SYNTHETIC_INPUT_GUARD_SECONDS

    def press_host_key(self, key_code, extended=False):
        self.allow_synthetic_keyboard_input()
        flags = KEYEVENTF_EXTENDEDKEY if extended else 0
        if not self.send_inputs([self.keyboard_input(key_code, flags)]):
            print(f"⚠️ SendInput 按键发送失败：err={ctypes.get_last_error()}")
            return
        time.sleep(HOST_INPUT_STEP_DELAY)
        if not self.send_inputs([self.keyboard_input(key_code, flags | KEYEVENTF_KEYUP)]):
            print(f"⚠️ SendInput 按键抬起失败：err={ctypes.get_last_error()}")

    @staticmethod
    def host_key_flags(key_code):
        return KEYEVENTF_EXTENDEDKEY if key_code in EXTENDED_HOST_KEYS else 0

    @staticmethod
    def keyboard_input(key_code, flags=0):
        return INPUT(
            type=INPUT_KEYBOARD,
            ki=KEYBDINPUT(wVk=key_code, wScan=0, dwFlags=flags, time=0, dwExtraInfo=0),
        )

    @staticmethod
    def mouse_input(flags, mouse_data=0):
        return INPUT(
            type=INPUT_MOUSE,
            mi=MOUSEINPUT(dx=0, dy=0, mouseData=mouse_data, dwFlags=flags, time=0, dwExtraInfo=0),
        )

    def send_inputs(self, inputs):
        if not inputs:
            return True
        input_array = (INPUT * len(inputs))(*inputs)
        sent = self.user32.SendInput(
            len(inputs),
            input_array,
            ctypes.sizeof(INPUT),
        )
        ok = sent == len(inputs)
        if GESTURE_DEBUG:
            print(
                "SendInput result: "
                f"requested={len(inputs)}, sent={sent}, ok={ok}"
            )
        return ok

    def send_host_hotkey(self, key_codes):
        self.allow_synthetic_keyboard_input()
        for key_code in key_codes:
            if not self.send_inputs([self.keyboard_input(key_code, self.host_key_flags(key_code))]):
                print(f"⚠️ SendInput 快捷键按下失败：err={ctypes.get_last_error()}")
                return
            time.sleep(HOST_INPUT_STEP_DELAY)
        time.sleep(HOST_HOTKEY_HOLD_DELAY)
        for key_code in reversed(key_codes):
            if not self.send_inputs(
                [
                    self.keyboard_input(
                        key_code,
                        self.host_key_flags(key_code) | KEYEVENTF_KEYUP,
                    )
                ]
            ):
                print(f"⚠️ SendInput 快捷键抬起失败：err={ctypes.get_last_error()}")
                return
            time.sleep(HOST_INPUT_STEP_DELAY)

    def open_windows_task_view(self):
        try:
            subprocess.Popen(
                [
                    "explorer.exe",
                    "shell:::{3080F90E-D7AD-11D9-BD98-0000947B0257}",
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=SUBPROCESS_FLAGS,
            )
            return True
        except Exception as exc:
            print(f"⚠️ 打开 Windows 任务视图失败：{exc}")
            return False

    def click_mouse_button(self, down_flag, up_flag, mouse_data=0, count=1, position=None):
        self.allow_synthetic_mouse_input()
        if position is not None:
            self.user32.SetCursorPos(int(position[0]), int(position[1]))
            time.sleep(HOST_INPUT_STEP_DELAY)
        for _ in range(max(1, count)):
            if not self.send_inputs(
                [
                    self.mouse_input(down_flag, mouse_data),
                    self.mouse_input(up_flag, mouse_data),
                ]
            ):
                print(f"⚠️ SendInput 鼠标动作发送失败：err={ctypes.get_last_error()}")
            time.sleep(HOST_MOUSE_REPEAT_DELAY)

    def _resolve_mouse_click_position(self, payload):
        x = int(payload.get("x", 0))
        y = int(payload.get("y", 0))
        class_name = str(payload.get("class_name") or "")
        caption = str(payload.get("caption") or "")
        is_regex = bool(payload.get("is_regex", False))
        if not class_name:
            return x, y

        hwnd = self._find_window_by_class(class_name, caption, is_regex)
        if not hwnd:
            if GESTURE_DEBUG:
                print(
                    "Mouse click target window not found: "
                    f"class={class_name}, caption={caption}, regex={is_regex}"
                )
            return x, y

        if payload.get("activate", False):
            self.activate_window_class(class_name, caption, is_regex, 0)

        if payload.get("client", False):
            try:
                return win32gui.ClientToScreen(hwnd, (x, y))
            except Exception as exc:
                if GESTURE_DEBUG:
                    print(f"ClientToScreen for mouse click failed: {exc!r}")
        return x, y

    def _shell_application(self):
        return win32com.client.Dispatch("Shell.Application")

    def minimize_all_windows(self):
        try:
            self._shell_application().MinimizeAll()
            return True
        except Exception as exc:
            if GESTURE_DEBUG:
                print(f"Shell MinimizeAll failed: {exc!r}")
            return False

    def undo_minimize_all_windows(self):
        try:
            self._shell_application().UndoMinimizeALL()
            time.sleep(0.12)
            return True
        except Exception as exc:
            if GESTURE_DEBUG:
                print(f"Shell UndoMinimizeALL failed: {exc!r}")
            return False

    def _foreground_is_desktop(self):
        try:
            foreground_hwnd = win32gui.GetForegroundWindow()
            foreground_class = win32gui.GetClassName(foreground_hwnd) if foreground_hwnd else ""
            return foreground_class in {"Progman", "WorkerW"}
        except Exception:
            return False

    def activate_window_class(self, class_name, caption="", is_regex=False, timeout_ms=0):
        deadline = time.monotonic() + max(0, int(timeout_ms or 0)) / 1000.0
        while True:
            hwnd = self._find_window_by_class(class_name, caption, is_regex)
            if hwnd:
                if self._foreground_is_desktop():
                    self.undo_minimize_all_windows()
                ok = self._confirm_foreground_window(hwnd)
                if GESTURE_DEBUG:
                    print(
                        "Activate window class: "
                        f"class={class_name}, caption={caption}, "
                        f"regex={is_regex}, ok={ok}, "
                        f"target=({self._describe_window(hwnd)}), "
                        f"foreground=({self._describe_window(win32gui.GetForegroundWindow())})"
                    )
                return ok
            if timeout_ms <= 0 or time.monotonic() >= deadline:
                break
            time.sleep(0.01)

        if GESTURE_DEBUG:
            print(
                "Activate window class: "
                f"class={class_name}, caption={caption}, regex={is_regex}, not found"
            )
        return False

    def _find_window_by_class(self, class_name, caption="", is_regex=False):
        class_name = str(class_name or "")
        caption = str(caption or "")
        if not is_regex:
            hwnd = self.user32.FindWindowW(class_name or None, caption or None)
            if hwnd:
                return hwnd

        matches = []

        def enum_windows_callback(candidate_hwnd, _):
            try:
                is_visible = win32gui.IsWindowVisible(candidate_hwnd)
                candidate_class = win32gui.GetClassName(candidate_hwnd) or ""
                candidate_caption = win32gui.GetWindowText(candidate_hwnd) or ""
                if is_regex:
                    class_ok = re.search(class_name, candidate_class, re.I | re.S) is not None
                    caption_ok = re.search(caption, candidate_caption, re.I | re.S) is not None
                else:
                    class_ok = not class_name or candidate_class == class_name
                    caption_ok = not caption or candidate_caption == caption
                if class_ok and caption_ok:
                    matches.append((0 if is_visible else 1, candidate_hwnd))
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception:
            pass
        if not matches:
            return 0
        matches.sort(key=lambda item: item[0])
        return matches[0][1]

    def _describe_window(self, hwnd):
        if not hwnd:
            return "hwnd=0"
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return (
                f"hwnd={int(hwnd)}, "
                f"class={win32gui.GetClassName(hwnd)!r}, "
                f"title={win32gui.GetWindowText(hwnd)!r}, "
                f"visible={bool(self.user32.IsWindowVisible(hwnd))}, "
                f"iconic={bool(self.user32.IsIconic(hwnd))}, "
                f"rect={win32gui.GetWindowRect(hwnd)}, "
                f"pid={pid}"
            )
        except Exception as exc:
            return f"hwnd={int(hwnd)}, describe_error={exc!r}"

    def _is_window_visibly_foreground(self, hwnd):
        if not hwnd or win32gui.GetForegroundWindow() != hwnd:
            return False
        if not self.user32.IsWindowVisible(hwnd) or self.user32.IsIconic(hwnd):
            return False
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        except Exception:
            return False
        return (right - left) > 16 and (bottom - top) > 16

    def _confirm_foreground_window(self, hwnd):
        deadline = time.monotonic() + HOST_WINDOW_ACTIVATE_CONFIRM_SECONDS
        stable_since = None

        while True:
            self._force_foreground_window(hwnd)
            time.sleep(HOST_WINDOW_ACTIVATE_RETRY_DELAY)

            if self._is_window_visibly_foreground(hwnd):
                if stable_since is None:
                    stable_since = time.monotonic()
                if time.monotonic() - stable_since >= HOST_WINDOW_ACTIVATE_STABLE_SECONDS:
                    return True
            else:
                stable_since = None

            if time.monotonic() >= deadline:
                return self._is_window_visibly_foreground(hwnd)

    def _nudge_foreground_permission(self):
        self.allow_synthetic_keyboard_input()
        self.send_inputs(
            [
                self.keyboard_input(VK_MENU, 0),
                self.keyboard_input(VK_MENU, KEYEVENTF_KEYUP),
            ]
        )
        time.sleep(HOST_INPUT_STEP_DELAY)

    def _briefly_raise_topmost(self, hwnd):
        flags = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
        self.user32.SetWindowPos(
            hwnd,
            wintypes.HWND(HWND_TOPMOST),
            0,
            0,
            0,
            0,
            flags,
        )
        self.user32.SetWindowPos(
            hwnd,
            wintypes.HWND(HWND_NOTOPMOST),
            0,
            0,
            0,
            0,
            flags,
        )

    def _force_foreground_window(self, hwnd):
        if not hwnd:
            return False

        current_thread = self.kernel32.GetCurrentThreadId()
        foreground_hwnd = self.user32.GetForegroundWindow()
        foreground_thread = 0
        target_process = wintypes.DWORD(0)
        target_thread = self.user32.GetWindowThreadProcessId(
            hwnd,
            ctypes.byref(target_process),
        )
        if foreground_hwnd:
            foreground_process = wintypes.DWORD(0)
            foreground_thread = self.user32.GetWindowThreadProcessId(
                foreground_hwnd,
                ctypes.byref(foreground_process),
            )

        attached_threads = []
        for other_thread in (foreground_thread, target_thread):
            if (
                other_thread
                and current_thread
                and other_thread != current_thread
                and other_thread not in attached_threads
                and self.user32.AttachThreadInput(current_thread, other_thread, True)
            ):
                attached_threads.append(other_thread)

        try:
            if self._foreground_is_desktop():
                self.undo_minimize_all_windows()
            allow_foreground = getattr(self.user32, "AllowSetForegroundWindow", None)
            if allow_foreground is not None:
                allow_foreground(wintypes.DWORD(ASFW_ANY))
            self._nudge_foreground_permission()
            self.user32.ShowWindowAsync(hwnd, SW_RESTORE)
            self.user32.ShowWindow(hwnd, SW_RESTORE)
            time.sleep(0.01)

            if self.switch_to_this_window is not None:
                self.switch_to_this_window(hwnd, True)
            self.user32.BringWindowToTop(hwnd)
            self.user32.SetWindowPos(
                hwnd,
                wintypes.HWND(0),
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
            )
            self.user32.SetActiveWindow(hwnd)
            self.user32.SetFocus(hwnd)
            result = bool(self.user32.SetForegroundWindow(hwnd))
            if not result:
                self.user32.BringWindowToTop(hwnd)
                self.user32.SetForegroundWindow(hwnd)

            if not self._is_window_visibly_foreground(hwnd):
                self._briefly_raise_topmost(hwnd)
                self.user32.SetForegroundWindow(hwnd)

            time.sleep(0.01)
            return self._is_window_visibly_foreground(hwnd)
        finally:
            for other_thread in reversed(attached_threads):
                self.user32.AttachThreadInput(current_thread, other_thread, False)

    def _schedule_activate_window_class_post_retry(self, class_name, caption, is_regex, action_seq):
        threading.Thread(
            target=self._activate_window_class_post_retry,
            args=(class_name, caption, is_regex, action_seq),
            daemon=True,
        ).start()

    def _activate_window_class_post_retry(self, class_name, caption, is_regex, action_seq):
        start = time.monotonic()
        for delay in HOST_WINDOW_ACTIVATE_POST_RETRY_DELAYS:
            time.sleep(max(0.0, start + delay - time.monotonic()))
            if action_seq != self.fixed_gesture_action_seq:
                if GESTURE_DEBUG:
                    print("Activate window post retry: cancelled by newer gesture action")
                return
            hwnd = self._find_window_by_class(class_name, caption, is_regex)
            if not hwnd:
                if GESTURE_DEBUG:
                    print(
                        "Activate window post retry: "
                        f"class={class_name}, caption={caption}, regex={is_regex}, not found"
                    )
                continue
            if self._is_window_visibly_foreground(hwnd):
                if GESTURE_DEBUG:
                    print(
                        "Activate window post retry: already visible, "
                        f"target=({self._describe_window(hwnd)})"
                    )
                continue
            if self._foreground_is_desktop():
                self.undo_minimize_all_windows()
            ok = self._confirm_foreground_window(hwnd)
            if GESTURE_DEBUG:
                print(
                    "Activate window post retry: "
                    f"ok={ok}, target=({self._describe_window(hwnd)}), "
                    f"foreground=({self._describe_window(win32gui.GetForegroundWindow())})"
                )

    def execute_fixed_gesture_action(self, action):
        self.fixed_gesture_action_seq += 1
        action_seq = self.fixed_gesture_action_seq
        action_type = action.action_type
        payload = action.payload or {}
        if GESTURE_DEBUG:
            print(f"Fixed gesture executing: {action_type}, payload={payload}")
        if action_type == "media_play_pause":
            self.press_host_key(VK_MEDIA_PLAY_PAUSE, extended=True)
        elif action_type == "next_window":
            self.send_host_hotkey([VK_MENU, VK_TAB])
        elif action_type == "previous_window":
            self.send_host_hotkey([VK_MENU, VK_SHIFT, VK_TAB])
        elif action_type == "task_view":
            self.open_windows_task_view()
        elif action_type == "hotkey":
            key_map = {"win": VK_LWIN, "tab": VK_TAB, "d": VK_D}
            keys = [str(key).lower() for key in payload.get("keys", [])]
            if keys == ["win", "tab"]:
                self.open_windows_task_view()
                return
            if keys == ["win", "d"] and self.minimize_all_windows():
                return
            key_codes = [key_map[key] for key in keys if key in key_map]
            if key_codes:
                self.send_host_hotkey(key_codes)
        elif action_type == "mouse_right_click":
            self.click_mouse_button(MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP)
        elif action_type == "mouse_x1_click":
            self.click_mouse_button(MOUSEEVENTF_XDOWN, MOUSEEVENTF_XUP, XBUTTON1)
        elif action_type == "scrcpy_middle_click_home":
            self.execute_scrcpy_middle_click_home()
        elif action_type in {"mouse_middle_click", "mouse_middle_double_click"}:
            count = 2 if action_type == "mouse_middle_double_click" else 1
            self.click_mouse_button(
                MOUSEEVENTF_MIDDLEDOWN,
                MOUSEEVENTF_MIDDLEUP,
                count=count,
                position=self._resolve_mouse_click_position(payload),
            )
        elif action_type == "volume_down":
            repeats = max(1, int(payload.get("percent", 10)) // VOLUME_STEP_PERCENT)
            for _ in range(repeats):
                self.press_host_key(VK_VOLUME_DOWN, extended=True)
        elif action_type == "volume_up":
            repeats = max(1, int(payload.get("percent", 10)) // VOLUME_STEP_PERCENT)
            for _ in range(repeats):
                self.press_host_key(VK_VOLUME_UP, extended=True)
        elif action_type == "activate_window_class":
            class_name = str(payload.get("class_name", ""))
            caption = str(payload.get("caption", ""))
            is_regex = bool(payload.get("is_regex", False))
            self.activate_window_class(
                class_name,
                caption,
                is_regex,
                int(payload.get("timeout_ms", 0) or 0),
            )
            self._schedule_activate_window_class_post_retry(
                class_name,
                caption,
                is_regex,
                action_seq,
            )

    def execute_scrcpy_middle_click_home(self):
        hwnd, device_id = get_foreground_scrcpy_context(self.manager)
        if not device_id:
            active_devices = self.manager.get_active_scrcpy_device_ids()
            if len(active_devices) == 1:
                hwnd = None
                device_id = active_devices[0]
            else:
                if GESTURE_DEBUG:
                    print("Scrcpy middle-click home ignored: no foreground scrcpy window")
                return False

        display_id = self.manager.get_virtual_display_id(device_id)
        if display_id is not None:
            self.enqueue_recover_virtual_display(device_id)
            return True

        keep_screen_off = bool(hwnd and is_window_mostly_black(hwnd))
        self.enqueue_scrcpy_home(device_id, keep_screen_off=keep_screen_off)
        return True

    def _send_scrcpy_o_shortcut(self, shift=False, device_id=None, attempts=2):
        if device_id:
            activated = self.activate_window_class("SDL_app", f"Superdex {device_id}", False, 300)
            if not activated and GESTURE_DEBUG:
                print(f"scrcpy shortcut target not focused for device: {device_id}")
            time.sleep(0.08)

        for _ in range(max(1, int(attempts or 1))):
            self.user32.keybd_event(VK_RCONTROL, 0, KEYEVENTF_EXTENDEDKEY, None)
            if shift:
                self.user32.keybd_event(VK_SHIFT, 0, 0, None)
            try:
                time.sleep(0.04)
                self.user32.keybd_event(VK_O, 0, 0, None)
                time.sleep(0.04)
                self.user32.keybd_event(VK_O, 0, KEYEVENTF_KEYUP, None)
                time.sleep(0.02)
            finally:
                if shift:
                    self.user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, None)
                self.user32.keybd_event(
                    VK_RCONTROL,
                    0,
                    KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP,
                    None,
                )
            time.sleep(0.12)

    def send_scrcpy_screen_off_shortcut(self):
        self._send_scrcpy_o_shortcut(shift=False, attempts=1)

    def send_scrcpy_screen_on_shortcut(self, device_id=None):
        self._send_scrcpy_o_shortcut(shift=True, device_id=device_id, attempts=3)

    def get_virtual_edge_metrics(self, width, height):
        shorter = max(1, min(width, height))
        return {
            "side_hot": max(
                VIRTUAL_EDGE_SIDE_HOT_MIN,
                min(VIRTUAL_EDGE_SIDE_HOT_MAX, int(width * VIRTUAL_EDGE_SIDE_HOT_RATIO)),
            ),
            "bottom_hot": max(
                VIRTUAL_EDGE_BOTTOM_HOT_MIN,
                min(VIRTUAL_EDGE_BOTTOM_HOT_MAX, int(height * VIRTUAL_EDGE_BOTTOM_HOT_RATIO)),
            ),
            "swipe": max(
                VIRTUAL_EDGE_SWIPE_MIN,
                min(VIRTUAL_EDGE_SWIPE_MAX, int(shorter * VIRTUAL_EDGE_SWIPE_RATIO)),
            ),
            "horizontal": max(
                VIRTUAL_EDGE_SWIPE_MIN,
                min(VIRTUAL_EDGE_SWIPE_MAX, int(width * VIRTUAL_EDGE_HORIZONTAL_RATIO)),
            ),
        }

    def classify_virtual_edge(self, client_x, client_y, width, height):
        metrics = self.get_virtual_edge_metrics(width, height)
        if client_y >= height - metrics["bottom_hot"]:
            return "bottom"
        if client_x <= metrics["side_hot"]:
            return "left"
        if client_x >= width - metrics["side_hot"]:
            return "right"
        return None

    def begin_virtual_edge_gesture(self, hwnd, device_id, display_id, mouse_struct):
        point = get_client_point(hwnd, mouse_struct.pt.x, mouse_struct.pt.y)
        if point is None:
            return False
        client_x, client_y, width, height = point
        edge = self.classify_virtual_edge(client_x, client_y, width, height)
        if edge is None:
            return False
        if edge == "bottom":
            return False
        if edge in ("left", "right") and not VIRTUAL_EDGE_ENABLE_LEFT_BUTTON_SIDE_GESTURES:
            return False

        with self.edge_gesture_lock:
            self.cancel_virtual_edge_hold_locked()
            self.edge_gesture = {
                "hwnd": hwnd,
                "device_id": device_id,
                "display_id": display_id,
                "edge": edge,
                "down_x": client_x,
                "down_y": client_y,
                "width": width,
                "height": height,
                "armed": False,
                "triggered": False,
                "timer": None,
            }
        return True

    def handle_virtual_edge_gesture_event(self, w_param, mouse_struct):
        with self.edge_gesture_lock:
            gesture = self.edge_gesture
            if not gesture:
                return False

            point = get_client_point(gesture["hwnd"], mouse_struct.pt.x, mouse_struct.pt.y)
            if point is None:
                client_x, client_y = screen_to_client_point(
                    gesture["hwnd"], mouse_struct.pt.x, mouse_struct.pt.y
                )
            else:
                client_x, client_y, _, _ = point

            dx = client_x - gesture["down_x"]
            dy = client_y - gesture["down_y"]
            metrics = self.get_virtual_edge_metrics(gesture["width"], gesture["height"])

            if w_param == WM_MOUSEMOVE:
                self.update_virtual_edge_gesture_locked(gesture, dx, dy, metrics)
                return True

            if w_param == WM_LBUTTONUP:
                self.finish_virtual_edge_gesture_locked(gesture, dx, dy, metrics)
                self.cancel_virtual_edge_hold_locked()
                self.edge_gesture = None
                return True

            return True

    def update_virtual_edge_gesture_locked(self, gesture, dx, dy, metrics):
        if gesture["triggered"]:
            return
        edge = gesture["edge"]
        abs_dx = abs(dx)
        abs_dy = abs(dy)

        if edge == "left" and dx >= metrics["swipe"] and abs_dx >= max(12, abs_dy * 0.55):
            gesture["triggered"] = True
            self.enqueue_virtual_display_back(gesture["device_id"], gesture["display_id"])
            return

        if edge == "right" and dx <= -metrics["swipe"] and abs_dx >= max(12, abs_dy * 0.55):
            gesture["triggered"] = True
            self.enqueue_virtual_display_back(gesture["device_id"], gesture["display_id"])
            return

        if edge != "bottom":
            return

        if abs_dx >= metrics["horizontal"] and abs_dx > abs_dy * 1.2:
            gesture["triggered"] = True
            direction = "right" if dx > 0 else "left"
            self.enqueue_virtual_display_horizontal_swipe(
                gesture["device_id"], gesture["display_id"], direction
            )
            return

        if dy <= -metrics["swipe"] and abs_dy >= max(12, abs_dx * 0.75) and not gesture["armed"]:
            gesture["armed"] = True
            timer = threading.Timer(
                VIRTUAL_EDGE_HOLD_DELAY,
                self.trigger_virtual_edge_bottom_hold,
                args=(gesture["device_id"], gesture["display_id"]),
            )
            timer.daemon = True
            gesture["timer"] = timer
            timer.start()

    def finish_virtual_edge_gesture_locked(self, gesture, dx, dy, metrics):
        if gesture["triggered"]:
            return
        edge = gesture["edge"]
        abs_dx = abs(dx)
        abs_dy = abs(dy)

        if edge == "left" and dx >= metrics["swipe"] and abs_dx >= max(12, abs_dy * 0.55):
            gesture["triggered"] = True
            self.enqueue_virtual_display_back(gesture["device_id"], gesture["display_id"])
            return

        if edge == "right" and dx <= -metrics["swipe"] and abs_dx >= max(12, abs_dy * 0.55):
            gesture["triggered"] = True
            self.enqueue_virtual_display_back(gesture["device_id"], gesture["display_id"])
            return

        if edge == "bottom" and dy <= -metrics["swipe"] and abs_dy >= max(12, abs_dx * 0.75):
            gesture["triggered"] = True
            self.enqueue_recover_virtual_display(gesture["device_id"])
            return

    def trigger_virtual_edge_bottom_hold(self, device_id, display_id):
        with self.edge_gesture_lock:
            gesture = self.edge_gesture
            if (
                not gesture
                or gesture["device_id"] != device_id
                or gesture["display_id"] != display_id
                or gesture["edge"] != "bottom"
                or gesture["triggered"]
            ):
                return
            gesture["triggered"] = True
        self.enqueue_virtual_display_task_view(device_id, display_id)

    def cancel_virtual_edge_hold_locked(self):
        gesture = self.edge_gesture
        if not gesture:
            return
        timer = gesture.get("timer")
        if timer:
            timer.cancel()
            gesture["timer"] = None

    def handle_mouse_event(self, w_param, l_param):
        if w_param not in (
            WM_MOUSEMOVE,
            WM_LBUTTONDOWN,
            WM_LBUTTONUP,
            WM_RBUTTONDOWN,
            WM_MBUTTONDOWN,
            WM_XBUTTONDOWN,
        ):
            return False

        if time.monotonic() < self.synthetic_mouse_until:
            return False

        mouse_struct = ctypes.cast(l_param, ctypes.POINTER(MSLLHOOKSTRUCT)).contents

        if w_param in (WM_MOUSEMOVE, WM_LBUTTONUP):
            if self.handle_virtual_edge_gesture_event(w_param, mouse_struct):
                return True
            if w_param == WM_MOUSEMOVE:
                return False

        hwnd, device_id = get_foreground_scrcpy_context(self.manager)
        if not device_id:
            return False

        display_id = self.manager.get_virtual_display_id(device_id)

        if w_param == WM_LBUTTONDOWN:
            if display_id is not None:
                return self.begin_virtual_edge_gesture(hwnd, device_id, display_id, mouse_struct)
            return False

        if w_param == WM_RBUTTONDOWN:
            self.enqueue_keyevent(device_id, "KEYCODE_BACK")
            return True

        if w_param == WM_MBUTTONDOWN:
            if display_id is not None:
                self.enqueue_recover_virtual_display(device_id)
                return True
            if not is_window_mostly_black(hwnd):
                return False
            self.enqueue_home_keep_screen_off(device_id)
            return True

        if w_param == WM_XBUTTONDOWN:
            xbutton = (mouse_struct.mouseData >> 16) & 0xFFFF
            if xbutton == XBUTTON1:
                self.enqueue_keyevent(device_id, "KEYCODE_APP_SWITCH")
                return True
            if xbutton == XBUTTON2:
                self.enqueue_notifications(device_id)
                return True
        return False

    def hook_loop(self):
        LOW_LEVEL_MOUSE_PROC = ctypes.WINFUNCTYPE(
            wintypes.LPARAM, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM
        )

        def callback(n_code, w_param, l_param):
            if n_code == HC_ACTION and self.handle_mouse_event(w_param, l_param):
                return 1
            return self.user32.CallNextHookEx(self.hook_handle, n_code, w_param, l_param)

        self.callback_ref = LOW_LEVEL_MOUSE_PROC(callback)
        self.hook_handle = self.user32.SetWindowsHookExW(
            WH_MOUSE_LL, self.callback_ref, None, 0
        )
        if not self.hook_handle:
            raise ctypes.WinError(ctypes.get_last_error())

        msg = wintypes.MSG()
        while self.user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
            self.user32.TranslateMessage(ctypes.byref(msg))
            self.user32.DispatchMessageW(ctypes.byref(msg))

    def worker_loop(self):
        while True:
            device_id, command_type, payload = self.command_queue.get()
            if command_type == "fixed_gesture_action":
                self.execute_fixed_gesture_action(payload)
                continue
            if not is_device_ready(device_id):
                continue
            if command_type == "keyevent":
                adb(["-s", device_id, "shell", "input", "keyevent", payload], capture_output=False)
            elif command_type == "notifications":
                adb(
                    ["-s", device_id, "shell", "cmd", "statusbar", "expand-notifications"],
                    capture_output=False,
                )
            elif command_type == "home_keep_screen_off":
                start_scrcpy_screen_off_helper(device_id)
                time.sleep(0.03)
                send_home_with_screen_off_guard(device_id)
                time.sleep(SCREEN_OFF_FALLBACK_DELAY)
                force_physical_screen_off(device_id)
                self.send_scrcpy_screen_off_shortcut()
            elif command_type == "recover_virtual_display":
                recover_virtual_display_activity(device_id, self.manager)
            elif command_type == "virtual_display_back":
                adb_input_keyevent_on_display(device_id, payload, "KEYCODE_BACK")
            elif command_type == "virtual_display_task_view":
                show_superdex_task_view_on_display(device_id, payload)
            elif command_type == "virtual_display_horizontal_swipe":
                display_id, direction = payload
                print(
                    f"Virtual display horizontal recent-app swipe reserved: "
                    f"{direction} ({device_id}, display {display_id})"
                )


class DeviceManager:
    def __init__(self):
        self.lock = threading.RLock()
        self.states = {}
        self.scrcpy_procs = {}
        self.pid_to_device = {}
        self.keyboard_modes = {}
        self.watchdog_procs = {}
        self.wifi_ids = {}
        self.pre_superdex_apps = {}
        self.last_states = {}
        self.last_orientations = {}
        self.virtual_display_ids = {}
        self.scrcpy_virtual_display_modes = {}
        self.restore_locks = {}
        self.starting = set()
        self.physical_serials = {}
        self.starting_physical_serials = set()

    def get_restore_lock(self, device_id):
        with self.lock:
            lock = self.restore_locks.get(device_id)
            if lock is None:
                lock = threading.Lock()
                self.restore_locks[device_id] = lock
            return lock

    def get_state(self, device_id):
        with self.lock:
            return self.states.get(device_id)

    def set_state(self, device_id, state):
        with self.lock:
            self.states[device_id] = state

    def get_wifi_id(self, device_id):
        with self.lock:
            return self.wifi_ids.get(device_id)

    def set_wifi_id(self, device_id, wifi_id):
        with self.lock:
            self.wifi_ids[device_id] = wifi_id

    def pop_wifi_id(self, device_id):
        with self.lock:
            return self.wifi_ids.pop(device_id, None)

    def get_active_usb_for_wifi_id(self, wifi_id):
        with self.lock:
            for usb_id, known_wifi_id in self.wifi_ids.items():
                if known_wifi_id != wifi_id:
                    continue
                proc = self.scrcpy_procs.get(usb_id)
                if usb_id in self.starting or is_process_running(proc):
                    return usb_id
        return None

    def set_pre_superdex_apps(self, device_id, packages):
        with self.lock:
            self.pre_superdex_apps[device_id] = set(packages or [])

    def pop_pre_superdex_apps(self, device_id):
        with self.lock:
            return self.pre_superdex_apps.pop(device_id, None)

    def set_virtual_display_id(self, device_id, display_id):
        with self.lock:
            if display_id is None:
                self.virtual_display_ids.pop(device_id, None)
            else:
                self.virtual_display_ids[device_id] = display_id

    def get_virtual_display_id(self, device_id):
        with self.lock:
            return self.virtual_display_ids.get(device_id)

    def pop_virtual_display_id(self, device_id):
        with self.lock:
            return self.virtual_display_ids.pop(device_id, None)

    def set_scrcpy_virtual_display_mode(self, device_id, use_virtual_display):
        with self.lock:
            self.scrcpy_virtual_display_modes[device_id] = bool(use_virtual_display)

    def get_scrcpy_virtual_display_mode(self, device_id):
        with self.lock:
            if device_id in self.scrcpy_virtual_display_modes:
                return self.scrcpy_virtual_display_modes[device_id]
        return USE_SCRCPY_VIRTUAL_DISPLAY

    def pop_scrcpy_virtual_display_mode(self, device_id):
        with self.lock:
            return self.scrcpy_virtual_display_modes.pop(device_id, None)

    def get_scrcpy_proc(self, device_id):
        with self.lock:
            return self.scrcpy_procs.get(device_id)

    def get_active_scrcpy_device_ids(self):
        with self.lock:
            return [
                device_id
                for device_id, proc in self.scrcpy_procs.items()
                if is_process_running(proc)
            ]

    def is_scrcpy_proc_current(self, device_id, proc):
        with self.lock:
            return proc is not None and self.scrcpy_procs.get(device_id) is proc

    def set_scrcpy_proc(self, device_id, proc):
        with self.lock:
            old_proc = self.scrcpy_procs.get(device_id)
            if old_proc is not None:
                self.pid_to_device.pop(old_proc.pid, None)
            if proc is None:
                self.scrcpy_procs.pop(device_id, None)
            else:
                self.scrcpy_procs[device_id] = proc
                self.pid_to_device[proc.pid] = device_id
            self.starting.discard(device_id)
            physical_serial = self.physical_serials.get(device_id)
            if physical_serial:
                self.starting_physical_serials.discard(physical_serial)

    def pop_scrcpy_proc(self, device_id):
        with self.lock:
            proc = self.scrcpy_procs.pop(device_id, None)
            if proc is not None:
                self.pid_to_device.pop(proc.pid, None)
            return proc

    def get_device_id_by_pid(self, pid):
        with self.lock:
            return self.pid_to_device.get(pid)

    def set_keyboard_mode(self, device_id, keyboard_mode):
        with self.lock:
            self.keyboard_modes[device_id] = keyboard_mode

    def get_keyboard_mode(self, device_id):
        with self.lock:
            return self.keyboard_modes.get(device_id)

    def get_watchdog_proc(self, device_id):
        with self.lock:
            return self.watchdog_procs.get(device_id)

    def set_watchdog_proc(self, device_id, proc):
        with self.lock:
            self.watchdog_procs[device_id] = proc

    def pop_watchdog_proc(self, device_id):
        with self.lock:
            return self.watchdog_procs.pop(device_id, None)

    def get_last_state(self, device_id, default=False):
        with self.lock:
            return self.last_states.get(device_id, default)

    def set_last_state(self, device_id, connected):
        with self.lock:
            self.last_states[device_id] = connected

    def known_devices(self):
        with self.lock:
            return list(self.last_states.keys())

    def get_last_orientation(self, device_id):
        with self.lock:
            return self.last_orientations.get(device_id)

    def set_last_orientation(self, device_id, orientation):
        with self.lock:
            self.last_orientations[device_id] = orientation

    def pop_last_orientation(self, device_id):
        with self.lock:
            self.last_orientations.pop(device_id, None)

    def set_physical_serial(self, device_id, physical_serial):
        with self.lock:
            if physical_serial:
                self.physical_serials[device_id] = physical_serial
            else:
                self.physical_serials.pop(device_id, None)

    def pop_physical_serial(self, device_id):
        with self.lock:
            physical_serial = self.physical_serials.pop(device_id, None)
            if physical_serial:
                self.starting_physical_serials.discard(physical_serial)
            return physical_serial

    def _active_device_for_physical_serial_unlocked(self, physical_serial, exclude_device_id=None):
        if not physical_serial:
            return None
        for known_device_id, known_physical_serial in self.physical_serials.items():
            if known_device_id == exclude_device_id:
                continue
            if known_physical_serial != physical_serial:
                continue
            proc = self.scrcpy_procs.get(known_device_id)
            if known_device_id in self.starting or is_process_running(proc):
                return known_device_id
        return None

    def get_active_device_for_physical_serial(self, physical_serial, exclude_device_id=None):
        with self.lock:
            return self._active_device_for_physical_serial_unlocked(
                physical_serial,
                exclude_device_id=exclude_device_id,
            )

    def begin_start(self, device_id, physical_serial=None):
        with self.lock:
            proc = self.scrcpy_procs.get(device_id)
            if device_id in self.starting or is_process_running(proc):
                return False
            if physical_serial:
                active_device_id = self._active_device_for_physical_serial_unlocked(
                    physical_serial,
                    exclude_device_id=device_id,
                )
                if active_device_id or physical_serial in self.starting_physical_serials:
                    return False
                self.physical_serials[device_id] = physical_serial
                self.starting_physical_serials.add(physical_serial)
            self.starting.add(device_id)
            return True

    def clear_starting(self, device_id):
        with self.lock:
            self.starting.discard(device_id)
            physical_serial = self.physical_serials.get(device_id)
            if physical_serial:
                self.starting_physical_serials.discard(physical_serial)


class DeviceScanThread(QtCore.QThread):
    scanned = QtCore.pyqtSignal(list)
    failed = QtCore.pyqtSignal(str)

    def run(self):
        try:
            devices = []
            for device_id in list_connected_devices():
                model = adb(
                    ["-s", device_id, "shell", "getprop", "ro.product.model"]
                ).strip()
                label = f"{device_id} ({model})" if model else device_id
                devices.append((device_id, label))
            self.scanned.emit(devices)
        except Exception as exc:
            self.failed.emit(str(exc))

class SuperdexGUI(QtWidgets.QWidget):
    screen_off_finished = QtCore.pyqtSignal(str)
    screen_on_finished = QtCore.pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setAttribute(QtCore.Qt.WA_NativeWindow, True)
        self.setWindowTitle("Superdex 设备管理器")
        self.setObjectName("rootWindow")
        self.resize(820, 860)
        self.setMinimumSize(720, 760)
        self.tray_icon = None
        self.tray_notice_shown = False
        self.exit_requested = False
        self.app_icon = self.build_app_icon()
        self.setWindowIcon(self.app_icon)

        self.manager = DeviceManager()
        self.keyboard_mapper = KeyboardMapper(self.manager)
        self.mouse_shortcut_mapper = MouseShortcutMapper(self.manager, self.keyboard_mapper)
        self.touch_gesture_bridge = TouchGestureBridge(
            self.mouse_shortcut_mapper.handle_fixed_touchpad_gesture
        )
        self.selection_lock = threading.Lock()
        self.selected_devices = set()
        self.refresh_thread = None

        self.device_list = QtWidgets.QListWidget()
        self.device_list.setObjectName("deviceList")
        self.device_list.setAlternatingRowColors(False)
        self.device_list.setSpacing(10)
        self.device_list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.device_list.setUniformItemSizes(True)
        self.display_mode_label = QtWidgets.QLabel("连接模式")
        self.display_mode_label.setObjectName("fieldLabel")
        self.display_mode_combo = QtWidgets.QComboBox()
        self.display_mode_combo.setObjectName("modeCombo")
        self.display_mode_combo.addItem("扩展模式（默认）", SCRCPY_MODE_VIRTUAL)
        self.display_mode_combo.addItem("镜像模式", SCRCPY_MODE_MIRROR)
        self.refresh_btn = QtWidgets.QPushButton("刷新设备")
        self.autostart_btn = QtWidgets.QPushButton("开启开机自启")
        self.screen_off_btn = QtWidgets.QPushButton("仅熄灭手机屏幕")
        self.screen_on_btn = QtWidgets.QPushButton("仅打开手机屏幕")
        self.connect_btn = QtWidgets.QPushButton("连接选中设备")
        self.connect_btn.setObjectName("primaryButton")
        for button in (
            self.refresh_btn,
            self.autostart_btn,
            self.screen_off_btn,
            self.screen_on_btn,
            self.connect_btn,
        ):
            button.setCursor(QtCore.Qt.PointingHandCursor)
        self.apply_modern_style()
        for widget in (
            self,
            self.device_list,
            self.display_mode_label,
            self.display_mode_combo,
            self.refresh_btn,
            self.autostart_btn,
            self.screen_off_btn,
            self.screen_on_btn,
            self.connect_btn,
        ):
            widget.setAttribute(QtCore.Qt.WA_AcceptTouchEvents, True)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)
        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.setFont(QtGui.QFont("Microsoft YaHei UI", 12))
            app.installEventFilter(self)

        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(32, 32, 32, 28)
        layout.setSpacing(22)

        header_layout = QtWidgets.QHBoxLayout()
        header_layout.setSpacing(18)
        icon_label = QtWidgets.QLabel()
        icon_label.setObjectName("appIcon")
        icon_label.setFixedSize(72, 72)
        icon_label.setAlignment(QtCore.Qt.AlignCenter)
        icon_label.setPixmap(self.app_icon.pixmap(50, 50))
        title_layout = QtWidgets.QVBoxLayout()
        title_layout.setSpacing(5)
        title = QtWidgets.QLabel("Superdex 设备管理器")
        title.setObjectName("titleLabel")
        subtitle = QtWidgets.QLabel("连接设备、管理连接模式，并快速控制手机屏幕")
        subtitle.setWordWrap(True)
        subtitle.setObjectName("subtitleLabel")
        title_layout.addWidget(title)
        title_layout.addWidget(subtitle)
        status_pill = QtWidgets.QLabel("● 设备管理")
        status_pill.setObjectName("statusPill")
        status_pill.setAlignment(QtCore.Qt.AlignCenter)
        header_layout.addWidget(icon_label)
        header_layout.addLayout(title_layout, 1)
        header_layout.addWidget(status_pill, 0, QtCore.Qt.AlignTop)
        layout.addLayout(header_layout)

        device_card = QtWidgets.QFrame()
        device_card.setObjectName("card")
        device_card_layout = QtWidgets.QVBoxLayout(device_card)
        device_card_layout.setContentsMargins(24, 24, 24, 24)
        device_card_layout.setSpacing(16)
        device_title = QtWidgets.QLabel("已连接设备")
        device_title.setObjectName("sectionTitle")
        device_card_layout.addWidget(device_title)
        device_card_layout.addWidget(self.device_list)
        layout.addWidget(device_card, 1)

        mode_card = QtWidgets.QFrame()
        mode_card.setObjectName("card")
        mode_layout = QtWidgets.QHBoxLayout(mode_card)
        mode_layout.setContentsMargins(24, 22, 24, 22)
        mode_layout.setSpacing(18)
        mode_layout.addWidget(self.display_mode_label)
        mode_layout.addWidget(self.display_mode_combo, 1)
        layout.addWidget(mode_card)

        secondary_actions = QtWidgets.QGridLayout()
        secondary_actions.setHorizontalSpacing(16)
        secondary_actions.setVerticalSpacing(16)
        secondary_actions.addWidget(self.refresh_btn, 0, 0)
        secondary_actions.addWidget(self.autostart_btn, 0, 1)
        secondary_actions.addWidget(self.screen_off_btn, 1, 0)
        secondary_actions.addWidget(self.screen_on_btn, 1, 1)
        layout.addLayout(secondary_actions)
        layout.addWidget(self.connect_btn)
        self.setLayout(layout)

        self.refresh_btn.clicked.connect(self.refresh_devices)
        self.display_mode_combo.currentIndexChanged.connect(self.on_display_mode_changed)
        self.autostart_btn.clicked.connect(self.toggle_autostart)
        self.screen_off_btn.clicked.connect(self.turn_selected_screens_off)
        self.screen_on_btn.clicked.connect(self.turn_selected_screens_on)
        self.connect_btn.clicked.connect(self.connect_selected_devices)
        self.device_list.itemSelectionChanged.connect(self.update_selected_devices)
        self.screen_off_finished.connect(self.on_screen_off_finished)
        self.screen_on_finished.connect(self.on_screen_on_finished)
        self.update_display_mode_combo()
        self.update_autostart_button()

        threading.Thread(
            target=main,
            args=(self.manager, self.get_selected_devices),
            daemon=True,
        ).start()
        self.setup_tray_icon()
        QtCore.QTimer.singleShot(0, self.refresh_devices)

    def apply_modern_style(self):
        self.setStyleSheet("""
            QWidget {
                background: #f3f7fb;
                color: #0f172a;
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
                font-size: 19px;
            }
            QWidget#rootWindow {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f8fbff, stop:0.52 #f4f8fd, stop:1 #eef7f6);
            }
            QLabel#titleLabel {
                color: #0f172a;
                font-size: 34px;
                font-weight: 700;
                background: transparent;
            }
            QLabel#subtitleLabel {
                color: #5b6b84;
                font-size: 17px;
                background: transparent;
            }
            QLabel#statusPill {
                background: #ecfdf5;
                border: 1px solid #bbf7d0;
                border-radius: 16px;
                color: #15803d;
                font-size: 15px;
                font-weight: 700;
                padding: 7px 12px;
            }
            QLabel#sectionTitle, QLabel#fieldLabel {
                color: #334155;
                font-size: 18px;
                font-weight: 700;
                background: transparent;
            }
            QLabel#appIcon {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #dff7ff, stop:1 #e7f8ef);
                border-radius: 22px;
                padding: 11px;
            }
            QFrame#card {
                background: rgba(255, 255, 255, 236);
                border: 1px solid #d9e3ef;
                border-radius: 24px;
            }
            QListWidget#deviceList {
                background: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 22px;
                padding: 11px;
                outline: none;
            }
            QListWidget#deviceList::item {
                min-height: 64px;
                padding: 12px 20px;
                border-radius: 16px;
                color: #1e293b;
            }
            QListWidget#deviceList::item:hover {
                background: #e0f2fe;
                color: #075985;
            }
            QListWidget#deviceList::item:selected {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #22c55e, stop:1 #06b6d4);
                color: #ffffff;
            }
            QScrollBar:vertical {
                background: transparent;
                border: 0;
                margin: 8px 4px 8px 0;
                width: 12px;
            }
            QScrollBar::handle:vertical {
                background: #cbd5e1;
                border-radius: 6px;
                min-height: 48px;
            }
            QScrollBar::handle:vertical:hover {
                background: #94a3b8;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
            }
            QComboBox {
                background: #ffffff;
                border: 1px solid #c8d6e6;
                border-radius: 18px;
                padding: 15px 20px;
                min-height: 34px;
            }
            QComboBox:hover, QComboBox:focus {
                border-color: #22c55e;
            }
            QComboBox::drop-down {
                border: 0;
                width: 44px;
            }
            QComboBox QAbstractItemView {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 12px;
                padding: 6px;
                selection-background-color: #dcfce7;
                selection-color: #166534;
            }
            QPushButton {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 18px;
                color: #334155;
                font-weight: 600;
                padding: 17px 22px;
                min-height: 34px;
            }
            QPushButton:hover {
                background: #ecfeff;
                border-color: #06b6d4;
                color: #0e7490;
            }
            QPushButton:pressed {
                background: #cffafe;
            }
            QPushButton:disabled {
                background: #e2e8f0;
                border-color: #cbd5e1;
                color: #94a3b8;
            }
            QPushButton#primaryButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #22c55e, stop:1 #06b6d4);
                border: 0;
                color: white;
                font-size: 20px;
                padding: 19px 24px;
            }
            QPushButton#primaryButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #16a34a, stop:1 #0891b2);
                color: white;
            }
        """)

    def eventFilter(self, obj, event):
        if event.type() in (
            QtCore.QEvent.TouchBegin,
            QtCore.QEvent.TouchUpdate,
            QtCore.QEvent.TouchEnd,
            QtCore.QEvent.TouchCancel,
        ):
            if self.touch_gesture_bridge.handle_qt_touch_event(event):
                return True
        return super().eventFilter(obj, event)

    def showEvent(self, event):
        super().showEvent(event)
        self.touch_gesture_bridge.register_window(self.winId())

    @staticmethod
    def native_msg_field(msg, *names):
        for name in names:
            if hasattr(msg, name):
                return getattr(msg, name)
        return None

    def nativeEvent(self, eventType, message):
        if eventType in (b"windows_generic_MSG", b"windows_dispatcher_MSG"):
            msg = ctypes.cast(int(message), ctypes.POINTER(wintypes.MSG)).contents
            hwnd = self.native_msg_field(msg, "hwnd", "hWnd")
            message_id = self.native_msg_field(msg, "message")
            w_param = self.native_msg_field(msg, "wParam", "wparam")
            l_param = self.native_msg_field(msg, "lParam", "lparam")
            if self.touch_gesture_bridge.handle_native_pointer_message(
                hwnd,
                message_id,
                w_param,
                l_param,
            ):
                return True, 0
            if self.touch_gesture_bridge.handle_native_touch_message(
                hwnd,
                message_id,
                w_param,
                l_param,
            ):
                return True, 0
        return super().nativeEvent(eventType, message)

    def build_app_icon(self):
        if os.path.exists(APP_ICON_PATH):
            icon = QtGui.QIcon(APP_ICON_PATH)
            if not icon.isNull():
                return icon
        return self.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon)

    def setup_tray_icon(self):
        if not QtWidgets.QSystemTrayIcon.isSystemTrayAvailable():
            return

        tray_menu = QtWidgets.QMenu(self)
        show_action = tray_menu.addAction("显示主窗口")
        refresh_action = tray_menu.addAction("刷新设备")
        tray_menu.addSeparator()
        quit_action = tray_menu.addAction("退出应用")

        show_action.triggered.connect(self.show_from_tray)
        refresh_action.triggered.connect(self.refresh_from_tray)
        quit_action.triggered.connect(self.quit_from_tray)

        self.tray_icon = QtWidgets.QSystemTrayIcon(self.app_icon, self)
        self.tray_icon.setToolTip("Superdex 设备管理器")
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()

    def refresh_from_tray(self):
        self.show_from_tray()
        self.refresh_devices()

    def show_from_tray(self):
        if self.isMinimized():
            self.setWindowState(self.windowState() & ~QtCore.Qt.WindowMinimized)
        self.show()
        self.raise_()
        self.activateWindow()

    def hide_to_tray(self):
        self.hide()
        if self.tray_icon and not self.tray_notice_shown:
            self.tray_icon.showMessage(
                "Superdex",
                "Superdex 已隐藏到系统托盘，双击托盘图标可重新打开。",
                QtWidgets.QSystemTrayIcon.Information,
                3000,
            )
            self.tray_notice_shown = True

    def quit_from_tray(self):
        self.exit_requested = True
        if self.tray_icon:
            self.tray_icon.hide()
        QtWidgets.QApplication.quit()

    def on_tray_activated(self, reason):
        if reason in (
            QtWidgets.QSystemTrayIcon.Trigger,
            QtWidgets.QSystemTrayIcon.DoubleClick,
        ):
            self.show_from_tray()

    def get_selected_devices(self):
        with self.selection_lock:
            return set(self.selected_devices)

    def refresh_devices(self):
        if self.refresh_thread and self.refresh_thread.isRunning():
            return

        self.refresh_btn.setEnabled(False)
        self.refresh_thread = DeviceScanThread()
        self.refresh_thread.scanned.connect(self.on_devices_scanned)
        self.refresh_thread.failed.connect(self.on_device_scan_failed)
        self.refresh_thread.finished.connect(self.on_refresh_finished)
        self.refresh_thread.start()

    def on_devices_scanned(self, devices):
        selected = self.get_selected_devices()
        self.device_list.clear()
        for device_id, label in devices:
            item = QtWidgets.QListWidgetItem(label)
            self.device_list.addItem(item)
            if device_id in selected:
                item.setSelected(True)
        self.update_selected_devices()

    def on_device_scan_failed(self, message):
        QtWidgets.QMessageBox.warning(self, "提示", f"设备扫描失败：{message}")

    def on_refresh_finished(self):
        self.refresh_btn.setEnabled(True)
        if self.refresh_thread:
            self.refresh_thread.deleteLater()
            self.refresh_thread = None

    def update_selected_devices(self):
        selected = {
            item.text().split(" ")[0] for item in self.device_list.selectedItems()
        }
        with self.selection_lock:
            self.selected_devices = selected

    def update_display_mode_combo(self):
        mode = get_scrcpy_display_mode()
        index = self.display_mode_combo.findData(mode)
        if index < 0:
            index = self.display_mode_combo.findData(SCRCPY_MODE_DEFAULT)
        self.display_mode_combo.blockSignals(True)
        self.display_mode_combo.setCurrentIndex(max(0, index))
        self.display_mode_combo.blockSignals(False)

    def on_display_mode_changed(self, index):
        mode = self.display_mode_combo.itemData(index)
        mode = set_scrcpy_display_mode(mode)
        label = "扩展模式" if mode == SCRCPY_MODE_VIRTUAL else "镜像模式"
        print(f"ℹ️ 已切换连接模式：{label}。后续连接和开机自启将使用该模式。")

    def get_autostart_shortcut_path(self):
        appdata = os.getenv("APPDATA")
        if not appdata:
            return None
        return os.path.join(
            appdata,
            "Microsoft\\Windows\\Start Menu\\Programs\\Startup",
            "Superdex.lnk",
        )

    def is_autostart_enabled(self):
        shortcut = self.get_autostart_shortcut_path()
        return bool(shortcut and os.path.exists(shortcut))

    def update_autostart_button(self):
        if self.is_autostart_enabled():
            self.autostart_btn.setText("关闭开机自启")
        else:
            self.autostart_btn.setText("开启开机自启")

    def toggle_autostart(self):
        if self.is_autostart_enabled():
            self.disable_autostart()
        else:
            self.enable_autostart()

    def enable_autostart(self):
        shortcut = self.get_autostart_shortcut_path()
        if not shortcut:
            QtWidgets.QMessageBox.warning(self, "提示", "无法获取开机启动目录。")
            return

        exe_path = os.path.abspath(sys.argv[0])
        startup_dir = os.path.dirname(shortcut)
        try:
            os.makedirs(startup_dir, exist_ok=True)
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut_obj = shell.CreateShortcut(shortcut)
            shortcut_obj.TargetPath = exe_path
            shortcut_obj.WorkingDirectory = os.path.dirname(exe_path)
            shortcut_obj.save()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "提示", f"开启开机自启失败：{exc}")
            return

        self.update_autostart_button()
        QtWidgets.QMessageBox.information(self, "提示", "已开启开机自启！")

    def disable_autostart(self):
        shortcut = self.get_autostart_shortcut_path()
        if not shortcut:
            QtWidgets.QMessageBox.warning(self, "提示", "无法获取开机启动目录。")
            return

        try:
            if os.path.exists(shortcut):
                os.remove(shortcut)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "提示", f"关闭开机自启失败：{exc}")
            return

        self.update_autostart_button()
        QtWidgets.QMessageBox.information(self, "提示", "已关闭开机自启！")

    def turn_selected_screens_off(self):
        selected_devices = self.get_selected_devices()
        if not selected_devices:
            selected_devices = set(self.manager.get_active_scrcpy_device_ids())
        if not selected_devices:
            QtWidgets.QMessageBox.information(self, "提示", "请先选择设备，或先连接设备。")
            return

        self.screen_off_btn.setEnabled(False)
        self.screen_off_btn.setText("正在熄屏...")
        threading.Thread(
            target=self._turn_selected_screens_off_worker,
            args=(selected_devices,),
            daemon=True,
        ).start()

    def _turn_selected_screens_off_worker(self, selected_devices):
        ok = []
        failed = []
        for device_id in selected_devices:
            target_device_id = self.get_screen_off_target_device_id(device_id)
            if not target_device_id:
                print(f"⚠️ 设备不在线，无法仅熄灭屏幕：{device_id}")
                failed.append(device_id)
                continue
            try:
                turn_physical_screen_off_keep_awake(target_device_id)
                ok.append(target_device_id)
                print(f"ℹ️ 已发送仅熄灭手机物理屏幕命令，并保持设备活跃：{target_device_id}")
            except Exception as exc:
                failed.append(f"{target_device_id} ({exc})")

        parts = []
        if ok:
            parts.append("已发送仅熄灭屏幕命令：\n" + "\n".join(ok))
        if failed:
            parts.append("失败或离线：\n" + "\n".join(failed))
        self.screen_off_finished.emit("\n\n".join(parts) or "没有可处理的设备。")

    def on_screen_off_finished(self, message):
        self.screen_off_btn.setEnabled(True)
        self.screen_off_btn.setText("仅熄灭手机屏幕")
        QtWidgets.QMessageBox.information(self, "提示", message)

    def turn_selected_screens_on(self):
        selected_devices = self.get_selected_devices()
        if not selected_devices:
            selected_devices = set(self.manager.get_active_scrcpy_device_ids())
        if not selected_devices:
            QtWidgets.QMessageBox.information(self, "提示", "请先选择设备，或先连接设备。")
            return

        self.screen_on_btn.setEnabled(False)
        self.screen_on_btn.setText("正在亮屏...")
        threading.Thread(
            target=self._turn_selected_screens_on_worker,
            args=(selected_devices,),
            daemon=True,
        ).start()

    def _turn_selected_screens_on_worker(self, selected_devices):
        ok = []
        failed = []
        for device_id in selected_devices:
            target_device_id = self.get_screen_off_target_device_id(device_id)
            if not target_device_id:
                print(f"⚠️ 设备不在线，无法仅打开屏幕：{device_id}")
                failed.append(device_id)
                continue
            try:
                # First try the scrcpy shortcut (Ctrl+Shift+O on the scrcpy
                # window).  This works when the scrcpy SDL_app window is
                # focused, but silently fails when the window is minimized
                # (e.g. after Show Desktop).
                self.mouse_shortcut_mapper.send_scrcpy_screen_on_shortcut(target_device_id)
                time.sleep(0.15)
                # Always follow up with the standard wake sequence, then fall
                # through to the force-wake path if the screen is still
                # reported off.  turn_physical_screen_on_force cycles Android's
                # POWER key — the same wake path the user sees when pressing the
                # physical power button — which works even when scrcpy's
                # virtual-display mode left the panel in a dark-but-logically-ON
                # state.
                screen_on = (
                    turn_physical_screen_on(target_device_id)
                    or turn_physical_screen_on_force(target_device_id)
                )
                if screen_on:
                    ok.append(target_device_id)
                    print(f"✅ 手机物理屏幕已亮起，并保持设备活跃：{target_device_id}")
                else:
                    state_debug = get_physical_screen_state_debug(target_device_id)
                    print(f"❌ 仅打开手机物理屏幕失败：{target_device_id}；状态：{state_debug}")
                    failed.append(f"{target_device_id} (物理屏仍为关闭状态；{state_debug})")
            except Exception as exc:
                failed.append(f"{target_device_id} ({exc})")

        parts = []
        if ok:
            parts.append("手机物理屏幕已亮起：\n" + "\n".join(ok))
        if failed:
            parts.append("失败或离线：\n" + "\n".join(failed))
        self.screen_on_finished.emit("\n\n".join(parts) or "没有可处理的设备。")

    def on_screen_on_finished(self, message):
        self.screen_on_btn.setEnabled(True)
        self.screen_on_btn.setText("仅打开手机屏幕")
        QtWidgets.QMessageBox.information(self, "提示", message)

    def get_screen_off_target_device_id(self, device_id):
        if is_device_ready(device_id):
            return device_id

        wifi_id = self.manager.get_wifi_id(device_id)
        if wifi_id and is_device_ready(wifi_id):
            return wifi_id

        active_device_id = self.manager.get_active_usb_for_wifi_id(device_id)
        if active_device_id and is_device_ready(active_device_id):
            return active_device_id

        return None

    def connect_selected_devices(self):
        selected_devices = self.get_selected_devices()
        threading.Thread(
            target=self._connect_selected_devices_worker,
            args=(selected_devices,),
            daemon=True,
        ).start()

    def _connect_selected_devices_worker(self, selected_devices):
        for device_id in selected_devices:
            if ":" not in device_id:
                print(f"ℹ️ 手动启动 USB Superdex：{device_id}")
                start_device_in_superdex(device_id, self.manager, require_wifi=False)
                continue

            start_device_in_superdex(device_id, self.manager, require_wifi=True)

    def closeEvent(self, event):
        if self.exit_requested or not self.tray_icon or not self.tray_icon.isVisible():
            super().closeEvent(event)
            return

        event.ignore()
        self.hide_to_tray()

def list_connected_devices():
    out = adb(["devices"])
    devices = []
    for line in out.splitlines():
        if "\tdevice" not in line:
            continue
        dev_id = line.split("\t", 1)[0]
        if dev_id.startswith("adb-"):
            continue
        devices.append(dev_id)
    return devices

def is_device_ready(serial):
    if not serial:
        return False
    out = adb(["devices"])
    return f"{serial}\tdevice" in out


def get_physical_device_serial(device_id):
    if not device_id:
        return None
    for prop in ("ro.serialno", "ro.boot.serialno"):
        try:
            value = adb(["-s", device_id, "shell", "getprop", prop], timeout=2).strip()
        except Exception:
            value = ""
        if value and value.lower() not in {"unknown", "null"}:
            return value
    if ":" not in device_id:
        return device_id
    return None


def normalize_setting_value(value):
    value = (value or "").strip()
    return value if value and value.lower() != "null" else None


def normalize_saved_setting(value):
    if value is None:
        return None
    return normalize_setting_value(str(value))


def normalize_screen_off_timeout(value, baseline_timeout=None):
    value = normalize_saved_setting(value)
    baseline_timeout = normalize_saved_setting(baseline_timeout)
    if value and value.isdigit() and value not in SUPERDEX_FORCED_SCREEN_OFF_TIMEOUTS:
        return value
    if (
        baseline_timeout
        and baseline_timeout.isdigit()
        and baseline_timeout not in SUPERDEX_FORCED_SCREEN_OFF_TIMEOUTS
    ):
        return baseline_timeout
    return DEFAULT_SCREEN_OFF_TIMEOUT_MS


def sanitize_initial_state(state):
    if not isinstance(state, dict):
        return state

    sanitized = dict(state)
    globals_state = sanitized.get("globals")
    if isinstance(globals_state, dict):
        globals_state = dict(globals_state)
        for key, forced_value in SUPERDEX_FORCED_GLOBALS.items():
            if normalize_saved_setting(globals_state.get(key)) == forced_value:
                globals_state[key] = None
        sanitized["globals"] = globals_state
    return sanitized


def read_setting(device_id, namespace, key):
    value = adb(["-s", device_id, "shell", "settings", "get", namespace, key])
    return normalize_setting_value(value)


def capture_settings(device_id, namespace, keys, baseline_values=None, forced_values=None):
    captured = {}
    baseline_values = baseline_values if isinstance(baseline_values, dict) else {}
    forced_values = forced_values if isinstance(forced_values, dict) else {}
    for key in keys:
        value = read_setting(device_id, namespace, key)
        forced_value = forced_values.get(key)
        baseline_value = baseline_values.get(key)
        if forced_value is not None and value == forced_value:
            if key in baseline_values and baseline_value != forced_value:
                value = baseline_value
            else:
                value = None
        captured[key] = value
    return captured


def clear_policy_control_with_retry(device_id, attempts=3, delay=0.25):
    for _ in range(attempts):
        adb(["-s", device_id, "shell", "settings", "delete", "global", "policy_control"], capture_output=False)
        time.sleep(delay)
        if read_setting(device_id, "global", "policy_control") != SUPERDEX_FORCED_GLOBALS["policy_control"]:
            return True
    return False


def restore_settings(device_id, namespace, settings_map):
    for key, val in (settings_map or {}).items():
        if val is None or str(val).strip() == "" or str(val).lower() == "null":
            adb(["-s", device_id, "shell", "settings", "delete", namespace, key], capture_output=False)
        else:
            adb(["-s", device_id, "shell", "settings", "put", namespace, key, str(val)], capture_output=False)


def has_complete_system_state(state):
    system_state = state.get("system") if isinstance(state, dict) else None
    return isinstance(system_state, dict) and all(key in system_state for key in SYSTEM_STATE_KEYS)


def has_superdex_runtime_marks(device_id):
    launcher = get_device_launcher(device_id)
    policy_control = read_setting(device_id, "global", "policy_control")
    timeout = read_setting(device_id, "system", "screen_off_timeout")
    return (
        launcher == SUPERDEX_ACTIVITY
        or policy_control == SUPERDEX_FORCED_GLOBALS["policy_control"]
        or timeout in SUPERDEX_FORCED_SCREEN_OFF_TIMEOUTS
    )


def is_soft_keyboard_visible(device_id):
    output = adb(["-s", device_id, "shell", "dumpsys", "input_method"])
    normalized = output.lower()
    if "minputshown=true" in normalized or "misinputviewshown=true" in normalized:
        return True

    match = re.search(r"mImeWindowVis=(\d+)", output)
    return bool(match and match.group(1) != "0")


def hide_soft_keyboard_if_visible(device_id):
    if not is_soft_keyboard_visible(device_id):
        return False
    adb(["-s", device_id, "shell", "input", "keyevent", "KEYCODE_BACK"], capture_output=False)
    time.sleep(0.15)
    return True


def _load_all_states_unlocked():
    if not os.path.exists(STATE_FILE):
        return {}

    with open(STATE_FILE, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except Exception:
            return {}


def save_initial_state(device_id, state):
    state = sanitize_initial_state(state)
    with STATE_LOCK:
        all_states = _load_all_states_unlocked()
        all_states[device_id] = state

        tmp_file = STATE_FILE + ".tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(all_states, f, ensure_ascii=False, indent=2)
        os.replace(tmp_file, STATE_FILE)


def load_initial_state(device_id):
    with STATE_LOCK:
        all_states = _load_all_states_unlocked()
    return sanitize_initial_state(all_states.get(device_id))


def parse_size(output):
    match = re.search(r"Physical size:\s*(\d+x\d+)", output or "")
    if match:
        return match.group(1)
    return None


def capture_full_state(device_id):
    size_output = adb(["-s", device_id, "shell", "wm", "size"])
    init_size = parse_size(size_output)

    baseline_state = load_initial_state(device_id)
    if not isinstance(baseline_state, dict):
        baseline_state = None
    baseline_dpi = None
    if (
        baseline_state
        and baseline_state.get("dpi")
        and str(baseline_state["dpi"]).isdigit()
    ):
        baseline_dpi = int(baseline_state["dpi"])
    init_dpi = get_device_dpi(device_id, baseline_dpi=baseline_dpi)

    init_launcher = get_device_launcher(device_id)
    init_orientation = get_device_orientation(device_id)

    globals_state = capture_settings(
        device_id,
        "global",
        GLOBAL_STATE_KEYS,
        baseline_values=(baseline_state or {}).get("globals"),
        forced_values=SUPERDEX_FORCED_GLOBALS,
    )
    system_state = capture_settings(
        device_id,
        "system",
        SYSTEM_STATE_KEYS,
        baseline_values=(baseline_state or {}).get("system"),
        forced_values=SUPERDEX_FORCED_SYSTEM,
    )
    secure_state = capture_settings(device_id, "secure", [HID_IME_SETTING])

    baseline_timeout = baseline_state.get("orig_timeout") if baseline_state else None
    orig_timeout = adb(
        ["-s", device_id, "shell", "settings", "get", "system", "screen_off_timeout"]
    )
    orig_timeout = normalize_screen_off_timeout(orig_timeout, baseline_timeout)

    return {
        "size": init_size,
        "dpi": init_dpi,
        "launcher": init_launcher,
        "orientation": init_orientation,
        "globals": globals_state,
        "system": system_state,
        "secure": secure_state,
        "orig_timeout": str(orig_timeout),
    }


def ensure_initial_state(device_id, manager, force_capture=False):
    state = manager.get_state(device_id)
    if isinstance(state, dict) and state and not force_capture:
        return sanitize_initial_state(state)

    state = load_initial_state(device_id)
    if not isinstance(state, dict):
        state = None

    if force_capture:
        captured_state = capture_full_state(device_id)
        if captured_state:
            state = captured_state
            save_initial_state(device_id, state)
    elif state is None or not has_complete_system_state(state):
        if state is not None and has_superdex_runtime_marks(device_id):
            manager.set_state(device_id, state)
            return state
        captured_state = capture_full_state(device_id)
        if captured_state:
            state = captured_state
            save_initial_state(device_id, state)
    if state:
        manager.set_state(device_id, state)
    return state


def refresh_initial_state(device_id, manager):
    return ensure_initial_state(device_id, manager, force_capture=True)


def scrcpy_mode_label(use_virtual_display):
    return "扩展模式" if use_virtual_display else "镜像模式"


def stop_running_superdex_for_mode_switch(device_id, manager, desired_virtual_display):
    proc = manager.get_scrcpy_proc(device_id)
    if not is_process_running(proc):
        return False
    current_virtual_display = manager.get_scrcpy_virtual_display_mode(device_id)
    if current_virtual_display == desired_virtual_display:
        return False
    print(
        "ℹ️ 连接模式已改变，正在重启 Superdex："
        f"{device_id} "
        f"{scrcpy_mode_label(current_virtual_display)} -> "
        f"{scrcpy_mode_label(desired_virtual_display)}"
    )
    handle_scrcpy_exit(device_id, manager)
    manager.set_last_state(device_id, False)
    return True


def start_device_in_superdex(device_id, manager, require_wifi):
    if not is_device_ready(device_id):
        if require_wifi:
            print(f"❌ 无线设备不在线：{device_id}")
        else:
            print(f"❌ USB 设备不在线：{device_id}")
        return False

    desired_virtual_display = use_scrcpy_virtual_display()
    physical_serial = get_physical_device_serial(device_id)
    manager.set_physical_serial(device_id, physical_serial)
    active_device_id = manager.get_active_device_for_physical_serial(
        physical_serial,
        exclude_device_id=device_id,
    )
    if not active_device_id and ":" in device_id:
        active_device_id = manager.get_active_usb_for_wifi_id(device_id)
    if active_device_id:
        if stop_running_superdex_for_mode_switch(
            active_device_id,
            manager,
            desired_virtual_display,
        ):
            active_device_id = None
        else:
            print(f"ℹ️ 同一台手机已通过 {active_device_id} 运行 Superdex，跳过：{device_id}")
            return False

    if is_process_running(manager.get_scrcpy_proc(device_id)):
        if not stop_running_superdex_for_mode_switch(
            device_id,
            manager,
            desired_virtual_display,
        ):
            print(f"ℹ️ 已在 Superdex 中，跳过：{device_id}")
            return False

    init_state = ensure_initial_state(
        device_id,
        manager,
        force_capture=not has_superdex_runtime_marks(device_id),
    )
    if not init_state:
        print(f"❌ 无法捕获初始状态：{device_id}")
        return False

    wifi_id = ensure_wifi_adb(device_id)
    if wifi_id and is_device_ready(wifi_id):
        manager.set_wifi_id(device_id, wifi_id)
    elif require_wifi:
        print(f"❌ 无线设备 Wi‑Fi adb 建立失败：{device_id}，不进入 Superdex")
        return False
    elif ":" not in device_id:
        print(f"⚠️ Wi‑Fi adb 建立失败：{device_id}（仅备用，不影响 USB Superdex）")

    switch_to_superdex(device_id, manager, physical_serial=physical_serial)
    return True

# -------------------------------
# DPI / Launcher 获取
# -------------------------------
def get_device_dpi(device_id, baseline_dpi=None):
    def read_once():
        dpi = adb(
            ["-s", device_id, "shell", "settings", "get", "secure", "display_density_forced"]
        )
        if dpi.isdigit():
            return int(dpi)
        out = adb(["-s", device_id, "shell", "wm", "density"])
        match = re.search(r"(\d+)", out)
        if match:
            return int(match.group(1))
        return None

    d1 = read_once()
    time.sleep(0.5)
    d2 = read_once()
    candidate = d1 if d1 == d2 else (d1 or d2)

    if candidate and 300 <= candidate <= 800:
        if baseline_dpi and abs(candidate - baseline_dpi) > baseline_dpi * 0.15:
            return str(baseline_dpi)
        return str(candidate)

    if baseline_dpi:
        return str(baseline_dpi)
    return "reset"


def get_device_launcher(device_id):
    out = adb(
        [
            "-s",
            device_id,
            "shell",
            "cmd",
            "package",
            "resolve-activity",
            "--brief",
            "-a",
            "android.intent.action.MAIN",
            "-c",
            "android.intent.category.HOME",
        ]
    ).strip()

    if out and re.fullmatch(r"[A-Za-z0-9_.]+/[A-Za-z0-9_.$]+", out):
        return out

    for launcher in COMMON_LAUNCHERS:
        pkg = launcher.split("/")[0]
        check = adb(["-s", device_id, "shell", "pm", "list", "packages", pkg])
        if pkg in check:
            return launcher

    return COMMON_LAUNCHERS[-1]

def calculate_safe_dpi(width, height, target_short_dp=960):
    short_px = min(width, height)
    return int(round(short_px * 160 / target_short_dp))

def get_device_orientation(device_id):
    out = adb(["-s", device_id, "shell", "dumpsys", "input"])
    match = re.search(r"SurfaceOrientation:\s*(\d+)", out)
    if match:
        return int(match.group(1))
    return 0

def get_pc_resolution():
    cmd = "Get-CimInstance Win32_VideoController | Select-Object -First 1 CurrentHorizontalResolution,CurrentVerticalResolution"
    output = run_host(cmd)
    match = re.search(r"(\d+)\s+(\d+)", output)
    if match:
        width, height = map(int, match.groups())
        return width, height
    return 1920, 1080


def get_virtual_display_resolution(target_short_dp=VIRTUAL_DISPLAY_TARGET_SHORT_DP):
    pc_w, pc_h = get_pc_resolution()
    width, height = max(pc_w, pc_h), min(pc_w, pc_h)
    if not VIRTUAL_DISPLAY_FORCE_LANDSCAPE:
        width, height = pc_w, pc_h
    dpi = calculate_safe_dpi(width, height, target_short_dp)
    return width, height, dpi


def get_display_ids(device_id):
    output = adb(["-s", device_id, "shell", "dumpsys", "display"])
    ids = {int(match) for match in re.findall(r"mDisplayId=(\d+)", output or "")}
    if not ids:
        ids = {int(match) for match in re.findall(r"\bDisplay\s+(\d+)\b", output or "")}
    ids.add(0)
    return ids


def is_package_on_display(device_id, display_id, package_name):
    if display_id is None or not package_name:
        return False

    output = adb(["-s", device_id, "shell", "dumpsys", "activity", "activities"])
    if not output:
        return False

    display_match = re.search(
        rf"Display #{re.escape(str(display_id))}\b(?P<section>.*?)(?=\nDisplay #|\Z)",
        output,
        re.S,
    )
    if not display_match:
        return False

    section = display_match.group("section")
    return f"{package_name}/" in section


def wait_for_virtual_display_id(
    device_id,
    previous_ids=None,
    timeout=VIRTUAL_DISPLAY_WAIT_TIMEOUT,
    poll_interval=VIRTUAL_DISPLAY_WAIT_INTERVAL,
):
    previous_ids = set(previous_ids or [])
    deadline = time.time() + timeout
    while time.time() < deadline:
        current_ids = get_display_ids(device_id)
        new_ids = sorted(current_ids - previous_ids)
        if new_ids:
            return new_ids[-1]
        time.sleep(poll_interval)
    return None


def launch_activity_on_display(device_id, display_id, activity):
    if display_id is None or not activity:
        return False

    package_name = activity.split("/", 1)[0]
    if is_package_on_display(device_id, display_id, package_name):
        return True

    base = ["-s", device_id, "shell", "am", "start", "--display", str(display_id)]
    extras = [
        "--ei",
        "launch_display_id",
        str(display_id),
        "--ez",
        "skip_superdex_launch_animation",
        "true",
        "-n",
        activity,
    ]
    attempts = [
        base + ["--activity-clear-top", "--activity-reset-task-if-needed"] + extras,
        base + ["--activity-new-task"] + extras,
        base + extras,
        base + ["-n", activity],
    ]

    for command in attempts:
        output = adb(command, capture_output=True, timeout=4)
        lowered = (output or "").lower()
        if "error" not in lowered and "exception" not in lowered and "not found" not in lowered:
            return True
        if is_package_on_display(device_id, display_id, package_name):
            return True

    return False


def is_device_awake(device_id):
    output = adb(["-s", device_id, "shell", "dumpsys", "power"])
    normalized = (output or "").lower()

    match = re.search(r"\bminteractive=(true|false)\b", normalized)
    if match:
        return match.group(1) == "true"

    match = re.search(r"\bmwakefulness=([a-z_]+)\b", normalized)
    if match:
        return match.group(1) == "awake"

    match = re.search(r"display power:\s*state=([a-z_]+)", normalized)
    if match:
        return match.group(1) not in {"off", "doze", "doze_suspend"}

    return True


def recover_virtual_display_activity(device_id, manager):
    display_id = manager.get_virtual_display_id(device_id)
    if display_id is None:
        print(f"No virtual display id; cannot recover Superdex desktop: {device_id}")
        return False

    start_scrcpy_screen_off_helper(device_id)
    start_physical_screen_off_guard(device_id)
    time.sleep(0.03)

    output = adb(
        [
            "-s",
            device_id,
            "shell",
            "am",
            "start",
            "--display",
            str(display_id),
            "--activity-clear-top",
            "--activity-single-top",
            "--activity-reset-task-if-needed",
            "--ez",
            "skip_superdex_launch_animation",
            "true",
            "-n",
            SUPERDEX_ACTIVITY,
        ],
        capture_output=True,
    )
    time.sleep(SCREEN_OFF_FALLBACK_DELAY)
    force_physical_screen_off(device_id)

    lowered = (output or "").lower()
    ok = "error" not in lowered and "exception" not in lowered
    if ok:
        print(f"Recovered Superdex desktop on virtual display {display_id}: {device_id}")
    else:
        print(f"Superdex virtual display recovery may have failed: {device_id}; output: {output}")
    return ok


def start_scrcpy_monitor_threads(proc, device_id, manager):
    threading.Thread(target=wait_for_scrcpy_exit, args=(proc, device_id, manager), daemon=True).start()
    threading.Thread(target=monitor_scrcpy_process, args=(proc, device_id, manager), daemon=True).start()


def start_virtual_display_keepalive_thread(proc, device_id, manager):
    threading.Thread(
        target=virtual_display_keepalive_loop,
        args=(proc, device_id, manager),
        daemon=True,
    ).start()


def virtual_display_keepalive_loop(proc, device_id, manager):
    needs_display_recover = False
    while is_process_running(proc) and manager.is_scrcpy_proc_current(device_id, proc):
        try:
            awake = is_device_awake(device_id)

            if not awake:
                keep_virtual_display_device_awake(device_id)
                time.sleep(VIRTUAL_DISPLAY_WAKE_RECOVER_DELAY)
                recover_virtual_display_activity(device_id, manager)
                needs_display_recover = False
            elif needs_display_recover:
                recover_virtual_display_activity(device_id, manager)
                needs_display_recover = False
        except Exception:
            pass
        time.sleep(VIRTUAL_DISPLAY_SCREEN_CHECK_INTERVAL)


def set_orientation_based_resolution(device_id, target_short_dp=960):
    orientation = get_device_orientation(device_id)
    pc_w, pc_h = get_pc_resolution()
    if orientation in (1, 3):
        width, height = max(pc_w, pc_h), min(pc_w, pc_h)
    else:
        width, height = min(pc_w, pc_h), max(pc_w, pc_h)

    adb(["-s", device_id, "shell", "wm", "size", f"{width}x{height}"], capture_output=False)
    dpi = calculate_safe_dpi(width, height, target_short_dp)
    adb(["-s", device_id, "shell", "wm", "density", str(dpi)], capture_output=False)
    return f"{width}x{height}", dpi

def set_launcher(device_id, launcher_activity):
    if not launcher_activity:
        return
    adb(
        ["-s", device_id, "shell", "cmd", "package", "set-home-activity", launcher_activity],
        capture_output=False,
    )
    adb(["-s", device_id, "shell", "input", "keyevent", "KEYCODE_HOME"], capture_output=False)
def get_device_sdk(device_id):
    try:
        sdk = adb(["-s", device_id, "shell", "getprop", "ro.build.version.sdk"])
        return int(sdk)
    except Exception:
        return None


def get_device_manufacturer(device_id):
    return adb(["-s", device_id, "shell", "getprop", "ro.product.manufacturer"]).strip().lower()


def should_force_sdk_keyboard(device_id, sdk):
    manufacturer = get_device_manufacturer(device_id)
    return manufacturer == "samsung" and sdk is not None and sdk >= 33


def get_keyboard_mode(device_id, sdk):
    if should_force_sdk_keyboard(device_id, sdk):
        return KEYBOARD_MAPPING_MODE
    if sdk is not None and sdk >= HID_KEYBOARD_MIN_SDK:
        return "uhid"
    return "sdk"


def launch_scrcpy(device_id, use_virtual_display=None):
    if use_virtual_display is None:
        use_virtual_display = use_scrcpy_virtual_display()
    os.environ["SCRCPY_CLIPBOARD_SYNC"] = "true"
    sdk = get_device_sdk(device_id)
    keyboard_mode = get_keyboard_mode(device_id, sdk)
    scrcpy_keyboard_mode = (
        MAPPED_MODE_SCRCPY_KEYBOARD if keyboard_mode == KEYBOARD_MAPPING_MODE else keyboard_mode
    )
    args = [
        SCRCPY_PATH,
        "-s",
        device_id,
        "--video-bit-rate",
        "20M",
        "--fullscreen",
        "--window-borderless",
        "--capture-orientation=0",
        "--mouse=sdk",
        "--mouse-bind=bhsn",
        f"--shortcut-mod={SCRCPY_SHORTCUT_MOD}",
        f"--window-title=Superdex {device_id}",
        "--keep-active",
    ]
    if not use_virtual_display:
        args += ["-S", "--turn-screen-off"]

    if use_virtual_display:
        vd_width, vd_height, vd_dpi = get_virtual_display_resolution()
        args += [f"--new-display={vd_width}x{vd_height}/{vd_dpi}"]
        if VIRTUAL_DISPLAY_DISABLE_SCREENSAVER:
            args += ["--disable-screensaver"]
        if VIRTUAL_DISPLAY_DISABLE_SYSTEM_DECORATIONS:
            args += ["--no-vd-system-decorations"]
        if VIRTUAL_DISPLAY_START_APP_PACKAGE:
            args += [f"--start-app={VIRTUAL_DISPLAY_START_APP_PACKAGE}"]
        if VIRTUAL_DISPLAY_IME_POLICY:
            args += [f"--display-ime-policy={VIRTUAL_DISPLAY_IME_POLICY}"]
    if sdk is None or sdk >= 29:
        args += ["--video-codec=h265"]
    args += [f"--keyboard={scrcpy_keyboard_mode}"]
    if scrcpy_keyboard_mode == "sdk":
        args += ["--prefer-text"]

    return subprocess.Popen(args, creationflags=SUBPROCESS_FLAGS)



def wait_for_scrcpy_exit(proc, device_id, manager):
    proc.wait()
    if manager.is_scrcpy_proc_current(device_id, proc):
        handle_scrcpy_exit(device_id, manager)


def monitor_scrcpy_process(proc, device_id, manager):
    try:
        process = psutil.Process(proc.pid)
        while True:
            if not manager.is_scrcpy_proc_current(device_id, proc):
                break
            if not process.is_running() or process.status() == psutil.STATUS_ZOMBIE:
                if manager.is_scrcpy_proc_current(device_id, proc):
                    handle_scrcpy_exit(device_id, manager)
                break
            time.sleep(2)
    except psutil.NoSuchProcess:
        if manager.is_scrcpy_proc_current(device_id, proc):
            handle_scrcpy_exit(device_id, manager)


def get_third_party_packages(device_id):
    packages = set()
    output = adb(["-s", device_id, "shell", "pm", "list", "packages", "-3"])
    for line in output.splitlines():
        line = line.strip()
        if line.startswith("package:"):
            package = line.split("package:", 1)[1].strip()
            if package:
                packages.add(package)
    return packages


def get_running_process_names(device_id):
    output = adb(["-s", device_id, "shell", "ps", "-A"])
    if not output:
        output = adb(["-s", device_id, "shell", "ps"])

    names = set()
    for line in output.splitlines():
        parts = line.split()
        if len(parts) < 2 or parts[0].upper() == "USER":
            continue
        name = parts[-1]
        if not name or name.startswith("["):
            continue
        names.add(name)
    return names


def get_running_third_party_packages(device_id):
    third_party_packages = get_third_party_packages(device_id)
    if not third_party_packages:
        return set()

    running = set()
    for process_name in get_running_process_names(device_id):
        package = process_name.split(":", 1)[0]
        if package in third_party_packages:
            running.add(package)
    return running - THIRD_PARTY_RESTART_EXCLUDE_PACKAGES


def get_current_input_method_package(device_id):
    ime = read_setting(device_id, "secure", "default_input_method")
    if not ime:
        return None
    return ime.split("/", 1)[0].strip() or None


def restart_layout_affected_apps(device_id, pre_superdex_apps, state):
    if not RESTART_THIRD_PARTY_APPS_ON_RESTORE:
        return

    running_apps = get_running_third_party_packages(device_id)
    if RESTART_ONLY_APPS_STARTED_DURING_SUPERDEX:
        if pre_superdex_apps is None:
            return
        targets = running_apps - set(pre_superdex_apps)
    else:
        targets = running_apps

    excluded = set(THIRD_PARTY_RESTART_EXCLUDE_PACKAGES)
    launcher = str((state or {}).get("launcher") or "")
    if "/" in launcher:
        excluded.add(launcher.split("/", 1)[0])
    ime_package = get_current_input_method_package(device_id)
    if ime_package:
        excluded.add(ime_package)

    for package in sorted(targets - excluded):
        stop_package_completely(device_id, package, attempts=2, delay=0.2)


def stop_package_completely(device_id, pkg, attempts=5, delay=0.5):
    for _ in range(attempts):
        adb(["-s", device_id, "shell", "am", "force-stop", pkg], capture_output=False)
        adb(["-s", device_id, "shell", "am", "kill", pkg], capture_output=False)
        adb(["-s", device_id, "shell", "cmd", "activity", "stop-app", pkg], capture_output=False)
        adb(["-s", device_id, "shell", "pkill", "-f", pkg], capture_output=False)

        pids = [
            pid
            for pid in adb(["-s", device_id, "shell", "pidof", pkg]).split()
            if pid.isdigit()
        ]
        if not pids:
            return True

        for pid in pids:
            adb(["-s", device_id, "shell", "kill", "-9", pid], capture_output=False)
        time.sleep(delay)

    remaining = [
        pid
        for pid in adb(["-s", device_id, "shell", "pidof", pkg]).split()
        if pid.isdigit()
    ]
    return not remaining


def kill_superdex_process(device_id):
    return stop_package_completely(device_id, SUPERDEX_PACKAGE)


def prepare_superdex_clean_launch(device_id, stage="prelaunch"):
    if not SUPERDEX_PRELAUNCH_CLEAN_RESTART:
        return
    kill_superdex_process(device_id)
    adb(
        [
            "-s",
            device_id,
            "shell",
            "am",
            "broadcast",
            "-a",
            "android.intent.action.CLOSE_SYSTEM_DIALOGS",
        ],
        capture_output=False,
        timeout=2,
    )
    time.sleep(SUPERDEX_PRELAUNCH_CLEAN_DELAY)
    print(f"Superdex clean launch prepared ({stage}): {device_id}")


def restart_launcher(device_id, launcher):
    if not launcher or "/" not in launcher:
        return
    pkg = launcher.split("/")[0]
    adb(["-s", device_id, "shell", "am", "force-stop", pkg], capture_output=False)
    time.sleep(0.5)

# -------------------------------
# 状态恢复 / Wi‑Fi ADB / Watchdog
# -------------------------------
def restore_state(device_id, state):
    kill_superdex_process(device_id)

    size = state.get("size")
    if size and re.match(r"^\d+x\d+$", str(size)):
        adb(["-s", device_id, "shell", "wm", "size", str(size)], capture_output=False)
    else:
        adb(["-s", device_id, "shell", "wm", "size", "reset"], capture_output=False)

    dpi = str(state.get("dpi") or "").strip()
    if dpi.isdigit():
        adb(["-s", device_id, "shell", "wm", "density", dpi], capture_output=False)
    else:
        adb(["-s", device_id, "shell", "wm", "density", "reset"], capture_output=False)

    launcher = str(state.get("launcher") or "").strip()
    if not launcher or "/" not in launcher or "accessibility" in launcher.lower():
        launcher = get_device_launcher(device_id)

    time.sleep(1)
    set_launcher(device_id, launcher)
    saved_launcher = str(state.get("launcher") or "").strip()
    if saved_launcher:
        restart_launcher(device_id, saved_launcher)

    system_state = dict(state.get("system") or {})
    # Match superdex2.py's stable restore behavior: keep rotation locked after exit,
    # otherwise Android's sensor rotation may immediately override user_rotation.
    system_state.pop("accelerometer_rotation", None)
    orientation = state.get("orientation")
    if orientation is not None and str(orientation).isdigit():
        system_state["user_rotation"] = str(orientation)
    restore_settings(device_id, "system", system_state)

    restore_settings(device_id, "global", state.get("globals"))

    restore_settings(device_id, "secure", state.get("secure"))

    clear_policy_control_with_retry(device_id)
    send_home_key(device_id)

    orig_timeout = normalize_screen_off_timeout(state.get("orig_timeout"))
    adb(
        ["-s", device_id, "shell", "settings", "put", "system", "screen_off_timeout", orig_timeout],
        capture_output=False,
    )

    # restore_settings() already restored stay_on_while_plugged_in from the saved state.
    adb(
        ["-s", device_id, "shell", "svc", "power", "stayon", "false"],
        capture_output=False,
    )

    clear_policy_control_with_retry(device_id)

def get_wlan_ip(device_id, retries=5, delay=1):
    for _ in range(retries):
        candidates = []
        for ifc in ["wlan0", "wlan1"]:
            out = adb(["-s", device_id, "shell", "ip", "-f", "inet", "addr", "show", ifc])
            m = re.search(r"inet (\d+\.\d+\.\d+\.\d+)/", out)
            if m:
                candidates.append(m.group(1))
        for key in ["dhcp.wlan0.ipaddress", "dhcp.wlan1.ipaddress"]:
            ip = adb(["-s", device_id, "shell", "getprop", key]).strip()
            if re.match(r"^\d+\.\d+\.\d+\.\d+$", ip):
                candidates.append(ip)
        if candidates:
            return candidates[0]
        time.sleep(delay)
    return None

def ensure_wifi_adb(device_id):
    if ":" in device_id:
        if is_device_ready(device_id):
            return device_id
        adb(["connect", device_id], capture_output=False)
        return device_id if is_device_ready(device_id) else None

    adb(["-s", device_id, "tcpip", "5555"], capture_output=False)
    ip = get_wlan_ip(device_id)
    if not ip:
        return None
    wifi_id = f"{ip}:5555"
    adb(["connect", wifi_id], capture_output=False)
    return wifi_id if is_device_ready(wifi_id) else None

def build_usb_watchdog_script(launcher):
    return (
        "while true; do "
        'STATE="$(dumpsys battery)"; '
        'case "$STATE" in '
        '*"USB powered: true"*) sleep 3 ;; '
        f'*) cmd package set-home-activity {launcher}; '
        f'am start -n {launcher}; '
        "input keyevent KEYCODE_HOME; "
        f'am force-stop {SUPERDEX_PACKAGE}; '
        f'am kill {SUPERDEX_PACKAGE}; '
        f'cmd activity stop-app {SUPERDEX_PACKAGE}; '
        f'for i in 1 2 3 4 5; do PIDS="$(pidof {SUPERDEX_PACKAGE})"; '
        '[ -z "$PIDS" ] && break; '
        'kill -9 $PIDS >/dev/null 2>&1; '
        'sleep 0.5; '
        "done; "
        "input keyevent KEYCODE_HOME; "
        "exit 0 ;; "
        "esac; "
        "done"
    )

def start_usb_watchdog(device_id, launcher, manager):
    if ":" in device_id or not launcher:
        return

    stop_process(manager.pop_watchdog_proc(device_id))
    proc = subprocess.Popen(
        [ADB_PATH, "-s", device_id, "shell", "sh", "-c", build_usb_watchdog_script(launcher)],
        creationflags=SUBPROCESS_FLAGS,
    )
    manager.set_watchdog_proc(device_id, proc)


def stop_watchdog(device_id, manager):
    stop_process(manager.pop_watchdog_proc(device_id))


def disable_charging_stay_awake(device_id):
    adb(
        ["-s", device_id, "shell", "settings", "put", "global", "stay_on_while_plugged_in", "0"],
        capture_output=False,
    )
    adb(
        ["-s", device_id, "shell", "svc", "power", "stayon", "false"],
        capture_output=False,
    )


def enable_never_sleep(device_id):
    adb(
        ["-s", device_id, "shell", "settings", "put", "system", "screen_off_timeout", "2147483647"],
        capture_output=False,
        timeout=2,
    )
    adb(
        ["-s", device_id, "shell", "settings", "put", "global", "stay_on_while_plugged_in", "7"],
        capture_output=False,
        timeout=2,
    )
    adb(
        ["-s", device_id, "shell", "svc", "power", "stayon", "true"],
        capture_output=False,
        timeout=2,
    )


def keep_virtual_display_device_awake(device_id):
    enable_never_sleep(device_id)
    adb(
        ["-s", device_id, "shell", "input", "keyevent", "KEYCODE_WAKEUP"],
        capture_output=False,
        timeout=2,
    )


def switch_to_superdex(device_id, manager, physical_serial=None):
    if not manager.begin_start(device_id, physical_serial=physical_serial):
        active_device_id = manager.get_active_device_for_physical_serial(
            physical_serial,
            exclude_device_id=device_id,
        )
        if active_device_id:
            print(f"ℹ️ 同一台手机已通过 {active_device_id} 运行 Superdex，跳过：{device_id}")
        else:
            print(f"ℹ️ 已在 Superdex 中或正在启动：{device_id}")
        return

    started = False
    use_virtual_display = use_scrcpy_virtual_display()
    manager.set_scrcpy_virtual_display_mode(device_id, use_virtual_display)
    try:
        init_state = ensure_initial_state(device_id, manager)
        if not init_state:
            print(f"❌ 无法捕获初始状态：{device_id}")
            return

        if RESTART_THIRD_PARTY_APPS_ON_RESTORE and not use_virtual_display:
            manager.set_pre_superdex_apps(device_id, get_running_third_party_packages(device_id))

        sdk = get_device_sdk(device_id)
        keyboard_mode = get_keyboard_mode(device_id, sdk)
        manager.set_keyboard_mode(device_id, keyboard_mode)

        if use_virtual_display:
            prepare_superdex_clean_launch(device_id, stage="before virtual display")
            display_ids_before_launch = get_display_ids(device_id)
            keep_virtual_display_device_awake(device_id)

            if keyboard_mode in KEYBOARD_LIKE_MODES:
                adb(
                    ["-s", device_id, "shell", "settings", "put", "secure", HID_IME_SETTING, "0"],
                    capture_output=False,
                )
                hide_soft_keyboard_if_visible(device_id)

            time.sleep(SUPERDEX_INPUT_SETTLE_DELAY)
            proc = launch_scrcpy(device_id, use_virtual_display=True)
            manager.set_scrcpy_proc(device_id, proc)
            started = True
            manager.set_last_orientation(device_id, None)

            start_virtual_display_keepalive_thread(proc, device_id, manager)

            virtual_display_id = wait_for_virtual_display_id(device_id, display_ids_before_launch)
            manager.set_virtual_display_id(device_id, virtual_display_id)
            if virtual_display_id is not None:
                time.sleep(0.25)
                if not launch_activity_on_display(device_id, virtual_display_id, SUPERDEX_ACTIVITY):
                    print(f"⚠️ Superdex 启动到虚拟屏失败，尝试恢复：{device_id}")
                    recover_virtual_display_activity(device_id, manager)
                turn_physical_screen_off_keep_awake(device_id)
                print(f"ℹ️ 扩展模式已自动仅熄灭手机物理屏幕：{device_id}")
                print(f"ℹ️ 已切换到 scrcpy 虚拟屏显示模式，显示 ID：{virtual_display_id}（{device_id}）")
            else:
                print(f"⚠️ scrcpy 虚拟屏已启动，但未能识别显示 ID：{device_id}")
        else:
            set_orientation_based_resolution(device_id, target_short_dp=960)
            set_launcher(device_id, SUPERDEX_ACTIVITY)
            adb(
                ["-s", device_id, "shell", "settings", "put", "system", "screen_off_timeout", "2147483647"],
                capture_output=False,
            )

            rotation = 1 if (sdk is not None and sdk < 34) else 1
            adb(
                ["-s", device_id, "shell", "settings", "put", "system", "user_rotation", str(rotation)],
                capture_output=False,
            )
            adb(
                ["-s", device_id, "shell", "settings", "put", "system", "accelerometer_rotation", "0"],
                capture_output=False,
            )
            adb(
                ["-s", device_id, "shell", "settings", "put", "global", "policy_control", "immersive.full=*"],
                capture_output=False,
            )
            adb(
                ["-s", device_id, "shell", "settings", "put", "global", "display_cutout_force_fullscreen", "1"],
                capture_output=False,
            )
            adb(
                ["-s", device_id, "shell", "settings", "put", "global", "force_fullscreen", "1"],
                capture_output=False,
            )
            if keyboard_mode in KEYBOARD_LIKE_MODES:
                adb(
                    ["-s", device_id, "shell", "settings", "put", "secure", HID_IME_SETTING, "0"],
                    capture_output=False,
                )
                hide_soft_keyboard_if_visible(device_id)

            time.sleep(SUPERDEX_INPUT_SETTLE_DELAY)
            start_usb_watchdog(device_id, init_state.get("launcher"), manager)

            proc = launch_scrcpy(device_id, use_virtual_display=False)
            manager.set_scrcpy_proc(device_id, proc)
            started = True
            manager.set_last_orientation(device_id, get_device_orientation(device_id))

        if keyboard_mode == "uhid":
            print(f"ℹ️ 键盘模式：UHID（物理键盘模拟）{device_id}")
        elif keyboard_mode == KEYBOARD_MAPPING_MODE:
            print(f"ℹ️ 键盘模式：映射转发（QtScrcpy 风格，附着 UHID 键盘）{device_id}")
            print(f"ℹ️ Ctrl 组合键交由附着键盘处理：{device_id}")
        else:
            print(f"ℹ️ 键盘模式：SDK 兼容模式（共享键盘优先）{device_id}")
        print(f"✅ scrcpy 已启动（{device_id}）。拖拽文件到 scrcpy 窗口即可传输到 /sdcard/Download")

        start_scrcpy_monitor_threads(proc, device_id, manager)
    except Exception:
        manager.clear_starting(device_id)
        raise
    finally:
        if not started:
            manager.pop_pre_superdex_apps(device_id)
            manager.pop_virtual_display_id(device_id)
            manager.pop_scrcpy_virtual_display_mode(device_id)
            manager.pop_physical_serial(device_id)
            manager.clear_starting(device_id)



def handle_scrcpy_exit(device_id, manager):
    restore_lock = manager.get_restore_lock(device_id)
    if not restore_lock.acquire(blocking=False):
        return

    try:
        stop_watchdog(device_id, manager)
        stop_process(manager.pop_scrcpy_proc(device_id))

        init_state = manager.get_state(device_id) or load_initial_state(device_id)
        wifi_id = manager.get_wifi_id(device_id)
        pre_superdex_apps = manager.pop_pre_superdex_apps(device_id)
        was_virtual_display = manager.get_scrcpy_virtual_display_mode(device_id)
        manager.pop_virtual_display_id(device_id)
        manager.pop_scrcpy_virtual_display_mode(device_id)
        restore_target = None

        if is_device_ready(device_id):
            restore_target = device_id
        elif wifi_id and is_device_ready(wifi_id):
            restore_target = wifi_id

        if init_state and restore_target:
            try:
                restore_state(restore_target, init_state)
                if not was_virtual_display:
                    restart_layout_affected_apps(restore_target, pre_superdex_apps, init_state)
            except Exception:
                pass

        manager.pop_wifi_id(device_id)
        manager.pop_last_orientation(device_id)
        manager.pop_physical_serial(device_id)
    finally:
        restore_lock.release()



def handle_usb_disconnect(device_id, manager):
    handle_scrcpy_exit(device_id, manager)


def main(manager, get_selected_devices=lambda: set()):
    while True:
        devices = list_connected_devices()
        current_set = set(devices)
        previous_set = set(manager.known_devices())

        for device_id in previous_set - current_set:
            if manager.get_last_state(device_id):
                handle_usb_disconnect(device_id, manager)
            manager.set_last_state(device_id, False)
            manager.pop_physical_serial(device_id)

        selected_devices = get_selected_devices()
        for device_id in devices:
            try:
                previously_connected = manager.get_last_state(device_id, False)

                if not previously_connected and ":" not in device_id:
                    start_device_in_superdex(device_id, manager, require_wifi=False)

                if ":" in device_id and device_id in selected_devices and not previously_connected:
                    start_device_in_superdex(device_id, manager, require_wifi=True)

                if (
                    is_process_running(manager.get_scrcpy_proc(device_id))
                    and not manager.get_scrcpy_virtual_display_mode(device_id)
                ):
                    current_orientation = get_device_orientation(device_id)
                    if current_orientation != manager.get_last_orientation(device_id):
                        set_orientation_based_resolution(device_id, target_short_dp=960)
                        manager.set_last_orientation(device_id, current_orientation)

                manager.set_last_state(device_id, True)
            except Exception as exc:
                manager.clear_starting(device_id)
                manager.set_last_state(device_id, False)
                print(f"⚠️ Superdex 主循环处理设备失败，稍后重试：{device_id} ({exc})")

        time.sleep(3)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(not QtWidgets.QSystemTrayIcon.isSystemTrayAvailable())
    if os.path.exists(APP_ICON_PATH):
        app_icon = QtGui.QIcon(APP_ICON_PATH)
        if not app_icon.isNull():
            app.setWindowIcon(app_icon)
    gui = SuperdexGUI()
    gui.show()
    sys.exit(app.exec_())
