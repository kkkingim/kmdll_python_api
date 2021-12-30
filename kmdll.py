import ctypes, _ctypes
import time, random

from enum import Enum


class MouseButton(Enum):
    MOUSE_LEFT = 0
    MOUSE_RIGHT = 1
    MOUSE_MIDDLE = 2


KEY_CODE_MAPPING = {
    # A-Z
    "a": 0x04,
    "b": 0x05,
    "c": 0x06,
    "d": 0x07,
    "e": 0x08,
    "f": 0x09,
    "g": 0x0A,
    "h": 0x0B,
    "i": 0x0C,
    "j": 0x0D,
    "k": 0x0E,
    "l": 0x0F,
    "m": 0x10,
    "n": 0x11,
    "o": 0x12,
    "p": 0x13,
    "q": 0x14,
    "r": 0x15,
    "s": 0x16,
    "t": 0x17,
    "u": 0x18,
    "v": 0x19,
    "w": 0x1A,
    "x": 0x1B,
    "y": 0x1C,
    "z": 0x1D,
    # 0-9
    "1": 0x1E,
    "2": 0x1F,
    "3": 0x20,
    "4": 0x21,
    "5": 0x22,
    "6": 0x23,
    "7": 0x24,
    "8": 0x25,
    "9": 0x26,
    "0": 0x27,
    # signal
    "-": 0x2D,
    "=": 0x2E,
    "[": 0x2F,
    "]": 0x30,
    "\\": 0x31,
    ";": 0x33,
    "'": 0x34,
    "`": 0x35,
    ",": 0x36,
    ".": 0x37,
    "/": 0x38,
    # F1-F12
    "f1": 0x3A,
    "f2": 0x3B,
    "f3": 0x3C,
    "f4": 0x3D,
    "f5": 0x3E,
    "f6": 0x3F,
    "f7": 0x40,
    "f8": 0x41,
    "f9": 0x42,
    "f10": 0x43,
    "f11": 0x44,
    "f12": 0x45,
    # function keys
    "leftctrl": 0xE0,
    "leftshift": 0xE1,
    "leftalt": 0xE2,
    "leftwin": 0xE3,
    "rightctrl": 0xE4,
    "rightshift": 0xE5,
    "rightalt": 0xE6,
    "rightwin": 0xE7,
    "ctrl": 0xE0,
    "shift": 0xE1,
    "alt": 0xE2,
    "win": 0xE3,
    "enter": 0x28,
    "esc": 0x29,
    "backspace": 0x2A,
    "delete": 0x4C,
    "tab": 0x2B,
    "space": 0x2C,
    " ": 0x2C,
    "home": 0x4A,
    "end": 0x4D,
    "pageup": 0x4B,
    "pagedown": 0x4E,
    "printscreen": 0x46,
    "pause": 0x48,
    "break": 0x48,
    "insert": 0x49,
    # arrow
    "right": 0x4F,
    "left": 0x50,
    "down": 0x51,
    "up": 0x52,
    # lock
    "capslock": 0x39,
    "scrolllock": 0x47,
    "numlock": 0x53,
    # num
    "num_/": 0x54,
    "num_*": 0x55,
    "num_-": 0x56,
    "num_+": 0x57,
    "num_enter": 0x58,
    "num_1": 0x59,
    "num_2": 0x5A,
    "num_3": 0x5B,
    "num_4": 0x5C,
    "num_5": 0x5D,
    "num_6": 0x5E,
    "num_7": 0x5F,
    "num_8": 0x60,
    "num_9": 0x61,
    "num_0": 0x62,
    "num_.": 0x63,
    "num_=": 0x67,
    # media
    "Menu": 0x65,
    "mute": 0x7F,
    "vol+": 0x80,
    "vol-": 0x81,
}


class KM:
    def __init__(self, dllpath="./kmdll.dll"):
        self.dllpath = dllpath

        self.km = ctypes.CDLL(dllpath)
        if not self.km.DllOpenDevice():
            print("Open Device Failed!")
            self.close(-1)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close(0)

    def close(self, signal):
        time.sleep(1)
        self.km.DllStop()
        _ctypes.FreeLibrary(self.km._handle)
        exit(signal)

    def ScreenShot(self, save_to, bounds):
        if not save_to.endswith(".bmp"):
            save_to += ".bmp"
        self.km.DllPrintScreenW(*bounds, save_to)

    def DelayRandom(self, min=50, max=100):
        s = random.randint(min, max) / 1000.0
        time.sleep(s)

    def KeyDown(self, key):
        if type(key) == str and key in KEY_CODE_MAPPING:
            key = KEY_CODE_MAPPING[key]
        if type(key == int):
            self.km.Dllkey_eventCode(1, key)

    def KeyUp(self, key):
        if type(key) == str and key in KEY_CODE_MAPPING:
            key = KEY_CODE_MAPPING[key]
        if type(key == int):
            self.km.Dllkey_eventCode(2, key)

    def KeyUpAll(self):
        self.km.Dllkey_eventCode(3)

    def KeyPress(self, key, delay=None):
        self.KeyDown(key)
        if not delay:
            self.DelayRandom()
        else:
            time.sleep(delay / 1000.0)
        self.KeyUp(key)

    def KeyPressHotkey(self, func_keys, keys):
        for key in func_keys:
            self.KeyDown(key)
        for key in keys:
            self.KeyPress(key)
        for key in func_keys:
            self.KeyUp(key)

    def MouseClickDown(self, button=MouseButton.MOUSE_LEFT):
        btn = button.value * 2 + 1
        self.km.Dllmouse_event(btn, 0, 0)

    def MouseClickUp(self, button=MouseButton.MOUSE_LEFT):
        btn = (button.value + 1) * 2
        self.km.Dllmouse_event(btn, 0, 0)

    def MouseClickUpAll(self):
        self.km.Dllmouse_event(7, 0, 0)

    def MouseClick(self, button=MouseButton.MOUSE_LEFT, delay=None):
        self.MouseClickDown(button)
        if not delay:
            self.DelayRandom()
        else:
            time.sleep(delay / 1000.0)
        self.MouseClickUp(button)

    def MouseMove(self, dx, dy):
        if dx not in range(-128, 128) or dy not in range(-128, 128):
            return
        self.km.Dllmouse_event(9, dx, dy)

    def MouseMoveTo(self, x, y, smoothly=False):
        code = 8
        if smoothly:
            code = 11
        self.km.Dllmouse_event(code, x, y)

    def MouseMoveClick(self, x, y, button=MouseButton.MOUSE_LEFT, smoothly=False):
        self.MouseMoveTo(x, y, smoothly)
        self.DelayRandom()
        self.MouseClick(button)

    def ScrollUp(self, val=1):
        self.km.Dllmouse_event(10, val, 0)

    def ScrollDown(self, val=-1):
        self.km.Dllmouse_event(10, val, 0)

    def InputStr(self, str):
        for c in str:
            self.DelayRandom()
            self.KeyPress(c)
