import ctypes
from ctypes import wintypes, windll

import win32api
import win32con
import win32gui

# PUL = ctypes.POINTER(ctypes.c_ulong)
PUL = ctypes.c_void_p


class KeyBdMsg(ctypes.Structure):
    """
    键盘回调函数用结构体
    """
    _fields_ = [
        ('vkCode', wintypes.DWORD),
        ('scanCode', wintypes.DWORD),
        ('flags', wintypes.DWORD),
        ('time', wintypes.DWORD),
        ('dwExtraInfo', PUL)]


class KeyBdInput(ctypes.Structure):
    """
    键盘输入用结构体
    """
    EXTENDEDKEY = 0x0001
    KEYUP = 0x0002
    SCANCODE = 0x0008
    UNICODE = 0x0004

    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class HardwareInput(ctypes.Structure):
    """
    硬件输入用结构体
    """
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]


class MouseInput(ctypes.Structure):
    """
    鼠标输入用结构体
    """
    MOVE = 0x0001
    LEFTDOWN = 0x0002
    LEFTUP = 0x0004
    RIGHTDOWN = 0x0008
    RIGHTUP = 0x0010
    MIDDLEDOWN = 0x0020
    MIDDLEUP = 0x0040
    XDOWN = 0x0080
    XUP = 0x0100
    WHEEL = 0x0800
    HWHEEL = 0x1000
    ABSOLUTE = 0x8000

    XBUTTON1 = 0x0001
    XBUTTON2 = 0x0002

    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class InputUnion(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                ("mi", MouseInput),
                ("hi", HardwareInput)]


class Input(ctypes.Structure):
    """
    SendInput函数用最终结构体
    """
    MOUSE = 0
    KEYBOARD = 1
    HARDWARE = 2

    _fields_ = [("type", ctypes.c_ulong),
                ("ii", InputUnion)]


# 键盘事件用回调函数
HookProc = ctypes.WINFUNCTYPE(
    wintypes.LPARAM,
    ctypes.c_int32, wintypes.WPARAM, ctypes.POINTER(KeyBdMsg))


# 消息队列发送函数
SendInput = windll.user32.SendInput
SendInput.argtypes = (
    wintypes.UINT,
    ctypes.POINTER(Input),
    ctypes.c_int)


# 获取并阻断消息队列
GetMessage = windll.user32.GetMessageA
GetMessage.argtypes = (
    wintypes.MSG,
    wintypes.HWND,
    wintypes.UINT,
    wintypes.UINT)


# 设置回调函数
SetWindowsHookEx = windll.user32.SetWindowsHookExA
SetWindowsHookEx.argtypes = (
    ctypes.c_int,
    HookProc,
    wintypes.HINSTANCE,
    wintypes.DWORD)


# 解除回调函数
UnhookWindowsHookEx = windll.user32.UnhookWindowsHookEx
UnhookWindowsHookEx.argtypes = (
    wintypes.HHOOK,)


# 将消息传递到钩子链下一函数
CallNextHookEx = windll.user32.CallNextHookEx
CallNextHookEx.argtypes = (
    wintypes.HHOOK,
    ctypes.c_int,
    wintypes.WPARAM,
    KeyBdMsg)


GetAsyncKeyState = windll.user32.GetAsyncKeyState
GetAsyncKeyState.argtypes = (
    ctypes.c_int,
)


GetMessageExtraInfo = windll.user32.GetMessageExtraInfo


SetMessageExtraInfo = windll.user32.SetMessageExtraInfo
SetMessageExtraInfo.argtypes = (
    wintypes.LPARAM,
)


def send_kb_event(v_key, is_pressed):
    """
    向消息队列发送键盘输入，指定dwExtraInfo为228，便于回调函数过滤此部分键盘输入
    :param v_key: 虚拟键号
    :param is_pressed: 是否按下
    :return:
    """
    extra = 228
    li = InputUnion()
    flag = KeyBdInput.KEYUP if not is_pressed else 0
    li.ki = KeyBdInput(v_key, 0x48, flag, 0, extra)
    input = Input(Input.KEYBOARD, li)
    return SendInput(1, ctypes.pointer(input), ctypes.sizeof(input))


def send_unicode(unicode):
    extra = 228
    li = InputUnion()
    flag = KeyBdInput.UNICODE
    li.ki = KeyBdInput(0, ord(unicode), flag, 0, extra)
    input = Input(Input.KEYBOARD, li)
    return SendInput(1, ctypes.pointer(input), ctypes.sizeof(input))


def change_language_layout(language):
    hwnd = win32gui.GetForegroundWindow()
    im_list = win32api.GetKeyboardLayoutList()
    im_list = list(map(hex, im_list))
    # print(im_list)

    if hex(language) not in im_list:
        win32api.LoadKeyboardLayout('0000' + hex(language)[-4:], 1)
        im_list = win32api.GetKeyboardLayoutList()
        im_list = list(map(hex, im_list))
        if hex(language) not in im_list:
            return False

    result = win32api.SendMessage(
        hwnd,
        win32con.WM_INPUTLANGCHANGEREQUEST,
        0,
        language)
    return result == 0
