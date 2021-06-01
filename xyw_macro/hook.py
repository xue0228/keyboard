import win32con
import time
from xyw_macro.notify import Notification
from xyw_macro.utils import SingletonType
from xyw_macro.win32 import *
from xyw_macro.contants import SLEEP_TIME


class HookConstants:
    """
    存储windows钩子内置常数，包括钩子类型，虚拟键号名与其值之间的相互映射，以及时间类型值与名称之间的映射
    """
    # 钩子类型
    WH_MIN = -1
    WH_MSGFILTER = -1
    WH_JOURNALRECORD = 0
    WH_JOURNALPLAYBACK = 1
    WH_KEYBOARD = 2
    WH_GETMESSAGE = 3
    WH_CALLWNDPROC = 4
    WH_CBT = 5
    WH_SYSMSGFILTER = 6
    WH_MOUSE = 7
    WH_HARDWARE = 8
    WH_DEBUG = 9
    WH_SHELL = 10
    WH_FOREGROUNDIDLE = 11
    WH_CALLWNDPROCRET = 12
    WH_KEYBOARD_LL = 13
    WH_MOUSE_LL = 14
    WH_MAX = 15

    # 鼠标事件类型
    WM_MOUSEFIRST = 0x0200
    WM_MOUSEMOVE = 0x0200
    WM_LBUTTONDOWN = 0x0201
    WM_LBUTTONUP = 0x0202
    WM_LBUTTONDBLCLK = 0x0203
    WM_RBUTTONDOWN = 0x0204
    WM_RBUTTONUP = 0x0205
    WM_RBUTTONDBLCLK = 0x0206
    WM_MBUTTONDOWN = 0x0207
    WM_MBUTTONUP = 0x0208
    WM_MBUTTONDBLCLK = 0x0209
    WM_MOUSEWHEEL = 0x020A
    WM_MOUSELAST = 0x020A

    # 键盘事件类型
    WM_KEYFIRST = 0x0100
    WM_KEYDOWN = 0x0100
    WM_KEYUP = 0x0101
    WM_CHAR = 0x0102
    WM_DEADCHAR = 0x0103
    WM_SYSKEYDOWN = 0x0104
    WM_SYSKEYUP = 0x0105
    WM_SYSCHAR = 0x0106
    WM_SYSDEADCHAR = 0x0107
    WM_KEYLAST = 0x0108

    # VK_0~VK_9与ASCII '0'~'9' (0x30 -' : 0x39)相同
    # VK_A~VK_Z与ASCII 'A'~'Z' (0x41 -' : 0x5A)相同

    # 虚拟键号名称与其数值id
    vk_to_id_dict = {
        'VK_LBUTTON': 0x01, 'VK_RBUTTON': 0x02, 'VK_CANCEL': 0x03, 'VK_MBUTTON': 0x04,
        'VK_BACK': 0x08, 'VK_TAB': 0x09, 'VK_CLEAR': 0x0C, 'VK_RETURN': 0x0D, 'VK_SHIFT': 0x10,
        'VK_CONTROL': 0x11, 'VK_MENU': 0x12, 'VK_PAUSE': 0x13, 'VK_CAPITAL': 0x14, 'VK_KANA': 0x15,
        'VK_HANGEUL': 0x15, 'VK_HANGUL': 0x15, 'VK_JUNJA': 0x17, 'VK_FINAL': 0x18, 'VK_HANJA': 0x19,
        'VK_KANJI': 0x19, 'VK_ESCAPE': 0x1B, 'VK_CONVERT': 0x1C, 'VK_NONCONVERT': 0x1D, 'VK_ACCEPT': 0x1E,
        'VK_MODECHANGE': 0x1F, 'VK_SPACE': 0x20, 'VK_PRIOR': 0x21, 'VK_NEXT': 0x22, 'VK_END': 0x23,
        'VK_HOME': 0x24, 'VK_LEFT': 0x25, 'VK_UP': 0x26, 'VK_RIGHT': 0x27, 'VK_DOWN': 0x28,
        'VK_SELECT': 0x29, 'VK_PRINT': 0x2A, 'VK_EXECUTE': 0x2B, 'VK_SNAPSHOT': 0x2C, 'VK_INSERT': 0x2D,
        'VK_DELETE': 0x2E, 'VK_HELP': 0x2F, 'VK_LWIN': 0x5B, 'VK_RWIN': 0x5C, 'VK_APPS': 0x5D,
        'VK_NUMPAD0': 0x60, 'VK_NUMPAD1': 0x61, 'VK_NUMPAD2': 0x62, 'VK_NUMPAD3': 0x63, 'VK_NUMPAD4': 0x64,
        'VK_NUMPAD5': 0x65, 'VK_NUMPAD6': 0x66, 'VK_NUMPAD7': 0x67, 'VK_NUMPAD8': 0x68, 'VK_NUMPAD9': 0x69,
        'VK_MULTIPLY': 0x6A, 'VK_ADD': 0x6B, 'VK_SEPARATOR': 0x6C, 'VK_SUBTRACT': 0x6D, 'VK_DECIMAL': 0x6E,
        'VK_DIVIDE': 0x6F, 'VK_F1': 0x70, 'VK_F2': 0x71, 'VK_F3': 0x72, 'VK_F4': 0x73, 'VK_F5': 0x74,
        'VK_F6': 0x75, 'VK_F7': 0x76, 'VK_F8': 0x77, 'VK_F9': 0x78, 'VK_F10': 0x79, 'VK_F11': 0x7A,
        'VK_F12': 0x7B, 'VK_F13': 0x7C, 'VK_F14': 0x7D, 'VK_F15': 0x7E, 'VK_F16': 0x7F, 'VK_F17': 0x80,
        'VK_F18': 0x81, 'VK_F19': 0x82, 'VK_F20': 0x83, 'VK_F21': 0x84, 'VK_F22': 0x85, 'VK_F23': 0x86,
        'VK_F24': 0x87, 'VK_NUMLOCK': 0x90, 'VK_SCROLL': 0x91, 'VK_LSHIFT': 0xA0, 'VK_RSHIFT': 0xA1,
        'VK_LCONTROL': 0xA2, 'VK_RCONTROL': 0xA3, 'VK_LMENU': 0xA4, 'VK_RMENU': 0xA5, 'VK_PROCESSKEY': 0xE5,
        'VK_ATTN': 0xF6, 'VK_CRSEL': 0xF7, 'VK_EXSEL': 0xF8, 'VK_EREOF': 0xF9, 'VK_PLAY': 0xFA,
        'VK_ZOOM': 0xFB, 'VK_NONAME': 0xFC, 'VK_PA1': 0xFD, 'VK_OEM_CLEAR': 0xFE, 'VK_BROWSER_BACK': 0xA6,
        'VK_BROWSER_FORWARD': 0xA7, 'VK_BROWSER_REFRESH': 0xA8, 'VK_BROWSER_STOP': 0xA9, 'VK_BROWSER_SEARCH': 0xAA,
        'VK_BROWSER_FAVORITES': 0xAB, 'VK_BROWSER_HOME': 0xAC, 'VK_VOLUME_MUTE': 0xAD, 'VK_VOLUME_DOWN': 0xAE,
        'VK_VOLUME_UP': 0xAF, 'VK_MEDIA_NEXT_TRACK': 0xB0, 'VK_MEDIA_PREV_TRACK': 0xB1, 'VK_MEDIA_STOP': 0xB2,
        'VK_MEDIA_PLAY_PAUSE': 0xB3, 'VK_LAUNCH_MAIL': 0xB4, 'VK_LAUNCH_MEDIA_SELECT': 0xB5, 'VK_LAUNCH_APP1': 0xB6,
        'VK_LAUNCH_APP2': 0xB7, 'VK_OEM_1': 0xBA, 'VK_OEM_PLUS': 0xBB, 'VK_OEM_COMMA': 0xBC, 'VK_OEM_MINUS': 0xBD,
        'VK_OEM_PERIOD': 0xBE, 'VK_OEM_2': 0xBF, 'VK_OEM_3': 0xC0, 'VK_OEM_4': 0xDB, 'VK_OEM_5': 0xDC,
        'VK_OEM_6': 0xDD, 'VK_OEM_7': 0xDE, 'VK_OEM_8': 0xDF, 'VK_OEM_102': 0xE2, 'VK_PACKET': 0xE7
    }

    id_to_vk_dict = dict([(v, k) for k, v in vk_to_id_dict.items()])

    # 消息类型id与其名称
    msg_to_name_dict = {WM_MOUSEMOVE: 'mouse move', WM_LBUTTONDOWN: 'mouse left down',
                        WM_LBUTTONUP: 'mouse left up', WM_LBUTTONDBLCLK: 'mouse left double',
                        WM_RBUTTONDOWN: 'mouse right down', WM_RBUTTONUP: 'mouse right up',
                        WM_RBUTTONDBLCLK: 'mouse right double', WM_MBUTTONDOWN: 'mouse middle down',
                        WM_MBUTTONUP: 'mouse middle up', WM_MBUTTONDBLCLK: 'mouse middle double',
                        WM_MOUSEWHEEL: 'mouse wheel', WM_KEYDOWN: 'key down',
                        WM_KEYUP: 'key up', WM_CHAR: 'key char', WM_DEADCHAR: 'key dead char',
                        WM_SYSKEYDOWN: 'key sys down', WM_SYSKEYUP: 'key sys up',
                        WM_SYSCHAR: 'key sys char', WM_SYSDEADCHAR: 'key sys dead char'}

    @classmethod
    def msg_to_name(cls, msg):
        """
        将消息类型id转化为名称
        :param msg: 消息类型id
        :return: 消息名称
        """
        return cls.msg_to_name_dict.get(msg)

    @classmethod
    def vkey_to_id(cls, vkey):
        """
        虚拟按键名称转化为id号
        :param vkey: 虚拟按键名称
        :return: id值
        """
        return cls.vk_to_id_dict.get(vkey)

    @classmethod
    def id_to_name(cls, code):
        """
        将给定的键号id转为键位名称
        :param code: 键号id
        :return: 键位名称
        """
        if (0x30 <= code <= 0x39) or (0x41 <= code <= 0x5A):
            text = chr(code)
        else:
            text = cls.id_to_vk_dict.get(code)
            if text is not None:
                text = text[3:].title()
        return text

    @classmethod
    def is_vk(cls, code):
        """
        判断是否为虚拟键码
        :param code:
        :return:
        """
        return (code in cls.vk_to_id_dict.values()
                or (0x30 <= code <= 0x39) or (0x41 <= code <= 0x5A))

    @classmethod
    def is_keydown(cls, w_param):
        """
        判断事件是否为按下按键
        :param w_param:
        :return:
        """
        return w_param == HookConstants.WM_KEYDOWN or w_param == HookConstants.WM_SYSKEYDOWN

    @classmethod
    def is_keyup(cls, w_param):
        """
        判断事件是否为松开按键
        :param w_param:
        :return:
        """
        return w_param == HookConstants.WM_KEYUP or w_param == HookConstants.WM_SYSKEYUP


# 数字与名称互转的两个快捷接口
VK = HookConstants.vkey_to_id
ID = HookConstants.id_to_name


class KbEvent:
    """
    键盘事件的python封装
    """

    def __init__(self, n_code, w_param, l_param):
        self.__n_code = n_code
        self.__w_param = w_param
        self.__vk_code = l_param.vkCode
        self.__scan_code = l_param.scanCode
        self.__flags = l_param.flags
        self.__time = l_param.time

    def get_n_code(self):
        return self.__n_code

    def get_w_param(self):
        return self.__w_param

    def get_vk_code(self):
        return self.__vk_code

    def get_scan_code(self):
        return self.__scan_code

    def get_flags(self):
        return self.__flags

    def get_time(self):
        return self.__time

    n_code = property(get_n_code)
    w_param = property(get_w_param)
    vk_code = property(get_vk_code)
    scan_code = property(get_scan_code)
    flags = property(get_flags)
    time = property(get_time)


class KbHook(metaclass=SingletonType):
    """
    键盘钩子类
    """

    def __init__(self):
        self.__hook = None
        self.__handler = lambda event: 1

    def install_hook(self, hook_proc):
        """
        安装钩子
        :param hook_proc:
        :return:
        """
        if self.__hook:
            raise RuntimeError('hook has been installed')
        self.__hook = SetWindowsHookEx(
            win32con.WH_KEYBOARD_LL,
            hook_proc,
            None,
            0
        )
        if not self.__hook:
            raise RuntimeError('install hook failed')

    def uninstall_hook(self):
        """
        卸载钩子
        :return:
        """
        if not self.__hook:
            return
        UnhookWindowsHookEx(self.__hook)
        self.__hook = None

    def start(self):
        """
        开始获取消息列表
        :return:
        """
        GetMessage(wintypes.MSG(), 0, 0, 0)

    @staticmethod
    @HookProc
    def __hook_proc(n_code, w_param, l_param):
        """
        键盘回调钩子
        :param n_code:
        :param w_param:
        :param l_param:
        :return:
        """
        # 获取指针所指数据
        l_param = l_param.contents
        # 根据额外信息判断是否为模块发出的键盘事件
        if l_param.dwExtraInfo == 228:
            return CallNextHookEx(0, n_code, w_param, l_param)

        # 定义键盘事件类
        event = KbEvent(n_code, w_param, l_param)
        # 将事件实例传入真正的键盘事件处理函数
        res = KbHook().__handler(event)
        # 根据函数返回值决定是否拦截该键盘事件
        if res == 0:
            return CallNextHookEx(0, n_code, w_param, l_param)
        else:
            return 1

    def set_handler(self, func):
        """
        安装键盘回调钩子函数
        :param func:
        :return:
        """
        self.__handler = func
        if self.__hook:
            self.uninstall_hook()
        self.install_hook(self.__hook_proc)


class Condition:
    """
    条件类，用来判断键盘事件是否符合命令触发条件
    """
    def __init__(self, vk=None, alias=None, repeat=False, show=True):
        # 虚拟键号
        self.vk = vk
        # 命令别名
        self.alias = alias
        # 是否支持按住连续触发（默认根据松开按键触发命令）
        self.repeat = repeat
        self.show = show

    def inspect(self, event):
        """
        检查条件是否满足
        :param event:
        :return:
        """
        if event.vk_code != self.vk:
            return False
        if self.repeat and HookConstants.is_keydown(event.w_param):
            return True
        elif (not self.repeat) and HookConstants.is_keyup(event.w_param):
            return True
        else:
            return False


class Command:
    """
    命令类，包含触发条件和命令函数两个属性
    """
    def __init__(self, condition, func):
        self.condition = condition
        self.func = func


class Configuration:
    """
    配置类，用于保存一组快捷键配置信息
    """

    # 小键盘区虚拟键码
    KEYPAD = tuple([VK('VK_NUMPAD{}'.format(i)) for i in range(10)])
    # FUNCTION区虚拟键码
    FUNCTION = tuple([VK('VK_F{}'.format(i + 1)) for i in range(24)])
    # 模式
    _MODE = ('part', 'all')

    def __init__(self, name, mode='part', parts=None, excepts=None):
        """
        初始化实例
        :param name: 配置名称
        :param mode: 配置模式（全键盘或部分按键）
        :param parts: 部分按键虚拟键码列表（模式为all时无效）
        """
        self.name = name
        self.__mode = mode
        if parts is None:
            self.__parts = self.FUNCTION
        else:
            self.__parts = parts
        if excepts is None:
            self.__excepts = []
        else:
            self.__excepts = excepts

        self.__commands = {}

    def add_command(self, condition, func):
        """
        添加命令
        :param condition: 触发条件
        :param func: 命令函数
        :return:
        """
        if not isinstance(condition, Condition):
            raise TypeError('the first param type must be <class Condition>')
        if not (callable(func) or hasattr(func, '__call__')):
            raise TypeError('the second param must be callable')
        if not HookConstants.is_vk(condition.vk):
            raise ValueError('condition\'s attr vk is not a vkey code')
        key = str(condition.vk)
        if condition.repeat:
            key = key + '_p'
        else:
            key = key + '_r'
        self.__commands[key] = Command(condition, func)

    def get_commands(self):
        """
        获取命令列表
        :return:
        """
        return self.__commands

    def get_mode(self):
        """
        获取配置模式
        :return:
        """
        return self.__mode

    def set_mode(self, mode):
        """
        设置配置模式
        :param mode:
        :return:
        """
        if mode in self._MODE:
            self.__mode = mode
        else:
            raise ValueError('mode must be part or all')

    def get_parts(self):
        """
        获取快捷键区列表
        :return:
        """
        return self.__parts

    def set_parts(self, parts):
        """
        设置快捷键区列表
        :param parts:
        :return:
        """
        if not isinstance(parts, (list, tuple)):
            raise TypeError('must be list or tuple')
        for key in parts:
            if not HookConstants.is_vk(key):
                raise ValueError('invalid key code {}'.format(key))
        self.__parts = tuple(parts)

    def get_excepts(self):
        """
        获取快捷键区列表
        :return:
        """
        return self.__excepts

    def set_excepts(self, excepts):
        """
        设置快捷键区列表
        :param excepts:
        :return:
        """
        if not isinstance(excepts, (list, tuple)):
            raise TypeError('must be list or tuple')
        for key in excepts:
            if not HookConstants.is_vk(key):
                raise ValueError('invalid key code {}'.format(key))
        self.__excepts = tuple(excepts)

    mode = property(get_mode, set_mode)
    parts = property(get_parts, set_parts)
    excepts = property(get_excepts, set_excepts)
    commands = property(get_commands)


class Core(metaclass=SingletonType):
    """
    键盘事件核心处理类，包含主要响应逻辑
    """
    def __init__(self, pool, window=None, switch_key=None, active_time=500):
        """
        初始化实例
        :param pool: 线程池
        :param window: 窗体实例
        :param switch_key: 切换键，长按进入退出，短按切换配置
        :param active_time: 长按时间
        """
        self.__pool = pool
        self.__window = window
        if switch_key is None:
            self.__switch_key = VK('VK_F1')
        else:
            self.__switch_key = switch_key
        self.__active_time = active_time
        self.__configs = {}
        self.__last_key = None
        self.__current_config = None

        self.__flag_run = False
        self.__flag_start = False
        self.__repeat_key = None

    def get_flag_run(self):
        return self.__flag_run

    def set_flag_run(self, flag):
        self.__flag_run = flag

    def get_current_config(self):
        return list(self.__configs.keys()).index(self.__current_config)

    def set_current_config(self, ind):
        if not isinstance(ind, int):
            raise TypeError('index must be integer')
        if ind >= len(self.__configs):
            raise RuntimeError('index out of range')
        self.__current_config = list(self.__configs.keys())[ind]

    def add_window(self, window):
        """
        添加窗体实例
        :param window:
        :return:
        """
        if not isinstance(window, Notification):
            raise TypeError('param window must be <class Notification>')
        self.__window = window

    def get_configs(self):
        """
        获取配置字典
        :return:
        """
        return self.__configs

    def add_config(self, config):
        """
        添加配置实例
        :param config:
        :return:
        """
        if not isinstance(config, Configuration):
            raise TypeError('the first param type must be <class Configuration>')
        self.__configs[config.name] = config

    def __next_config(self):
        """
        获取下一配置的index
        :return:
        """
        if not len(self.__configs):
            raise RuntimeError('there is no config')
        keys = tuple(self.__configs.keys())
        if self.__current_config is None:
            self.__current_config = keys[0]
        else:
            try:
                index = keys.index(self.__current_config)
            except ValueError:
                index = len(keys)
            index = (index + 1) % len(keys)
            self.__current_config = keys[index]

    def __call__(self, event):
        """
        回调函数
        :param event:
        :return:
        """
        # 设置__last_key属性
        if self.__last_key is None:
            self.__last_key = event

        # 切换按键相关逻辑
        if event.vk_code == self.__switch_key:
            # print('x')
            if HookConstants.is_keyup(event.w_param):
                if self.__last_key.vk_code == self.__switch_key \
                        and (event.time - self.__last_key.time) >= self.__active_time:
                    self.__flag_start = not self.__flag_start
                    self.__last_key = event
                    if self.__flag_start:
                        if self.__current_config is None:
                            self.__next_config()
                        self.__show_message('配置\n{}'.format(self.__configs[self.__current_config].name))
                    else:
                        self.__show_message('普通键盘')
                    return 1

        # 更改__last_key属性
        if self.__last_key.vk_code != event.vk_code:
            self.__last_key = event
        elif self.__last_key.vk_code == event.vk_code and self.__last_key.w_param != event.w_param:
            self.__last_key = event

        # 判断是否已进入宏模式
        if not self.__flag_start:
            if event.vk_code == self.__switch_key:
                # print('pass')
                return 1
            return 0

        config = self.__configs[self.__current_config]

        # 部分模式下放行所有非快捷区域键盘输入
        if config.mode == 'part' and (event.vk_code not in config.parts) and event.vk_code != self.__switch_key:
            # print('no')
            return 0

        # 全部模式下放行所有非快捷区域键盘输入
        if config.mode == 'all' and (event.vk_code in config.excepts) and event.vk_code != self.__switch_key:
            return 0

        # 判断是否有命令正在执行中
        if self.__flag_run:
            if event.vk_code == self.__repeat_key and HookConstants.is_keyup(event.w_param):
                self.__repeat_key = None
                if (str(event.vk_code) + '_r') not in config.commands.keys():
                    return 1
            self.__show_message('执行中,\n请等待!')
            return 1

        # 切换配置相关逻辑
        if event.vk_code == self.__switch_key and HookConstants.is_keyup(event.w_param):
            # print('change')
            self.__next_config()
            self.__show_message('配置\n{}'.format(self.__configs[self.__current_config].name))
            return 1
        elif event.vk_code == self.__switch_key and HookConstants.is_keydown(event.w_param):
            return 1

        # 判断按键是否触发相关命令函数
        for command in config.commands.values():
            if command.condition.inspect(event):
                if command.condition.show:
                    if command.condition.alias is None:
                        self.__show_message('命令\n{}'.format(ID(command.condition.vk)))
                    else:
                        self.__show_message('命令\n{}'.format(command.condition.alias))
                if command.condition.repeat:
                    self.__repeat_key = event.vk_code

                def callback():
                    self.__flag_run = True
                    try:
                        command.func()
                    except Exception as e:
                        print(repr(e))
                        self.__show_message('Error')
                    self.__flag_run = False

                self.__pool.submit(callback)
                return 1
            elif command.condition.vk == event.vk_code and HookConstants.is_keyup(event.w_param) \
                    and ((str(event.vk_code) + '_r') not in config.commands.keys()):
                return 1
        if HookConstants.is_keyup(event.w_param):
            self.__show_message('命令\n空')
        return 1

    def __show(self, text, duration):
        """
        设置窗体显示文字并显示一段时间
        :param text:
        :param duration:
        :return:
        """
        self.__window.text = text
        self.__window.show()
        time.sleep(duration)
        self.__window.hide()

    def __show_message(self, text, duration=SLEEP_TIME):
        """
        后台线程运行窗体显示函数
        :param text:
        :param duration:
        :return:
        """
        self.__pool.submit(self.__show, text, duration)

    flag_run = property(get_flag_run, set_flag_run)
    current_config = property(get_current_config, set_current_config)
