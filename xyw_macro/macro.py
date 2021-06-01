from concurrent.futures import ThreadPoolExecutor
import time

from xyw_macro.hook import KbHook, Core
from xyw_macro.notify import Notification
from xyw_macro.utils import SingletonType


class Macro(metaclass=SingletonType):
    """
    最终封装好的宏命令类，只能在主线程中定义使用
    """
    def __init__(self, switch_key=None, max_workers=20, start_message='xyw_macro\n已启动'):
        """
        初始化实例
        :param switch_key: 模式切换键
        :param max_workers: 线程池最大线程数
        :param start_message: 启动时的显示语
        """
        self.pool = ThreadPoolExecutor(max_workers=max_workers)
        self.__message = start_message
        self.__core = Core(self.pool, switch_key=switch_key)

    def __sub_thread(self):
        """
        监听拦截键盘输入的子进程
        :return:
        """
        kb = KbHook()
        kb.set_handler(self.__core)
        kb.start()

    def __start_gui(self, window):
        """
        启动特效
        :param window:
        :return:
        """
        window.text = self.__message
        window.show()
        time.sleep(2)
        window.hide()

    def add_command_to_all(self, condition, func):
        """
        向所有现有配置中添加同一命令
        :param condition:
        :param func:
        :return:
        """
        for config in self.__core.get_configs().values():
            config.add_command(condition, func)

    def add_config(self, config):
        """
        添加配置
        :param config:
        :return:
        """
        self.__core.add_config(config)

    def run(self, fg='white', bg='black'):
        """
        启动键盘宏
        :return:
        """
        window = Notification(fg=fg, bg=bg)
        self.__core.add_window(window)
        self.pool.submit(self.__start_gui, window)
        self.pool.submit(self.__sub_thread)
        window.run()

    def get_flag_run(self):
        """
        获取运行状态
        :return:
        """
        return self.__core.flag_run

    def set_flag_run(self, flag):
        """
        设置运行状态
        :param flag:
        :return:
        """
        self.__core.flag_run = flag

    def get_current_config(self):
        """
        获取当前配置index
        :return:
        """
        return self.__core.current_config

    def set_current_config(self, ind):
        """
        设置配置index
        :param ind:
        :return:
        """
        self.__core.current_config = ind

    flag_run = property(get_flag_run, set_flag_run)
    current_config = property(get_current_config, set_current_config)
