import time
import random
import uuid
import ctypes
import wave

import win32clipboard
import win32con
import win32gui
import win32api
import webbrowser
import simpleaudio as sa
import os
import subprocess
from xyw_macro.macro import Macro
from xyw_macro.win32 import send_kb_event, send_unicode, change_language_layout
from xyw_macro.hook import VK
from xyw_macro.contants import SLEEP_TIME, ZH_LANGUAGE, EN_LANGUAGE, BAIDU_LANGUAGE


def press_key(v_key):
    """
    按下某个按键
    :param v_key:
    :return:
    """
    return send_kb_event(v_key, True)


def release_key(v_key):
    """
    松开某个按键
    :param v_key:
    :return:
    """
    return send_kb_event(v_key, False)


def touch_key(v_key):
    """
    按键一次
    :param v_key:
    :return:
    """
    press_key(v_key)
    random_sleep(0.01)
    release_key(v_key)


def random_sleep(seconds=SLEEP_TIME, range_percent=None):
    """
    睡眠指定秒数，随机浮动指定比例
    :param seconds: 睡眠时间，单位s
    :param range_percent: 浮动比例
    :return: 实际睡眠时间，单位s
    """
    if range_percent is None:
        range_percent = 0
    ran = seconds * range_percent
    sleep_time = random.uniform(seconds - ran, seconds + ran)
    time.sleep(sleep_time)
    return sleep_time


def press_copy():
    press_key(VK('VK_CONTROL'))
    press_key(ord('C'))
    release_key(VK('VK_CONTROL'))
    release_key(ord('C'))
    time.sleep(0.2)


def get_clipboard_files():
    """
    获取剪切板中文件的文件地址
    :return: 地址元组
    """
    press_copy()
    if win32clipboard.IsClipboardFormatAvailable(win32con.CF_HDROP) == 0:
        return
    win32clipboard.OpenClipboard()
    data = win32clipboard.GetClipboardData(win32con.CF_HDROP)
    win32clipboard.CloseClipboard()
    return data


def run_cmd(cmd):
    """
    运行cmd命令，可以用来打开应用快捷方式
    :param cmd:
    :return:
    """
    # os.system(cmd)
    try:
        subprocess.run(cmd, shell=True, timeout=1)
    except subprocess.TimeoutExpired:
        pass


def open_file(path):
    """
    使用系统默认程序打开文件或文件夹
    :param path:
    :return:
    """
    os.system('explorer.exe %s' % path)


def input_chars(chars):
    """
    模拟键盘输入字符串
    :param chars:
    :return:
    """
    for char in chars:
        send_unicode(char)
        random_sleep(0.01)


def open_url(url, browser_path=None):
    """
    使用浏览器打开网址
    :param url: 网址
    :param browser_path: 非默认浏览器程序地址
    :return:
    """
    if browser_path is not None:
        webbrowser.register('browser', None, webbrowser.BackgroundBrowser(browser_path))
        webbrowser.get('browser').open(url, new=0, autoraise=True)
    else:
        webbrowser.open(url, new=0, autoraise=True)


def get_mac():
    """
    获取本机mac地址（单网卡）
    :return: 12位mac地址
    """
    mac = uuid.uuid1().hex[-12:].upper()
    tem = [mac[i * 2: i * 2 + 2] for i in range(len(mac) // 2)]
    mac = '-'.join(tem)
    return mac


def confirm_mac(mac):
    """
    验证所给mac地址是否与本机匹配
    :param mac:
    :return:
    """
    if isinstance(mac, str):
        return mac.upper() == get_mac()
    elif isinstance(mac, (list, tuple)):
        local_mac = get_mac()
        for item in mac:
            if item.upper() == local_mac:
                return True
        return False
    else:
        raise TypeError('param mac must be string or list')


def change_to_zh():
    """
    切换为中文输入法，优先切换为百度输入法
    :return:
    """
    tem = change_language_layout(BAIDU_LANGUAGE)  # 适配win7系统
    if not tem:
        change_language_layout(ZH_LANGUAGE)


def change_to_en():
    """
    切换为英文输入法
    :return:
    """
    change_language_layout(EN_LANGUAGE)


def change_to_config(index):
    """
    切换到指定配置
    :param index: 配置索引（0开始）
    :return:
    """
    Macro().current_config = index


def enable_network_adapter(name):
    """
    启用指定的网络适配器
    :param name: 网卡名称
    :return:
    """
    subprocess.run('netsh interface set interface "%s" enabled' % name)


def disable_network_adapter(name):
    """
    禁用指定的网络适配器
    :param name: 网卡名称
    :return:
    """
    subprocess.run('netsh interface set interface "%s" disabled' % name)


def minimize_window():
    """
    最小化当前窗口
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)


def maximize_window():
    """
    最大化当前窗口
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)


def close_window():
    """
    关闭当前窗口
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)


def restore_window():
    """
    还原当前窗口
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)


def change_window_state():
    """
    更改窗口的状态，更改顺序为最小化、正常、最大化循环
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    if win32gui.IsIconic(hwnd) != 0:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
    elif ctypes.windll.user32.IsZoomed(hwnd) != 0:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)
    else:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)


def topmost_window():
    """
    置顶当前窗口
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                          win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)


def notopmost_window():
    """
    取消置顶当前窗口
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                          win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)


def change_window_topmost():
    """
    更改当前窗口置顶状态
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    if win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOPMOST == 0:
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                              win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)
    else:
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)


def set_window_alpha(alpha):
    """
    设置当前窗口透明度
    :param alpha: 0-255之间
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                           win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
    win32gui.SetLayeredWindowAttributes(hwnd, 0, alpha, win32con.LWA_ALPHA)


def change_window_alpha():
    """
    改变当前窗口透明度，四档可调
    :return:
    """
    hwnd = win32gui.GetForegroundWindow()
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                           win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_LAYERED)
    try:
        tem = win32gui.GetLayeredWindowAttributes(hwnd)
    except Exception:
        tem = (0, 0, 0)
    if tem[2] == win32con.LWA_ALPHA:
        alpha = tem[1]
    else:
        alpha = 255
    alpha = (alpha + 64) % 256
    win32gui.SetLayeredWindowAttributes(hwnd, 0, alpha, win32con.LWA_ALPHA)


def play_wave(wave_path):
    """
    播放音频
    :param wave_path: 音频地址
    :return:
    """
    wave_obj = sa.WaveObject.from_wave_file(wave_path)
    play_obj = wave_obj.play()
    play_obj.wait_done()


if __name__ == '__main__':
    # print(random_sleep(0.2, 0.25))
    # print(get_clipboard_files())
    # print(type(get_clipboard_files()))
    # press_copy()
    # print(run_cmd(r"C:\Program Files (x86)\Thunder Network\Thunder\Program\ThunderStart.exe"))
    # print(open_file('hook.py'))
    # input_chars('今天是个好人797987/*/-*-*+')
    # open_url('www.baidu.com', r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
    # open_url('www.baidu.com')
    # print(run_cmd(r"D:\Program Files (x86)\Everything\Everything.exe"))
    # print(run_cmd(r"C:\Program Files (x86)\WXWork\WXWork.exe"))
    # print(get_mac())
    # print(confirm_mac('0C-9D-92-0F-42-6E'))
    # change_to_en()
    # change_language_layout(-0x1fddf7fc)
    # enable_network_adapter('本地连接')
    # disable_network_adapter('本地连接')
    # import threading
    # threading.Thread(target=play_wave, args=('../警报声.wav',)).start()
    # play_wave('../警报声.wav')
    # play_wave('../警报声.wav')
    # with open('../警报声.wav', 'rb') as f:
    #     data = f.read()[::-1]
    # with open('test.wav', 'wb') as f:
    #     f.write(data)
    # play_wave('test.wav')
    # wav = wave.open('../警报声.wav')
    # wav_obj = sa.WaveObject.from_wave_read(wav)
    # wav_obj.play().wait_done()
    # wav_obj.play().wait_done()
    # wav_obj.play().wait_done()
    pass
