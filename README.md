# 说明文档

## 快速开始

这个库的灵感来自于一个全键可编程的小键盘，那个键盘可以通过按键切换不同的配置，9个按键可以当成36个快捷键使用。本来想买的，但是想想对我真没啥用，就干脆写了这么一个库模拟体验一下。

```python
import time
from xyw_macro import *


def func1():
    """命令函数1"""
    time.sleep(4)
    print('func1')


def func2():
    """命令函数2"""
    print('func2')

    
def func3():
    """命令函数3"""
    print('func3')

# 创建宏实例
macro = Macro()  # 默认配置切换按键为F1，长按可以进入或退出键盘宏模式，宏模式下短按可以循环切换配置
# 创建配置实例1
config1 = Configuration('01', 'part', Configuration.FUNCTION)
# 向配置1中添加命令
config1.add_command(Condition(VK('VK_F3'), '打印'), func1)
# 创建配置实例2
config2 = Configuration('02', 'all', Configuration.FUNCTION)
# 向配置2中添加命令
config2.add_command(Condition(VK('VK_F2'), '打印'), func2)

# 向宏实例中添加之前定义的配置实例
macro.add_config(config1)
macro.add_config(config2)

# 向宏实例中所有现有的配置实例中添加同一命令
macro.add_command_to_all(Condition(VK('VK_F4'), '打印'), func3)

# 运行键盘宏
macro.run()

```

## command模块

### 模拟键盘输入

以按下Ctrl+C为例，代码如下：

```python
from xyw_macro.command import *

press_key(VK('VK_CONTROL'))
press_key(ord('C'))
release_key(VK('VK_CONTROL'))
release_key(ord('C'))
time.sleep(0.2)  # 对于这种功能快捷键最好加上0.2s的睡眠时间，等待系统响应完成
```

注意：press_key与release_key函数必须成对使用，同时请勿使用其他第三方库中的键盘输入函数代替，非本模块中的键盘输入函数会被拦截当做普通的键盘输入处理。有时游戏中为了避免脚本检测，需要模拟真人按键，此时可以在每个按键的按松动作之间插入random_sleep函数，示例如下：

```python
press_key(ord('C'))
random_sleep(0.2, 0.5)  # 休眠时间波动范围为0.2*(1-0.5)~0.2*(1+0.5)
release_key(ord('C'))
```

### 输入字符串

模拟键盘输入unicode字符串：

```python
input_chars('pip install xyw-macro --upgrade')
```

注意：为安全起见，请勿使用此函数输入账号密码，如必须使用，可以参见安全使用一节。

### 获取选中文件绝对地址

直接调用win32的api很难直接获得资源管理器中选中文件的绝对地址，所以在这里采用了一种折中的做法，首先调用Ctrl+C复制选中文件到系统剪切板，然后读取剪切板中的文件名信息：

```python
print(get_clipboard_files())

# 单个文件夹
>>>('C:\\Users\\Administrator\\Desktop\\keyboard',)
# 单个文件
>>>('C:\\Users\\Administrator\\Desktop\\keyboard\\setup.py',)
# 多个文件或文件夹
>>>('C:\\Users\\Administrator\\Desktop\\keyboard\\xyw_macro', 'C:\\Users\\Administrator\\Desktop\\keyboard\\LICENSE', 'C:\\Users\\Administrator\\Desktop\\keyboard\\README.md', 'C:\\Users\\Administrator\\Desktop\\keyboard\\setup.py')
```

注意：此函数设计的初衷是为了方便实现一键处理文件的宏功能。

### 打开应用快捷方式

以打开企业微信为例：

```python
run_cmd(r"C:\Program Files (x86)\WXWork\WXWork.exe")
```

注意：此处的地址可以通过快捷方式的属性查看。

### 打开文件或文件夹

```python
open_file('hook.py')
```

### 打开网址

```python
open_url('www.baidu.com', r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
```

注意：默认情况下只需要第一个参数即可，此时会以系统默认浏览器打开地址。第二个为可选参数，可以指定需要使用的浏览器的程序地址。

### 切换输入法

```python
change_to_en()  # 切换到英文键盘布局
change_to_zh()  # 切换到中文键盘布局
```

### 切换配置

如果总共有两套配置，即通过add_config方法添加了两次配置信息，可以通过配置的index手动切换配置：

```python
change_to_config(0)  # 切换到第一套配置
```

### 更改当前窗口状态

```python
minimize_window()  # 最小化当前窗口
maximize_window()  # 最大化当前窗口
restore_window()  # 还原当前窗口
close_window()  # 关闭当前窗口

topmost_window()  # 置顶当前窗口
notopmost_window()  # 取消置顶当前窗口

set_window_alpha(180)  # 更改当前窗口透明度，0-255

change_window_state()  # 更改当前窗口状态，更改顺序为最小化、正常、最大化循环
change_window_topmost()  # 更改当前窗口置顶状态
change_window_alpha()  # 循环切换当前窗口透明度，有63、127、191、255四档
```

### 播放wav格式音频

此功能是对simpleaudio库的简单封装，只支持wav格式音频文件，如需更多文件格式支持及更复杂的播放功能可以使用pyaudio等第三方库自行实现：

```python
play_wave('test.wav')
```

## 交互框

### 确认框

tkinter中本身就有一系列消息框，在编写命令函数时都可以正常使用。对于确认框，在引入本模块后可以通过装饰器更方便的使用：

```python
from xyw_macro import *


@confirm_box('确定打开本网站吗？')
def open_sec_url(url):
    open_url(url)
    change_to_config(0)


config3 = Configuration('地址', 'all', excepts=[VK('VK_RETURN'), VK('VK_ESCAPE')])
config3.add_command(
	Condition(ord('B'), '百度'),
	lambda: open_sec_url(r'https://www.baidu.com/')
)
```

### 输入框

对于命令函数中的参数输入也可以使用装饰器轻松地完成：

```python
@input_box(
    InputField(name='输入框', type='entry', default='默认值', focus=True),
    InputField('文件选择', 'file', '默认值'),
    InputField('文件夹选择', 'dir', '默认值'),
    InputField(name='下拉框', type='combobox', default=0, options=['1', '2'])
)
def test(*args):
    print(args)


config3.add_command(
    Condition(VK('VK_F7'), '测试'),
    test
)
```

注意：input_box装饰器目前仅支持四种类型的输入框，且所有输入框最终输出的参数皆为字符串形式，数据类型需要在命令函数中自行转换。InputField共有5个参数，name：参数名（提示用），type：输入框类型，default：默认值（combobox类型的默认值为int类型，代表options参数中值的index），options：下拉框选项（combobox类型独占参数），focus：是否获得焦点（用于设置初始焦点位置）。

## 安全使用

如果脚本中包含部分隐私信息，可以在脚本中添加mac地址验证，同时将脚本打包编译为pwd格式文件，脚本内容示例如下：

```python
# macro.py
from xyw_macro import *
from xyw_macro.command import *


macro = Macro()

if confirm_mac('0C-9D-92-0F-42-65'):
    def func():
        print('True function')

    config = Configuration('true', 'part', Configuration.FUNCTION)
    config.add_command(Condition(VK('VK_F2'), '打印'), func)
    macro.add_config(config)
else:
    config = Configuration('false', 'all')
    macro.add_config(config)

```

将macro.py文件编译为pwd格式后需要再写一个入口脚本，重命名为pyw后缀后即可无命令窗后台运行：

```python
# run.pyw
from macro3 import macro

if __name__ == '__main__':
    macro.run()

```

注意：此方法并不是完全安全，只能说是降低了安全风险，不过一个宏命令脚本也不会有人专门去研究。说到底也就是和门锁一样，防君子不防小人。

## 实例演示

```python
# 内置模块
import os
import urllib
import tkinter.messagebox

# 第三方库
import img2pdf
import fitz
import pyperclip
import win32gui

# 本模块
from xyw_macro import *
from xyw_macro.command import *


# 创建实例，设置切换按键为右alt
macro = Macro(VK('VK_RMENU'))

# 检测电脑mac地址是否匹配
if get_mac() == '84-E6-F6-0D-0E-D5':
    def search_selected_text():
        """
        在浏览器中百度搜索选定的文本
        :return:
        """
        touch_keys([VK('VK_CONTROL'), ord('C')])
        random_sleep(0.2)
        text = pyperclip.paste()
        search_text = urllib.parse.quote(text.encode('gbk'))
        url = 'https://www.baidu.com/s?wd=' + search_text
        open_nor_url(url)


    def touch_keys(keys):
        """
        模拟同时按下多个按键
        :param keys: 按键虚拟键号列表
        :return:
        """
        if isinstance(keys, int):
            keys = [keys]
        if not isinstance(keys, (list, tuple)):
            raise TypeError('param keys must be int or list')
        for key in keys:
            press_key(key)
        for key in keys:
            release_key(key)


    def input_text(text):
        """
        输入文本后切换回第一套配置
        :param text:
        :return:
        """
        input_chars(text)
        change_to_config(0)


    def password_q():
        """
        输入邮箱1174543101@qq.com
        :return:
        """
        change_to_en()
        touch_key(VK('VK_NUMPAD1'))
        touch_key(VK('VK_NUMPAD1'))
        touch_key(VK('VK_NUMPAD7'))
        touch_key(VK('VK_NUMPAD4'))
        touch_key(VK('VK_NUMPAD5'))
        touch_key(VK('VK_NUMPAD4'))
        touch_key(VK('VK_NUMPAD3'))
        touch_key(VK('VK_NUMPAD1'))
        touch_key(VK('VK_NUMPAD0'))
        touch_key(VK('VK_NUMPAD1'))
        press_key(VK('VK_SHIFT'))
        random_sleep(0.01)
        touch_key(ord('2'))
        random_sleep(0.01)
        release_key(VK('VK_SHIFT'))
        touch_key(ord('Q'))
        touch_key(ord('Q'))
        touch_key(VK('VK_DECIMAL'))
        touch_key(ord('C'))
        touch_key(ord('O'))
        touch_key(ord('M'))
        change_to_zh()
        change_to_config(0)


    @confirm_box('确定打开本网站吗？')
    def open_sec_url(url):
        """
        打开网址，确认框确认后打开
        :param url:
        :return:
        """
        open_url(url)
        change_to_config(0)


    def open_nor_url(url):
        """
        打开网址
        :param url:
        :return:
        """
        open_url(url)
        change_to_config(0)


    def open_path(path):
        """
        在资源管理器中打开指定地址
        :param path:
        :return:
        """
        open_file(path)
        change_to_config(0)


    def img_to_pdf():
        """
        将文件夹中所有图片合成到一个pdf文件中，并以文件夹名称命名
        :return:
        """
        random_sleep()
        target = get_clipboard_files()
        if target is None:
            target = InputBox('输入框', InputField('图片', 'dir')).show()
            if target is None:
                raise RuntimeError('there is no file selected')
        if len(target) == 1 and os.path.isdir(target[0]):
            dir_path = target[0]
        else:
            raise RuntimeError('target must be a dir')

        imgs = []
        for root, dirs, files in os.walk(dir_path):
            for file in files:
                filename, ext = os.path.splitext(file)
                ext = ext.lower()
                if ext in ['.png', '.jpg', '.jpeg']:
                    imgs.append(filename + ext)
            if len(imgs) == 0:
                raise RuntimeError('there is no image in this dir')
        imgs = [os.path.join(dir_path, img) for img in imgs]
        imgs.sort()

        save_dir, pdf_filename = os.path.split(dir_path)
        if os.path.isfile(os.path.join(save_dir, pdf_filename + '.pdf')):
            index = 1
            while True:
                if not os.path.isfile(os.path.join(save_dir, pdf_filename + '_%d.pdf' % index)):
                    pdf_filename = pdf_filename + '_%d.pdf' % index
                    break
                else:
                    index += 1
        else:
            pdf_filename = pdf_filename + '.pdf'

        page_size = (img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297))
        layout_fun = img2pdf.get_layout_fun(page_size)
        try:
            with open(os.path.join(save_dir, pdf_filename), 'wb') as f:
                f.write(img2pdf.convert(imgs, layout_fun=layout_fun))
        except img2pdf.AlphaChannelError:
            os.remove(os.path.join(save_dir, pdf_filename))
            raise RuntimeError('Image contains transparency which cannot be retained in PDF')
        except Exception as e:
            raise RuntimeError(str(e))

        change_to_config(0)
        tkinter.messagebox.showinfo('提示', '合并pdf成功！')


    def pdf_to_img():
        """
        将pdf文件每一页保存为同一文件夹中的jpg图片，文件夹名为pdf文件名
        :return:
        """
        random_sleep()
        target = get_clipboard_files()
        if target is None:
            target = InputBox('输入框', InputField('PDF文件', 'file')).show()
            if target is None:
                raise RuntimeError('there is no file selected')
        if len(target) == 1 and os.path.isfile(target[0]):
            file_path = target[0]
        else:
            raise RuntimeError('target must be a file')

        dir_path, file = os.path.split(file_path)
        filename, ext = os.path.splitext(file)
        if ext.lower() != '.pdf':
            raise RuntimeError('target file type must be pdf')
        save_dir = os.path.join(dir_path, filename)
        if os.path.isdir(save_dir):
            index = 1
            while True:
                save_dir = os.path.join(dir_path, filename + '_%d' % index)
                if not os.path.isdir(save_dir):
                    break
                else:
                    index += 1
                    print(index)
        os.makedirs(save_dir)

        try:
            trans = fitz.Matrix(2, 2)
            with fitz.open(file_path) as pdf:
                for pg in range(pdf.pageCount):
                    page = pdf[pg]
                    pix = page.get_pixmap(alpha=False, matrix=trans)
                    img_name = filename + '_%s.jpg' % str(pg+1).zfill(4)
                    pix.save(os.path.join(save_dir, img_name))
        except Exception:
            os.removedirs(save_dir)

        change_to_config(0)
        tkinter.messagebox.showinfo('提示', 'pdf转换图片成功！')


    @confirm_box('确定要切换到工作网络吗？')
    def change_to_work_mode():
        """
        切换到工作网络，即禁用无线网卡，启用有线网卡
        win7可用，其他未测试
        :return:
        """
        enable_network_adapter('本地连接')
        disable_network_adapter('无线网络连接')
        change_to_config(0)


    @confirm_box('确定要切换到个人网络吗？')
    def change_to_personal_mode():
        """
        切换到个人网络，即启用有线网卡，禁用无线网卡
        :return:
        """
        enable_network_adapter('无线网络连接')
        disable_network_adapter('本地连接')
        change_to_config(0)


    config1 = Configuration('默认', 'part', Configuration.FUNCTION)
    config1.add_command(
        Condition(VK('VK_F5'), '音量+', True),
        lambda: touch_keys(VK('VK_VOLUME_UP'))
    )
    config1.add_command(
        Condition(VK('VK_F6'), '音量-', True),
        lambda: touch_keys(VK('VK_VOLUME_DOWN'))
    )
    config1.add_command(
        Condition(VK('VK_F7'), '上一桌面', True),
        lambda: touch_keys([VK('VK_CONTROL'), VK('VK_LWIN'), VK('VK_LEFT')])
    )
    config1.add_command(
        Condition(VK('VK_F8'), '下一桌面', True),
        lambda: touch_keys([VK('VK_CONTROL'), VK('VK_LWIN'), VK('VK_RIGHT')])
    )
    config1.add_command(
        Condition(VK('VK_F9'), '关闭窗口'),
        close_window
    )
    config1.add_command(
        Condition(VK('VK_F10'), '窗口状态'),
        change_window_state
    )
    config1.add_command(
        Condition(VK('VK_F11'), '窗口置顶'),
        change_window_topmost
    )
    config1.add_command(
        Condition(VK('VK_F12'), '窗口透明'),
        change_window_alpha
    )

    config2 = Configuration('地址', 'all', excepts=[VK('VK_RETURN'), VK('VK_ESCAPE')])
    config2.add_command(
        Condition(ord('B'), '百度'),
        lambda: open_nor_url(r'https://www.baidu.com/')
    )
    config2.add_command(
        Condition(ord('L'), 'bili'),
        lambda: open_nor_url(r'https://www.bilibili.com/')
    )
    config2.add_command(
        Condition(ord('T'), '翻译'),
        lambda: open_nor_url(r'https://translate.google.cn/')
    )
    config2.add_command(
        Condition(ord('Z'), '知乎'),
        lambda: open_nor_url(r'https://www.zhihu.com/')
    )
    config2.add_command(
        Condition(ord('D'), '斗鱼'),
        lambda: open_nor_url(r'https://www.douyu.com/room/my_follow/')
    )
    config2.add_command(
        Condition(ord('N'), 'NGA'),
        lambda: open_nor_url(r'https://bbs.nga.cn/')
    )
    config2.add_command(
        Condition(ord('P'), 'pypi'),
        lambda: open_nor_url(r'https://pypi.org/')
    )
    config2.add_command(
        Condition(ord('M'), 'bimi'),
        lambda: open_sec_url(r'http://www.bimiacg.com/')
    )
    config2.add_command(
        Condition(ord('G'), 'G盘'),
        lambda: open_path(r'G:')
    )
    config2.add_command(
        Condition(ord('F'), 'F盘'),
        lambda: open_path(r'F:')
    )

    config3 = Configuration('文字', 'all', excepts=[VK('VK_RETURN'), VK('VK_ESCAPE')])
    config3.add_command(
        Condition(ord('P')),
        lambda: input_text('python setup.py sdist bdist_wheel')
    )
    config3.add_command(
        Condition(ord('T')),
        lambda: input_text('twine upload dist/*')
    )
    config3.add_command(
        Condition(ord('N')),
        lambda: input_text('nuitka --mingw64 --module --show-progress --output-dir=o ')
    )
    config3.add_command(
        Condition(ord('Q')),
        password_q
    )

    config4 = Configuration('脚本', 'all', excepts=[VK('VK_RETURN'), VK('VK_ESCAPE')])
    config4.add_command(
        Condition(ord('P'), 'PDF'),
        img_to_pdf
    )
    config4.add_command(
        Condition(ord('I'), 'IMAGE'),
        pdf_to_img
    )
    config4.add_command(
        Condition(ord('S'), '搜索'),
        search_selected_text
    )
    config4.add_command(
        Condition(ord('W'), '工作网络'),
        change_to_work_mode
    )
    config4.add_command(
        Condition(ord('Q'), '个人网络'),
        change_to_personal_mode
    )

    macro.add_config(config1)
    macro.add_config(config2)
    macro.add_config(config3)
    macro.add_config(config4)

    macro.add_command_to_all(Condition(VK('VK_F1'), '配置默认'), lambda index=0: change_to_config(index))
    macro.add_command_to_all(Condition(VK('VK_F2'), '配置地址'), lambda index=1: change_to_config(index))
    macro.add_command_to_all(Condition(VK('VK_F3'), '配置文本'), lambda index=2: change_to_config(index))
    macro.add_command_to_all(Condition(VK('VK_F4'), '配置脚本'), lambda index=3: change_to_config(index))

else:
    config = Configuration('功能', 'all')
    macro.add_config(config)

```

注：为了方便切换，此处将切换键设置为了右alt，如果有可编程鼠标的话，可以将右alt映射到鼠标上，切换起来会更加方便。

# Release Notes

## 0.0.1

- 初次发布，实现了键盘宏的基本功能

## 0.0.2

- 将GUI框架由pyqt5换为tkinter，提高兼容性
- 新增command模块，包含部分常用键盘宏命令函数
- 新增部分源码注释
- 修复了部分bug

## 0.0.3

- 优化了Macro类
- 新增部分源码注释
- command模块中新增了部分常用函数

## 0.0.4

- 更改了提示窗口尺寸，缩小为原来的一半
- 更改了提示窗口显示时间，减少为原来的一半
- 添加部分说明文档
- 修复了command模块中的部分bug
- 修复了键盘回调函数中的部分bug

## 0.0.5

- 修复了不能自定义切换键的bug

## 0.0.6

- 修复了长按触发命令中的部分bug

## 0.0.7

- 修复了不能设置同一个键按下和松开的逻辑bug
- 修复了命令运行时非功能区按键无法使用的bug
- 修复了切换键的bug
- Condition类添加了show参数，可以控制提示框是否出现
- Configition类添加了excepts参数，可以排除all模式下的部分按键
- notify模块中添加了确认框装饰器、参数输入框装饰器

## 0.0.8

- 优化了command模块中部分函数
- 新增了Macro类中的两个属性，可以更改当前配置和运行状态
- 新增了部分注释

## 0.0.9

- 修复了Configition类中excepts参数的bug
- 修复了command模块中部分函数bug

## 0.0.10

- 更改了notify模块中InputBox的参数形式，增加了combobox类型输入框
- Macro类新增add_command_to_all方法，用于向所有配置中批量统一添加命令

## 0.0.11

- 为touch_key函数增加了按下和松开的间隔时间
- command模块新增了两个函数，change_to_zh以及change_to_en，用于切换输入法状态

## 0.0.12

- 修复了一个显示bug

## 0.0.13

- command模块新增了两个函数enable_network_adapter以及disable_network_adapter，用于启用和禁用网络适配器（网卡）
- command模块中新增了部分窗口操作相关的函数

## 0.0.14

- 修复了win7上更改窗体透明度报错的bug
- 新增了播放音频的函数
- 更新了说明文档

## 0.0.15

- 新增了命令提示框的颜色设置选项，可以设置前景色和背景色

## 0.0.16

- Macro类中pool属性变为可访问