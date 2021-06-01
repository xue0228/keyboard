import os
import img2pdf
import fitz
import urllib
import pyperclip
import win32gui
import tkinter.messagebox
from xyw_macro import *
from xyw_macro.command import *


macro = Macro(VK('VK_RMENU'))

if get_mac() == '94-E6-F7-0D-0E-D9':
    def search_selected_text():
        touch_keys([VK('VK_CONTROL'), ord('C')])
        random_sleep(0.2)
        text = pyperclip.paste()
        search_text = urllib.parse.quote(text.encode('gbk'))
        url = 'https://www.baidu.com/s?wd=' + search_text
        open_nor_url(url)

    def touch_keys(keys):
        if isinstance(keys, int):
            keys = [keys]
        if not isinstance(keys, (list, tuple)):
            raise TypeError('param keys must be int or list')
        for key in keys:
            press_key(key)
        for key in keys:
            release_key(key)


    def input_text(text):
        input_chars(text)
        change_to_config(0)


    def password_q():
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
        open_url(url)
        change_to_config(0)


    def open_nor_url(url):
        open_url(url)
        change_to_config(0)


    def open_path(path):
        open_file(path)
        change_to_config(0)


    def img_to_pdf(direction):
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

        if direction == 'v':
            page_size = (img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297))
        else:
            page_size = (img2pdf.mm_to_pt(297), img2pdf.mm_to_pt(210))
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
        enable_network_adapter('本地连接')
        disable_network_adapter('无线网络连接')
        change_to_config(0)


    @confirm_box('确定要切换到个人网络吗？')
    def change_to_personal_mode():
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

    # config1.add_command(
    #     Condition(VK('VK_F9'), '战网'),
    #     lambda: run_cmd(r'"C:\Program Files (x86)\Battle.net\Battle.net Launcher.exe"')
    # )
    # config1.add_command(
    #     Condition(VK('VK_F10'), 'Steam'),
    #     lambda: run_cmd(r'"C:\Program Files (x86)\Steam\steam.exe"')
    # )
    # config1.add_command(
    #     Condition(VK('VK_F11'), '切换应用', True, show=False),
    #     lambda: touch_keys([VK('VK_CONTROL'), VK('VK_MENU'), VK('VK_TAB')])
    # )
    # config1.add_command(
    #     Condition(VK('VK_F12'), '任务视图'),
    #     lambda: touch_keys([VK('VK_LWIN'), VK('VK_TAB')])
    # )
    # config1.add_command(
    #     Condition(VK('VK_F11'), 'F盘'),
    #     lambda: open_file(r'F:')
    # )

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
        lambda: open_nor_url(r'http://www.bimiacg.com/')
    )
    config2.add_command(
        Condition(ord('H'), 'hen'),
        lambda: open_sec_url(r'https://hentaidh.cc/')
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
        # lambda: input_text('1174543101@qq.com')
        password_q
    )

    config4 = Configuration('脚本', 'all', excepts=[VK('VK_RETURN'), VK('VK_ESCAPE')])
    config4.add_command(
        Condition(ord('P'), 'PDF'),
        lambda: img_to_pdf('v')
    )
    config4.add_command(
        Condition(ord('O'), 'PDF'),
        lambda: img_to_pdf('h')
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
