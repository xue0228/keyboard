import tkinter as tk
import tkinter.font as tf
from tkinter import ttk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, askdirectory
import time
import threading
from functools import wraps


from xyw_macro.utils import SingletonType
from xyw_macro.contants import SLEEP_TIME


class Notification(metaclass=SingletonType):
    def __init__(self, text='xyw', fg='white', bg='black'):
        self.__text = text
        self.__fg = fg
        self.__bg = bg

        self.__visible = False
        self.__vnum = 0
        self.__window, self.__label, self.__width = self.__init__window()
        self.set_visible(self.__visible)

    def show(self):
        if self.__vnum == 0:
            self.set_visible(True)
        self.__vnum = self.__vnum + 1

    def hide(self):
        self.__vnum = self.__vnum - 1
        if self.__vnum == 0:
            self.set_visible(False)

    def __init__window(self):
        window = tk.Tk()
        window.wm_attributes('-topmost', True)
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        width = round(screen_width / 10)
        height = round(screen_width / 10)
        window.geometry('{}x{}+{}+{}'.format(width, height, (screen_width - width) // 2, (screen_height - height) // 2))
        window.overrideredirect(True)
        window.configure(background=self.__bg)
        window.attributes('-alpha', 0.7)
        font_size = self.__get_font_size(width)
        outer_border_size = round(font_size * 0.08)
        inner_border_size = round(font_size * 0.05)
        font = tf.Font(size=font_size, weight=tf.BOLD)
        label_border = tk.LabelFrame(window, background=self.__fg, relief='flat')
        label = tk.Label(label_border, text=self.__text, font=font, bg=self.__bg, fg=self.__fg,
                              height=height, width=width, justify='center', anchor='center',
                              borderwidth=0, relief='flat')
        label_border.pack(fill='both', expand=True, padx=outer_border_size, pady=outer_border_size)
        label.pack(fill='both', expand=True, padx=inner_border_size, pady=inner_border_size)
        return window, label, width

    def get_text(self):
        """
        ??????????????????
        :return:
        """
        return self.__text

    def __get_font_size(self, width):
        # ???????????????????????????
        texts = self.__text.split('\n')
        # ?????????????????????
        alnum = r'abcdefghijklmnopqrstuvwxyz0123456789+-*/=`~!@#$%^&*()_\|?><.,'
        # ??????????????????????????????
        length = [1]
        for item in texts:
            tem = 0
            for i in item:
                if i.lower() in alnum:
                    # ???????????????????????????????????????
                    tem = tem + 0.5
                else:
                    # ?????????????????????????????????
                    tem = tem + 1
            length.append(tem)
        length = max(length)
        # ??????????????????????????????????????????
        font_size = round(width * 0.6 / length)
        return font_size

    def set_text(self, text):
        """
        ??????????????????
        :param text:
        :return:
        """
        self.__text = text
        font_size = self.__get_font_size(self.__width)
        # ??????????????????
        font = tf.Font(size=font_size, weight=tf.BOLD)
        self.__label.config(text=self.__text, font=font)

    def get_visible(self):
        """
        ?????????????????????
        :return:
        """
        return self.__visible

    def set_visible(self, visible):
        """
        ?????????????????????
        :param visible:
        :return:
        """
        self.__visible = visible
        if self.__visible:
            self.__window.update()
            self.__window.deiconify()
        else:
            self.__window.withdraw()

    def run(self):
        """
        ?????????????????????
        :return:
        """
        self.__window.mainloop()

    text = property(get_text, set_text)
    visible = property(get_visible, set_visible)


class InputField:
    def __init__(self, name, type='entry', default=None, options=None, focus=False):
        self.name = name
        self.type = type
        self.default = default
        self.options = options
        self.focus = focus

    @staticmethod
    def select_file(var):
        filepath = askopenfilename()
        var.set(filepath)

    @staticmethod
    def select_dir(var):
        dirpath = askdirectory()
        var.set(dirpath)

    def draw_frame(self, window):
        var = tk.StringVar()
        frame = tk.Frame(window, takefocus=True)
        frame.pack(fill=tk.X, padx=10, pady=2, expand=1)
        tk.Label(frame, text=self.name).pack(side=tk.TOP, anchor=tk.W)
        if self.type == 'entry':
            widget = tk.Entry(frame, show=None, textvariable=var)
            widget.pack(fill=tk.X, side=tk.TOP)
            if self.default is not None:
                var.set(self.default)
        elif self.type == 'file':
            widget = tk.Entry(frame, show=None, textvariable=var, state=tk.DISABLED)
            widget.pack(fill=tk.X, side=tk.LEFT, expand=1)
            tk.Button(frame, text='????????????', command=lambda var=var: self.select_file(var)) \
                .pack(fill=tk.X, side=tk.LEFT)
            if self.default is not None:
                var.set(self.default)
        elif self.type == 'dir':
            widget = tk.Entry(frame, show=None, textvariable=var, state=tk.DISABLED)
            widget.pack(fill=tk.X, side=tk.LEFT, expand=1)
            tk.Button(frame, text='???????????????', command=lambda var=var: self.select_dir(var)) \
                .pack(fill=tk.X, side=tk.LEFT)
            if self.default is not None:
                var.set(self.default)
        elif self.type == 'combobox':
            widget = ttk.Combobox(frame, textvariable=var)
            widget['values'] = self.options
            widget.pack(fill=tk.X, side=tk.TOP)
            if self.default is None:
                widget.current(0)
            else:
                widget.current(self.default)
        else:
            raise ValueError('there is no such type,select in "entry","file","dir" or "combobox"')
        if self.focus:
            widget.focus_set()
        return var


class InputBox:
    """
    ??????????????????
    """
    def __init__(self, title='?????????', *args):
        """
        ???????????????
        :param title: ???????????????
        """
        self.title = title
        self.__args = args

        self.top = None
        self.vars = []

        self.values = []

    def show(self):
        """
        ?????????????????????
        :return: ?????????????????????
        """
        return self.top_window()

    def clear_all(self):
        for var in self.vars:
            var.set('')

    def close_window(self, flag=False):
        if flag:
            self.values = None
        else:
            self.values = [var.get() for var in self.vars]
        self.top.destroy()

    def top_window(self):
        self.top = tk.Toplevel()
        self.top.withdraw()
        self.top.update()
        self.top.wm_attributes('-topmost', True)
        self.top.attributes('-toolwindow', True)
        self.top.title(self.title)
        self.top.grab_set()

        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        width = 300
        height = (len(self.__args) * 2 + 1) * 30
        self.top.geometry('{}x{}+{}+{}'
                          .format(width, height, (screen_width - width) // 2, (screen_height - height) // 2))

        for field in self.__args:
            if not isinstance(field, InputField):
                raise TypeError('args must be <class InputField>')
            self.vars.append(field.draw_frame(self.top))

        frame = tk.Frame(self.top, takefocus=True)
        frame.pack(fill=tk.X, padx=10, pady=2, expand=1)
        button1 = tk.Button(frame, text='??????', command=lambda: self.close_window(False))
        button1.pack(side=tk.LEFT, fill=tk.X, expand=1)
        button2 = tk.Button(frame, text='??????', command=self.clear_all)
        button2.pack(side=tk.LEFT, fill=tk.X, expand=1)

        self.top.protocol("WM_DELETE_WINDOW", lambda: self.close_window(True))
        self.top.bind('<Return>', lambda event: self.close_window(False))
        self.top.bind('<Escape>', lambda event: self.close_window(True))

        self.top.deiconify()
        self.top.focus_force()
        self.top.focus_set()

        self.top.wait_window()
        return self.values


def input_box(*ags, title='?????????'):
    """
    ????????????????????????
    :param title: ???????????????
    :return:
    """
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            time.sleep(SLEEP_TIME)
            res = InputBox(title, *ags).show()
            if res is not None:
                return f(*res)
        return decorated
    return decorator


def confirm_box(message='???????????????????????????'):
    """
    ????????????????????????
    :param message: ????????????
    :return:
    """
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            time.sleep(SLEEP_TIME)
            if messagebox.askokcancel('??????', message):
                return f(*args, **kwargs)
        return decorated
    return decorator


if __name__ == '__main__':
    def sub():
        time.sleep(2)
        notify.text = 'xue'
        notify.show()
        time.sleep(2)
        notify.hide()
    #     notify = Notification()
    #     threading.Thread(target=auto_hide).start()
    #     notify.start()
    # thd = threading.Thread(target=sub)
    # thd.start()
    # def auto_hide():
    #     time.sleep(2)
    #     # notify.destroy()
    #     # flag = False
    #     notify.hide()
    notify = Notification('xyw_macro\n?????????')
    threading.Thread(target=sub).start()
    notify.run()
    # notify.show(0.2)
    # print('end')
    # time.sleep(2)
    # notify.set_text('changed')
    # notify.show()

    # notify.start()
    # print('xue')
    # print(type(notify.get_window()))
    # notify.start()
    # flag = True
    # while flag:
    #     # notify.get_window().update_idletasks()
    #     notify.get_window().update()
