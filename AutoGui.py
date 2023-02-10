#源头作者b站水哥：不高兴就喝水
#author:LOCKSAIKO
#date:2022/04/10
#mail:2537334539@qq.com

import pyautogui
import xlrd
import pyperclip
import time
import json
import os
import sys
from tkinter import *
from tkinter import messagebox
from multiprocessing import Process
import multiprocessing
import threading
import datetime
import tkinter.filedialog
from functools import partial
import ttkbootstrap
from pynput import keyboard
import keyboard as KB

from Func import AutoRun, NarrowIcon


def set_win_center(root, curWidth='', curHight=''):
    if not curWidth:
        curWidth = root.winfo_width()
    if not curHight:
        curHight = root.winfo_height()
    scn_w, scn_h = root.maxsize()
    cen_x = (scn_w - curWidth) / 2
    cen_y = (scn_h - curHight) / 2
    size_xy = '%dx%d+%d+%d' % (650, 500, cen_x - 250, cen_y - 250)
    root.geometry(size_xy)


class Main_window(object):
    theme_list = ['vapor', 'darkly', 'superhero', 'solar', 'cyborg', 'cerculean', 'simplex', 'journal', 'morph',
                  'united', 'pulse', 'yeti', 'sandstone', 'lumen', 'minty', 'litera', 'flatly', 'cosmo']

    def __init__(self, window):
        self.window = window
        self.process_pid_list = []

        # 缩小至系统托盘方法
        icons = os.getcwd() + r'./k.ico'
        hover_text = "自动化脚本"  # 悬浮于图标上方时的提示
        menu_options = ()
        self.sysTrayIcon = NarrowIcon.SysTrayIcon(icons, hover_text, menu_options, on_quit=self.on_closing,
                                                  default_menu_index=1, window=self.window)

    def create_page(self):
        self.window.title('自动化脚本')
        self.window.iconbitmap('./config/k.ico')
        self.state = self.window.state()

        with open('./config/ThemeStyle.json', 'r') as f:
            dt = json.load(f)
        style = ttkbootstrap.Style(theme=dt)
        self.window = style.master

        self.fream = Frame(self.window, width=650, height=500)
        self.fream.pack_propagate(0)

        lb_tx = Label(self.fream, text='运行日志：', justify='center', width=40, background='white')
        self.text_ = Text(self.fream, width=40, height=20, background='white')

        self.window.bind("<Unmap>", lambda event: self.Unmap() if self.window.state() == 'iconic' else False)

        self.el_b1_tx = StringVar()
        self.default_filename()
        self.el_b2_tx = StringVar()
        self.el_b2_tx.set('辅助脚本(未选择)')
        self.my_auxiliray_script = None
        self.el_b3_tx = StringVar()
        self.el_b3_tx.set('辅助脚本2(未选择)')
        self.my_auxiliray_script2 = None

        b_choose1 = ttkbootstrap.Button(self.fream, style='success', textvariable=self.el_b1_tx, width=20,
                                        command=partial(self.choose_file, self.el_b1_tx))
        b_choose2 = ttkbootstrap.Button(self.fream, style='success', textvariable=self.el_b2_tx, width=20,
                                        command=partial(self.choose_auxiliray_file, self.el_b2_tx, '辅助脚本(未选择)'))
        b_choose3 = ttkbootstrap.Button(self.fream, style='success', textvariable=self.el_b3_tx, width=20,
                                        command=partial(self.choose_auxiliray_file2, self.el_b3_tx, '辅助脚本2(未选择)'))
        self.scale = Scale(self.fream, from_=1, orient=HORIZONTAL, to_=10, tickinterval=1, resolution=1, length=200)
        b1 = ttkbootstrap.Button(self.fream, text='运行脚本', width=15,
                                 command=partial(self.add_script_process, AutoRun.use_mainWork))
        b2 = ttkbootstrap.Button(self.fream, text='一直运行', width=15,
                                 command=partial(self.add_script_process, AutoRun.loop_use_mainWork))
        b3 = ttkbootstrap.Button(self.fream, text='结束运行', width=15, command=self.kill_script)
        b4 = ttkbootstrap.Button(self.fream, text='清空文本框', width=15, command=self.clean_text)

        # config_page配置
        self.page_flag = 0
        self.default_var = StringVar()
        if self.page_flag == 0:
            self.default_var.set('1')

        self.create_menu()

        self.fream.place(x=0, y=0)
        lb_tx.place(x=260, y=0)
        self.text_.place(x=260, y=20)

        b_choose1.place(x=20, y=30)
        b_choose2.place(x=20, y=70)
        b_choose3.place(x=20, y=110)
        b1.place(x=45, y=225)
        b2.place(x=45, y=270)
        b3.place(x=45, y=315)
        b4.place(x=45, y=360)
        self.scale.place(x=20, y=170)

        KB.add_hotkey('ctrl+alt+a', self.hide_window)
        KB.add_hotkey('ctrl+alt+b', self.show_window)

        #horizontal, vertical
        VScroll_tb = Scrollbar(self.text_, orient='vertical', command=self.text_.yview)
        VScroll_tb.place(relx=0.95, rely=0, relheight=1)
        self.text_.configure(yscrollcommand=VScroll_tb.set)

        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.window.mainloop()


    def create_menu(self):
        self.menubar = Menu(self.window)
        self.config_menu = Menu(self.menubar, tearoff=0)
        check_theme = Menu(self.config_menu, tearoff=0)
        self.config_menu.insert_cascade(0, label='主题', menu=check_theme)
        for theme in Main_window.theme_list:
            check_theme.add_command(label=theme, command=partial(self.switch_theme, theme))
        self.menubar.add_cascade(label='配置', menu=self.config_menu)
        self.config_menu.add_command(label='无响应等待时长', command=self.config_window)
        self.window.config(menu=self.menubar)

    def switch_theme(self, theme):
        tx = theme
        with open('./config/ThemeStyle.json', 'w')as f:
            json.dump(f, tx)
        self.create_page()

    def hide_window(self):
        self.window.iconify()

    def show_window(self):
        self.window.deiconify()

    def kill_sys(self, id):
        command = 'taskkill /f /pid %s' % id
        os.system(command)

    def kill_script(self):
        if self.process_pid_list != '':
            for id in self.process_pid_list:
                self.kill_sys(id)
            self.process_pid_list = []
        open('./config/temp_log.txt', 'w').close()
        self.text_.insert(END, '脚本已关闭')
        with open('./config/run_log.txt', 'a+')as run_log:
            text = self.text_.get('1.0', END)
            run_log.write(text + '\n' + '-' * 5 + str(datetime.datetime.now()) + '-' * 5)
        self.thread_flag = False

    def Unmap(self):
        self.window.withdraw()
        self.sysTrayIcon.show_icon()

    def on_closing(self):
        if tkinter.messagebox.askokcancel("退出", "你想要退出么?"):
            with open('./config/run_log.txt', 'a+')as run_log:
                text = self.text_.get('1.0', END)
                run_log.write(text + '\n' + '-'*5 + str(datetime.datetime.now()) + '-'*5)
                if os.path.exists('./config/temp_log.txt'):
                    os.remove('./config/temp_log.txt')
                if self.process_pid_list != '':
                    for id in self.process_pid_list:
                        self.kill_sys(id)
            sys.exit(0)

    def default_filename(self):
        if os.path.exists('./config/filepath.json'):
            with open('./config/filepath.json', 'r')as f:
                file_path = json.load(f)
            if file_path != '':
                self.my_script = file_path
                filename = str(self.my_script).split('/')
                self.el_b1_tx.set(filename[-1])
            else:
                self.el_b1_tx.set('主脚本（未选择）')
        else:
            self.el_b1_tx.set('主脚本（未选择）')

    def choose_file(self, button_var):
        filename = tkinter.filedialog.askopenfilename()
        if filename != '':
            with open('./config/filepath.json', 'w')as f:
                json.dump(filename, f)
            self.my_script = filename
            filename = str(self.my_script).split('/')
            button_var.set(filename[-1])
        else:
            button_var.set('主脚本（未选择）')

    def choose_auxiliray_file(self, button_var, tips):
        filename = tkinter.filedialog.askopenfilename()
        if filename != '':
            self.my_auxiliray_script = filename
            filename = str(self.my_auxiliray_script).split('/')
            button_var.set(filename[-1])
        else:
            button_var.set(tips)

    def choose_auxiliray_file2(self, button_var, tips):
        filename = tkinter.filedialog.askopenfilename()
        if filename != '':
            self.my_auxiliray_script2 = filename
            filename = str(self.my_auxiliray_script2).split('/')
            button_var.set(filename[-1])
        else:
            button_var.set(tips)

    def switch_theme(self, theme):
        tx = theme
        with open('./config/ThemeStyle.json', 'w')as f:
            json.dump(tx, f)
        self.create_page()

    def config_window(self):
        self.page = Toplevel()
        self.page.geometry('200x200')
        self.page.iconbitmap('./config/k.ico')

        global ts_entry

        ts_entry = Entry(self.page, textvariable=self.default_var, justify='center', width=16)
        self.ts_button = Button(self.page, text='确定', width=12, command=self.set_tsvar)

        ts_entry.place(x=25, y=50)
        self.ts_button.place(x=42, y=100)

        self.page.mainloop()

    def set_tsvar(self):
        self.page_flag += 1
        self.default_var.set(ts_entry.get())
        self.page.destroy()

    def suspension_window(self):
        global sus_window
        self.sus_window = Toplevel()
        self.sus_window.geometry('100x100')
        self.sus_window.overrideredirect(True)
        sus_window = self.sus_window

        Label(self.sus_window, justify='center', font=(None, 12), text='正在运行').place(x=40, y=50)
        self.sus_window.mainloop()

    def insert_message(self):
        while self.thread_flag:
            if self.thread_flag:
                time.sleep(float(self.default_var.get()))
                if os.path.exists('./config/temp_log.txt'):
                    with open('./config/temp_log.txt', 'r')as templog:
                        for message in templog:
                            self.text_.insert(END, message)
                    open('./config/temp_log.txt', 'w').close()
            else:
                break

    def add_script_process(self, name):
        add = Process(target=name,
                      args=(self.default_var.get(), self.my_script, self.scale.get(), self.my_auxiliray_script,
                            self.my_auxiliray_script2))
        add.start()
        self.thread_flag = True
        self.add_thread()
        id = add.pid
        self.process_pid_list.append(id)

    def clean_text(self):
        self.text_.delete('0.0', END)

    def add_thread(self):
        add = threading.Thread(target=self.insert_message)
        add.setDaemon(True)
        add.start()


if __name__ == '__main__':
    multiprocessing.freeze_support()
    window = Tk()
    set_win_center(window)
    main = Main_window(window)
    main.create_page()
