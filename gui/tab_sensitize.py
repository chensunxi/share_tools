import tkinter as tk
from tkinter import ttk

import file_processing

from utils.logger_utils import LoggerUtils

def setup_tab2(self, tab, out_logger):
    self.logger = LoggerUtils.get_logger(None, out_logger)
    # 设置主界面背景色
    self.tab = tab
    style = ttk.Style()
    style.configure('Tab.TFrame', background='#F5F5F5')
    self.tab.configure(style='Tab.TFrame')
    # 文件处理TAB：3个输入框 + 1个按钮
    label3 = tk.Label(tab, text="社保城市：")
    label3.grid(row=0, column=0, padx=5, pady=5)
    entry3 = tk.Entry(tab)
    entry3.grid(row=0, column=1, padx=5, pady=5)

    label4 = tk.Label(tab, text="文件目录：")
    label4.grid(row=1, column=0, padx=5, pady=5)
    entry4 = tk.Entry(tab)
    entry4.grid(row=1, column=1, padx=5, pady=5)


    def on_button_click():
        file_processing.test_backend(self, entry3.get(), entry4.get(), entry5.get())
        self.logger.info(f"文件处理：{entry3.get()}, {entry4.get()}, {entry5.get()}")

    button2 = tk.Button(tab, text="提 交", command=on_button_click)
    button2.grid(row=3, column=0, columnspan=2, pady=10)