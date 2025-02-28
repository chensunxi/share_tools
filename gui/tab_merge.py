import tkinter as tk
from tkinter import ttk

import file_merge
from utils.logger_utils import LoggerUtils


def setup_tab3(self, tab, out_logger):
    logger = LoggerUtils.get_logger(None, out_logger)
    # 设置主界面背景色
    self.tab = tab
    style = ttk.Style()
    style.configure('Tab.TFrame', background='#F5F5F5')
    self.tab.configure(style='Tab.TFrame')
    # 文件合并TAB：4个输入框 + 1个按钮
    label6 = tk.Label(tab, text="入场名单：")
    label6.grid(row=0, column=0, padx=5, pady=5)
    entry6 = tk.Entry(tab)
    entry6.grid(row=0, column=1, padx=5, pady=5)

    label7 = tk.Label(tab, text="文件目录：")
    label7.grid(row=1, column=0, padx=5, pady=5)
    entry7 = tk.Entry(tab)
    entry7.grid(row=1, column=1, padx=5, pady=5)



    def on_button_click():
        file_merge.test_backend(self, entry6.get(), entry7.get(), entry8.get(), entry9.get())
        logger.info(f"文件合并：{entry6.get()}, {entry7.get()}, {entry8.get()}, {entry9.get()}")

    button3 = tk.Button(tab, text="提 交", command=on_button_click)
    button3.grid(row=4, column=0, columnspan=2, pady=10)