import os
import sys
import tkinter as tk
from datetime import datetime
from tkinter import ttk
import logging
from gui.tab_download import FileDownloadUI
from gui.tab_biddingdocx import BiddingDocxUI
from gui.tab_sensitize import setup_tab2
from gui.tab_merge import setup_tab3
from gui.tab_mail_monitor import MailMonitorUI

# 获取当前脚本的绝对路径
project_root = os.path.dirname(os.path.abspath(__file__))
# 将 social_download 目录添加到 sys.path 中
social_download_dir = os.path.join(project_root, 'social_download')
sys.path.append(social_download_dir)
# 将 bidding_docx 目录添加到 sys.path 中
bidding_docx_dir = os.path.join(project_root, 'bidding_docx')
sys.path.append(bidding_docx_dir)
# 将 utils 目录添加到 sys.path 中
utils_dir = os.path.join(project_root, 'utils')
sys.path.append(utils_dir)

def resource_path(relative_path):
    """获取资源的绝对路径"""
    try:
        # PyInstaller创建临时文件夹,将路径存储于_MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class ShareToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("运营共享中心办公效能提升工具箱 - 当前版本：v1.0.0 - 2025.02.17")
        # 获取当前文件的绝对路径
        current_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = resource_path(os.path.join("resources", "icon.ico"))

        # 设置窗口图标
        self.root.iconbitmap(icon_path)
        # 禁用最大化，限制窗口大小调整
        self.root.resizable(False, False)  # 禁止水平和垂直方向的大小调整
        # 获取屏幕的宽度和高度
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # 设置窗口的大小
        window_width = 550
        window_height = 500

        # 计算窗口的起始位置，使窗口居中
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2

        # 设置窗口的位置和大小
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # 创建并配置样式
        style = ttk.Style()
        style.theme_use('alt')  # 使用alt主题以支持更多样式自定义

        # 配置标签页和整体样式
        # style.configure("Main.TNotebook", background="#2196F3")  # 主背景色
        style.configure("Main.TNotebook.Tab",
                        background="#f0f0f0",  # 未选中时的背景色
                        padding=[5, 5],
                        width=8)

        # 设置选中标签页的样式
        style.map("Main.TNotebook.Tab",
                  background=[("selected", "#1565C0"),  # 选中时的深蓝色背景
                              ("active", "#FFC107")],  # 鼠标悬停时的颜色
                  foreground=[("selected", "white"),  # 选中时的白色字体
                              ("active", "black")])  # 鼠标悬停时的字体颜色

        # 创建Tab控制器
        self.tab_control = ttk.Notebook(self.root, style="Main.TNotebook")

        # 创建Tab页
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab2 = ttk.Frame(self.tab_control)
        self.tab3 = ttk.Frame(self.tab_control)
        self.tab4 = ttk.Frame(self.tab_control)
        self.tab5 = ttk.Frame(self.tab_control)

        # 添加Tab页到Tab控制器
        self.tab_control.add(self.tab1, text="社保下载")
        self.tab_control.add(self.tab2, text="社保脱敏")
        self.tab_control.add(self.tab3, text="文件合并")
        self.tab_control.add(self.tab4, text="投标应用")
        self.tab_control.add(self.tab5, text="邮件监测")

        # 设置Tab控制器的样式
        style = ttk.Style()

        # 设置未选中Tab页的背景色
        style.configure("TNotebook.Tab", background="#f0f0ff", padding=[10, 5])

        # 设置选中Tab页的背景色和前景色（高亮）
        style.configure("TNotebook.Tab.selected", background="#1565C0", foreground="white", padding=[10, 5])

        # 配置Tab控制器背景
        self.tab_control.pack(expand=1, fill="both")

        self.tab_control.pack(padx=10, pady=10)

        # 放置底部状态栏，居中显示状态信息
        self.status_bar = tk.Label(self.root, text=f"Copyright (c) 2002-{datetime.now().strftime('%Y')}, 深圳市长亮科技股份有限公司, All Rights Reserved", relief=tk.SUNKEN, font=("Arial", 9), anchor="center", bd=1, height=1)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))

        # 初始化日志控件
        # self.logger = LoggerUtils.get_logger()
        self.init_logging()

        # 调用子界面设置方法
        file_download = FileDownloadUI(self.tab1)  # 创建 FileDownloadUI 实例
        file_download.init_page_display(self.tab1, self.log_output)
        bidding_docxui = BiddingDocxUI(self.tab4)  # 创建 BiddingDocxUI 实例
        bidding_docxui.init_page_display(self.tab4, self.log_output)
        mail_monitorui = MailMonitorUI(self.tab5)  # 创建 MailMonitorUI 实例
        mail_monitorui.init_page_display(self.tab5, self.log_output)
        setup_tab2(self, self.tab2, self.log_output)
        setup_tab3(self, self.tab3, self.log_output)


    def init_logging(self):
        # 放置清空日志按钮在底部，居中
        clear_button = ttk.Button(self.root, text="清空日志", command=self.clear_log)
        clear_button.pack(side=tk.BOTTOM, pady=(0, 10))

        # 放置日志控件在底部，居中，填充宽度
        scrollbar = ttk.Scrollbar(self.root)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_output = tk.Text(self.root, height=10, width=50, font=("宋体", 9), wrap=tk.WORD, yscrollcommand=scrollbar.set)
        self.log_output.pack(side=tk.BOTTOM, padx=10, pady=(10, 5), fill=tk.X)
        # self.log_output.config(state=tk.DISABLED, bg="#F5F5DC", fg="#4682B4")  # 米色背景 + 蓝色字体
        self.log_output.config(bg="#001F3D", fg="#87CEFA")  # 深蓝色背景 + 蓝色字体
        scrollbar.config(command=self.log_output.yview)

    def log_message(self, message, level="INFO"):
        """日志记录功能"""
        # 输出到日志窗口
        self.log_area.config(state=tk.NORMAL)  # 允许编辑
        self.log_area.insert(tk.END,
                             f"{logging.Formatter().formatTime(logging.LogRecord('', 0, '', 0, '', '', '', ''))} - [{level}] {message}\n")
        self.log_area.config(state=tk.DISABLED)  # 设置为只读

        # 使用日志级别常量（INFO, DEBUG, ERROR）
        level_dict = {
            "DEBUG": logging.DEBUG,
            "INFO": logging.INFO,
            "WARNING": logging.WARNING,
            "ERROR": logging.ERROR,
            "CRITICAL": logging.CRITICAL
        }

        # 如果日志级别无效，使用INFO作为默认
        log_level = level_dict.get(level.upper(), logging.INFO)

        # 在控制台和日志中同时输出日志
        logging.log(log_level, message)

    def clear_log(self):
        """清空日志框"""
        # self.log_output.config(state=tk.NORMAL)  # 允许编辑
        self.log_output.delete(1.0, tk.END)  # 清空日志
        # self.log_output.config(state=tk.DISABLED)  # 设置为只读


if __name__ == "__main__":
    root = tk.Tk()
    app = ShareToolApp(root)
    root.mainloop()
