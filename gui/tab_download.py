import os
import sys
import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, END  # 导入 ttk 模块
import tkinter.messagebox as messagebox
import traceback
import threading  # 添加线程支持

from utils.logger_utils import LoggerUtils
from utils.common_utils import get_script_directory
from social_download.beijing_social import BeijingSocialSecurity
from social_download.hangzhou_social import HangzhouSocialSecurity
from social_download.guangzhou_social import GuangzhouSocialSecurity
from social_download.shenzhen_social import ShenzhenSocialSecurity
from social_download.shanghai_social import ShanghaiSocialSecurity
from social_download.changsha_social import ChangshaSocialSecurity
from social_download.nanjing_social import NanjingSocialSecurity

class FileDownloadUI:
    logger = None
    excel_file = None
    def __init__(self,tab):
        # 初始化时可以创建实例变量
        self.tab = tab
        # 设置主界面背景色
        style = ttk.Style()
        style.configure('Tab.TFrame', background='#F5F5F5')
        self.tab.configure(style='Tab.TFrame')
        self.is_downloading = False  # 添加下载状态标志
        self.radio_var = tk.StringVar()

    def init_page_display(self, tab, out_logger):
        self.out_logger = out_logger
        self.logger = LoggerUtils.get_logger(None, out_logger)
        x_position = 25
        y_position = 30

        label1 = tk.Label(tab, text="社保下载名单：")
        label1.place(x=x_position, y=35, width=110, height=20, anchor="w")
        self.excel_file_entry = tk.Entry(tab, width=40)
        self.excel_file_entry.place(x=x_position+133, y=25, width=210, height=23)
        upload_button = tk.Button(tab, text="选择文件", command=self.select_excel_file)
        upload_button.place(x=370, y=23, width=60, height=25)
        y_position += 45

        label2 = tk.Label(tab, text="社保下载城市：")
        label2.place(x=x_position, y=y_position, width=110, height=20, anchor="w")
        # 下拉框
        city_options = ["beijing-北京", "shanghai-上海", "guangzhou-广州", "shenzhen-深圳", "hangzhou-杭州", "changsha-长沙", "nanjing-南京"]
        self.city_combobox = ttk.Combobox(tab, values=city_options, state="readonly")
        self.city_combobox.place(x=x_position+133, y=y_position, width=210, height=23, anchor="w")
        y_position += 45
        # 绑定选择事件
        self.city_combobox.bind("<<ComboboxSelected>>", self.city_selected)

        label3 = tk.Label(tab, text="社保下载模式：")
        label3.place(x=x_position, y=y_position, width=110, height=20, anchor="w")
        # 设置默认选中的按钮为 Radio 1
        self.radio_var.set("1")
        # 创建两个单选按钮并放置
        radio1 = ttk.Radiobutton(
            tab,
            text="按个人",
            variable=self.radio_var,
            value="1",
            command=self.on_radio_change,
            style="Custom.TRadiobutton"
        )
        radio2 = ttk.Radiobutton(
            tab,
            text="按单位",
            variable=self.radio_var,
            value="2",
            command=self.on_radio_change,
            style="Custom.TRadiobutton"
        )
        radio1.place(x=x_position+133, y=y_position-10)
        radio2.place(x=215, y=y_position-10)

        label4 = tk.Label(tab, text="单次下载上限：")
        label4.place(x=295, y=y_position, width=85, height=20, anchor="w")
        self.download_limit_entry = tk.Entry(tab, width=40, justify='right')
        self.download_limit_entry.place(x=380, y=y_position-10, width=39, height=23)
        self.download_limit_entry.insert(0, "200")
        self.download_limit_entry.config(state='disabled')

        self.submit_button = tk.Button(tab, text="开始下载", command=self.start_download_thread)
        self.submit_button.place(x=182, y=175, width=90)

    # 选择 Excel 文件
    def select_excel_file(self):
        # 打开文件选择对话框
        self.excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        # 如果选择了文件，更新 Entry 组件内容
        if self.excel_file:
            self.excel_file_entry.delete(0, END)  # 清空现有内容
            self.excel_file_entry.insert(0, self.excel_file)  # 插入新文件路径

    def city_selected(self, event):
        """处理下拉框选择事件"""
        selected_city = self.city_combobox.get()  # 获取选择的城市
        log_message = f"已选择城市: {selected_city}"
        # self.log_output.insert(tk.END, log_message + "\n")  # 打印到日志中
        self.logger.info(log_message)

    def start_download_thread(self):
        """在新线程中启动下载"""
        if self.is_downloading:
            messagebox.showinfo("提示", "下载正在进行中，请等待...")
            return
            
        # 创建新线程执行下载
        download_thread = threading.Thread(target=self.on_button_click)
        download_thread.daemon = True  # 设置为守护线程
        
        self.is_downloading = True
        self.submit_button.config(state="disabled", text="下载处理中...")
        download_thread.start()
        
        # 启动检查线程状态的方法
        self.check_thread_status(download_thread)

    def check_thread_status(self, thread):
        """检查下载线程的状态"""
        if thread.is_alive():
            # 如果线程还在运行，继续检查
            self.tab.after(100, lambda: self.check_thread_status(thread))
        else:
            # 线程结束，恢复按钮状态
            self.submit_button.config(state="normal", text="开始下载")
            self.is_downloading = False

    def on_button_click(self):
        """执行下载操作"""
        try:
            # 获取中英文城市名称
            if self.city_combobox.get():
                eng_city, chn_city = self.city_combobox.get().split('-')
            else:
                messagebox.showinfo("提示", "请选择下载的城市.")
                return
                
            if not self.excel_file:
                messagebox.showinfo("提示", "请选择下载的名单.")
                return

            # 根据用户输入返回对应的城市类
            city_mapping = {
                "beijing": BeijingSocialSecurity,
                "hangzhou": HangzhouSocialSecurity,
                "guangzhou": GuangzhouSocialSecurity,
                "shenzhen": ShenzhenSocialSecurity,
                "shanghai": ShanghaiSocialSecurity,
                "changsha": ChangshaSocialSecurity,
                "nanjing": NanjingSocialSecurity
            }
            # 获取当前脚本或执行文件所在的目录
            script_dir = get_script_directory()
            # 根据需要选择城市，举例选择北京
            source_dir = os.path.join(script_dir, 'download', eng_city, f"{datetime.now().strftime('%Y%m%d%H%M%S')}")
            if not os.path.exists(source_dir):
                os.makedirs(source_dir)
            if eng_city in city_mapping:
                city_social_security = city_mapping[eng_city](eng_city, source_dir, chn_city, self.excel_file, self.radio_var.get(), self.download_limit_entry.get(), self.out_logger)  # 返回对应的城市子类
                city_social_security.run()
            else:
                print("无效的选择，请重新选择。")
                messagebox("请选择下载的城市.")
        except Exception as e:
            self.logger.error(f"数据处理失败: {str(e)}")
            traceback.print_exc()
            messagebox.showerror("错误", f"下载过程出错：{str(e)}")
        finally:
            self.submit_button.config(state="normal")  # 启用提交按钮

    def on_radio_change(self):
        """处理单选按钮状态变化"""
        if self.radio_var.get() == "1":  # 按个人
            self.download_limit_entry.config(state='disabled')
        else:  # 按单位
            self.download_limit_entry.config(state='normal')