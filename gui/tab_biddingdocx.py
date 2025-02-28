import os
import sys
import tkinter as tk
from tkinter import ttk  # 添加这行导入
from datetime import datetime
from tkinter import filedialog, END
import tkinter.messagebox as messagebox
import traceback
import threading
from bidding_docx.bidding_docx import BiddingDocument
from bidding_docx.social_docx import SocialDocument
from utils.logger_utils import LoggerUtils
from utils.common_utils import get_script_directory

class BiddingDocxUI:
    logger = None
    excel_file = None
    pdf_dir = None
    is_processing = False  # 添加处理状态标志

    def __init__(self,tab):
        # 初始化时可以创建实例变量
        self.tab = tab

        # 设置主界面背景色
        style = ttk.Style()
        style.configure('Tab.TFrame', background='#F5F5F5')
        self.tab.configure(style='Tab.TFrame')

        # 配置LabelFrame样式
        # style.configure("TLabelframe", background='#F5F5F5')
        # style.configure("TLabelframe.Label", background='#F5F5F5')
        
        self.radio_var = tk.StringVar()
        self.checkbox_var = tk.IntVar()
        self.switch_var = tk.IntVar(value=1)
        # 多选框变量
        self.checkbox_vars = [tk.IntVar(value=1) for _ in range(6)]

    def init_page_display(self, tab, out_logger):
        self.out_logger = out_logger
        self.logger = LoggerUtils.get_logger(None, out_logger)
        
        # 修改标签样式，使背景色与主界面一致
        label_style = {
            'font': ('SimSun', 9, 'bold'), 
            'bg': '#F5F5F5',  # 与主界面背景色一致
            'fg': '#333333'   # 深灰色文字
        }

        # 修改按钮样式
        button_style = {
            'bg': '#4A90E2',  # 适中的蓝色
            'fg': 'white',
            'relief': 'flat',
            'font': ('SimSun', 9, 'bold')
        }

        # 定义按钮的正常与选中状态的边框颜色
        highlight_color = "blue"  # 选中时的边框颜色
        normal_color = "gray"  # 正常时的边框颜色
        normal_bg_color = "#f0f0f0"  # 背景色正常
        style = ttk.Style()
        # 配置单选按钮样式
        style.configure(
            "Custom.TRadiobutton",
            background=normal_bg_color,  # 通过样式设置背景色
            font=("SimSun", 9)
        )

        # 配置多选框样式
        style.configure(
            "Custom.TCheckbutton",
            background=normal_bg_color,  # 通过样式设置背景色
            font=("SimSun", 9)
        )

        # 配置选中和未选中状态的样式
        style.map(
            "Custom.TRadiobutton",
            background=[("disabled", normal_bg_color)],
            foreground=[("disabled", "gray")],
            indicatorrelief=[("selected", "solid"), ("!selected", "solid")],
            borderwidth=[("selected", 2), ("!selected", 1)],
            bordercolor=[("selected", highlight_color), ("!selected", normal_color)]
        )
        style.map(
            "Custom.TCheckbutton",
            background=[("disabled", normal_bg_color)],
            foreground=[("disabled", "gray")],
            indicatorrelief=[("selected", "solid"), ("!selected", "solid")],
            borderwidth=[("selected", 2), ("!selected", 1)],
            bordercolor=[("selected", highlight_color), ("!selected", normal_color)]
        )

        # 创建一个Frame作为背景
        bg_frame = tk.Frame(tab, bg='#F5F5F5')
        bg_frame.place(x=0, y=0, relwidth=1, relheight=1)

        # 所有组件改为使用bg_frame作为父容器
        label1 = tk.Label(bg_frame, text="投标名单：", **label_style)
        label1.place(x=23, y=32, width=65, height=20, anchor="w")
        self.down_list_entry = tk.Entry(bg_frame, bg='white')
        self.down_list_entry.place(x=85, y=22, width=250, height=23)
        upload_button = tk.Button(bg_frame, text="选择文件", **button_style, command=self.select_excel_file)
        upload_button.place(x=353, y=21, width=65, height=23)

        label2 = tk.Label(bg_frame, text="社保目录：", **label_style)
        label2.place(x=23, y=70, width=65, height=20, anchor="w")
        self.pdf_dir_entry = tk.Entry(bg_frame, bg='white')
        self.pdf_dir_entry.place(x=85, y=60, width=250, height=20)
        directory_button = tk.Button(bg_frame, text="选择目录", **button_style, command=self.select_pdf_directory)
        directory_button.place(x=353, y=59, width=65, height=23)
        self.cb_social = tk.Checkbutton(
            bg_frame, 
            text='只处理社保', 
            variable=self.checkbox_var,
            command=self.toggle_entry_state,
            bg='#F5F5F5'
        )
        self.cb_social.place(x=423, y=58)

        label3 = tk.Label(bg_frame, text="ERP用户：", **label_style)
        label3.place(x=23, y=110, width=63, height=30, anchor="w")
        self.erp_user_entry = tk.Entry(bg_frame)
        self.erp_user_entry.place(x=85, y=98, width=110, height=23)
        label4 = tk.Label(bg_frame, text="ERP密码：", **label_style)
        label4.place(x=239, y=110, width=62, height=20, anchor="w")
        self.erp_pwd_entry = tk.Entry(bg_frame, show="*")
        self.erp_pwd_entry.place(x=305, y=98, width=110, height=23)
        # 创建一个按钮来模拟开关
        self.button_switch = tk.Button(bg_frame, text="正式环境", fg="white", bg="green", relief="flat", command=self.toggle_switch_button)
        self.button_switch.place(x=428, y=97, width=73, height=23)

        label5 = tk.Label(bg_frame, text="标书模板：", **label_style)
        label5.place(x=23, y=150, width=63, height=20, anchor="w")
        # 设置默认选中的按钮为 Radio 1
        self.radio_var.set("1")
        # 创建两个单选按钮并放置
        radio1 = ttk.Radiobutton(
            bg_frame, 
            text="模板<bidding_template1.docx>", 
            variable=self.radio_var, 
            value="1", 
            style="Custom.TRadiobutton"
        )
        radio2 = ttk.Radiobutton(
            bg_frame, 
            text="模板<bidding_template2.docx>", 
            variable=self.radio_var, 
            value="2", 
            style="Custom.TRadiobutton"
        )
        radio1.place(x=80, y=138)
        radio2.place(x=300, y=138)

        label6 = tk.Label(bg_frame, text="下载范围：", **label_style)
        label6.place(x=23, y=188, width=63, height=20, anchor="w")
        # 自定义每个多选框的显示名称
        checkbox_labels = ["身份证", "毕业证", "学位证", "学信网", "劳动合同", "资质证书"]
        # 创建多选框
        for i in range(6):
            cb = ttk.Checkbutton(
                bg_frame, 
                text=checkbox_labels[i], 
                variable=self.checkbox_vars[i],
                style="Custom.TCheckbutton"
            )
            if i < 5:
                cb.place(x=88 + i * 65, y=175)
            else:
                cb.place(x=92 + i * 67, y=175)

        self.submit_button = tk.Button(bg_frame, text="开始处理", 
                                     bg="#4A90E2",
                                     fg="white",
                                     relief="flat", 
                                     font=('SimSun', 10, 'bold'), 
                                     command=self.start_process_thread)
        self.submit_button.place(x=220, y=215, width=90, height=25)

    # 选择下载名单文件
    def select_excel_file(self):
        # 打开文件选择对话框
        self.excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        # 如果选择了文件，更新 Entry 组件内容
        if self.excel_file:
            self.down_list_entry.delete(0, END)  # 清空现有内容
            self.down_list_entry.insert(0, self.excel_file)  # 插入新文件路径

    # 选择社保文件夹
    def select_pdf_directory(self):
        self.pdf_dir = filedialog.askdirectory()
        if self.pdf_dir:
            # self.label_pdf.config(text=f"Selected: {self.pdf_dir}")
            self.pdf_dir_entry.delete(0, END)  # 清空现有内容
            self.pdf_dir_entry.insert(0, self.pdf_dir)  # 插入新文件路径

    # 回调函数，用于切换输入框的状态
    def toggle_entry_state(self):
        is_disabled = self.checkbox_var.get() == 1

        # 获取 bg_frame 中的所有控件
        bg_frame = [w for w in self.tab.winfo_children() if isinstance(w, tk.Frame)][0]
        radio_buttons = [w for w in bg_frame.winfo_children() if isinstance(w, ttk.Radiobutton)]
        checkbuttons = [w for w in bg_frame.winfo_children() if isinstance(w, ttk.Checkbutton)]

        # 禁用/启用输入框
        self.erp_user_entry.config(state="disabled" if is_disabled else "normal")
        self.erp_pwd_entry.config(state="disabled" if is_disabled else "normal")

        # 禁用/启用单选按钮
        for radio in radio_buttons:
            radio.state(['disabled'] if is_disabled else ['!disabled'])

        # 禁用/启用多选框
        for checkbox in checkbuttons:
            checkbox.state(['disabled'] if is_disabled else ['!disabled'])

    def toggle_switch_button(self):
        if self.switch_var.get() == 0:
            self.switch_var.set(1)
            self.button_switch.config(text="正式环境", bg="green")
        else:
            self.switch_var.set(0)
            self.button_switch.config(text="测试环境", bg="red")

    def start_process_thread(self):
        """在新线程中启动处理"""
        if self.is_processing:
            messagebox.showinfo("提示", "文件处理中，请等待...")
            return
            
        # 创建新线程执行处理
        process_thread = threading.Thread(target=self.process_files)
        process_thread.daemon = True  # 设置为守护线程
        
        self.is_processing = True
        self.submit_button.config(state="disabled", text="正在处理中...")
        process_thread.start()
        
        # 启动检查线程状态的方法
        self.check_thread_status(process_thread)

    def check_thread_status(self, thread):
        """检查处理线程的状态"""
        if thread.is_alive():
            # 如果线程还在运行，继续检查
            self.tab.after(100, lambda: self.check_thread_status(thread))
        else:
            # 线程结束，恢复按钮状态
            self.submit_button.config(state="normal", text="开始处理")
            self.is_processing = False

    def process_files(self):
        """处理文件的主要逻辑"""
        try:
            # 判断输入是否为空
            if not self.pdf_dir:
                messagebox.showinfo("提示", "请选择社保的目录.")
                self.logger.error("请选择社保的目录.")
                return
            if not self.excel_file:
                messagebox.showinfo("提示", "请选择下载的名单.")
                self.logger.error("请选择下载的名单.")
                return

            if self.checkbox_var.get() != 1:
                # 获取多选框的选中状态，生成一个字符串表示
                chkbox_values = ''.join([str(var.get()) for var in self.checkbox_vars])

                # 获取当前脚本或执行文件所在的目录
                current_dir = get_script_directory()
                self.logger.info(f"当前目录：{current_dir}")

                download_dir = os.path.join(current_dir, 'download', f"{datetime.now().strftime('%Y%m%d%H%M%S')}")
                if not os.path.exists(download_dir):
                    os.makedirs(download_dir)
                bidding_docx = BiddingDocument(self.erp_user_entry.get(), self.erp_pwd_entry.get(), chkbox_values, download_dir, self.pdf_dir, self.excel_file, self.radio_var.get(), self.out_logger)
                bidding_docx.run(self.switch_var.get())
            else:
                social_docx = SocialDocument(self.pdf_dir, self.excel_file, self.out_logger)
                social_docx.insert_socialdocx_to_word()

        except Exception as e:
            self.logger.error(f"数据处理失败: {str(e)}")
            traceback.print_exc()
            messagebox.showerror("错误", f"处理过程出错：{str(e)}")
        finally:
            self.submit_button.config(state="normal")  # 启用提交按钮