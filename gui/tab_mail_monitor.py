import tkinter as tk
from tkinter import ttk, messagebox
import requests
import json
import os
import datetime
import sys

from cryptography.fernet import Fernet
from mail_monitor.mail_process import ENCRYPTION_KEY  # 导入密钥常量
import traceback
from .styles import AppStyles

def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        # PyInstaller创建临时文件夹,将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

class MailMonitorUI:
    def __init__(self, tab, master=None):
        print("初始化 MailMonitorUI...")  # 添加日志
        self.tab = tab
        self.password_manager = PasswordManager()
        self.recv_monitor_mail = []  # 添加这行，用于保存邮箱配置
        self.business_types = []  # 添加业务类型列表
        
        # 预加载所有图片资源
        self.images = {
            'add': tk.PhotoImage(file=get_resource_path("resources/add.png")),
            'del': tk.PhotoImage(file=get_resource_path("resources/del.png")),
            'run': tk.PhotoImage(file=get_resource_path("resources/run.png")),
            'stop': tk.PhotoImage(file=get_resource_path("resources/stop.png"))
        }
        
        # 获取 Notebook 实例
        parent = tab.master
        if isinstance(parent, ttk.Notebook):
            self.notebook = parent
            parent.bind('<<NotebookTabChanged>>', self.on_tab_changed)

    def on_tab_changed(self, event):
        """当标签页切换时触发"""
        current = self.notebook.select()
        current_tab = self.notebook.index(current)
        my_tab = self.notebook.index(self.tab)
        
        # 只有切换到邮件监测标签页时才加载配置
        if current_tab == my_tab:
            print("切换到邮件监控标签页，开始加载配置...")
            if  self.load_server_config() : # 加载服务器配置
                self.load_config()
                # 初始化按钮状态 - 移到最后
                self.check_monitor_status()

    def init_page_display(self, tab, out_logger):
        """初始化页面显示"""
        # 应用全局样式
        AppStyles.configure_styles()
        self.tab.configure(style=AppStyles.FRAME_STYLE)
        
        # 创建主布局
        main_frame = ttk.Frame(tab, style=AppStyles.FRAME_STYLE)
        main_frame.place(x=5, y=5, width=530, height=300)

        # 左侧布局
        left_frame = ttk.Frame(main_frame, style=AppStyles.FRAME_STYLE)
        left_frame.place(x=5, y=0, width=250, height=270)

        # 监测邮箱配置（左侧）
        self.create_email_password_group(left_frame, 0)

        # 右侧布局
        right_frame = ttk.Frame(main_frame, style=AppStyles.FRAME_STYLE)
        right_frame.place(x=260, y=0, width=260, height=270)
        # 抄送邮箱（右侧上方）
        self.create_list_group(right_frame, "抄送邮箱", "send_copy_mail", 223, 48, 0, 0)

        # 左下布局
        lbottom_frame = ttk.Frame(main_frame, style=AppStyles.FRAME_STYLE)
        lbottom_frame.place(x=5, y=80, width=250, height=300)
        # 准入邮箱
        self.create_list_group(lbottom_frame, "准入邮箱", "auth_enter_mail",203,88,0, 0)

        # 关键词输入框
        keyword_frame = ttk.Frame(main_frame, style=AppStyles.FRAME_STYLE)
        keyword_frame.place(x=260, y=83, width=240, height=30)
        ttk.Label(keyword_frame, text="关键词:", style=AppStyles.LABEL_STYLE).place(x=0, y=3)
        self.keyword_var = tk.StringVar()
        ttk.Entry(keyword_frame, textvariable=self.keyword_var, style=AppStyles.ENTRY_STYLE).place(x=48, y=0, width=190)

        # 回复内容输入框
        reply_content_frame = ttk.Frame(main_frame, style=AppStyles.FRAME_STYLE)
        reply_content_frame.place(x=260, y=110, width=240, height=98)
        ttk.Label(reply_content_frame, text="回 复", style=AppStyles.LABEL_STYLE).place(x=0, y=20)
        ttk.Label(reply_content_frame, text="正 文:", style=AppStyles.LABEL_STYLE).place(x=0, y=38)

        # 创建文本框和滚动条
        auto_reply_text = tk.Text(reply_content_frame, 
                             wrap='word',  # 改为 'word' 以实现自动换行
                             undo=True,
                             font=(AppStyles.FONT_FAMILY, AppStyles.FONT_SIZE),
                             borderwidth=AppStyles.ENTRY_BORDER_WIDTH,
                             relief=AppStyles.ENTRY_RELIEF)
        auto_reply_text.place(x=48, y=0, width=175, height=83)

        # 创建垂直滚动条
        scrollbar = ttk.Scrollbar(reply_content_frame, orient='vertical', command=auto_reply_text.yview)
        scrollbar.place(x=55 + 169, y=0, width=15, height=82)

        # 绑定文本框和滚动条
        auto_reply_text['yscrollcommand'] = scrollbar.set

        # 保存对文本框的引用
        self.reply_content_text = auto_reply_text

        # 添加状态显示框架
        status_frame = ttk.Frame(main_frame, style=AppStyles.FRAME_STYLE)
        status_frame.place(x=420, y=205, width=100, height=25)
        
        # 状态标签
        ttk.Label(status_frame, text="监听状态:", style=AppStyles.LABEL_STYLE).pack(side=tk.LEFT)
        
        # 状态图标（使用预加载的图片）
        self.status_label = ttk.Label(status_frame, image=self.images['stop'], style=AppStyles.LABEL_STYLE)
        self.status_label.pack(side=tk.LEFT, padx=2)

        # 按钮组（底部）
        btn_frame = ttk.Frame(main_frame, style=AppStyles.FRAME_STYLE)
        btn_frame.place(x=5, y=200, width=310, height=35)
        
        # 保存配置按钮
        self.save_btn = ttk.Button(btn_frame, text="保存配置", 
                                  style=AppStyles.PRIMARY_BUTTON_STYLE,
                                  width=10, 
                                  command=self.save_config)
        self.save_btn.pack(side=tk.LEFT, padx=1)
        
        # 监测控制按钮（合并启动和停止）
        self.monitor_btn = ttk.Button(btn_frame, text="启动监测",
                                   style=AppStyles.SUCCESS_BUTTON_STYLE,
                                   width=10,
                                   command=self.toggle_monitor)
        self.monitor_btn.pack(side=tk.LEFT, padx=5)
        
        # 下载按钮
        download_btn = ttk.Button(btn_frame, text="下载报表",
                                style=AppStyles.PRIMARY_BUTTON_STYLE,
                                width=10,
                                command=self.show_download_dialog)
        download_btn.pack(side=tk.LEFT, padx=5)

    def create_email_password_group(self, parent, y_position):
        """创建邮箱和密码配对的组件组"""
        # 创建容器Frame
        group_frame = ttk.Frame(parent, style=AppStyles.FRAME_STYLE)
        group_frame.place(x=0, y=y_position, width=240, height=120)
        
        # 标签和添加按钮容器
        header_frame = ttk.Frame(group_frame, style=AppStyles.FRAME_STYLE)
        header_frame.place(x=0, y=0, width=240, height=25)
        
        # 标签
        label = ttk.Label(header_frame, text="监测邮箱:", style=AppStyles.LABEL_STYLE)
        label.place(x=0, y=2)  # 调整标签位置
        
        # 添加按钮（使用预加载的图片）
        add_btn = ttk.Label(header_frame, image=self.images['add'], cursor="hand2", style=AppStyles.LABEL_STYLE)
        add_btn.place(x=200, y=2)  # 使用place来精确定位
        add_btn.bind("<Button-1>", lambda e: self.add_email_password())
        
        # Listbox容器Frame
        list_frame = ttk.Frame(group_frame, style=AppStyles.FRAME_STYLE)
        list_frame.place(x=0, y=25, width=220, height=48)
        
        # 创建Listbox和Scrollbar
        self.email_list = tk.Listbox(list_frame, **AppStyles.get_listbox_config())
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.email_list.yview)
        self.email_list.configure(yscrollcommand=scrollbar.set)
        
        # Listbox和Scrollbar布局
        self.email_list.place(x=0, y=0, width=203, height=48)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定鼠标移动事件来显示/隐藏删除图标
        def on_motion(event):
            index = self.email_list.nearest(event.y)
            if index >= 0:
                bbox = self.email_list.bbox(index)
                if bbox:  # 确保item可见
                    x = bbox[2] + 6  # 调整删除图标位置
                    y = bbox[1] - 5   # 垂直方向微调，使图标居中
                    # 创建或更新删除按钮
                    if not hasattr(self.email_list, 'del_button'):
                        self.email_list.del_button = ttk.Label(self.email_list, image=self.images['del'], cursor="hand2", style=AppStyles.LABEL_STYLE)
                        self.email_list.del_button.bind("<Button-1>", lambda e: self.delete_email_password())
                    self.email_list.del_button.place(x=x, y=y)
                else:
                    if hasattr(self.email_list, 'del_button'):
                        self.email_list.del_button.place_forget()
        
        def on_leave(event):
            if hasattr(self.email_list, 'del_button'):
                self.email_list.del_button.place_forget()
        
        self.email_list.bind('<Motion>', on_motion)
        self.email_list.bind('<Leave>', on_leave)

    def create_list_group(self, parent, label_text, list_name, list_width, list_heigh, x_position, y_position):
        """创建列表组件组"""
        # 创建容器Frame
        group_frame = ttk.Frame(parent, style=AppStyles.FRAME_STYLE)
        group_frame.place(x=x_position, y=y_position, width=250, height=150)
        
        # 标签和添加按钮容器
        header_frame = ttk.Frame(group_frame, style=AppStyles.FRAME_STYLE)
        header_frame.place(x=0, y=0, width=240, height=25)
        
        # 标签
        label = ttk.Label(header_frame, text=f"{label_text}:", style=AppStyles.LABEL_STYLE)
        label.place(x=0, y=2)  # 调整标签位置
        
        # 添加按钮（使用预加载的图片）
        add_btn = ttk.Label(header_frame, image=self.images['add'], cursor="hand2", style=AppStyles.LABEL_STYLE)
        if list_name == 'send_copy_mail':
            add_btn.place(x=220, y=2)  # 使用place来精确定位
        else:
            add_btn.place(x=200, y=2)  # 使用place来精确定位
        add_btn.bind("<Button-1>", lambda e: self.add_item(label_text, list_name))
        
        # Listbox容器Frame
        list_frame = ttk.Frame(group_frame, style=AppStyles.FRAME_STYLE)
        list_frame.place(x=0, y=25, width=list_width+20, height=list_heigh)
        
        # 创建Listbox和Scrollbar
        listbox = tk.Listbox(list_frame, **AppStyles.get_listbox_config())
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        
        # Listbox和Scrollbar布局 - 使用place代替pack
        listbox.place(x=0, y=0, width=list_width, height=list_heigh)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar.place(x=list_width+1, y=0, width=15, height=list_heigh)
        
        # 保存引用
        setattr(self, f"{list_name}_list", listbox)
        
        # 绑定鼠标移动事件来显示/隐藏删除图标
        def on_motion(event):
            index = listbox.nearest(event.y)
            if index >= 0:
                bbox = listbox.bbox(index)
                if bbox:  # 确保item可见
                    x = bbox[2] + 6  # 调整删除图标位置
                    y = bbox[1] - 5   # 垂直方向微调，使图标居中
                    # 创建或更新删除按钮
                    if not hasattr(listbox, 'del_button'):
                        listbox.del_button = ttk.Label(listbox, image=self.images['del'], cursor="hand2", style=AppStyles.LABEL_STYLE)
                        listbox.del_button.bind("<Button-1>", lambda e: self.delete_item(list_name))
                    listbox.del_button.place(x=x, y=y)
                else:
                    if hasattr(listbox, 'del_button'):
                        listbox.del_button.place_forget()
        
        def on_leave(event):
            if hasattr(listbox, 'del_button'):
                listbox.del_button.place_forget()
        
        listbox.bind('<Motion>', on_motion)
        listbox.bind('<Leave>', on_leave)

    def add_item(self, label_text, list_name):
        """添加项目到列表"""
        # 确保业务类型列表已加载
        if not self.business_types or self.business_types == ["获取数据失败"]:
            try:
                # 重新加载配置
                config_file = os.path.join('resources', 'mail_client_cfg.json')
                if os.path.exists(config_file):
                    with open(config_file, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                        self.business_types = config.get('business_types', ["数金总部", "长亮金服", "长亮合度", "长亮科技", "长亮控股", "数据总部"])
                else:
                    self.business_types = ["数金总部", "长亮金服", "长亮合度", "长亮科技", "长亮控股", "数据总部"]
            except Exception as e:
                print(f"加载业务类型失败: {str(e)}")
                self.business_types = ["数金总部", "长亮金服", "长亮合度", "长亮科技", "长亮控股", "数据总部"]

        dialog = self.create_modal_dialog(f"添加{label_text}", 280, 250)
        
        # 创建Frame容器
        input_frame = ttk.Frame(dialog, style=AppStyles.FRAME_STYLE)
        input_frame.pack(pady=10, padx=20, fill=tk.X)
        
        # 业务主体
        ttk.Label(input_frame, text="业务主体:", style=AppStyles.LABEL_STYLE).pack(anchor=tk.W)
        business_var = tk.StringVar()
        business_combo = ttk.Combobox(input_frame, 
                                    textvariable=business_var,
                                    values=self.business_types,
                                    state='readonly',
                                    style=AppStyles.COMBOBOX_STYLE)
        business_combo.pack(fill=tk.X, pady=(5, 10))
        
        # 设置默认值
        if self.business_types:
            business_combo.set(self.business_types[0])
            print(f"设置默认业务类型: {self.business_types[0]}")  # 添加调试日志
        
        # 用户姓名输入
        ttk.Label(input_frame, text="用户姓名:", style=AppStyles.LABEL_STYLE).pack(anchor=tk.W)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(input_frame, textvariable=name_var, style=AppStyles.ENTRY_STYLE)
        name_entry.pack(fill=tk.X, pady=(5, 10))
        
        # 邮箱地址输入
        ttk.Label(input_frame, text="邮箱地址:", style=AppStyles.LABEL_STYLE).pack(anchor=tk.W)
        email_var = tk.StringVar()
        email_entry = ttk.Entry(input_frame, textvariable=email_var, style=AppStyles.ENTRY_STYLE)
        email_entry.pack(fill=tk.X, pady=(5, 10))
        
        def confirm():
            # 获取输入值
            business = business_var.get().strip()
            name = name_var.get().strip()
            email = email_var.get().strip()
            
            # 验证输入
            if not all([business, name, email]):
                messagebox.showwarning("警告", "所有字段都必须填写")
                return
            
            # 添加到列表
            listbox = getattr(self, f"{list_name}_list")
            item_text = f"{business}|{name}|{email}"
            listbox.insert(tk.END, item_text)
            
            # 关闭对话框
            dialog.overlay.destroy()
            dialog.destroy()
        
        # 按钮Frame
        btn_frame = ttk.Frame(dialog, style=AppStyles.FRAME_STYLE)
        btn_frame.pack(side=tk.BOTTOM, pady=10)
        
        ttk.Button(btn_frame, text="确 定", style=AppStyles.PRIMARY_BUTTON_STYLE,
                   command=confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取 消", style=AppStyles.DANGER_BUTTON_STYLE,
                   command=lambda: (dialog.overlay.destroy(), dialog.destroy())).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        dialog.bind('<Return>', lambda e: confirm())
        business_combo.focus_set()

    def delete_item(self, list_name):
        listbox = getattr(self, f"{list_name}_list")
        selection = listbox.curselection()
        if selection:
            listbox.delete(selection[0])

    # def on_mode_changed(self):
    #     if self.mode_var.get() == "scheduled":
    #         self.time_frame.place(x=10, y=345, width=390, height=25)  # 显示时间选择框
    #     else:
    #         self.time_frame.place_forget()  # 隐藏时间选择框

    @staticmethod
    def handle_request_errors(show_config_dialog=True, timeout=5):
        """处理网络请求错误的装饰器
        
        Args:
            show_config_dialog (bool): 发生错误时是否显示服务器配置对话框
            timeout (int): 请求超时时间(秒)
        """
        def decorator(func):
            def wrapper(self, *args, **kwargs):
                try:
                    # 将timeout参数注入到被装饰函数中
                    if 'timeout' not in kwargs:
                        kwargs['timeout'] = timeout
                    return func(self, *args, **kwargs)
                except requests.exceptions.Timeout:
                    error_msg = f"{func.__name__}超时"
                    print(error_msg)
                    messagebox.showerror("错误", "连接服务器超时，请检查服务器是否启动")
                    if show_config_dialog:
                        self.show_server_config()
                except requests.exceptions.ConnectionError:
                    error_msg = f"{func.__name__}失败: 无法连接到服务器"
                    print(error_msg)
                    messagebox.showerror("错误", "无法连接到服务器，请确保服务已启动")
                    if show_config_dialog:
                        self.show_server_config()
                except Exception as e:
                    error_msg = f"{func.__name__}失败: {str(e)}"
                    print(error_msg)
                    if isinstance(e, requests.exceptions.RequestException):
                        messagebox.showerror("错误", f"网络请求失败: {str(e)}")
                    else:
                        messagebox.showerror("错误", error_msg)
                    if show_config_dialog:
                        self.show_server_config()
                return None
            return wrapper
        return decorator

    @handle_request_errors()
    def load_config(self, timeout=5):
        """加载配置"""
        print("开始加载配置...")
        response = requests.get(f"{self.api_base_url}/config", timeout=timeout)
        if response.status_code == 200:
            config = response.json()['data']
            print(f"从服务器获取的配置: {config}")
            
            # 加载监测邮箱配置
            self.recv_monitor_mail = config.get('recv_monitor_mail', [])
            self.email_list.delete(0, tk.END)
            for email_config in self.recv_monitor_mail:
                self.email_list.insert(tk.END, email_config['email'])
            
            # 加载抄送地址
            self.send_copy_mail_list.delete(0, tk.END)
            for cc in config.get('send_copy_mail', []):
                display_text = f"{cc['business']}|{cc['name']}|{cc['email']}"
                self.send_copy_mail_list.insert(tk.END, display_text)
            
            # 加载准入邮箱
            self.auth_enter_mail_list.delete(0, tk.END)
            for auth in config.get('auth_enter_mail', []):
                display_text = f"{auth['business']}|{auth['name']}|{auth['email']}"
                self.auth_enter_mail_list.insert(tk.END, display_text)
            
            # 加载关键词
            self.keyword_var.set(config.get('keywords', ''))
            
            # 加载回复正文
            self.reply_content_text.delete('1.0', tk.END)
            self.reply_content_text.insert('1.0', config.get('auto_reply_text', ''))
            
            # 加载监控状态
            self.check_monitor_status()
        else:
            print(f"API请求失败，状态码: {response.status_code}")
            messagebox.showwarning("警告", "无法从服务器加载配置，将使用默认配置")
            self.show_server_config()

    @handle_request_errors()
    def save_config(self, timeout=5):
        """保存配置"""
        # 准备新的配置
        config = {
            'recv_monitor_mail': self.get_recv_monitor_mail(),
            'send_copy_mail': [
                {
                    'business': item.split('|')[0],
                    'name': item.split('|')[1],
                    'email': item.split('|')[2]
                }
                for item in self.send_copy_mail_list.get(0, tk.END)
            ],
            'auth_enter_mail': [
                {
                    'business': item.split('|')[0],
                    'name': item.split('|')[1],
                    'email': item.split('|')[2]
                }
                for item in self.auth_enter_mail_list.get(0, tk.END)
            ],
            'keywords': self.keyword_var.get(),
            'auto_reply_text': self.reply_content_text.get('1.0', 'end-1c')
        }
        
        print("准备保存的配置:", config)
        response = requests.post(f"{self.api_base_url}/config", json=config, timeout=timeout)
        
        if response.status_code == 200:
            result = response.json()
            if result.get('status') == 'success':
                messagebox.showinfo("成功", "配置已保存")
                print("配置保存成功")
            else:
                error_msg = result.get('message', '未知错误')
                messagebox.showerror("错误", f"保存配置失败: {error_msg}")
                print(f"保存配置失败: {error_msg}")
        else:
            messagebox.showerror("错误", f"保存配置失败，状态码: {response.status_code}")
            print(f"保存配置失败，状态码: {response.status_code}")

    @handle_request_errors(show_config_dialog=False)
    def check_monitor_status(self, timeout=5):
        """检查监控状态并更新界面"""
        response = requests.get(f"{self.api_base_url}/monitor/status", timeout=timeout)
        if response.status_code == 200:
            status = response.json().get('status', 'stopped')
            # 更新按钮状态和显示
            if status == 'running':
                self.monitor_btn.configure(
                    text="停止监测",
                    style=AppStyles.DANGER_BUTTON_STYLE,
                    command=self.stop_monitor
                )
                self.status_label.configure(image=self.images['run'])
            else:
                self.monitor_btn.configure(
                    text="启动监测",
                    style=AppStyles.SUCCESS_BUTTON_STYLE,
                    command=self.start_monitor
                )
                self.status_label.configure(image=self.images['stop'])
        # 如果请求失败,装饰器会处理异常,这里不需要else分支

    @handle_request_errors()
    def start_monitor(self, timeout=5):
        """启动监控"""
        response = requests.post(f"{self.api_base_url}/monitor/start", timeout=timeout)
        if response.status_code == 200:
            messagebox.showinfo("成功", "监测已启动")
            self.check_monitor_status()
        else:
            messagebox.showerror("错误", "启动监测失败")

    @handle_request_errors()
    def stop_monitor(self, timeout=5):
        """停止监控"""
        response = requests.post(f"{self.api_base_url}/monitor/stop", timeout=timeout)
        if response.status_code == 200:
            messagebox.showinfo("成功", "监测已停止")
            self.check_monitor_status()
        else:
            messagebox.showerror("错误", "停止监测失败")

    def get_recv_monitor_mail(self):
        """获取所有邮箱配置"""
        try:
            # 获取当前显示的邮箱列表
            new_emails = [self.email_list.get(i) for i in range(self.email_list.size())]
            
            # 构建新的配置列表
            new_configs = []
            for email in new_emails:
                # 从实例变量中查找配置
                existing_config = next(
                    (cfg.copy() for cfg in self.recv_monitor_mail if cfg['email'] == email),
                    None
                )
                
                if existing_config:
                    # 确保配置中包含所有必要的字段
                    if 'erpuser' not in existing_config:
                        existing_config['erpuser'] = ""
                    if 'erppwd' not in existing_config:
                        existing_config['erppwd'] = ""
                    new_configs.append(existing_config)
                else:
                    # 如果在实例变量中找不到，创建新的配置
                    new_config = {
                        'email': email,
                        'password': '',  # 空密码，因为没有配置
                        'erpuser': '',
                        'erppwd': ''
                    }
                    new_configs.append(new_config)
            
            return new_configs
                
        except Exception as e:
            print(f"获取邮箱配置失败: {str(e)}")
            print(f"错误详情: {traceback.format_exc()}")
            messagebox.showerror("错误", f"获取邮箱配置失败: {str(e)}")
            return []

    def is_encrypted(self, password):
        """检查密码是否已经是加密形式"""
        try:
            # 尝试解密，如果成功则说明是加密的密码
            self.password_manager.decrypt(password)
            return True
        except:
            return False

    def create_modal_dialog(self, title, width=300, height=200):
        """创建带遮罩的模态对话框"""
        # 获取主窗口
        root = self.tab.winfo_toplevel()
        root_x = root.winfo_x()
        root_y = root.winfo_y()
        root_width = root.winfo_width()
        root_height = root.winfo_height()
        
        # 创建遮罩 - 使用浅灰色半透明效果
        overlay = tk.Frame(root)  # 附加到主窗口
        overlay.place(x=0, y=0, width=root_width, height=root_height)
        overlay.configure(bg='gray75')  # 使用浅灰色
        
        # 创建对话框
        dialog = tk.Toplevel(root)
        dialog.title(title)
        dialog.configure(bg='white')
        
        # 计算对话框位置使其居中
        dialog_x = root_x + (root_width - width) // 2
        dialog_y = root_y + (root_height - height) // 2
        dialog.geometry(f"{width}x{height}+{dialog_x}+{dialog_y}")
        
        # 设置对话框属性
        dialog.transient(root)
        dialog.grab_set()
        dialog.focus_set()
        dialog.resizable(False, False)
        
        # 保存遮罩的引用到对话框
        dialog.overlay = overlay
        
        def on_dialog_close():
            overlay.destroy()
            dialog.destroy()
        
        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        return dialog

    def load_server_config(self):
        """加载服务器配置"""
        try:
            config_file = os.path.join('resources', 'mail_client_cfg.json')
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.server_host = config.get('server_host', 'localhost')
                    self.server_port = config.get('server_port', '5000')
                    self.server_protocol = config.get('server_protocol', 'http')
                    # 加载业务类型列表
                    self.business_types = config.get('business_types', ["获取数据失败"])
            else:
                self.server_host = 'localhost'
                self.server_port = '5000'
                self.server_protocol = 'http'
                self.business_types = ["获取数据失败"]
                self.save_server_config()
            
            self.api_base_url = f"{self.server_protocol}://{self.server_host}:{self.server_port}/api"
            print(f"API基础URL: {self.api_base_url}")
        except Exception as e:
            print(f"加载服务器配置失败: {str(e)}")
            messagebox.showerror("错误", f"加载服务器配置失败: {str(e)}")

    def save_server_config(self):
        """保存服务器配置"""
        try:
            config = {
                'server_host': self.server_host,
                'server_port': self.server_port,
                'server_protocol': self.server_protocol,
                'business_types': self.business_types
            }
            config_file = os.path.join('resources', 'mail_client_cfg.json')
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"保存服务器配置失败: {str(e)}")
            messagebox.showerror("错误", f"保存服务器配置失败: {str(e)}")

    def show_server_config(self):
        """显示服务器配置对话框"""
        dialog = self.create_modal_dialog("服务器配置", width=330, height=250)
        
        # 创建Frame容器
        input_frame = ttk.Frame(dialog, style=AppStyles.FRAME_STYLE)
        input_frame.pack(pady=10, padx=20, fill=tk.X)
        
        # 服务器地址
        ttk.Label(input_frame, text="服务器地址:", style=AppStyles.LABEL_STYLE).pack(anchor=tk.W)
        host_var = tk.StringVar(value=self.server_host)
        host_entry = ttk.Entry(input_frame, textvariable=host_var, style=AppStyles.ENTRY_STYLE)
        host_entry.pack(fill=tk.X, pady=(5, 10))
        
        # 端口
        ttk.Label(input_frame, text="端口:", style=AppStyles.LABEL_STYLE).pack(anchor=tk.W)
        port_var = tk.StringVar(value=self.server_port)
        port_entry = ttk.Entry(input_frame, textvariable=port_var, style=AppStyles.ENTRY_STYLE)
        port_entry.pack(fill=tk.X, pady=(5, 10))
        
        # 协议选择
        ttk.Label(input_frame, text="协议:", style=AppStyles.LABEL_STYLE).pack(anchor=tk.W)
        protocol_var = tk.StringVar(value=self.server_protocol)
        protocol_combo = ttk.Combobox(input_frame, textvariable=protocol_var,
                                     values=['http', 'https'], state='readonly',
                                     style=AppStyles.COMBOBOX_STYLE)
        protocol_combo.pack(fill=tk.X, pady=(5, 10))
        
        def test_and_save_config():
            try:
                # 获取输入值
                host = host_var.get().strip()
                port = port_var.get().strip()
                protocol = protocol_var.get().strip()
                
                if not all([host, port, protocol]):
                    messagebox.showwarning("警告", "所有字段都必须填写")
                    return
                
                # 构建测试用的API URL
                test_api_url = f"{protocol}://{host}:{port}/api/monitor/status"
                
                try:
                    # 测试连接
                    response = requests.get(test_api_url, timeout=5)  # 添加超时设置
                    if response.status_code == 200:
                        # 连接成功，保存配置
                        self.server_host = host
                        self.server_port = port
                        self.server_protocol = protocol
                        self.api_base_url = f"{protocol}://{host}:{port}/api"
                        
                        # 保存配置到文件
                        self.save_server_config()
                        
                        # 加载配置到主界面
                        self.load_config()
                        
                        messagebox.showinfo("成功", "服务器连接成功，配置已保存")
                        dialog.overlay.destroy()
                        dialog.destroy()
                    else:
                        messagebox.showerror("错误", f"服务器连接失败，状态码: {response.status_code}")
                except requests.exceptions.ConnectionError:
                    messagebox.showerror("错误", "无法连接到服务器，请检查服务器地址和端口")
                except requests.exceptions.Timeout:
                    messagebox.showerror("错误", "连接服务器超时")
                except Exception as e:
                    messagebox.showerror("错误", f"连接测试失败: {str(e)}")
                    
            except Exception as e:
                messagebox.showerror("错误", f"保存配置失败: {str(e)}")
        
        # 按钮Frame
        btn_frame = ttk.Frame(dialog, style=AppStyles.FRAME_STYLE)
        btn_frame.pack(side=tk.BOTTOM, pady=10)
        
        ttk.Button(btn_frame, text="连 接", style=AppStyles.PRIMARY_BUTTON_STYLE,
                   command=test_and_save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取 消", style=AppStyles.DANGER_BUTTON_STYLE,
                   command=lambda: (dialog.overlay.destroy(), dialog.destroy())).pack(side=tk.LEFT, padx=5)

    def toggle_monitor(self):
        """切换监控状态"""
        try:
            response = requests.get(f"{self.api_base_url}/monitor/status")
            if response.status_code == 200:
                status = response.json().get('status', 'stopped')
                if status == 'running':
                    # 如果正在运行，则停止
                    self.stop_monitor()
                else:
                    # 如果已停止，则启动
                    self.start_monitor()
            else:
                messagebox.showerror("错误", "获取监控状态失败")
        except Exception as e:
            messagebox.showerror("错误", f"切换监控状态失败: {str(e)}")

    def show_download_dialog(self):
        """显示下载报表对话框"""
        dialog = self.create_modal_dialog("下载报表", width=300, height=200)
        
        # 创建Frame容器
        input_frame = ttk.Frame(dialog, style=AppStyles.FRAME_STYLE)
        input_frame.pack(pady=10, padx=20, fill=tk.X)
        
        # 年月选择Frame
        date_frame = ttk.Frame(input_frame, style=AppStyles.FRAME_STYLE)
        date_frame.pack(fill=tk.X, pady=5)
        
        # 年份选择
        ttk.Label(date_frame, text="年份:", style=AppStyles.LABEL_STYLE).pack(side=tk.LEFT, padx=(0, 5))
        current_year = datetime.datetime.now().year
        year_values = [str(year) for year in range(current_year-5, current_year+1)]
        year_var = tk.StringVar(value=str(current_year))
        year_combo = ttk.Combobox(date_frame, textvariable=year_var, 
                                 values=year_values, width=10, style=AppStyles.COMBOBOX_STYLE)
        year_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        # 月份选择
        ttk.Label(date_frame, text="月份:", style=AppStyles.LABEL_STYLE).pack(side=tk.LEFT, padx=(15, 5))
        current_month = datetime.datetime.now().month
        month_values = [f"{month:02d}" for month in range(1, 13)]
        month_var = tk.StringVar(value=f"{current_month:02d}")
        month_combo = ttk.Combobox(date_frame, textvariable=month_var, 
                                  values=month_values, width=8, style=AppStyles.COMBOBOX_STYLE)
        month_combo.pack(side=tk.LEFT)
        
        # 下载目录选择Frame
        dir_frame = ttk.Frame(input_frame, style=AppStyles.FRAME_STYLE)
        dir_frame.pack(fill=tk.X, pady=12)
        
        ttk.Label(dir_frame, text="下载目录:", style=AppStyles.LABEL_STYLE).pack(anchor=tk.W)
        
        # 目录输入和选择按钮Frame
        dir_select_frame = ttk.Frame(dir_frame, style=AppStyles.FRAME_STYLE)
        dir_select_frame.pack(fill=tk.X, pady=5)
        
        # 默认下载目录
        current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        default_dir = os.path.join(current_dir, "download")
        dir_var = tk.StringVar(value=default_dir)
        dir_entry = ttk.Entry(dir_select_frame, textvariable=dir_var, style=AppStyles.ENTRY_STYLE)
        dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        def select_directory():
            directory = tk.filedialog.askdirectory(initialdir=dir_var.get())
            if directory:
                dir_var.set(directory)
        
        browse_btn = ttk.Button(dir_select_frame, text="浏览", 
                               style=AppStyles.PRIMARY_BUTTON_STYLE,
                               command=select_directory)
        browse_btn.pack(side=tk.RIGHT)
        
        @handle_request_errors(timeout=10)  # 下载需要更长的超时时间
        def download_report(year_month, timeout=10):
            """实际执行下载的函数"""
            response = requests.get(
                f"{self.api_base_url}/mail_monitor/download_report",
                params={"year_month": year_month},
                timeout=timeout
            )
            
            if response.status_code == 404:
                error_msg = "未找到报表文件"
                try:
                    error_msg = response.json().get('message', error_msg)
                except:
                    pass
                messagebox.showwarning("警告", error_msg)
                return None
            elif response.status_code != 200:
                error_msg = "下载报表失败"
                try:
                    error_msg = response.json().get('message', error_msg)
                except:
                    pass
                messagebox.showerror("错误", error_msg)
                return None
            
            return response

        def download():
            try:
                year_month = f"{year_var.get()}{month_var.get()}"
                download_dir = dir_var.get()
                
                # 检查目录是否存在
                if not os.path.exists(download_dir):
                    try:
                        os.makedirs(download_dir)
                    except Exception as e:
                        messagebox.showerror("错误", f"创建下载目录失败: {str(e)}")
                        return
                
                # 发送下载请求
                response = download_report(year_month)
                if not response:
                    return
                
                # 检查响应头，确保是Excel文件
                content_type = response.headers.get('content-type', '')
                if 'spreadsheet' not in content_type.lower():
                    messagebox.showerror("错误", "服务器返回的不是有效的Excel文件")
                    return
                
                # 保存文件
                file_path = os.path.join(download_dir, f"mail_report_{year_month}.xlsx")
                with open(file_path, "wb") as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                
                messagebox.showinfo("成功", f"报表已下载到: {file_path}")
                
                # 打开下载目录
                os.startfile(download_dir)
                
                # 关闭对话框
                dialog.overlay.destroy()
                dialog.destroy()
                    
            except Exception as e:
                error_msg = f"下载报表时出错: {str(e)}"
                print(error_msg)
                messagebox.showerror("错误", error_msg)
        
        # 按钮Frame
        btn_frame = ttk.Frame(dialog, style=AppStyles.FRAME_STYLE)
        btn_frame.pack(side=tk.BOTTOM, pady=10)
        
        ttk.Button(btn_frame, text="下 载", style=AppStyles.PRIMARY_BUTTON_STYLE,
                   command=download).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取 消", style=AppStyles.DANGER_BUTTON_STYLE,
                   command=lambda: (dialog.overlay.destroy(), dialog.destroy())).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        dialog.bind('<Return>', lambda e: download())
        year_combo.focus_set()

class PasswordManager:
    def __init__(self):
        self.key = ENCRYPTION_KEY  # 使用相同的密钥
        self.cipher_suite = Fernet(self.key)

    def encrypt(self, password):
        """加密密码"""
        return self.cipher_suite.encrypt(password.encode()).decode()

    def decrypt(self, encrypted_password):
        """解密密码"""
        try:
            return self.cipher_suite.decrypt(encrypted_password.encode()).decode()
        except:
            return encrypted_password  # 如果解密失败，返回原始值 