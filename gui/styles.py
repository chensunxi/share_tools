from tkinter import ttk
import tkinter as tk

class AppStyles:
    # 颜色常量
    BG_COLOR = '#F5F5F5'  # 主背景色
    PRIMARY_COLOR = '#4A90E2'  # 主要按钮颜色
    DANGER_COLOR = '#E74C3C'  # 危险操作按钮颜色
    SUCCESS_COLOR = '#2ECC71'  # 成功/确认按钮颜色
    TEXT_COLOR = '#333333'  # 主要文字颜色
    DISABLED_COLOR = '#CCCCCC'  # 禁用状态颜色
    
    # 字体配置
    FONT_FAMILY = 'SimSun'
    FONT_SIZE = 9
    
    # 组件尺寸
    BUTTON_WIDTH = 10
    BUTTON_HEIGHT = 2
    PADDING = 5
    
    # 边框配置
    ENTRY_BORDER_WIDTH = 1
    ENTRY_RELIEF = 'solid'  # 可选值: flat, raised, sunken, ridge, solid, groove
    
    # 组件样式配置
    FRAME_STYLE = 'Custom.TFrame'
    LABEL_STYLE = 'Custom.TLabel'
    TITLE_LABEL_STYLE = 'Title.TLabel'
    PRIMARY_BUTTON_STYLE = 'Primary.TButton'
    SUCCESS_BUTTON_STYLE = 'Success.TButton'
    DANGER_BUTTON_STYLE = 'Danger.TButton'
    ENTRY_STYLE = 'Custom.TEntry'
    COMBOBOX_STYLE = 'Custom.TCombobox'
    NOTEBOOK_STYLE = 'Custom.TNotebook'
    
    @classmethod
    def configure_styles(cls):
        """配置全局样式"""
        style = ttk.Style()
        
        # 配置Frame样式
        style.configure(cls.FRAME_STYLE, 
                       background=cls.BG_COLOR)
        
        # 配置Label样式
        style.configure(cls.LABEL_STYLE, 
                       background=cls.BG_COLOR,
                       foreground=cls.TEXT_COLOR,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE))
        
        # 配置标题Label样式
        style.configure(cls.TITLE_LABEL_STYLE,
                       background=cls.BG_COLOR,
                       foreground=cls.TEXT_COLOR,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE, 'bold'))
        
        # 配置按钮样式
        style.configure(cls.PRIMARY_BUTTON_STYLE,
                       background=cls.PRIMARY_COLOR,
                       foreground='white',
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE, 'bold'),
                       width=cls.BUTTON_WIDTH)
        
        style.configure(cls.SUCCESS_BUTTON_STYLE,
                       background=cls.SUCCESS_COLOR,
                       foreground='white',
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE, 'bold'),
                       width=cls.BUTTON_WIDTH)
        
        style.configure(cls.DANGER_BUTTON_STYLE,
                       background=cls.DANGER_COLOR,
                       foreground='white',
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE, 'bold'),
                       width=cls.BUTTON_WIDTH)
        
        # 配置Entry样式
        style.configure(cls.ENTRY_STYLE,
                       background='white',
                       foreground=cls.TEXT_COLOR,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE),
                       borderwidth=cls.ENTRY_BORDER_WIDTH,
                       relief=cls.ENTRY_RELIEF)
        
        # 配置Combobox样式
        style.configure(cls.COMBOBOX_STYLE,
                       background='white',
                       foreground=cls.TEXT_COLOR,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE))
        
        # 配置Notebook样式
        style.configure(cls.NOTEBOOK_STYLE,
                       background=cls.BG_COLOR)
        style.configure(f"{cls.NOTEBOOK_STYLE}.Tab",
                       background=cls.BG_COLOR,
                       font=(cls.FONT_FAMILY, cls.FONT_SIZE))
    
    @classmethod
    def get_listbox_config(cls):
        """获取Listbox配置"""
        return {
            'bg': 'white',
            'fg': cls.TEXT_COLOR,
            'font': (cls.FONT_FAMILY, cls.FONT_SIZE),
            'selectbackground': cls.PRIMARY_COLOR,
            'selectforeground': 'white',
            'relief': cls.ENTRY_RELIEF,
            'borderwidth': cls.ENTRY_BORDER_WIDTH
        }
    
