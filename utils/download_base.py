# -*- coding: utf-8 -*-
# Author:   chensx
# At    :   2024/11/15
# Email :   chensx@sunline.com
# About :   社保下载基类
import os
import re
import sys
import time
import traceback
import io

import pandas as pd
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pickle
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
import tkinter.messagebox as messagebox
from utils.logger_utils import LoggerUtils
from utils.common_utils import extract_chinese
from chromedriver_autoinstaller import install as install_chromedriver

# 自动安装或更新 ChromeDriver
install_chromedriver()

class SharedState:
    def __init__(self):
        self.downloaded_file = None

# 用一个共享对象管理 downloaded_file
shared_state = SharedState()

class DownloadHandler(FileSystemEventHandler):
    def __init__(self, shared_state):
        self.shared_state = shared_state
    def on_created(self, event):
        print(f"Received file event: {event.src_path}")

        # 如果创建的是文件夹，直接返回
        if event.is_directory or "downloads.htm" in event.src_path:
            print(f"跳过临时下载")
            return

        # 记录下载的文件
        self.shared_state.downloaded_file = event.src_path

def load_user_data(user_template, city_name=''):
    # user_template = f'{city_map[city_name]}社保证明打印模版.xlsx'
    if not os.path.exists(user_template):
        return

    def row_to_dict(row):
        return {
            "start_date": row.get("起始日期", ""),
            "end_date": row.get("终止日期", ""),
            "id_card": row.get("证件号码", ""),
            "staff_name": extract_chinese(row.get("工作名", "")),
            "seq_id": row.get("序号", ""),
            "id_social": row.get("电脑号", ""),
        }

    # 读取所有工作表的名称
    xl = pd.ExcelFile(user_template)
    sheet_names = xl.sheet_names

    # 使用正则表达式进行模糊匹配
    matched_sheets = [sheet for sheet in sheet_names if re.search(city_name, sheet, re.IGNORECASE)]
    if matched_sheets:
        data = pd.read_excel(user_template,sheet_name=matched_sheets[0], dtype={'证件号码': str}, skiprows=1).fillna('')
        data = data.apply(row_to_dict, axis=1).tolist()
    else:
        return
    return data

class CrawlerDownloadBase(object):
    def __init__(self, eng_city='', download_path='', chn_city='', name_list='', download_flag='', download_limit='', out_logger=''):
        self.logger = LoggerUtils.get_logger(None, out_logger)  # 调用静态方法获取 logger
        self.city = eng_city
        self.out_logger = out_logger
        self.download_path = download_path or os.path.join(os.getcwd(), 'source')
        self.chn_city = chn_city
        self.name_list = name_list
        self.download_flag = download_flag
        self.download_limit = download_limit
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)
        self.driver = self.init_browser()
        self.driver.implicitly_wait(1)
        self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": self.download_path
        })
        self.vars = {}
        self.user_data = None  # 初始化 user_data
        self.shared_state = shared_state
        # 启动监控
        event_handler = DownloadHandler(shared_state)
        self.observer = Observer()
        self.observer.schedule(event_handler, self.download_path, recursive=False)
        self.observer.start()

    def init_browser(self):
        self.logger.info(f'下载目录：{self.download_path}')
        # try:
        #     # 尝试创建一个 ChromeDriver 实例，检查是否已经安装了 ChromeDriver
        #     webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        #     print("ChromeDriver 已成功启动!")
        # except WebDriverException as e:
        #     print(f"WebDriverException: {e}")
        #     print("未检测到可用的 ChromeDriver，正在自动下载...")
        #
        #     # 如果没有找到ChromeDriver，自动下载并启动
        #     webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        #     print("ChromeDriver 已下载并成功启动!")
        options = webdriver.ChromeOptions()
        options.add_experimental_option('prefs', {
            "download.default_directory": self.download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True  # 始终使用外部程序打开PDF
        })
        # mywebdriver = webdriver.Chrome(options=options)
        # print(f"Chrome Driver安装路径: {mywebdriver.service.path}")

        return webdriver.Chrome(options=options)

    def web_url(self, type='beijing'):
        urls = {
            'beijing': 'https://fuwu.rsj.beijing.gov.cn/zhrs/yltc/yltc-home',
            'shanghai': 'https://zwdt.sh.gov.cn/govPortals/index.do',
            'guangzhou': 'https://ggfw.hrss.gd.gov.cn/gdggfw/index',
            'shenzhen': 'https://sipub.sz.gov.cn/hsoms/',
            'hangzhou': 'https://oauth.zjzwfw.gov.cn/login.jsp',
            'changsha': 'https://www.cs12333.com/neaf-ui-cs/#/login',
            'nanjing': 'https://rs.jshrss.jiangsu.gov.cn/index/',
            'sunerp': 'https://iboss.sunline.cn/#/login',
            'erptest': 'https://test.erp.sunline.cn:9086/#/login'
        }
        return urls[type]

    def get_cookie(self):
        with open(f"cookies_{self.city}.pkl", "wb") as file:
            pickle.dump(self.driver.get_cookies(), file)

    def load_cookie(self):
        with open(f"cookies_{self.city}.pkl", "rb") as file:
            cookies = pickle.load(file)
            for cookie in cookies:
                self.driver.add_cookie(cookie)

        self.driver.refresh()

    def wait(self, ele, timout=20, interval=2, retries=3):
        # logger.info(f'等待元素加载,{ele}')
        attempt = 0
        while attempt < retries:
            try:
                WebDriverWait(self.driver, timout, interval).until(
                    EC.presence_of_element_located(ele)
                )
                # 确保返回的是一个 Web 元素对象
                if isinstance(ele, WebElement):
                    print(f"页面元素【{ele.text}】已找到")
                return True  # 找到了元素，返回True
            except TimeoutException:
                attempt += 1
                self.logger.warning(f"第{attempt}次重试：等待元素超时，重新尝试...")
                time.sleep(2)  # 等待2秒后重试
        self.logger.error("所有重试均失败")
        return False  # 如果重试了多次都没找到元素，返回False
    
    def click_custom(self, ele, click_flag=True):
        """点击元素的自定义方法"""
        attempt = 0
        retries = 15
        while attempt < retries:
            try:
                # 如果 ele 是元组，直接使用
                if isinstance(ele, tuple):
                    by, value = ele
                    element = self.driver.find_element(by=by, value=value)
                # 如果 ele 是字典，从字典中获取值
                elif isinstance(ele, dict):
                    by = ele.get('by')
                    value = ele.get('value')
                    element = self.driver.find_element(by=by, value=value)
                else:
                    element = ele
                
                if click_flag:
                    element.click()
                else:
                    self.driver.execute_script("arguments[0].click()", element)
                return True
            except Exception as e:
                attempt += 1
                # 只获取异常类型和简要信息
                error_type = e.__class__.__name__
                error_msg = str(e).split('\n')[0]  # 只取第一行
                self.logger.warning(f"第{attempt}次重试：点击元素失败 - {error_type}: {error_msg}")
                time.sleep(3)
            
        error_msg = f"所有重试均失败，无法点击元素: {ele}"
        self.logger.error(error_msg)
        raise Exception(error_msg)  # 抛出异常而不是返回 False

    def login(self, city):
        url = self.web_url(city)
        # url = self.web_url('home')
        self.driver.get(url)
        # self.get_cookie()
        # self.load_cookie()
        # logger.info(f'您有60秒的时间进行登录，超时后需重新执行程序')
        # time.sleep(8)
        # self.wait((By.CSS_SELECTOR, 'p[title="单位参保证明查询打印"]'))

    def quit(self):
        if self.observer:  # 先检查 observer 是否已经初始化
            self.observer.stop()
            self.observer.join()
        # self.driver.quit()

    def click(self, ele):
        self.driver.execute_script("arguments[0].click()", ele)

    def generate_unique_filename(self, new_file_path):
        base, ext = os.path.splitext(new_file_path)
        counter = 1

        # 如果文件已经存在，则递增后缀
        while os.path.exists(new_file_path):
            new_file_path = f"{base}_{counter}{ext}"
            counter += 1

        return new_file_path

    def rename_downfile(self, new_name):
        # 等待文件下载完成
        timeout = 10  # 最大等待时间10秒
        elapsed_time = 0
        old_file_path = None
        while self.shared_state.downloaded_file is None and elapsed_time < timeout:
            self.logger.info(f'未开始下载文件，重试第{elapsed_time + 1}次')
            time.sleep(3)
            elapsed_time += 1

        if self.shared_state.downloaded_file is None:
            print("Download timed out or file not finished.")
        elif self.shared_state.downloaded_file.endswith('.crdownload'):
            # print(f"Downloading file: {self.shared_state.downloaded_file}")
            # 去掉 .crdownload 后缀
            old_file_path = self.shared_state.downloaded_file[:-11]
            elapsed_time = 0
            while not os.path.exists(old_file_path) and elapsed_time < timeout:
                self.logger.info(f'下载文件未完成，重试第{elapsed_time + 1}次')
                old_file_path = self.shared_state.downloaded_file[:-11]
                time.sleep(7)
                elapsed_time += 1
        else:
            print(f"File downloaded: {self.shared_state.downloaded_file}")
            old_file_path = self.shared_state.downloaded_file
        # 检查目标文件是否存在
        _, ext_name = os.path.splitext(old_file_path)
        new_file_path = f"{new_name}{ext_name}"
        if os.path.exists(new_file_path):
            print(f"Target file '{new_file_path}' already exists.")
            # 你可以选择覆盖文件，或者选择其他操作
            # 比如删除目标文件（谨慎操作）
            # os.remove(new_file_path)
            # 或者选择生成一个新的文件名
            # 调用生成唯一文件名的函数
            new_file_path = self.generate_unique_filename(new_file_path)
            print(f"Renaming to new file path: {new_file_path}")
        if old_file_path is not None:
            os.rename(old_file_path, new_file_path)
            self.shared_state.downloaded_file = None

        return new_file_path

    def download_detail(self):
        # 这是个通用方法，子类可以根据实际情况重写
        raise NotImplementedError("This method must be implemented by subclasses.")

    def run(self):
        self.logger.info(f"{'+' * 10}程序开始运行{'+' * 10}")

        self.user_data = load_user_data(self.name_list, self.chn_city)
        if not self.user_data:
            self.logger.error('请先配置用户数据')
            raise Exception('请先配置用户数据')
        try:
            self.login(self.city)
            self.download_detail()
            self.logger.info(f"{'+' * 10}下载处理成功{'+' * 10}")
            messagebox.showinfo("提示", "程序处理成功!")
        except Exception as e:
            # 获取完整的异常堆栈信息
            with io.StringIO() as buf:
                traceback.print_exc(file=buf)
                stack_trace = buf.getvalue()
                
            # 确保两种输出方式都能看到完整信息
            error_msg = f"异常详细信息:\n{stack_trace}"
            self.logger.error(error_msg)
            
            # 添加提示框
            messagebox.showerror("错误", "程序处理失败!")
            
            # 重新抛出异常
            raise
        finally:
            self.quit()
