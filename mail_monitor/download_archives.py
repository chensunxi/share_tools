import sys
import os
# 获取当前文件所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))
# 将当前目录添加到 Python 路径
sys.path.append(current_dir)
# 将项目根目录添加到 Python 路径
project_root = os.path.dirname(current_dir)
sys.path.append(project_root)

import time
import traceback
import re
from datetime import datetime
from tkinter import messagebox

import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from utils.logger_utils import LoggerUtils
from utils.download_base import CrawlerDownloadBase
from utils.common_utils import find_pdf_files

class DownloadArchives(CrawlerDownloadBase):
    def __init__(self, erp_user, erp_pwd, download_dir, user_data, out_logger):
        super().__init__(None, download_dir, None, None, None)
        self.erp_name = erp_user
        self.erp_pwd = erp_pwd
        self.download_path = download_dir or os.path.join(os.getcwd(), 'source')
        self.user_data = user_data
        self.out_logger = out_logger
        self.logger = LoggerUtils.get_logger(None, out_logger)
        # 初始化一个空的列表来存储每个人的档案数据
        self.employees_data = []
        # 初始化下载结果字典，用于跟踪每个员工的下载结果
        self.download_results = {}
        for user in user_data:
            self.download_results[user['job_no']] = '000000000'  # 初始化为全部未下载

    def download_detail(self):
        self.driver.set_window_size(1552, 832)

        # 输入用户密码
        self.wait((By.XPATH, "//input[@name='username']"))
        user_name = self.driver.find_element(by=By.XPATH, value="//input[@name='username']")
        user_name.send_keys(self.erp_name)
        user_pwd = self.driver.find_element(by=By.XPATH, value="//input[@name='password']")
        user_pwd.send_keys(self.erp_pwd)
        try:
            # 等待最多1.5秒，检查特定元素是否可见
            element = WebDriverWait(self.driver, 1.5).until(
                EC.presence_of_element_located((By.NAME, "verifyCode"))  # 使用name属性定位元素
            )
            messagebox.showinfo("提示", "请在网页输入验证码,然后点击【登录】按钮.")
            time.sleep(5)  # 如果元素存在，等待5秒
        except:
            # 点击[登录]按钮
            # login_btn = self.driver.find_element(by=By.XPATH, value="//button[contains(@class, 'el-button') and contains(@class, 'login-button')]")
            login_btn = self.driver.find_element(By.CSS_SELECTOR, "button.el-button.login-button")
            # 等待元素可点击
            WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable(login_btn))
            login_btn.click() # 标准点击方法
            try:
                # 等待弹出框的出现，最长等待3秒
                error_message_element = WebDriverWait(self.driver, 3).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, ".el-message-box__message p"))
                )
                # 如果需要检查是否是用户名或密码错误
                if "用户名或密码错误" in error_message_element.text:
                    # 关闭弹出框
                    element = self.driver.find_element(by=By.XPATH, value="//button[span[contains(text(),'确定')]]")
                    self.driver.execute_script("arguments[0].click()", element)
                    self.logger.info(f"用户【{self.erp_name}】登录失败，等待输入验证码!")
                    messagebox.showinfo("提示", "请在网页输入验证码,然后点击【登录】按钮.")
                    time.sleep(5)  # 如果元素存在，等待5秒
            except:
                self.logger.info(f"用户【{self.erp_name}】登录成功!")

        # 等待登录成功后，点击"服务中心"标签页
        self.wait((By.XPATH, '//*[@id="app"]/div[1]/section/header/div[2]/div[1]/ul/li[2]/span'))
        element = self.driver.find_element(by=By.XPATH, value='//*[@id="app"]/div[1]/section/header/div[2]/div[1]/ul/li[2]/span')
        self.driver.execute_script("arguments[0].click()", element)

        # 点击"员工档案"菜单
        self.wait((By.XPATH, "//span[text()='员工档案']"))
        element = self.driver.find_element(by=By.XPATH, value="//span[text()='员工档案']")
        self.driver.execute_script("arguments[0].click()", element)

        # 点击"员工档案管理"按钮
        self.wait((By.XPATH, "//span[text()='员工档案管理']"))
        element = self.driver.find_element(by=By.XPATH, value="//span[text()='员工档案管理']")
        self.driver.execute_script("arguments[0].click()", element)
        # 展开所有输入条件
        element = self.driver.find_element(by=By.XPATH,
                                           value='//*[@id="container"]/div[1]/div/form/div/div/div[2]/div/div/div[3]/span')
        self.driver.execute_script("arguments[0].click()", element)

        idown = 0  # 下载计数器
        total_count = len(self.user_data)

        for input_context in self.user_data:
            idown += 1
            self.logger.info(f"开始下载{idown}/{total_count}：{input_context['job_no']} - {input_context['job_name']}")
            # 初始化当前员工的下载结果跟踪
            current_download_scope = list(input_context['download_scope'])
            download_success_scope = ['0'] * 9  # 初始化为全部未下载
            
            # 点击"档案类型"
            element = self.driver.find_element(by=By.XPATH,
                                               value='//*[@id="container"]/div[1]/div/form/div/div/div[1]/div/div[2]/div[2]/div/div/div/div[1]/input')
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(element))
            self.driver.execute_script("arguments[0].click()", element)

            # 选择"个人信息类"
            element = self.driver.find_elements(by=By.CLASS_NAME, value='el-cascader-node__label')[0]
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(element))
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)

            # 检查第1位 (索引为0)
            if input_context['download_scope'][0] == '1':
                # 勾选"身份证"
                element = self.driver.find_element(by=By.XPATH, value='//div[2]/div/ul/li/label/span/span')
                self.driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)

            # 检查第2位 (索引为1)
            if input_context['download_scope'][1] == '1':
                # 勾选"学历证书"
                element = self.driver.find_element(by=By.XPATH, value='//div[2]/div/ul/li[2]/label/span/span')
                self.driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)

            # 检查第3位 (索引为2)
            if input_context['download_scope'][2] == '1':
                # 勾选"学位证书"
                element = self.driver.find_element(by=By.XPATH, value='//div[2]/div/ul/li[3]/label/span/span')
                self.driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)

            # 检查第4位 (索引为3)
            if input_context['download_scope'][3] == '1':
                # 勾选"学历验证报告"
                element = self.driver.find_element(by=By.XPATH, value='//div[2]/div/ul/li[10]/label/span/span')
                self.driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)

            # 检查第5位 (索引为4)
            if input_context['download_scope'][4] == '1':
                # 选择"合同协议类"
                element = self.driver.find_element(by=By.XPATH, value='//div[4]/div/div/div/ul/li[2]/span')
                self.driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)
                # 勾选"劳动合同"
                element = self.driver.find_element(by=By.XPATH, value='//div[2]/div/ul/li/label/span/span')
                self.driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)

            # 检查第6位 (索引为5)
            if input_context['download_scope'][5] == '1':
                # 选择"资质证书类"
                element = self.driver.find_element(by=By.XPATH, value='//li[3]/label/span/span')
                self.driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)

            # 再次点击"档案类型"，关闭多级选框
            element = self.driver.find_element(by=By.XPATH,
                                               value='//*[@id="container"]/div[1]/div/form/div/div/div[1]/div/div[2]/div[2]/div/div/div/div[1]/input')
            self.driver.execute_script("arguments[0].click()", element)

            # "档案状态"选择"有效"
            element = self.driver.find_element(by=By.XPATH,
                                               value="//div[@id='container']/div/div/form/div/div/div/div/div[2]/div[3]/div/div/div/div[2]/div/div/ul/li[2]")
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)

            # 填入"人事编号"
            textarea = self.driver.find_element(by=By.XPATH, value='//*[@id="container"]/div[1]/div/form/div/div/div[1]/div/div[3]/div[3]/div/div/div/input')
            textarea.clear()
            textarea.send_keys(input_context['job_no'])

            # 点击【查询】按钮
            element = self.driver.find_element(by=By.XPATH, value='//*[@id="container"]/div[1]/div/form/div/div/div[2]/div/div/div[1]/button/span')
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(2)
            if idown == 1:
                # 调整分页条数为100条
                element = self.driver.find_element(By.CSS_SELECTOR, ".el-pagination__sizes .el-input__inner")
                self.driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)
                self.wait((By.XPATH, '//div[5]/div/div/ul/li[4]/span'))
                element = self.driver.find_element(By.XPATH, "//div[5]/div/div/ul/li[4]/span")
                self.driver.execute_script("arguments[0].click()", element)
            # 获取查询结果记录条数
            element = self.driver.find_element(by=By.XPATH, value='//*[@id="container"]/div[2]/div[3]/span[1]')
            record_cnt = re.search(r'\d+', element.text).group()
            print(f'记录数:{record_cnt}')

            # 弹出提示框等待用户的确认
            if int(record_cnt) > 100:
                response = messagebox.askyesno("提示",
                                               f"员工：<{input_context['job_no']}-{input_context['job_name']}>的档案数量超过100,请确认！\n\n点击【是】则继续下载，点击【否】跳过该员工。")
                if response == False:  # 点击【否】跳过当前项
                    self.logger.info("跳过当前员工的档案下载处理，请确认档案数据是否标准！")
                    continue

            # 创建文档类型到索引的映射
            doc_type_map = {
                '身份证': 0,
                '学历证书': 1,
                '学位证书': 2,
                '学历验证报告': 3,
                '劳动合同': 4,
                '保密协议': 5,
                '离职证明': 6,
                '离职协议': 7,
                '照片': 8
            }

            for i in range(int(record_cnt)):
                i += 1
                docx_type = self.driver.find_element(by=By.XPATH, value=f'//*[@id="container"]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[7]/div')
                doc_type_text = docx_type.text.strip()
                
                # 点击'下载'按钮
                download_btn = self.driver.find_element(by=By.XPATH, value=f'//*[@id="container"]/div[2]/div[2]/div[4]/div[2]/table/tbody/tr[{i}]/td[14]/div/button[2]/span')
                self.driver.execute_script("arguments[0].click()", download_btn)
                time.sleep(3)
                new_file_name = f'{input_context['job_no']}_{input_context['job_name']}_{doc_type_text}'
                new_file_path = os.path.join(self.download_path, new_file_name)
                self.logger.info(f"下载档案{i}：{doc_type_text}")
                
                # 重命名下载文件并检查是否成功
                if self.rename_downfile(new_file_path):
                    # 更新下载成功标记
                    for doc_type, index in doc_type_map.items():
                        if doc_type in doc_type_text:
                            download_success_scope[index] = '1'
                            break
            
            # 更新下载结果
            self.download_results[input_context['job_no']] = ''.join(download_success_scope)
            self.logger.info(f"员工 {input_context['job_no']} 下载结果: {''.join(download_success_scope)}")

            #重置输入
            reset_btn = self.driver.find_element(By.XPATH, "//span[text()='重置']")
            self.driver.execute_script("arguments[0].click()", reset_btn)
            time.sleep(2)

    def get_download_result(self, job_no):
        """获取指定员工的下载结果"""
        return self.download_results.get(job_no, '000000000')
    
    def is_download_complete(self, job_no):
        """检查指定员工的下载是否完整"""
        if job_no not in self.download_results:
            return False
        
        # 获取该员工的下载范围和下载结果
        for user in self.user_data:
            if user['job_no'] == job_no:
                required_scope = user['download_scope']
                actual_scope = self.download_results[job_no]
                
                # 检查每个需要下载的文档是否都已成功下载
                for i in range(9):
                    if required_scope[i] == '1' and actual_scope[i] != '1':
                        return False
                return True
        
        return False

    def run(self, env_switch):
        self.logger.info(f"{'+' * 10}程序开始运行{'+' * 10}")
        try:
            if env_switch == 0:
                self.login("erptest")
            else:
                self.login("sunerp")
            self.download_detail()
            self.logger.info(f"{'+' * 10}下载处理成功{'+' * 10}")
            return self.download_results

        except Exception as e:
            self.logger.error(f'程序运行处理出错：{str(e)}')
            traceback.print_exc()
            return {}
        finally:
            self.quit()


if __name__ == '__main__':
    try:
        download_dir = os.path.join(os.getcwd(), 'download')
        download_archives = DownloadArchives("admin", "sunline", "111111",download_dir,None)  # 返回对应的子类
        download_archives.run(1)
    except Exception as e:
        traceback.print_exc()
        print(f'下载处理出错：{str(e)}')