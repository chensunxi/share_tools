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

from bidding_docx.docx_generator import EmployeeDocumentGenerator
from utils.logger_utils import LoggerUtils
from utils.download_base import CrawlerDownloadBase
from utils.common_utils import find_pdf_files

class BiddingDocument(CrawlerDownloadBase):
    def __init__(self, erp_user, erp_pwd, chkbox_values, download_dir, pdf_dir, excel_file, tpl_no, out_logger):
        super().__init__(None, download_dir, None, None, None)
        self.erp_name = erp_user
        self.erp_pwd = erp_pwd
        self.docx_scope = chkbox_values
        self.download_path = download_dir or os.path.join(os.getcwd(), 'source')
        self.social_dir = pdf_dir
        self.name_list = excel_file
        self.out_logger = out_logger
        self.logger = LoggerUtils.get_logger(None, out_logger)
        self.tpl_no = tpl_no
        # 初始化一个空的列表来存储每个人的档案数据
        self.employees_data = []

    def load_job_data(self, user_template):
        if not os.path.exists(user_template):
            return

        def row_to_dict(row):
            return {
                "job_no": row.get("工号", ""),
                "job_name": row.get("工作名", ""),
                "identity_card": row.get("证件号码", "")
            }

        data = pd.read_excel(user_template, dtype={'工号': str, '证件号码': str}, skiprows=1).fillna('')
        data = data.apply(row_to_dict, axis=1).tolist()

        return data

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

        # 点击"档案类型"
        element = self.driver.find_element(by=By.XPATH,
                                           value='//*[@id="container"]/div[1]/div/form/div/div/div[1]/div/div[2]/div[2]/div/div/div/div[1]/input')
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(element))
        self.driver.execute_script("arguments[0].click()", element)

        # 选择"个人信息类"
        element = self.driver.find_elements(by=By.CLASS_NAME, value = 'el-cascader-node__label')[0]
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(element))
        self.driver.execute_script("arguments[0].click()", element)
        time.sleep(0.5)

        # 检查第1位 (索引为0)
        if self.docx_scope[0] == '1':
            # 勾选"身份证"
            element = self.driver.find_element(by=By.XPATH, value = '//div[2]/div/ul/li/label/span/span')
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)

        # 检查第2位 (索引为1)
        if self.docx_scope[1] == '1':
            # 勾选"学历证书"
            element = self.driver.find_element(by=By.XPATH, value = '//div[2]/div/ul/li[2]/label/span/span')
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)

        # 检查第3位 (索引为2)
        if self.docx_scope[2] == '1':
            # 勾选"学位证书"
            element = self.driver.find_element(by=By.XPATH, value='//div[2]/div/ul/li[3]/label/span/span')
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)

        # 检查第4位 (索引为3)
        if self.docx_scope[3] == '1':
            # 勾选"学历验证报告"
            element = self.driver.find_element(by=By.XPATH, value='//div[2]/div/ul/li[10]/label/span/span')
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)

        # 勾选"学位验证报告"
        # element = self.driver.find_element(by=By.XPATH, value='//div[2]/div/ul/li[14]/label/span/span')
        # self.driver.execute_script("arguments[0].click()", element)

        # 检查第5位 (索引为4)
        if self.docx_scope[4] == '1':
            # 选择"合同协议类"
            element = self.driver.find_element(by=By.XPATH, value='//div[4]/div/div/div/ul/li[2]/span')
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)
            # 勾选"劳动合同"
            element = self.driver.find_element(by=By.XPATH, value='//div[2]/div/ul/li/label/span/span')
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)

        # 检查第6位 (索引为5)
        if self.docx_scope[5] == '1':
            # 选择"资质证书类"
            element = self.driver.find_element(by=By.XPATH, value='//li[3]/label/span/span')
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)

        # 再次点击"档案类型"，关闭多级选框
        element = self.driver.find_element(by=By.XPATH,
                                           value='//*[@id="container"]/div[1]/div/form/div/div/div[1]/div/div[2]/div[2]/div/div/div/div[1]/input')
        self.driver.execute_script("arguments[0].click()", element)

        # "档案状态"选择"有效"
        element = self.driver.find_element(by=By.XPATH, value="//div[@id='container']/div/div/form/div/div/div/div/div[2]/div[3]/div/div/div/div[2]/div/div/ul/li[2]")
        self.driver.execute_script("arguments[0].click()", element)
        time.sleep(0.5)
        # 展开所有输入条件
        element = self.driver.find_element(by=By.XPATH, value='//*[@id="container"]/div[1]/div/form/div/div/div[2]/div/div/div[3]/span')
        self.driver.execute_script("arguments[0].click()", element)
        idown = 0  # 下载计数器
        total_count = len(self.user_data)

        # 预设档案类型及对应的字段名
        docx_types = [
            ('身份证', 'id_card_front'),
            ('学历证书', 'graduation_cert'),
            ('学位证书', 'degree_cert'),
            ('学历验证报告', 'china_edu_screenshot'),
            ('劳动合同', 'labor_contract')
        ]
        for input_context in self.user_data:
            idown += 1
            self.logger.info(f"开始下载{idown}/{total_count}：{input_context['job_no']} - {input_context['job_name']}")
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

            # 查找本地目录的社保证明路径
            find_social_security = find_pdf_files(self.social_dir, input_context['job_name'])
            if find_social_security:
                social_security_path = os.path.normpath(find_social_security)
            else:
                social_security_path = None
            # 初始化员工数据，预设所有档案类型为 None
            employee_data = {
                'name': input_context['job_name'],
                'id_no': input_context['identity_card'],
                'id_card_front': None,
                'id_card_back': None,
                'graduation_cert': None,
                'degree_cert': None,
                'china_edu_screenshot': None,
                'labor_contract': None,
                'qualification_certs': {},  # 初始化为空字典，用于存储资质证书
                'social_security_proof': social_security_path  # 初始化社保证明路径
            }

            for i in range(int(record_cnt)):

                i += 1
                docx_type = self.driver.find_element(by=By.XPATH, value=f'//*[@id="container"]/div[2]/div[2]/div[3]/table/tbody/tr[{i}]/td[7]/div')
                # 点击'下载'按钮
                download_btn = self.driver.find_element(by=By.XPATH, value=f'//*[@id="container"]/div[2]/div[2]/div[4]/div[2]/table/tbody/tr[{i}]/td[14]/div/button[2]/span')
                self.driver.execute_script("arguments[0].click()", download_btn)
                time.sleep(3)
                new_file_name = f'{input_context['job_no']}_{input_context['job_name']}_{docx_type.text.strip()}'
                new_file_path = os.path.join(self.download_path, new_file_name)
                self.logger.info(f"下载档案{i}：{docx_type.text}")
                down_file_path = self.rename_downfile(new_file_path)

                # 遍历档案类型并填充实际数据
                for docx_type_name, key in docx_types:
                    # 假设 docx_type.text 是档案类型名称，new_file_path 是文件路径
                    if docx_type.text == docx_type_name:
                        employee_data[key] = down_file_path  # 如果找到该档案类型，则更新路径
                        break

                # 处理资质证书（注意可能有多个）
                if docx_type.text not in [x[0] for x in docx_types]:
                    # 将资质证书路径根据档案类型加入到 qualification_certs
                    employee_data['qualification_certs'][docx_type.text] = down_file_path

                # # 如果该员工有任何资质证书，确保 'qualification_certs' 字段非空
                # if not employee_data['qualification_certs']:
                #     employee_data['qualification_certs'] = None

            # 将当前员工的档案信息添加到 employees_data 列表中
            self.employees_data.append(employee_data)

        # 打印最终的结果，检查结构
        self.logger.info(f"员工数据：{self.employees_data}")

    def run(self, env_switch):
        self.logger.info(f"{'+' * 10}程序开始运行{'+' * 10}")
        self.user_data = self.load_job_data(self.name_list)
        if not self.user_data:
            self.logger.error('请先配置用户数据')
            raise Exception('请先配置用户数据')
        try:
            if env_switch == 0:
                self.login("erptest")
            else:
                self.login("sunerp")
            self.download_detail()
            self.logger.info(f"{'+' * 10}下载处理成功{'+' * 10}")

            self.logger.info(f"{'+' * 10}开始合成WORD{'+' * 10}")
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            template_path = os.path.join(os.getcwd(), 'template', f'bidding_template{self.tpl_no}.docx')
            output_path = os.path.join(os.getcwd(), 'output', f'bidding_documents_{current_time}.docx')

            generator = EmployeeDocumentGenerator(template_path, self.out_logger)
            generator.generate_document(self.employees_data, output_path)
            self.logger.info(f"{'+' * 10}合成WORD成功{'+' * 10}")
            # 打开文件
            os.startfile(output_path)

        except Exception as e:
            self.logger.error(f'程序运行处理出错：{str(e)}')
            traceback.print_exc()
        finally:
            self.quit()


if __name__ == '__main__':
    try:
        # 根据需要选择城市
        download_dir = os.path.join(os.getcwd(), 'download')
        pdf_dir = r'D:\project\share_tools\bidding_docx\test'
        excel_file = r'D:\project\share_tools\bidding_docx\test\投标档案名单.xlsx'
        bidding_download = BiddingDocument("admin", "sunline", "111111",download_dir,pdf_dir,excel_file,1,None)  # 返回对应的子类
        bidding_download.run()
    except Exception as e:
        traceback.print_exc()
        print(f'下载处理出错：{str(e)}')