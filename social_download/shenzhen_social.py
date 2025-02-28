# -*- coding: utf-8 -*-
# Author:   zhucong1
# At    :   2024/8/14
# Email :   zhucong1@sunline.com
# About :
import argparse
import json
import os
import re
import shutil
import time
import traceback

import fitz
import pandas as pd
from pymupdf import pymupdf
from selenium.common import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils.download_base import CrawlerDownloadBase

class Base(object):
    def __init__(self, source, target, *args, **kwargs):
        self._source = source
        self._target = target

    def get_rect(self):
        raise NotImplementedError()

    def process(self):
        for pdf in os.listdir(self._source):
            if pdf.endswith('.pdf'):
                doc = pymupdf.Document(os.path.join(self._source, pdf))
                self.cover_by_doc(doc)
                doc.save(f'{self._target}/{pdf}')

    def cover_by_doc(self, doc):
        for page in doc:
            self.cover_by_page(page)

    def cover_by_page(self, page):
        row_count = page.find_tables().tables[0].row_count
        end_y = self.get_y_end(row_count)
        for o in self.get_rect():
            o = list(o)
            o[-1] = end_y
            page.draw_rect(o, color=(1, 1, 1), fill=(1, 1, 1), width=0)
        self.add_last(page, end_y)

    def get_y_end(self, row_count):
        raise NotImplementedError()

    def add_last(self, page, end_y):
        pass


class PersonalDetail(Base):
    def get_rect(self):
        return (
            # 起点x：89，单独计算并固定写死,起点y：96，固定写死每列起点为96，终点x：单独计算并固定写死，终点y：按行计算，以y起始为基础，每增加一行增加12
            # 养老保险
            (88, 96, 104, 144),  # 基数
            (129, 96, 140, 144),  # 单位
            (169, 96, 180, 144),  # 个
            # # 医疗保险
            (237, 96, 250, 144),  # 基数
            (264, 96, 275, 144),  # 单位
            (304, 96, 316, 144),  # 个人
            # # 生育保险
            (373, 96, 383, 144),  # 基数
            (400, 96, 411, 144),  # 单位
            # # 工伤保险
            (433, 96, 443, 144),  # 基数
            (460, 96, 470, 144),  # 单位
            # # 失业保险
            (493, 96, 505, 144),  # 基数
            (520, 96, 530, 144),  # 单位
            (555, 96, 563, 144),  # 个人
        )

    def get_y_end(self, row_count):
        return 96 + 12 * (row_count - 2)

    def add_last(self, page, end_y):
        page.draw_rect((129, end_y + 8, 580, end_y + 16), color=(1, 1, 1), fill=(1, 1, 1), width=0)


class CompanyCertify(Base):

    def get_rect(self):
        return (
            (200, 204, 220, 755),  # 养老保险
            (258, 204, 278, 755),  # 医疗保险
            (328, 204, 348, 755),  # 生育保险/生育医疗
            (420, 204, 440, 755),  # 工伤保险
            (473, 204, 493, 755),  # 失业保险
        )

    def get_y_end(self, row_count):
        return 205 + 18 * (row_count - 1)


class InsuranceDetail(Base):
    def get_rect(self):
        return (
            # 养老保险
            (148, 163, 165, 375),  # 基数
            (189, 163, 203, 375),  # 个人
            (230, 163, 244, 375),  # 单位
            # 医疗保险
            (270, 163, 282, 375),  # 基数
            (311, 163, 326, 375),  # 个人
            (350, 163, 366, 375),  # 单位
            # # 生育保险
            (391, 163, 404, 375),  # 基数
            (433, 163, 444, 375),  # 单位
            # # 工伤保险
            (474, 163, 492, 375),  # 基数
            (514, 163, 523, 375),  # 单位
            # 失业保险
            (555, 163, 572, 375),  # 基数
            (595, 163, 602, 375),  # 个人
            (636, 163, 646, 375),  # 单位
            # 个人小计
            (677, 163, 695, 375),
            # 单位小计
            (717, 163, 736, 375),
            # 合计
            (758, 163, 776, 375),
        )

    def get_y_end(self, row_count):
        return 163 + 12 * (row_count - 3)

    def add_last(self, page, end_y):
        y_top = end_y + 1
        y_bottom = end_y + 10
        rect_x_range = ((190, 204), (230, 248),
                        (312, 327), (350, 365),
                        (433, 449), (515, 525),
                        (597, 608), (638, 648),
                        (678, 695), (718, 740),
                        (757, 774))
        for x_left, x_right in rect_x_range:
            page.draw_rect((x_left, y_top, x_right, y_bottom), color=(1, 1, 1), fill=(1, 1, 1), width=0)

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            if idx == len(doc) - 1:
                self.cover_latest_page(page)
            else:
                self.cover_by_page(page)

    def cover_latest_page(self, page):
        y_top = 97
        y_bottom = 110
        rect_x_range = ((63, 89), (179, 203),
                        (287, 305), (394, 410),
                        (498, 508), (608, 620),
                        (715, 732), (820, 835),
                        (892, 914),)
        for x_left, x_right in rect_x_range:
            page.draw_rect((x_left, y_top, x_right, y_bottom), color=(1, 1, 1), fill=(1, 1, 1), width=0)

class ShenzhenSocialSecurity(CrawlerDownloadBase):
    def switch_page(self):
        self.logger.info('切换到查询服务')
        time.sleep(10)
        
        # 处理所有弹窗
        max_retries = 3
        for _ in range(max_retries):
            try:
                # 等待弹窗出现
                self.wait((By.CSS_SELECTOR, '.dw-textButton-level1-ext > a'))
                # 获取所有关闭按钮
                close_buttons = self.driver.find_elements(
                    By.CSS_SELECTOR, 
                    '.dw-textButton-level1-ext > a'
                )
                
                # 从后往前关闭弹窗
                for ele in close_buttons[::-1]:
                    try:
                        # 确保元素可见和可交互
                        WebDriverWait(self.driver, 10).until(
                            EC.element_to_be_clickable(ele)
                        )
                        btn_title = ele.get_attribute("title")
                        print(f"按钮名字：{btn_title}")
                        if btn_title != '确认无误':
                            # 使用 JavaScript 点击
                            self.driver.execute_script("arguments[0].click();", ele)
                        time.sleep(2)  # 等待弹窗关闭动画
                    except StaleElementReferenceException:
                        # 如果元素已失效，重新获取元素
                        self.logger.warning("元素已失效，重新获取...")
                        continue
                    except Exception as e:
                        self.logger.warning(f"关闭弹窗时出错: {str(e)}")
                        continue
                
                break  # 如果成功处理所有弹窗，跳出重试循环
                
            except Exception as e:
                self.logger.warning(f"处理弹窗时出错，重试中: {str(e)}")
                time.sleep(2)
        
        # 切换到查询服务
        self.wait((By.XPATH, '//*[@id="wsfw_nav"]/ul/li[2]'))
        self.driver.find_element(by=By.XPATH, value='//*[@id="wsfw_nav"]/ul/li[2]').click()
        # 点击菜单[单位信息查询]
        self.wait((By.XPATH, '//*[@id="wsfw_cont_left_xxcx"]/dl/dt[1]/p'))
        self.driver.find_element(by=By.XPATH, value='//*[@id="wsfw_cont_left_xxcx"]/dl/dt[1]/p').click()

    def wait_download(self, file_type, user_name='', code=''):
        # self.logger.info(f'等待文件下载：{file_name}')
        # 等待下载
        time.sleep(3)
        if file_type == 'company_certify':
            file_name = '深圳社保_单位参保证明'
        elif file_type == 'insurance_detail':
            file_name = '深圳社保_缴交明细表'
        else:
            file_name = 'unknown'
        new_path = os.path.join(self.download_path, file_type)
        if not os.path.exists(new_path):
            os.makedirs(new_path)
        if user_name and code:
            file_name = f'{user_name}_{code}_{self.city}_个人参保证明'
        new_file_path = os.path.join(new_path, file_name)
        self.rename_downfile(new_file_path)

    def download_company_certify(self, start_date, end_date):
        self.logger.info(f'开始下载【单位参保证明】，起始年月：{start_date},终止年月：{end_date}')
        # self.switch_page()
        # 点击菜单[单位参保证明查询打印]
        self.driver.find_element(by=By.XPATH, value='//*[@id="wsfw_cont_left_xxcx"]/dl/dd[1]/ul/li[7]').click()
        # 输入起始年月和终止年月
        self.wait((By.CLASS_NAME, 'dw-fieldSet'))
        self.wait((By.NAME, 'qsny'))
        self.driver.find_element(by=By.NAME, value='qsny').send_keys(start_date)  # 起始年月
        self.wait((By.NAME, 'zzny'))
        js = f"""document.querySelector('input[name="zzny"]').value='{end_date}'"""
        self.driver.execute_script(js)
        # 点击下拉框[预览打印类别]
        self.driver.find_element(by=By.CSS_SELECTOR,
                                 value='.dw-sform-extdropdwon-widget.dw_sform_widget_border_ext').click()
        # 选中下拉项[五险]
        self.wait((By.CSS_SELECTOR, '.dw-sform-extdropdwon-showbox-items-item-text'))
        time.sleep(1)
        self.driver.find_elements(by=By.CSS_SELECTOR, value='.dw-sform-extdropdwon-showbox-items-item-text')[
            1].click()

        self.driver.find_element(by=By.NAME, value='btn_doshow').click()  # 预览
        self.driver.find_element(by=By.NAME, value='btnPrint').click()  # 下载

        self.wait_download('company_certify')

    def download_insurance_detail(self, jfny):
        self.logger.info(f'开始下载【社会保险缴交明细表】，缴费年月：{jfny}')
        # self.switch_page()
        # 点击菜单[打印深圳市参保单位职工社会保险月缴交明细表]
        self.driver.find_element(by=By.XPATH, value='//*[@id="wsfw_cont_left_xxcx"]/dl/dd[1]/ul/li[8]').click()
        self.driver.find_element(by=By.NAME, value='jfny').send_keys(jfny)  # 缴费年月
        # 点击下拉框[预览打印类别]
        self.driver.find_element(by=By.CSS_SELECTOR,
                                 value='.dw-sform-extdropdwon-widget.dw_sform_widget_border_ext').click()
        # 选中下拉项[五险]
        self.wait((By.CSS_SELECTOR, '.dw-sform-extdropdwon-showbox-items-item-text'))
        time.sleep(1)
        self.driver.find_elements(by=By.CSS_SELECTOR, value='div.dw-sform-extdropdwon-showbox-items-item-text[title="五险"]')[
            1].click()
        time.sleep(2)
        self.driver.find_element(by=By.NAME, value='btnPrint').click()  # 下载
        self.wait_download('insurance_detail')

    def download_personal_detail(self, user_data):
        self.logger.info(f'开始下载【个人参保明细表】')
        # self.switch_page()
        # 收缩菜单[单位信息查询]
        self.wait((By.XPATH, '//*[@id="wsfw_cont_left_xxcx"]/dl/dt[1]/p'))
        self.driver.find_element(by=By.XPATH, value='//*[@id="wsfw_cont_left_xxcx"]/dl/dt[1]/p').click()
        # 切换至员工信息查询
        self.click_custom((By.XPATH, '//*[@id="wsfw_cont_left_xxcx"]/dl/dt[2]/p'))
        # 点击菜单[查询、打印企业员工在本企业参保情况]
        self.click_custom((By.XPATH, '//*[@id="wsfw_cont_left_xxcx"]/dl/dd[2]/ul/li[3]'))

        self.wait((By.NAME, 'btnPrint'))
        for input_context in user_data:  # yxzjhm, xm, dnh
            # self.logger.info(f'开始下载个人参保明细表，用户信息：{input_context}')
            self.driver.find_element(by=By.NAME, value='yxzjhm').clear()
            self.driver.find_element(by=By.NAME, value='yxzjhm').send_keys(input_context.get('id_card', ''))  # 有效证件号码
            # self.driver.find_element(by=By.NAME, value='xm').clear()
            # self.driver.find_element(by=By.NAME, value='xm').send_keys(input_context.get('staff_name', ''))  # 姓名
            self.driver.find_element(by=By.NAME, value='dnh').clear()
            self.driver.find_element(by=By.NAME, value='dnh').send_keys(input_context.get('id_social', ''))  # 电脑号

            js = f"""document.querySelector('input[name="qsny"]').value='{input_context.get('start_date', '')}'"""
            self.driver.execute_script(js)

            js = f"""document.querySelector('input[name="zzny"]').value='{input_context.get('end_date', '')}'"""
            self.driver.execute_script(js)

            self.driver.find_element(by=By.NAME, value='btnPrint').click()  # 下载
            self.wait_download('personal_detail',
                               user_name=input_context.get('staff_name', ''),
                               code=input_context.get('id_card', '') or input_context.get('id_social', ''))

    def process(self):
        # 加载 Excel 文件
        df = pd.read_excel(self.name_list, sheet_name=self.chn_city, header=1)

        # 获取需要的列数据并确保是字符串类型
        start_date = str(df['单位参保起始日期'][0]) if pd.notna(df['单位参保起始日期'][0]) else ''
        end_date = str(df['单位参保终止日期'][0]) if pd.notna(df['单位参保终止日期'][0]) else ''
        pay_date = str(df['缴交日期'][0]) if pd.notna(df['缴交日期'][0]) else ''
        corp_num = str(df['单位编号'][0]) if pd.notna(df['单位编号'][0]) else ''

        query_data = {
            'company_certify': {
                'start_date': start_date,
                'end_date': end_date
            },
            'insurance_detail': {
                'jfny': pay_date
            },
            'personal_detail': {
                'users': []
            },
            'corp_num': corp_num
        }

        if self.user_data:
            query_data['personal_detail']['users'] = self.user_data

        msg = []
        # 只在日期不为空时验证格式
        if query_data.get('company_certify'):
            dates = query_data['company_certify'].values()
            non_empty_dates = [d for d in dates if d.strip()]  # 过滤掉空日期
            if non_empty_dates and not all(re.match(r'^\d{4}-\d{2}$', d) for d in non_empty_dates):
                msg.append('日期格式错误')

        if query_data.get('insurance_detail'):
            dates = query_data['insurance_detail'].values()
            non_empty_dates = [d for d in dates if d.strip()]  # 过滤掉空日期
            if non_empty_dates and not all(re.match(r'^\d{4}-\d{2}$', d) for d in non_empty_dates):
                msg.append('日期格式错误')

        if query_data.get('personal_detail') \
                and not (query_data['personal_detail'].get('users', [])
                         and all([any(i.values()) for i in query_data['personal_detail']['users']])):
            msg.append('有效证件号码,姓名,电脑号不能同时为空')

        if msg:
            self.logger.error(';'.join(msg))
            raise Exception("输入数据[证件、姓名、电脑号、日期]校验不通过")

        # 只在日期不为空时执行相应的下载
        if query_data.get('company_certify') and start_date and end_date:
            self.download_company_certify(**query_data['company_certify'])
        if query_data.get('insurance_detail') and pay_date:
            self.download_insurance_detail(**query_data['insurance_detail'])
        if query_data.get('personal_detail'):
            self.download_personal_detail(query_data['personal_detail']['users'])


    def download_detail(self):
        self.driver.set_window_size(1552, 832)
        # 关闭提示框
        self.driver.find_element(By.XPATH, '//*[@id="popBox"]/div[2]/a').click()
        
        # 输入单位编号
        element = self.driver.find_element(by=By.XPATH, value='//*[@id="userNameInput"]')
        element.send_keys('60022152')
        
        # 勾选'同意'选择框
        self.driver.find_element(by=By.XPATH, value='//*[@id="ysxy"]').click()
        
        # 等待用户手动登录
        time.sleep(10)

        # 检查是否有登录失败的提示框
        try:
            alert = self.driver.switch_to.alert
            alert_text = alert.text
            self.logger.error(f"登录失败: {alert_text}")
            alert.accept()  # 关闭提示框
        except:
            pass
        
        # 检查登录状态
        self.wait((By.CLASS_NAME, 'dw-mainFrame'), 30)
        self.logger.info("登录成功，开始处理数据...")

        # 切换到查询服务
        self.switch_page()
        
        # 开始下载处理
        self.process()
        # 开始脱敏处理
        target_dir = os.path.join(self.download_path, 'cover')
        self.cover(self.download_path, target_dir)

    def cover(self, source_dir, target_dir):
        self.logger.info(f'{"=" * 10}开始进行PDF数据覆盖{"=" * 10}')
        for i in os.listdir(source_dir):
            source = os.path.join(source_dir, i)
            target = os.path.join(target_dir, i)
            if not os.path.exists(target):
                os.makedirs(target)
            category_map = {
                # 'company_certify': CompanyCertify,
                'insurance_detail': InsuranceDetail,
                'personal_detail': PersonalDetail,
            }
            if not category_map.get(i):
                continue
            processor = category_map[i](source, target)
            processor.process()
        self.logger.info(f'{"=" * 10}PDF数据覆盖完成{"=" * 10}')


    if __name__ == '__main__':
        description = """脚本使用说明"""
        parser = argparse.ArgumentParser(description=description)
        parser.add_argument("--username", type=str, default='30247009', help='社保网站用户名')
        parser.add_argument("--password", type=str, default='Sunline009', help='社保网站密码')
        parser.add_argument("--data", type=str, default='data.json', help='''查询数据，可以是文件，也可以是字符串，json格式，示例：{
        "company_certify": {
        "qsny": "2021-01",
        "zzny": "2021-08"
        },
        "insurance_detail": {
        "jfny": "2021-08"
        },
        "personal_detail": {
        "users": [
          {
            "yxzjhm": "",
            "xm": "韦海乐",
            "dnh": "646798105",
            "qsny": "2021-01",
            "zzny": "2021-08"
          }
        ],
        "qsny": "2021-01",
        "zzny": "2021-08"
        }
        }''')
        args = parser.parse_args()

        # self.logger.info(f"{'+' * 10}开始处理社保数据{'+' * 10}")
        if os.path.exists(args.data) or os.path.exists(os.path.join(os.getcwd(), args.data)):
            with open(args.data, encoding='utf8') as f:
                json_data = f.read()
        else:
            json_data = args.data

        # self.logger.info('输入参数：')
        # self.logger.info(f"{' ' * 8}用户名: {args.username}")
        # self.logger.info(f"{' ' * 8}密码:{args.password}")
        try:
            query_data = json.loads(json_data)

            # user_data = self.load_user_data(self.name_list, self.chn_city)
            # if user_data:
            #     query_data['personal_detail']['users'] = user_data

            # self.logger.info(f"{' ' * 8}数据查询参数:{json.dumps(query_data)}")
        except Exception as e:
            # self.logger.error(f"数据格式错误: {str(e)}")
            traceback.print_exc()
            exit(1)

        input_data = {
            'username': args.username,
            'password': args.password,
            **query_data
        }
        source_dir = os.path.join(os.getcwd(), 'source')
        target_dir = os.path.join(os.getcwd(), 'target')
        try:
            # crawl(input_data, source_dir)
            cover(source_dir, target_dir)
        except Exception as e:
            traceback.print_exc()
            # self.logger.error(f"数据处理失败: {str(e)}")
        # self.logger.info(f"{'+' * 10}社保数据处理完成{'+' * 10}")