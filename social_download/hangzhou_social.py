# -*- coding: utf-8 -*-
# Author:   zhucong1
# At    :   2024/8/14
# Email :   zhucong1@sunline.com
# About :
import os
import time
import traceback
import logging
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils.download_base import CrawlerDownloadBase

class HangzhouSocialSecurity(CrawlerDownloadBase):
    def start_download(self, download_flag, download_user):
        # 点击"在线办理"按钮
        element = self.driver.find_elements(By.CLASS_NAME, 'title-foot-left')[1].find_element(By.TAG_NAME, "button")
        self.driver.execute_script("arguments[0].click()", element)
        # 3、新标签，切换。id = 'recept-footer-Notice'的div, 下的button, 点击
        WebDriverWait(self.driver, 10).until(EC.number_of_windows_to_be(3))
        time.sleep(2)
        all_windows = self.driver.window_handles
        self.driver.switch_to.window(all_windows[-1])

        # 点击"进入办事"按钮
        self.wait((By.XPATH, '//*[@id="recept-footer-Notice"]/button'))
        self.click_custom((By.XPATH, '//*[@id="recept-footer-Notice"]/button'), click_flag=False)

        # 4、id = 'recept-footer-SceneGuide' 下的第二个button,点击【请选择办理情况】的"确定"按钮
        self.wait((By.XPATH, '//*[@id="recept-footer-SceneGuide"]/button[2]'))
        self.click_custom((By.XPATH, '//*[@id="recept-footer-SceneGuide"]/button[2]'), click_flag=False)

        # 5、placeholder = '开始时间' / '结束时间' 的input value设置时间：2024-08
        self.wait((By.CSS_SELECTOR, 'input[placeholder="开始时间"]'))
        ele_start_date = self.driver.find_element(By.CSS_SELECTOR, 'input[placeholder="开始时间"]')
        # 等待元素可点击
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(ele_start_date))
        self.click_custom(ele_start_date, click_flag=True)

        # self.wait((By.CSS_SELECTOR, 'input[placeholder="结束时间"]'))
        # ele_end_date = self.driver.find_element(By.CSS_SELECTOR, 'input[placeholder="结束时间"]')
        # # 等待元素可点击
        # WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(ele_end_date))
        # self.click_custom(ele_end_date, click_flag=True)

        start_rows = self.driver.find_elements(By.XPATH, '//table[contains(@class, "next-calendar-table")]')[
            0].find_elements(By.XPATH, './tbody/tr')
        end_rows = self.driver.find_elements(By.XPATH, '//table[contains(@class, "next-calendar-table")]')[
            1].find_elements(By.XPATH, './tbody/tr')

        # 定义你需要的 <td> 元素索引
        if download_flag == '1':
            start_date = download_user['start_date']
            end_date = download_user['end_date']
        else:
            start_date = download_user[0]['start_date']
            end_date = download_user[0]['end_date']
        s_year, s_month = map(int, start_date.split('-'))
        e_year, e_month = map(int, end_date.split('-'))

        # 处理开始年份
        default_start_date = ele_start_date.get_attribute("value")
        # self.logger.info(f"开始默认日期：{default_start_date}")
        if default_start_date and default_start_date.strip():
            current_year = int(default_start_date[:4])
        else:
            current_year = datetime.now().year
        if s_year < current_year:
            # 左跳差额年份
            element = self.driver.find_elements(By.CSS_SELECTOR, "button[title='上一年']")[0]
            for i in range(current_year - s_year):
                time.sleep(0.5)
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(element))
                self.driver.execute_script("arguments[0].click()", element)
        if s_year > current_year:
            # 右跳差额年份
            element = self.driver.find_elements(By.CSS_SELECTOR, "button[title='下一年']")[0]
            for i in range(s_year - current_year):
                time.sleep(0.5)
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(element))
                self.driver.execute_script("arguments[0].click()", element)
        # 处理开始月份
        time.sleep(1)

        tr_index = (s_month - 1) // 3
        td_index = (s_month - 1) % 3

        target_td = start_rows[tr_index].find_elements(By.TAG_NAME, 'td')[td_index]
        self.driver.execute_script("arguments[0].click()", target_td)

        # 处理结束年份
        # 获取日期控件的默认日期
        ele_end_date = self.driver.find_element(By.CSS_SELECTOR, 'input[placeholder="结束时间"]')
        default_end_date = ele_end_date.get_attribute("value")
        if default_end_date and default_end_date.strip():
            current_year = int(default_end_date[:4])
        else:
            current_year = datetime.now().year
        if e_year < current_year:
            # 左跳差额年份
            element = self.driver.find_elements(By.CSS_SELECTOR, "button[title='上一年']")[1]
            for i in range(current_year - e_year):
                time.sleep(0.5)
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(element))
                self.click_custom(element, click_flag=False)
        if e_year > current_year:
            # 右跳差额年份
            element = self.driver.find_elements(By.CSS_SELECTOR, "button[title='下一年']")[1]
            for i in range(e_year - current_year):
                time.sleep(0.5)
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(element))
                self.click_custom(element, click_flag=False)
        # 处理结束月份
        tr_index = (e_month - 1) // 3
        td_index = (e_month - 1) % 3

        target_td = end_rows[tr_index].find_elements(By.TAG_NAME, 'td')[td_index]
        self.click_custom(target_td, click_flag=False)

        # 确定时间
        element = \
        self.driver.find_element(By.CLASS_NAME, 'next-date-picker-panel-footer').find_elements(By.TAG_NAME,
                                                                                               'button')[2]
        self.click_custom(element, click_flag=False)

        # 6、class = 'vc-container'的div下的第9个div,下的第3个div,下的span,点击下拉框"需证明的人员选择"
        element = self.driver.find_element(By.XPATH, '//*[@class="vc-container"]/div[9]/div/div[2]/span')
        self.click_custom(element, click_flag=False)

        # 7、class='options-item'的第一个div,点击
        element = self.driver.find_elements(By.CLASS_NAME, 'options-item')[0]
        self.click_custom(element, click_flag=False)
        # 8、textarea，赋值查询人员身份证
        self.wait((By.CLASS_NAME, 'next-input-textarea'))
        textarea = \
        self.driver.find_element(By.CLASS_NAME, 'next-input-textarea').find_elements(By.TAG_NAME, 'textarea')[0]
        # 等待元素可交互
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(textarea))
        if download_flag == '1':
            textarea.clear()
            input_user = download_user['id_card']
        else:
            input_user = ','.join([i['id_card'] for i in download_user])
        textarea.send_keys(input_user)
        # 9、点击"确认注意事项"
        element = self.driver.find_element(By.XPATH,
                                           '//*[@id="render-engine-page-container"]/div/div[2]/div[2]/div[2]/div/div/span/label/span[1]/input')
        self.click_custom(element, click_flag=False)
        # 10、id = 'recept-footer-Form'的div下的第3个button,点击"查询"按钮
        button = self.driver.find_element(By.XPATH, '//*[@id="recept-footer-Form"]/button[3]')
        self.click_custom(button, click_flag=False)
        # 11、点击，完成下载
        try:
            self.wait((By.XPATH, '//span[text()="下载文件"]'), 10)
        except Exception:
            self.driver.refresh()
            time.sleep(3)
        self.click_custom((By.XPATH, '//span[text()="下载文件"]'))
        time.sleep(5)
        if download_flag == '1':
            new_file_name = f"{download_user['staff_name']}_{download_user['id_card']}_{self.city}_个人参保证明"
        else:    
            new_file_name = '杭州分公司_单位汇总参保证明'
        new_file_path = os.path.join(self.download_path, new_file_name)
        self.rename_downfile(new_file_path)
        # 关闭当前标签页
        self.driver.close()
        time.sleep(1)
        # 获取当前所有标签页的句柄
        all_windows = self.driver.window_handles
        # 切换回上一个标签页
        self.driver.switch_to.window(all_windows[1])
        # 可选：等待一段时间，查看返回效果
        time.sleep(2)

    def download_detail(self):
        self.driver.set_window_size(1552, 832)
        # 0、切换到"扫码登录"标签页
        time.sleep(3)
        self.click_custom((By.XPATH, '//*[@id="zlbQrLoginTab"]'))

        # 1、title = '单位参保证明查询打印'的p标签 点击
        self.wait((By.CSS_SELECTOR, 'p[title="单位参保证明查询打印"]'))
        self.click_custom((By.CSS_SELECTOR, 'p[title="单位参保证明查询打印"]'))

        # 2、新标签，切换。class = 'title-foot-left'的第一个div, 下的button, 点击
        WebDriverWait(self.driver, 10).until(EC.number_of_windows_to_be(2))
        time.sleep(2)
        all_windows = self.driver.window_handles
        self.driver.switch_to.window(all_windows[-1])

        total_count = len(self.user_data)
        self.logger.info(f"总共需要下载 {total_count} 条数据")
        
        if self.download_flag == '1':
            # 个人下载模式
            idown = 0  # 下载计数器
            for input_context in self.user_data:
                idown += 1
                self.logger.info(f"开始下载{idown}/{total_count}：{input_context['staff_name']}")
                self.start_download('1', input_context)
        else:
            # 单位合并下载模式
            batch_size = int(self.download_limit)  # 每批次处理的最大数量
            total_batches = (total_count + batch_size - 1) // batch_size  # 向上取整得到总批次
            
            for batch_num in range(total_batches):
                start_idx = batch_num * batch_size
                end_idx = min((batch_num + 1) * batch_size, total_count)
                current_batch = self.user_data[start_idx:end_idx]
                
                self.logger.info(f"开始下载第 {batch_num + 1}/{total_batches} 批次 \n"
                               f"本批次含 {len(current_batch)} 条数据 "
                               f"({start_idx + 1}~{end_idx})")
                
                try:
                    self.start_download('2', current_batch)
                    # self.logger.info(f"第 {batch_num + 1} 批次下载完成")
                except Exception as e:
                    self.logger.error(f"第 {batch_num + 1} 批次下载失败: {str(e)}")
                    raise
                
                # 如果不是最后一批，等待一段时间再继续
                if batch_num < total_batches - 1:
                    wait_time = 2  # 每批次之间等待2秒
                    self.logger.info(f"等待 {wait_time} 秒后继续下一批次...")
                    time.sleep(wait_time)


