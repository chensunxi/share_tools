# -*- coding: utf-8 -*-
# Author:   chensx
# At    :   2025/02/10
# Email :   chensx@sunline.com
# About :   南京社保下载
import os
import time
import traceback
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils.download_base import CrawlerDownloadBase

class NanjingSocialSecurity(CrawlerDownloadBase):
    def download_detail(self):
        # 打开页面
        self.driver.set_window_size(1552, 832)

        # 选择社保地区
        self.click_custom((By.XPATH, '//a[@id="320100_1" and text()="南京"]'))
        time.sleep(2)

        # 关闭弹出窗口
        self.click_custom((By.XPATH, '//a[text()="关闭"]'))

        # 点击登录入口页面
        self.click_custom((By.XPATH, '//a[text()="您好！请登录"]'))
        time.sleep(2)

        # 切换单位登录页面
        self.click_custom((By.XPATH, '//span[text()="单位登录"]'))

        # 切换扫码登录页面
        element = self.driver.find_elements(By.XPATH, '//span[text()="扫码登录"]')[1]
        self.click_custom(element)

        # 扫码登录等待
        time.sleep(10)
        # 等待登录成功后，点击"权益查询"菜单
        self.wait((By.XPATH, '//span[text()="单位权益单"]'))
        self.click_custom((By.XPATH, '//span[text()="单位权益单"]'), click_flag=False)

        # 点击【起始年月】控件
        self.wait((By.CSS_SELECTOR, 'input[placeholder="选择年月"]'))
        elements = self.driver.find_elements(By.CSS_SELECTOR, 'input[placeholder="选择年月"]')
        if elements:
            ele_start_date = elements[0]  # 获取第一个匹配的元素
            # 等待元素可点击
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(ele_start_date))
            ele_start_date.click()  # 使用标准的click()方法
        else:
            print("没有找到匹配的元素")

        # 定义你需要的 <td> 元素索引
        start_date = self.user_data[0]['start_date']
        end_date = self.user_data[0]['end_date']

        s_year, s_month = map(int, start_date.split('-'))
        e_year, e_month = map(int, end_date.split('-'))

        # 处理开始年份
        current_year = int(time.strftime("%Y", time.localtime()))
        if s_year < current_year:
            # 左跳差额年份
            element = self.driver.find_elements(By.XPATH, '//a[@role="button" and contains(@title, "上一年")]')[0]
            for i in range(current_year - s_year):
                time.sleep(0.5)
                self.driver.execute_script("arguments[0].click()", element)
        if s_year > current_year:
            # 右跳差额年份
            element = self.driver.find_element(By.XPATH, '//a[@role="button" and contains(@title, "下一年")]')
            for i in range(s_year - current_year):
                time.sleep(0.5)
                self.driver.execute_script("arguments[0].click()", element)
        # 处理开始月份,根据当前月份获取对应的行和列,设每行包含三个月份
        time.sleep(1)
        tr_index = (s_month - 1) // 3  # 计算所在的行，0 表示第一行
        td_index = (s_month - 1) % 3  # 计算所在的列，0 表示第一列
        # 查找对应的 <a> 元素，点击对应月份
        # 定位到开始日期的表格
        start_rows = self.driver.find_element(By.XPATH, '//table[@class="ant-calendar-month-panel-table"]')
        target_td = start_rows.find_element(By.XPATH, f'.//tr[{tr_index + 1}]/td[{td_index + 1}]/a')
        self.driver.execute_script("arguments[0].click()", target_td)
        time.sleep(2)

        # 点击【结束年月】控件
        self.wait((By.CSS_SELECTOR, 'input[placeholder="选择年月"]'))
        elements = self.driver.find_elements(By.CSS_SELECTOR, 'input[placeholder="选择年月"]')
        if len(elements) > 1:
            ele_end_date = elements[1]  # 获取第二个匹配的元素
            # 等待元素可点击
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(ele_end_date))
            ele_end_date.click()  # 使用标准的click()方法
        else:
            print("没有找到第二个匹配的元素")

        # 处理结束年份
        current_year = int(time.strftime("%Y", time.localtime()))
        if e_year < current_year:
            # 左跳差额年份
            element = self.driver.find_elements(By.XPATH, '//a[@role="button" and contains(@title, "上一年")]')[1]
            for i in range(current_year - e_year):
                time.sleep(0.5)
                self.driver.execute_script("arguments[0].click()", element)
        if e_year > current_year:
            # 右跳差额年份
            element = self.driver.find_element(By.XPATH, '//a[@role="button" and contains(@title, "下一年")]')
            for i in range(e_year - current_year):
                time.sleep(0.5)
                self.driver.execute_script("arguments[0].click()", element)
        # 处理结束月份,根据当前月份获取对应的行和列,设每行包含三个月份
        tr_index = (e_month - 1) // 3  # 计算所在的行，0 表示第一行
        td_index = (e_month - 1) % 3  # 计算所在的列，0 表示第一列
        # 查找对应的 <a> 元素，点击对应月份
        # 定位到结束日期的表格
        end_rows = self.driver.find_element(By.XPATH, '//table[@class="ant-calendar-month-panel-table"]')
        target_td = end_rows.find_element(By.XPATH, f'.//tr[{tr_index + 1}]/td[{td_index + 1}]/a')
        self.driver.execute_script("arguments[0].click()", target_td)
        time.sleep(2)

        idown = 0  # 下载计数器
        total_count = len(self.user_data)
        for input_context in self.user_data:
            idown += 1
            self.logger.info(f"开始下载{idown}/{total_count}：{input_context['staff_name']}")

            # 社会保障号
            textarea = self.driver.find_elements(By.CSS_SELECTOR, 'input[type="text"].ant-input[style="width: 95%;"]')[0]

            # 删除已选内容
            textarea.send_keys(Keys.DELETE)
            textarea.clear()
            textarea.send_keys(input_context['id_card'])

            # 先点击查询按钮
            element = self.driver.find_element(By.XPATH, '//button/span[text()="查 询"]')
            self.driver.execute_script("arguments[0].click()", element)
            # 等待查询的遮盖层消失
            time.sleep(3)

            # 再点击勾选打印
            element = self.driver.find_elements(By.XPATH, '//input[@type="checkbox" and @class="ant-checkbox-input"]')[0]
            self.driver.execute_script("arguments[0].click()", element)
            # 移入打印列表
            self.driver.find_elements(By.CSS_SELECTOR, 'svg[data-icon="right"]')[1].click()
    
            # 然后点击预览按钮
            self.shared_state.downloaded_file = None
            element = self.driver.find_element(By.XPATH, '//button[span[text()="预览权益单"]]')
            self.driver.execute_script("arguments[0].click()", element)
            time.sleep(3)

            # 最后点击下载按钮
            self.click_custom((By.XPATH, '//button/span[text()="下载"]'), click_flag=False)
            time.sleep(3)

            new_file_name = f'{input_context['staff_name']}_{input_context['id_card']}_{self.city}_个人参保证明'
            new_file_path = os.path.join(self.download_path, new_file_name)
            self.rename_downfile(new_file_path)

            # 关闭下载页面
            self.driver.find_element(By.CSS_SELECTOR, 'svg[data-icon="close"]').click()

            # 将姓名移出"需打印单位权益单人员列表"
            element = self.driver.find_element(By.XPATH, '//*[@id="components-layout-demo-top-side"]/main/section/main/div/div/form/div[2]/div/div/div[3]/div[2]/div/div/div/div[2]/div/div[1]/table/thead/tr/th[1]/span/div/span[1]/div/label/span/input')
            self.driver.execute_script("arguments[0].click()", element)
            self.driver.find_elements(By.CSS_SELECTOR, 'svg[data-icon="left"]')[1].click()

if __name__ == '__main__':
    try:
        # 根据需要选择城市
        source_dir = os.path.join(os.getcwd(), 'source')
        city_social_security = NanjingSocialSecurity("nanjing", source_dir)  # 返回对应的城市子类
        city_social_security.run()
    except Exception as e:
        traceback.print_exc()