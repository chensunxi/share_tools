# -*- coding: utf-8 -*-
# Author:   chensx
# At    :   2024/11/15
# Email :   chensx@sunline.com
# About :   北京社保下载
import os
import time
import traceback

from selenium.common import ElementClickInterceptedException, TimeoutException
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils.download_base import CrawlerDownloadBase

class BeijingSocialSecurity(CrawlerDownloadBase):
    def download_detail(self):
        # 打开页面
        # self.driver.get("https://hrss.sz.gov.cn/szsi/")
        # self.driver.get("https://tyrz.gd.gov.cn/pscp/sso/static/?redirect_uri=https%3A%2F%2Fsipub.sz.gov.cn%2Fhspms%2Flogon.do%3Fmethod%3DgdCasCallback%26sxbm%3D%26gdbstoken%3Dnull&client_id=szsbgrwy")
        self.driver.set_window_size(1552, 832)

        # 保存当前窗口句柄
        self.vars["window_handles"] = self.driver.window_handles

        # 点击单位登录页面
        self.driver.find_element(By.XPATH, "//div[@class='login-unit']").click()

        # 扫码登录等待
        time.sleep(10)
        # 等待登录成功后，点击"权益查询"菜单
        self.click_custom((By.XPATH, '//span[@class="item-text" and contains(text(), "权益查询")]'), click_flag=False)

        # 等待新窗口加载
        WebDriverWait(self.driver, 10).until(EC.number_of_windows_to_be(2))
        time.sleep(2)
        all_windows = self.driver.window_handles
        self.driver.switch_to.window(all_windows[-1])

        # 点击“单位职工缴费信息查询”按钮
        self.click_custom((By.CSS_SELECTOR, 'div[data-url="unitRights/inquiryOfPayment"]'), click_flag=False)

        idown = 0  # 下载计数器
        total_count = len(self.user_data)
        for input_context in self.user_data:
            idown += 1
            self.logger.info(f"开始下载{idown}/{total_count}：{input_context['staff_name']}")
            # 点击【起始年月】控件
            self.wait((By.CSS_SELECTOR, 'input[placeholder="选择日期"]'))
            elements = self.driver.find_elements(By.CSS_SELECTOR, 'input[placeholder="选择日期"]')
            if elements:
                ele_start_date = elements[0]  # 获取第一个匹配的元素
                # 等待元素可点击
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(ele_start_date))
                ele_start_date.click()  # 使用标准的click()方法
            else:
                print("没有找到匹配的元素")

            # 点击【结束年月】控件
            self.wait((By.CSS_SELECTOR, 'input[placeholder="选择日期"]'))
            elements = self.driver.find_elements(By.CSS_SELECTOR, 'input[placeholder="选择日期"]')
            if len(elements) > 1:
                ele_end_date = elements[1]  # 获取第二个匹配的元素
                # 等待元素可点击
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(ele_end_date))
                ele_end_date.click()  # 使用标准的click()方法
            else:
                print("没有找到第二个匹配的元素")

            # 定位到开始日期和结束日期控件的月份的表格
            start_rows = self.driver.find_elements(By.XPATH, '//table[contains(@class, "el-month-table")]')[0]
            end_rows = self.driver.find_elements(By.XPATH, '//table[contains(@class, "el-month-table")]')[1]

            # 定义你需要的 <td> 元素索引
            start_date = input_context['start_date']
            end_date = input_context['end_date']

            s_year, s_month = map(int, start_date.split('-'))
            e_year, e_month = map(int, end_date.split('-'))

            # 处理开始年份
            default_start_date = ele_start_date.find_element(By.XPATH, '..').get_attribute("data")
            if default_start_date is None:
                current_year = int(time.strftime("%Y", time.localtime()))
            else:
                current_year = int(default_start_date[:4])
            if s_year < current_year:
            # 左跳差额年份
                element = self.driver.find_elements(By.CSS_SELECTOR, 'button[aria-label="前一年"]')[0]
                for i in range(current_year - s_year):
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click()", element)
            if s_year > current_year:
                # 右跳差额年份
                element = self.driver.find_elements(By.CSS_SELECTOR, 'button[aria-label="后一年"]')[0]
                for i in range(s_year - current_year):
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click()", element)
            # 处理开始月份,根据当前月份获取对应的行和列,设每行包含四个月份
            time.sleep(1)
            tr_index = (s_month - 1) // 4  # 计算所在的行，0 表示第一行
            td_index = (s_month - 1) % 4  # 计算所在的列，0 表示第一列
            # 查找对应的 <a> 元素，点击对应月份
            target_td = start_rows.find_element(By.XPATH, f'.//tr[{tr_index + 1}]/td[{td_index + 1}]/div/a')
            self.driver.execute_script("arguments[0].click()", target_td)

            # 处理结束年份
            # 获取日期控件的默认日期
            default_end_date = ele_end_date.find_element(By.XPATH, '..').get_attribute("data")
            if default_end_date is None:
                current_year = int(time.strftime("%Y", time.localtime()))
            else:
                current_year = int(default_end_date[:4])
            if e_year < current_year:
                # 左跳差额年份
                element = self.driver.find_elements(By.CSS_SELECTOR, 'button[aria-label="前一年"]')[1]
                for i in range(current_year - e_year):
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click()", element)
            if e_year > current_year:
                # 右跳差额年份
                element = self.driver.find_elements(By.CSS_SELECTOR, 'button[aria-label="后一年"]')[1]
                for i in range(e_year - current_year):
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click()", element)
            # 处理结束月份,根据当前月份获取对应的行和列,设每行包含四个月份
            tr_index = (e_month - 1) // 4  # 计算所在的行，0 表示第一行
            td_index = (e_month - 1) % 4  # 计算所在的列，0 表示第一列
            # 查找对应的 <a> 元素，点击对应月份
            target_td = end_rows.find_element(By.XPATH, f'.//tr[{tr_index + 1}]/td[{td_index + 1}]/div/a')
            self.driver.execute_script("arguments[0].click()", target_td)

            # 社会保障号
            textarea = self.driver.find_element(By.XPATH, "//textarea[@placeholder='输入社会保障号码']")

            # 删除已选内容
            textarea.send_keys(Keys.DELETE)
            textarea.clear()
            textarea.send_keys(input_context['id_card'])

            # 先点击查询按钮
            self.click_custom((By.XPATH, '//button[span[contains(text(),"查询")]]'), click_flag=True)
            # 等待查询的遮盖层消失
            time.sleep(5)

            # 再点击下载按钮
            self.shared_state.downloaded_file = None
            self.wait((By.XPATH, '//button[span[contains(text(),"下载打印")]]'))
            download_button = self.driver.find_element(by=By.XPATH, value='//button[span[contains(text(),"下载打印")]]')
            # 判断 disabled 属性
            is_disabled = download_button.get_attribute("disabled") is not None
            if is_disabled:
                print("查询社保记录不存在，不能下载")
                continue
            else:
                # 等待元素可点击
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[span[contains(text(),"下载打印")]]'))
                )
                try:
                    download_button.click()
                except ElementClickInterceptedException:
                    WebDriverWait(self.driver, 20).until(
                        EC.invisibility_of_element_located((By.CLASS_NAME, "el-loading-mask"))
                    )
                    download_button.click()
            time.sleep(3)

            new_file_name = f'{input_context['staff_name']}_{input_context['id_card']}_{self.city}_个人参保证明'
            new_file_path = os.path.join(self.download_path, new_file_name)
            self.rename_downfile(new_file_path)

if __name__ == '__main__':
    try:
        # 根据需要选择城市
        source_dir = os.path.join(os.getcwd(), 'source')
        city_social_security = BeijingSocialSecurity("beijing", source_dir)  # 返回对应的城市子类
        city_social_security.run()
    except Exception as e:
        traceback.print_exc()