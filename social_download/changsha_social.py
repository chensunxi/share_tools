# -*- coding: utf-8 -*-
# Author:   chensx
# At    :   2025/01/15
# Email :   chensx@sunline.com
# About :   长沙社保下载
import os
import time
import traceback
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from utils.download_base import CrawlerDownloadBase

class ChangshaSocialSecurity(CrawlerDownloadBase):
    def download_detail(self):
        self.driver.set_window_size(1552, 832)

        self.click_custom((By.XPATH, '//span[contains(text(),"CA账户登录")]'), click_flag=False)
        # 登录等待
        time.sleep(10)
        # 等待登录成功后，点击"社保服务"菜单
        self.wait((By.CSS_SELECTOR, 'span[title="社保服务"]'), 300)
        self.click_custom((By.CSS_SELECTOR, 'span[title="社保服务"]'), click_flag=False)

        # 点击"查询打印"菜单
        self.click_custom((By.CSS_SELECTOR, 'span[title="查询打印"]'), click_flag=False)

        # 点击"单位花名册打印"菜单
        self.click_custom((By.CSS_SELECTOR, 'span[title="单位花名册打印"]'), click_flag=False)
        time.sleep(3)

        # 显式等待元素加载完成
        # printarea = WebDriverWait(self.driver, 10).until(
        #     EC.visibility_of_element_located((By.XPATH, '//input[@placeholder="请输入打印用途"]'))
        # )
        # 查找所有的 iframe 元素
        iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
        # 确保页面中至少有两个 iframe
        if len(iframes) > 0:
            self.driver.switch_to.frame(iframes[0])
        time.sleep(1)
        # 填写"打印用途"输入框
        printarea = self.driver.find_element(By.XPATH, '//input[@placeholder="请输入打印用途"]')
        printarea.send_keys('员工社保证明')

        idown = 0  # 下载计数器
        total_count = len(self.user_data)
        for input_context in self.user_data:
            idown += 1
            self.logger.info(f"开始下载{idown}/{total_count}：{input_context['staff_name']}")

            # 证件号码
            textarea = self.driver.find_element(By.CSS_SELECTOR, 'input[minlength="15"][maxlength="18"]')

            # 删除已选内容
            textarea.send_keys(Keys.DELETE)
            textarea.clear()
            textarea.send_keys(input_context['id_card'])
            # 先点击查询按钮
            self.driver.find_element(By.XPATH, '//button[span[text()="查询"]]').click()
            # 等待查询的遮盖层消失
            time.sleep(3)

            # 点击勾选按钮
            element = self.driver.find_elements(By.CSS_SELECTOR, 'input[type="checkbox"][aria-hidden="false"]')[0]
            self.driver.execute_script("arguments[0].click()", element)

            # 再点击下载按钮
            self.shared_state.downloaded_file = None
            download_button = self.driver.find_element(By.XPATH, '//span[text()="打印"]')
            download_button.click()
            time.sleep(3)

            new_file_name = f'{input_context['staff_name']}_{input_context['id_card']}_{self.city}_个人参保证明'
            new_file_path = os.path.join(self.download_path, new_file_name)
            self.rename_downfile(new_file_path)

if __name__ == '__main__':
    try:
        # 根据需要选择城市
        source_dir = os.path.join(os.getcwd(), 'source')
        city_social_security = ChangshaSocialSecurity("changsha", source_dir)  # 返回对应的城市子类
        city_social_security.run()
    except Exception as e:
        traceback.print_exc()