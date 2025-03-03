# -*- coding: utf-8 -*-
# Author:   zhucong1
# At    :   2024/8/14
# Email :   zhucong1@sunline.com
# About :
import os
import time
import traceback

import fitz
from pymupdf import pymupdf
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils.download_base import CrawlerDownloadBase
from social_desensitize.social_desensitize import GuangzhouCover

class GuangzhouSocialSecurity(CrawlerDownloadBase):
    def download_file(self, file_type):
        idown = 0  # 下载计数器
        total_count = len(self.user_data)
        for input_context in self.user_data:
            idown += 1
            self.logger.info(f"开始下载{idown}/{total_count}：{input_context['staff_name']}")
            self.wait((By.XPATH, '//*[@id="app"]/div[1]/div[3]/div[2]/div/div[2]/div[2]/div/div/div[2]/iframe'))
            element = self.driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div[3]/div[2]/div/div[2]/div[2]/div/div/div[2]/iframe')
            self.driver.switch_to.frame(element)

            # 点击【起始年月】控件
            ele_start_date = self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/div[1]/form/div/div[2]/div/div/div/input')
            # 等待元素可点击
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(ele_start_date))
            ele_start_date.click()  # 使用标准的click()方法

            # 定位到开始日期和结束日期控件的月份的表格
            start_rows = self.driver.find_elements(By.XPATH, '//table[contains(@class, "el-month-table")]')[0]

            # 定义你需要的 <td> 元素索引
            start_date = input_context['start_date']
            end_date = input_context['end_date']

            s_year, s_month = map(int, start_date.split('-'))
            e_year, e_month = map(int, end_date.split('-'))

            # 处理开始年份
            # 获取日期控件的默认日期
            default_start_date = ele_start_date.find_element(By.XPATH, '..').get_attribute("data")
            if default_start_date is None:
                current_year = int(time.strftime("%Y", time.localtime()))
            else:
                current_year = int(default_start_date[:4])
            if s_year > current_year:
                # 右跳差额年份
                element = self.driver.find_elements(By.CSS_SELECTOR, 'button[aria-label="后一年"]')[0]
                for i in range(s_year - current_year):
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click()", element)
            if s_year < current_year:
                # 左跳差额年份
                element = self.driver.find_elements(By.CSS_SELECTOR, 'button[aria-label="前一年"]')[0]
                for i in range(current_year - s_year):
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click()", element)

            # 处理开始月份,根据当前月份获取对应的行和列,设每行包含四个月份
            time.sleep(1)
            tr_index = (s_month - 1) // 4  # 计算所在的行，0 表示第一行
            td_index = (s_month - 1) % 4  # 计算所在的列，0 表示第一列
            # 查找对应的 <a> 元素，点击对应月份
            target_td = start_rows.find_element(By.XPATH, f'.//tr[{tr_index + 1}]/td[{td_index + 1}]/div/a')
            self.driver.execute_script("arguments[0].click()", target_td)

            # 点击【结束年月】控件
            ele_end_date = self.driver.find_element(By.XPATH,
                                               '//*[@id="app"]/div/div[2]/div[2]/div[2]/div[1]/form/div/div[3]/div/div/div/input')
            # 等待元素可点击
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(ele_end_date))
            ele_end_date.click()  # 使用标准的click()方法
            end_rows = self.driver.find_elements(By.XPATH, '//table[contains(@class, "el-month-table")]')[1]
            # 处理结束年份
            # 获取日期控件的默认日期
            default_end_date = ele_end_date.find_element(By.XPATH, '..').get_attribute("data")
            if default_end_date is None:
                current_year = int(time.strftime("%Y", time.localtime()))
            else:
                current_year = int(default_end_date[:4])
            if e_year > current_year:
                # 右跳差额年份
                element = self.driver.find_elements(By.CSS_SELECTOR, 'button[aria-label="后一年"]')[1]
                for i in range(e_year - current_year):
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click()", element)
            if e_year < current_year:
                # 左跳差额年份
                element = self.driver.find_elements(By.CSS_SELECTOR, 'button[aria-label="前一年"]')[1]
                for i in range(current_year - e_year):
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click()", element)

            # 处理结束月份,根据当前月份获取对应的行和列,设每行包含四个月份
            tr_index = (e_month - 1) // 4  # 计算所在的行，0 表示第一行
            td_index = (e_month - 1) % 4  # 计算所在的列，0 表示第一列
            # 查找对应的 <a> 元素，点击对应月份
            target_td = end_rows.find_element(By.XPATH, f'.//tr[{tr_index + 1}]/td[{td_index + 1}]/div/a')
            self.driver.execute_script("arguments[0].click()", target_td)

            # 输入证件号码
            textarea = self.driver.find_element(By.XPATH,
                                                '//*[@id="app"]/div/div[2]/div[2]/div[2]/div[1]/form/div/div[4]/div/div/div[1]/input')
            # 删除已选内容
            textarea.clear()
            textarea.send_keys(input_context['id_card'])

            # 输入姓名
            textarea = self.driver.find_element(By.XPATH,
                                                '//*[@id="app"]/div/div[2]/div[2]/div[2]/div[1]/form/div/div[5]/div/div/div/input')
            # 删除已选内容
            textarea.clear()
            textarea.send_keys(input_context['staff_name'])

            # 先点击查询按钮
            self.wait((By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/div[1]/form/div/div[6]/div/div/button[2]'))
            self.driver.find_element(by=By.XPATH,
                                     value='//*[@id="app"]/div/div[2]/div[2]/div[2]/div[1]/form/div/div[6]/div/div/button[2]').click()
            time.sleep(3)
            # 再点击【打印凭证】按钮
            self.wait(
                (By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div[3]/table/tbody/tr/td[8]/div/a'))
            print_button = self.driver.find_element(by=By.XPATH,
                                                       value='//*[@id="app"]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div[3]/table/tbody/tr/td[8]/div/a')
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(print_button))
            print_button.click()

            # 等待新窗口加载
            WebDriverWait(self.driver, 30).until(EC.number_of_windows_to_be(2))
            time.sleep(2)
            all_windows = self.driver.window_handles
            self.driver.switch_to.window(all_windows[-1])

            # 再点击【下载】按钮
            self.wait((By.XPATH, '//*[@id="download"]'))
            download_button = self.driver.find_element(by=By.XPATH, value='//*[@id="download"]')
            download_button.click()
            time.sleep(3)

            # 重命名下载的文件
            if file_type == "1":
                new_file_name = f"{input_context['staff_name']}_{input_context['id_card']}_{self.city}_个人参保证明"
                source_dir = os.path.join(self.download_path, "参保证明")
                if not os.path.exists(source_dir):
                    os.makedirs(source_dir)
                new_file_path = os.path.join(source_dir, new_file_name)
            else:
                new_file_name = f"{input_context['staff_name']}_{input_context['id_card']}_{self.city}_个人缴费明细"
                source_dir = os.path.join(self.download_path, "缴费明细")
                if not os.path.exists(source_dir):
                    os.makedirs(source_dir)
                new_file_path = os.path.join(source_dir, new_file_name)

            source_file_path = self.rename_downfile(new_file_path)

            if file_type == "2":
                base, ext = os.path.splitext(source_file_path)
                target_cover_file = base + "_cover" + ext  # 例如添加 '_cover' 后缀
                target_dir = os.path.join(self.download_path, "缴费明细_脱敏")
                if not os.path.exists(target_dir):
                    os.makedirs(target_dir)
                target_file_path = os.path.join(target_dir, target_cover_file)
                doc = pymupdf.Document(source_file_path)
                city_cover = GuangzhouCover(source_file_path, target_file_path, self.out_logger)
                city_cover.cover_by_doc(doc)
                doc.save(target_file_path, encryption=fitz.PDF_ENCRYPT_KEEP)
                self.logger.info(f"{'+' * 10}【{input_context['staff_name']}】社保数据脱敏处理完成{'+' * 10}")

            # 执行完操作后关闭当前标签页
            self.driver.close()
            time.sleep(0.5)
            # 切换回之前的标签页
            all_windows = self.driver.window_handles
            self.driver.switch_to.window(all_windows[0])

    def download_detail(self):
        # 打开页面
        # self.driver.get("https://hrss.sz.gov.cn/szsi/")
        # self.driver.get("https://tyrz.gd.gov.cn/pscp/sso/static/?redirect_uri=https%3A%2F%2Fsipub.sz.gov.cn%2Fhspms%2Flogon.do%3Fmethod%3DgdCasCallback%26sxbm%3D%26gdbstoken%3Dnull&client_id=szsbgrwy")
        self.driver.set_window_size(1552, 832)

        # 保存当前窗口句柄
        self.vars["window_handles"] = self.driver.window_handles

        # 点击“登录”进入登录页面
        element = self.driver.find_element(by=By.XPATH, value="//div[@class='btn-login']")
        self.driver.execute_script("arguments[0].click()", element)

        # 在新窗口中切换"单位登录"TAB页
        self.wait((By.XPATH, '//*[@id="body"]/div/div/div/div/div[4]/div/div/div[1]/div[2]'))
        element = self.driver.find_element(by=By.XPATH,
                                           value='//*[@id="body"]/div/div/div/div/div[4]/div/div/div[1]/div[2]')
        self.driver.execute_script("arguments[0].click()", element)

        # 输入账号
        element = self.driver.find_element(by=By.XPATH, value='//*[@id="uesrname"]')
        element.send_keys("Sun300348")

        # 扫码登录等待
        time.sleep(10)

        # 等待“个人权益记录（参保证明）查询打印”元素可见
        # menu_element = WebDriverWait(self.driver, 10).until(
        #     EC.visibility_of_element_located((By.XPATH,
        #                                       '//div[@class="menu-name menu-more" and contains(text(), "个人权益记录（参保证明）查询打印")]'))
        # )
        # # 使用 ActionChains 执行鼠标悬停操作
        # actions = ActionChains(self.driver)
        # actions.move_to_element(menu_element).perform()

        # 选择填入【搜索输入框】
        self.wait((By.CSS_SELECTOR, "input[placeholder='请输入您要办理的事项或服务']"))
        element = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='请输入您要办理的事项或服务']")
        self.driver.execute_script("arguments[0].click()", element)
        time.sleep(1)
        # 页面存在跳转，需再次选中输入框，填写"个人参保证明查询打印"
        self.wait((By.CSS_SELECTOR, "input[placeholder='请输入您要办理的事项或服务']"))
        searchtext = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='请输入您要办理的事项或服务']")
        searchtext.send_keys("个人参保证明查询打印")
        # 点击【搜索】按钮
        element = self.driver.find_element(By.XPATH, "//button[span[text()='搜索']]")
        self.driver.execute_script("arguments[0].click()", element)
        # 点击【单位办理】按钮
        self.wait((By.XPATH, "//button/span[starts-with(text(),'单位办理')]"))
        element = self.driver.find_element(By.XPATH, "//button/span[starts-with(text(),'单位办理')]")
        self.driver.execute_script("arguments[0].click()", element)

        self.logger.info("下载文件类型:《《《参保证明》》》")
        self.download_file("1")

        # 执行返回上一页的操作
        self.driver.back()
        # 可选：等待一段时间，查看返回效果
        time.sleep(1)

        # 选择填入【搜索输入框】
        searchtext = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='请输入您要办理的事项或服务']")
        self.driver.execute_script("arguments[0].click()", searchtext)
        time.sleep(1)
        # 填写"个人参保证明查询打印"
        searchtext.clear()
        searchtext.send_keys("个人缴费证明查询打印")
        # 点击【搜索】按钮
        element = self.driver.find_element(By.XPATH, "//button[span[text()='搜索']]")
        self.driver.execute_script("arguments[0].click()", element)
        # 点击【单位办理】按钮
        self.wait((By.XPATH, "//button/span[starts-with(text(),'单位办理')]"))
        element = self.driver.find_element(By.XPATH, "//button/span[starts-with(text(),'单位办理')]")
        self.driver.execute_script("arguments[0].click()", element)

        self.logger.info("下载文件类型:《《《缴费明细》》》")
        self.download_file("2")

if __name__ == '__main__':
    try:
        # 根据需要选择城市
        source_dir = os.path.join(os.getcwd(), 'download')
        city_social_security = GuangzhouSocialSecurity("guangzhou", source_dir)  # 返回对应的城市子类
        city_social_security.run()
    except Exception as e:
        traceback.print_exc()
