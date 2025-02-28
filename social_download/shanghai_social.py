# -*- coding: utf-8 -*-
# Author:   zhucong1
# At    :   2024/8/14
# Email :   zhucong1@sunline.com
# About :
import os
import time
from selenium.webdriver.common.by import By
from utils.download_base import CrawlerDownloadBase

class ShanghaiSocialSecurity(CrawlerDownloadBase):
    def start_download(self, download_flag, download_user):
        time.sleep(2)
        table = self.driver.find_element(By.ID, 'page-001').find_elements(By.CLASS_NAME, 'mini-grid-rowstable')[1]
        _print = self.driver.find_element(By.ID, 'page-001').find_elements(By.CLASS_NAME, 'btn_bg')[2]

        rows = table.find_elements(By.TAG_NAME, 'tr')[1:]

        # user_data = [
        #     {'id': 1, 'name': '彭敬北', 'id_card': '340824198909126836'},
        #     {'id': 14, 'name': '杨财松', 'id_card': '352231199212202412'},
        #     {'id': 94, 'name': '谢达', 'id_card': '210503199208290331'},
        #     {'id': 738, 'name': '刘蕊', 'id_card': '342422199402211689'},
        # ]
        if download_flag == '1':
            start_id = int(download_user['seq_id'])
        else:
            start_id = int(download_user[0]['seq_id'])

        start_idx = 0
        for k, v in enumerate(rows):
            cells = v.find_elements(By.TAG_NAME, "td")
            if int(cells[2].text) == start_id:
                start_idx = k
                break

        rows = rows[start_idx:]  # 序号601，对应 idx 600
        unknown_data = []

        if download_flag == '1':
            cells = rows[0].find_elements(By.TAG_NAME, "td")
            name = cells[3].text
            id_card = cells[4].text
            if download_user['staff_name'] == name and download_user['id_card'] == id_card:
                cells[1].click()
        else:
            for input_context in download_user:
                j = 0
                while j < len(rows):
                    cells = rows[j].find_elements(By.TAG_NAME, "td")

                    name = cells[3].text
                    id_card = cells[4].text
                    if input_context['staff_name'] == name and input_context['id_card'] == id_card:
                        cells[1].click()
                        rows.pop(j)
                        break
                    j += 1
                else:
                    # 全部找完一遍之后，就是excel有但是网站没有的人
                    unknown_data.append(f'{cells[2].text}：{name}')

        with open('unknown.txt', 'w') as file:
            file.write("，".join(unknown_data))  # 将数据写入文件

        _print.click()

        new_file_name = '上海分公司_单位汇总参保证明'
        new_file_path = os.path.join(self.download_path, new_file_name)
        self.rename_downfile(new_file_path)

        # 本次下载完成，点击[确认]，去除已勾选项
        self.click_custom((By.XPATH, '//*[@id="psnl-ryjfqk"]/tbody/tr/td[4]/a'))

    def download_detail(self):
        self.driver.set_window_size(1552, 832)
        time.sleep(3)
        self.wait((By.ID, '_loginBtn'))
        self.click_custom((By.ID, '_loginBtn'))

        self.wait((By.XPATH, '//*[@id="login-hearder"]/a[2]'))
        self.click_custom((By.XPATH, '//*[@id="login-hearder"]/a[2]'))

        self.wait((By.ID, 'qr-tab'))
        self.click_custom((By.ID, 'qr-tab'))

        keyword = '企业职工就业参保登记 新办'
        self.wait((By.ID, 'minsearch'))
        # 1、ID minsearch  企业职工就业参保登记  searchBtn 点击
        self.driver.find_element(by=By.ID, value='minsearch').send_keys(keyword)
        self.click_custom((By.ID, 'searchBtn'))

        self.driver.implicitly_wait(5)
        # 切换到新窗口
        all_windows =self.driver.window_handles

        self.driver.switch_to.window(all_windows[1])
        self.wait((By.ID, 'results'))
        # 2、ID results 下 第二个DIV 下的button 点击
        self.driver.find_element(by=By.ID, value='results').find_element(By.CSS_SELECTOR, f"button[data-target-name='{keyword}']").click()
        # 切换到新窗口
        all_windows = self.driver.window_handles
        self.driver.switch_to.window(all_windows[2])
        # 3、class layui-nav layui-nav-tree下 第一个li 下的a标签 点击
        # driver.implicitly_wait(60)
        self.driver.switch_to.frame('right')
        # 弹出的框进行证书验证【## 扫码登录不需要 ##】
        # wait(driver, (By.ID, 'layui-layer2'))
        # driver.find_element(by=By.ID, value='layui-layer2').find_element(By.CLASS_NAME, 'layui-layer-input').send_keys('12345678')
        # driver.find_element(by=By.ID, value='layui-layer2').find_element(By.CLASS_NAME, 'layui-layer-btn0').click()

        # driver.switch_to.frame('right')
        self.driver.find_element(by=By.ID, value='main-layout').find_element(By.CSS_SELECTOR, 'a[data-text="社会保险"]').click()
        self.driver.switch_to.frame('iframe201')


        # driver.execute_script("arguments[0].scrollIntoView();", element)  滚动到元素的位置
        # 4、ID menuList下 第一个class sx_one 下的 class more

        self.wait((By.XPATH, '//*[@id="menuList"]/div[1]/p[2]/a'))
        self.click_custom((By.XPATH, '//*[@id="menuList"]/div[1]/p[2]/a'))

        # 5、data-id="S0427"  点击
        self.driver.find_element(By.ID, 'menuList').find_element(By.CSS_SELECTOR, 'li[data-id="S0427"]').find_element(By.CSS_SELECTOR, '.lanmu a').click()


        # driver.switch_to.window(all_windows[2])
        self.driver.switch_to.default_content()
        self.driver.switch_to.frame('right')
        self.driver.switch_to.frame('iframeAF_S0427')

        # 6、page-005-sq 下的ul 第2个li
        self.wait((By.XPATH, '//*[@id="page-005-sq"]/div/div/ul/li[2]'))
        self.click_custom((By.XPATH, '//*[@id="page-005-sq"]/div/div/ul/li[2]'))

        # 7、ID alls  ID psnl-ryjfqk下的 第一个tr下的第4个td下的a标签
        self.click_custom((By.XPATH, '//*[@id="psnl-ryjfqk"]/tbody/tr/td[4]/a'))

        # 分批次下载模式
        batch_size = int(self.download_limit)  # 将字符串转换为整数
        total_count = len(self.user_data)
        total_batches = (total_count + batch_size - 1) // batch_size  # 向上取整得到总批次
        self.logger.info(f"总共需要下载 {total_count} 条数据")
        if self.download_flag == '1':
            # 个人下载模式
            idown = 0  # 下载计数器
            for input_context in self.user_data:
                idown += 1
                self.logger.info(f"开始下载{idown}/{total_count}：{input_context['staff_name']}")
                self.start_download("1",input_context)
        else:
            for batch_num in range(total_batches):
                start_idx = batch_num * batch_size
                end_idx = min((batch_num + 1) * batch_size, total_count)
                current_batch = self.user_data[start_idx:end_idx]

                self.logger.info(f"开始下载第 {batch_num + 1}/{total_batches} 批次 \n"
                                 f"本批次含 {len(current_batch)} 条数据 "
                                 f"({start_idx + 1}~{end_idx})")

                try:
                    self.start_download("2",current_batch)
                except Exception as e:
                    self.logger.error(f"第 {batch_num + 1} 批次下载失败: {str(e)}")
                    raise

                # 如果不是最后一批，等待一段时间再继续
                if batch_num < total_batches - 1:
                    wait_time = 2  # 每批次之间等待2秒
                    self.logger.info(f"等待 {wait_time} 秒后继续下一批次...")
                    time.sleep(wait_time)



