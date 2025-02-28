# -*- coding: utf-8 -*-
# Author:   zhouju
# At    :   2024/8/7
# Email :   zhouju@sunline.com
# About :   https://blog.codingcat.net


import argparse
import json
import logging
import os
import re
import shutil
import time
import traceback
import sys

import pymupdf
import fitz
from utils.logger_utils import LoggerUtils

def get_script_directory():
    if getattr(sys, 'frozen', False):
        # 如果是打包后的可执行文件
        return os.path.dirname(sys.executable)
    else:
        # 如果是源代码脚本
        return os.path.dirname(os.path.abspath(__file__))


# os.chdir(get_script_directory())
#
# console = logging.StreamHandler()
# console.setLevel(logging.INFO)
# logger = logging.getLogger(__name__)
# logger.setLevel(level=logging.INFO)
# handler = logging.FileHandler(f"{datetime.now().strftime('%Y%m%d%H%M%S')}.log")
# handler.setLevel(logging.INFO)
# formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# handler.setFormatter(formatter)
#
# console.setLevel(logging.INFO)
# console.setFormatter(formatter)
# logger.addHandler(handler)
# logger.addHandler(console)


class Base(object):
    def __init__(self, source, target, out_logger, *args, **kwargs):
        self._source = source
        self._target = target
        self.logger = LoggerUtils.get_logger(None, out_logger)


    def process(self):
        for pdf in os.listdir(self._source):
            if pdf.endswith('.pdf'):
                doc = pymupdf.Document(os.path.join(self._source, pdf))
                self.cover_by_doc(doc)
                doc.save(f'{self._target}/{pdf}', encryption=fitz.PDF_ENCRYPT_KEEP)

    def cover_by_doc(self, doc):
        for page in doc:
            self.cover_by_page(page)

    def cover_by_page(self, page):
        ...


class XianCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):

        # 根据实际情况找block
        blocks = page.get_text("blocks")[9:21]

        if not blocks:
            return

        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[120, 43], [255, 30], [320, 40], [445, 40]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()


class ChangshaCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        if page.number == 0:
            i = 0
            start = end = 0
            while i < len(blocks):
                block = blocks[i]
                if '缴费明细' in block[4]:
                    start = i + 2
                if '盖章处' in block[4]:
                    end = i
                i += 1
            blocks = blocks[start: end]
        else:
            i = 0
            start = end = 0
            while i < len(blocks):
                if '盖章处' in blocks[i][4]:
                    end = i
                i += 1
            blocks = blocks[start: end]

        if not blocks:
            return

        y0 = blocks[1][1]

        if page.number == 0:
            y0 = blocks[2][1]

        y1 = blocks[-1][3]
        bias = [[180, 25], [235, 25], [295, 25]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()

        # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖3块
        #     bias = [[180, 25], [235, 25], [295, 25]]
        #     for i in bias:
        #         x0 = i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


class ZhengzhouCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_special(self, page, block):
        page_rect = page.rect

        x0, _, x1, _ = page_rect
        _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        # 覆盖3块
        bias = [[142, 162], [255, 140]]
        for i in bias:
            x0 += i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        if page.number == 0:
            i = 0
            start = end = 0
            while i < len(blocks):
                block = blocks[i]
                if '险种' in block[4]:
                    self.cover_by_special(page, blocks[i + 1])
                    i += 1
                    continue

                if '缴费基数' in block[4]:
                    start = i + 1
                    end = i + 12 + 1
                    break

                i += 1

            blocks = blocks[start: end]

        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[75, 30], [235, 30], [410, 30]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()

        # page_rect = page.rect
        #
        # # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖3块
        #     bias = [[60, 70], [175, 50], [170, 55]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


class FuzhouCover(Base):
    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        i = 0
        start = end = 0
        while i < len(blocks):
            block = blocks[i]
            if '单位管理码' in block[4]:
                start = i + 1

            if '打印日期' in block[4]:
                end = i

            i += 1

        blocks = blocks[start: end] + blocks[end + 1: end + 2]

        if not blocks:
            return

        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[540, 25]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()


        # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖3块
        #     bias = [[540, 50]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


class XiamenCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        if '社会保险参保缴费情况证明(单位)' in blocks[1][4]:
            blocks, bias = self.deal_prove(page)
        else:
            blocks, bias = self.deal_attach(page)

        page_rect = page.rect

        y0 = blocks[0][1]
        y1 = blocks[-1][3]

        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()


        for block in blocks:
            if '合计' in block[4]:
                bias = [[380, 400]]
                x0, _, x1, _ = page_rect
                _, y0, _, y1 = block[:4]  # 文本块的位置坐标
                for i in bias:
                    x0 += i[0]
                    x1 = x0 + i[1]

                    # 创建黑色的矩形区域
                    rect = fitz.Rect(x0, y0, x1, y1)
                    page.add_redact_annot(rect, fill=(1, 1, 1))
                    page.apply_redactions()

    # 证明
    def deal_prove(self, page):
        blocks = page.get_text("blocks")
        i = 0
        start = end = 0
        while i < len(blocks):
            block = blocks[i]
            if '职业年金' in block[4]:
                start = i + 1

            if '说明：1.依据社保费规则' in block[4]:
                end = i

            i += 1

        blocks = blocks[start: end]

        return blocks, [[222, 30], [385, 28], [625, 23], [665, 22], [710, 25]]

    # 附表
    def deal_attach(self, page):
        blocks = page.get_text("blocks")
        i = 0
        start = 0
        while i < len(blocks):
            block = blocks[i]
            if '职业\n年金' in block[4]:
                start = i + 1

            i += 1

        blocks = blocks[start: ]

        return blocks, [
            [315, 20], [355, 20], [395, 20], [425, 20], [450, 20],
            [580, 20], [615, 20], [650, 20], [710, 20]
        ]


class NankingCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        start = end = 1

        if page.number == 0:
            start = 23

        i = start
        while i < len(blocks):
            block = blocks[i]
            if '说明' in block[4] or '盖章' in block[4]:
                end = i
                break

            i += 1

        blocks = blocks[start: end]

        if not blocks:
            return

        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[235, 25], [295, 18], [345, 20], [405, 20], [450, 25]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()



        # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖3块
        #     bias = [[215, 270]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


class ChengduCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        start = end = 0

        i = 0
        while i < len(blocks):
            block = blocks[i]
            if '企业缴费人数' in block[4]:
                start = i + 1

            if '欠费情况（从单位' in block[4]:
                end = i
                break

            i += 1

        blocks = blocks[start: end]

        if not blocks:
            return
        y0 = blocks[0][1] # x0 y0  x1 y1
        y1 = blocks[-1][3]
        bias = [[170, 28], [250, 18], [300, 25], [420, 27], [485, 18], [525, 23]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()


        # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖3块
        #     bias = [[155, 165], [260, 140]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


class ShenyangCover(Base):
    # PDF中是图片，无法处理

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        ...


class YantaiCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        if page.number == 0:
            return

        blocks = page.get_text("blocks")
        start = end = 0

        i = 0
        while i < len(blocks):
            block = blocks[i]
            if '缴费基数' in block[4]:
                start = i + 1

            if '打印流水号' in block[4]:
                end = i
                break

            i += 1

        blocks = blocks[start: end]

        if not blocks:
            return
        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[295, 25]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()

        # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖1块
        #     bias = [[295, 55]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


class TaiyuanCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        start = end = 0

        i = 0
        while i < len(blocks):
            block = blocks[i]
            if '基本养老保险' in block[4]:
                start = i + 1

            if '说明' in block[4]:
                end = i
                break

            i += 1

        blocks = blocks[start: end]

        if not blocks:
            return
        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[190, 25], [255, 22]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()

        # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖1块
        #     bias = [[190, 110]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


class FoshanCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        start = end = 0

        i = 0
        while i < len(blocks):
            block = blocks[i]
            if '个人缴费单位缴费' in block[4]:
                start = i + 1

            if '1、表中' in block[4]:
                end = i
                break

            i += 1

        blocks = blocks[start: end]

        if not blocks:
            return
        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[160, 15], [200, 23], [250, 15], [290, 15], [340, 18], [380, 18], [430, 18], [480, 18]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()

        # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖1块
        #     bias = [[150, 360]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


class GuangzhouCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        start = end = 0

        i = 0
        while i < len(blocks):
            block = blocks[i]
            if '个人缴费单位缴费' in block[4]:
                start = i + 1

            if '1、表中' in block[4]:
                end = i
                break

            i += 1
            end = i

        if page.number > 0:
            start = 1

        blocks = blocks[start: end]

        # 根据实际情况找block
        # _, y0, _, y1 = block[:4]
        if not blocks:
            return
        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[160, 15], [200, 23], [250, 15], [290, 15], [340, 18], [380, 18], [430, 18], [480, 18]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1) # 左下x,左下y,右上x,右上y
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     # _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖1块
        #     bias = [[160, 15], [45, 15], [45, 15], [35, 15], [50, 15], [45, 15], [50, 15], [50, 15]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         # 可以让被遮盖的内容不能被复制
        #         page.add_redact_annot(rect, fill=(1, 1, 1))
        #         page.apply_redactions()


class XinjiangCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        start = end = 0

        i = 0
        while i < len(blocks):
            block = blocks[i]
            if '缴费\n标志' in block[4]:
                start = i + 1

            if '注：1、该单据' in block[4]:
                end = i
                break

            i += 1

        blocks = blocks[start: end]
        # page_rect = page.rect

        if not blocks:
            return
        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[425, 17], [475, 17], [525, 15]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()

        # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖1块
        #     bias = [[405, 140]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         # 可以让被遮盖的内容不能被复制
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


class NanNingCover(Base):

    def cover_by_doc(self, doc):
        for idx, page in enumerate(doc):
            self.cover_by_page(page)

    def cover_by_page(self, page):
        blocks = page.get_text("blocks")
        start = end = 0

        i = 0
        while i < len(blocks):
            block = blocks[i]
            if '缴费基数\n缴费' in block[4]:
                start = i + 2

            if '备注' in block[4]:
                end = i
                break

            i += 1

        blocks = blocks[start: end]

        if not blocks:
            return
        y0 = blocks[0][1]
        y1 = blocks[-1][3]
        bias = [[200, 17], [270, 17], [340, 17], [425, 20]]
        for i in bias:
            x0 = i[0]
            x1 = x0 + i[1]

            # 创建黑色的矩形区域
            rect = fitz.Rect(x0, y0, x1, y1)
            # 可以让被遮盖的内容不能被复制
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions()

        # 根据实际情况找block
        # for block in blocks:
        #     x0, _, x1, _ = page_rect
        #     _, y0, _, y1 = block[:4]  # 文本块的位置坐标
        #     # 覆盖1块
        #     bias = [[405, 140]]
        #     for i in bias:
        #         x0 += i[0]
        #         x1 = x0 + i[1]
        #
        #         # 创建黑色的矩形区域
        #         rect = fitz.Rect(x0, y0, x1, y1)
        #         # 可以让被遮盖的内容不能被复制
        #         page.add_redact_annot(rect, fill=(0.7, 0.7, 0.7))
        #         page.apply_redactions()


def cover(source_dir, target_dir):
    logger.info(f'{"=" * 10}开始进行PDF数据覆盖{"=" * 10}')
    for i in os.listdir(source_dir):
        source = os.path.join(source_dir, i)
        target = os.path.join(target_dir, i)
        if not os.path.exists(target):
            os.makedirs(target)
        category_map = {
            # 'xian': XianCover,
            # 'changsha': ChangshaCover,
            # 'zhengzhou': ZhengzhouCover,
            # 'fuzhou': FuzhouCover,
            # 'xiamen': XiamenCover,
            # 'nanjing': NankingCover,
            # 'chengdu': ChengduCover,
            # 'shenyang': ShenyangCover, # 暂无法处理
            # 'yantai': YantaiCover,
            # 'taiyuan': TaiyuanCover,
            # 'foshan': GuangzhouCover,
            'guangzhou': GuangzhouCover,
            # 'xinjiang': XinjiangCover,
            # 'nanning': NanNingCover,
        }
        if not category_map.get(i):
            continue
        processor = category_map[i](source, target)
        processor.process()
    logger.info(f'{"=" * 10}PDF数据覆盖完成{"=" * 10}')


if __name__ == '__main__':
    source_dir = os.path.join(os.getcwd(), 'source')
    target_dir = os.path.join(os.getcwd(), 'target')
    try:
        cover(source_dir, target_dir)
    except Exception as e:
        traceback.print_exc()
        logger.error(f"数据处理失败: {str(e)}")
    logger.info(f"{'+' * 10}社保数据处理完成{'+' * 10}")
