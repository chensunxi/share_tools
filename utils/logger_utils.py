import logging
import os
from logging.handlers import RotatingFileHandler
from datetime import datetime


class LoggerUtils:
    # 用于控制控制台处理器是否已被添加的标记
    _console_handler_added = False

    @staticmethod
    def get_logger(name=None, log_output=None):
        # 使用传入的名称，默认使用 'logger_utils'
        logger_name = name if name else 'logger_utils'

        # 获取 logger 实例
        logger = logging.getLogger(logger_name)

        # 删除所有已存在的处理器，防止重复添加
        if logger.hasHandlers():
            logger.handlers.clear()

        # 设置日志文件保存路径
        log_dir = os.path.join(os.getcwd(), 'logs')
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        log_file = os.path.join(log_dir, f"share.tools.{datetime.now().strftime('%Y%m%d')}.log")

        # 配置日志格式化器
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s')

        # 配置日志文件处理器
        max_log_size = 10 * 1024 * 1024  # 最大日志文件大小为10MB
        backup_count = 5  # 保留最多5个备份
        file_handler = RotatingFileHandler(log_file, maxBytes=max_log_size, backupCount=backup_count)
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)

        # 配置日志文件处理器（只添加一次）
        if not any(isinstance(handler, RotatingFileHandler) for handler in logger.handlers):
            logger.addHandler(file_handler)

        # 配置控制台处理器（避免重复添加）
        # if not LoggerUtils._console_handler_added:
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.DEBUG)
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
            LoggerUtils._console_handler_added = True  # 控制台处理器已添加，标记为已添加

        # 配置文本输出处理器（如果有的话）
        if log_output:
            def log_to_text_handler(record):
                # log_output.config(state='normal')  # 允许修改
                log_output.insert('end',
                                  f"{record.asctime} - {record.levelname} - [{record.filename}:{record.lineno}] - {record.message}\n")
                log_output.yview_pickplace("end")
                # log_output.config(state='disabled')  # 禁止修改

            text_handler = logging.Handler()
            text_handler.emit = log_to_text_handler
            logger.addHandler(text_handler)

        # 设置logger级别
        logger.setLevel(logging.DEBUG)

        return logger
