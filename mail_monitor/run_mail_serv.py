from datetime import datetime
import os
import sys
import time
import threading
from mail_monitor.mail_process import AutoReplyMail
from mail_monitor.mail_controller import app
from colorama import Fore, Style, init

class MailMonitor:
    def __init__(self):
        self.mail_service = None
        self.running = False

    def start(self):
        try:
            # 启动邮件服务
            self.mail_service = AutoReplyMail([], [], '')  # 初始化时使用空配置
            self.mail_service.start_monitoring()
            self.running = True
            print("✅ 邮件监听服务已启动")
        except Exception as e:
            print(f"❌ 启动失败: {str(e)}")
            raise e

    def stop(self):
        if self.mail_service:
            try:
                self.mail_service.stop_monitoring()
                self.running = False
                print("✅ 邮件监听服务已停止")
            except Exception as e:
                print(f"❌ 停止失败: {str(e)}")
                raise e

    def is_running(self):
        return self.running

    def get_config(self):
        if self.mail_service:
            return {
                "监听状态": "运行中" if self.running else "已停止",
                "监听邮箱数量": len(self.mail_service.recv_monitor_mail) if hasattr(self.mail_service, 'recv_monitor_mail') else 0,
                "抄送邮箱数量": len(self.mail_service.send_copy_mail) if hasattr(self.mail_service, 'send_copy_mail') else 0
            }
        return {"监听状态": "未初始化"}

# 定义emoji和颜色常量
EMOJI = {
    "MENULIST": "📋",
    "MENU1": "📧",
    "MENU2": "🛑",
    "MENU3": "📊",
    "MENU4": "⚙️",
    "MENU5": "🚪",
    "SUCCESS": "✅",
    "ERROR": "❌",
    "WARN": "⚠️",
    "ARROW": "➜",
    "LANG": "🌐",
    "UPDATE": "🔄"
}

class MailServUI:
    def __init__(self):
        self.mail_monitor = None
        self.running = True
        self.flask_thread = None
        self.flask_host = '0.0.0.0'
        self.flask_port = 5000
        # Initialize colorama
        init()

    def start_flask_server(self):
        """启动Flask服务器"""
        from waitress import serve
        print(f"服务器已启动，访问 http://{self.flask_host}:{self.flask_port}")
        serve(app, host=self.flask_host, port=self.flask_port)

    def clear_screen(self):
        os.system('cls' if os.name == 'nt' else 'clear')

    def show_logo(self):
        MAIL_MONITOR_LOGO = f"""
{Fore.GREEN} ═══════════════════════════════════════════════════════════════════════════════════════════
{Fore.CYAN}
  ███╗   ███╗ █████╗ ██╗██╗      ███╗   ███╗ ██████╗ ███╗   ██╗██╗████████╗ ██████╗ ██████╗ 
  ████╗ ████║██╔══██╗██║██║      ████╗ ████║██╔═══██╗████╗  ██║██║╚══██╔══╝██╔═══██╗██╔══██╗
  ██╔████╔██║███████║██║██║      ██╔████╔██║██║   ██║██╔██╗ ██║██║   ██║   ██║   ██║██████╔╝
  ██║╚██╔╝██║██╔══██║██║██║      ██║╚██╔╝██║██║   ██║██║╚██╗██║██║   ██║   ██║   ██║██╔══██╗
  ██║ ╚═╝ ██║██║  ██║██║███████╗ ██║ ╚═╝ ██║╚██████╔╝██║ ╚████║██║   ██║   ╚██████╔╝██║  ██║
  ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝╚══════╝ ╚═╝     ╚═╝ ╚═════╝ ╚═╝  ╚═══╝╚═╝   ╚═╝    ╚═════╝ ╚═╝  ╚═╝
{Fore.YELLOW}
                                 邮件监听服务控制台
{Fore.GREEN}
                            Mail Service Control Panel
{Fore.RED}
     Copyright (c) 2002-{datetime.now().strftime('%Y')}, Shenzhen Sunline Tech Co., Ltd, All Rights Reserved
{Style.RESET_ALL}
{Fore.GREEN} ═══════════════════════════════════════════════════════════════════════════════════════════
"""
        print(MAIL_MONITOR_LOGO)

    def show_menu(self):
        print(f" {Fore.CYAN}{EMOJI['MENULIST']} 系统操作菜单:{Style.RESET_ALL}")
        print(f" {Fore.YELLOW}{'─' * 40}{Style.RESET_ALL}")
        print(f" {Fore.GREEN}1{Style.RESET_ALL}. {EMOJI['MENU1']} {Fore.WHITE}启动邮件监听")
        print(f" {Fore.GREEN}2{Style.RESET_ALL}. {EMOJI['MENU2']} {Fore.WHITE}停止邮件监听")
        print(f" {Fore.GREEN}3{Style.RESET_ALL}. {EMOJI['MENU3']} {Fore.WHITE}显示监听状态")
        print(f" {Fore.GREEN}4{Style.RESET_ALL}. {EMOJI['MENU4']} {Fore.WHITE}查看服务配置")
        print(f" {Fore.GREEN}5{Style.RESET_ALL}. {EMOJI['MENU5']} {Fore.WHITE}退出程序")
        print(f" {Fore.YELLOW}{'─' * 40}{Style.RESET_ALL}")

    def start_monitor(self):
        if self.mail_monitor is None or not self.mail_monitor.is_running():
            try:
                self.mail_monitor = MailMonitor()
                self.mail_monitor.start()
                print(f"{EMOJI['SUCCESS']} 邮件监听服务已启动")
            except Exception as e:
                print(f"{EMOJI['ERROR']} 启动失败: {str(e)}")
        else:
            print(f"{EMOJI['WARN']} 监听服务已在运行中")
        input("\n按回车键继续...")

    def stop_monitor(self):
        if self.mail_monitor and self.mail_monitor.is_running():
            try:
                self.mail_monitor.stop()
                print(f"{EMOJI['SUCCESS']} 邮件监听服务已停止")
            except Exception as e:
                print(f"{EMOJI['ERROR']} 停止失败: {str(e)}")
        else:
            print(f"{EMOJI['WARN']} 监听服务未运行")
        input("\n按回车键继续...")

    def check_status(self):
        if self.mail_monitor:
            status = "运行中" if self.mail_monitor.is_running() else "已停止"
            print(f"📊 当前状态: {status}")
            if self.flask_thread and self.flask_thread.is_alive():
                print(f"🌐 Web服务: http://{self.flask_host}:{self.flask_port}")
        else:
            print("📊 当前状态: 未启动")
        input("\n按回车键继续...")

    def show_config(self):
        print("\n⚙️ 服务配置信息:")
        print("-" * 40)

        # 显示Web服务器配置
        print(f"Web服务器配置:")
        print(f"  - 主机: {self.flask_host}")
        print(f"  - 端口: {self.flask_port}")
        print(f"  - 状态: {'运行中' if (self.flask_thread and self.flask_thread.is_alive()) else '未启动'}")
        print("-" * 40)

        # 显示邮件监控配置
        if self.mail_monitor:
            config = self.mail_monitor.get_config()
            print("邮件监控配置:")
            for key, value in config.items():
                print(f"  - {key}: {value}")
        else:
            print("邮件监控: 未初始化")
        input("\n按回车键继续...")

    def run(self):
        # 启动Flask服务器
        if not self.flask_thread or not self.flask_thread.is_alive():
            self.flask_thread = threading.Thread(target=self.start_flask_server)
            self.flask_thread.daemon = True
            self.flask_thread.start()
            time.sleep(1)  # 等待服务器启动
        while self.running:
            self.clear_screen()
            # self.show_banner()
            self.show_logo()
            self.show_menu()

            choice = input(f" {Fore.CYAN}{EMOJI['ARROW']} 请输入选项 (1-5): ")

            if choice == '1':
                self.start_monitor()
            elif choice == '2':
                self.stop_monitor()
            elif choice == '3':
                self.check_status()
            elif choice == '4':
                self.show_config()
            elif choice == '5':
                if self.mail_monitor and self.mail_monitor.is_running():
                    self.mail_monitor.stop()
                print("👋 感谢使用，再见！")
                self.running = False
            else:
                print(f"{EMOJI['ERROR']} 无效的选项，请重试")
                time.sleep(1)

if __name__ == "__main__":
    try:
        ui = MailServUI()
        ui.run()
    except KeyboardInterrupt:
        print("\n程序被用户中断")
        sys.exit(0)
    except Exception as e:
        print(f"程序异常退出: {str(e)}")
        sys.exit(1) 