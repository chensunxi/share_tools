from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import json
import os
from .mail_process import AutoReplyMail
import threading
import traceback
import datetime

app = Flask(__name__)
CORS(app)  # 允许跨域请求

# 修改配置文件路径的定义
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_FILE = os.path.join(BASE_DIR, 'resources', 'mail_server_cfg.json')

# 确保 resources 目录存在
os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)

monitor_thread = None
auto_reply = None

def load_config():
    """加载配置"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"加载配置文件失败: {str(e)}")
    
    # 返回默认配置
    return {
        'recv_monitor_mail': [],
        'send_copy_mail': [],
        'auth_enter_mail': [],
        'keywords': '',
        'auto_reply_text': """您好！

我们已收到您发送的入场资料，我们会尽快处理。
如有任何问题，请随时与我们联系。

此邮件为自动回复，请勿直接回复。

祝好！""",
        'monitor_mode': 'realtime',
        'schedule_time': '09:00'
    }

@app.route('/api/config', methods=['POST'])
def save_config():
    """保存配置"""
    try:
        config = request.get_json()
        
        # 验证配置格式
        if not isinstance(config, dict):
            return jsonify({'status': 'error', 'message': '无效的配置格式'})
        
        # 确保目录存在
        os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
        
        # 确保配置包含所有必要字段
        if 'auth_enter_mail' not in config:
            config['auth_enter_mail'] = []
            
        # 加载现有配置，以保留可能未在新配置中包含的字段
        existing_config = load_config()
        
        # 特别处理recv_monitor_mail，确保erpuser和erppwd字段被保留
        if 'recv_monitor_mail' in config:
            # 创建一个邮箱到配置的映射，方便查找
            existing_mail_map = {mail_cfg['email']: mail_cfg for mail_cfg in existing_config.get('recv_monitor_mail', [])}
            
            for mail_cfg in config['recv_monitor_mail']:
                email = mail_cfg.get('email')
                # 如果新配置中没有erpuser或erppwd字段，但现有配置中有，则保留现有值
                if email in existing_mail_map:
                    if 'erpuser' not in mail_cfg and 'erpuser' in existing_mail_map[email]:
                        mail_cfg['erpuser'] = existing_mail_map[email]['erpuser']
                    if 'erppwd' not in mail_cfg and 'erppwd' in existing_mail_map[email]:
                        mail_cfg['erppwd'] = existing_mail_map[email]['erppwd']
        
        # 保存到文件
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        
        # 更新全局配置
        global auto_reply
        if auto_reply:
            auto_reply.update_config(config)
        
        return jsonify({'status': 'success'})
    except Exception as e:
        error_msg = f"保存配置失败: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/config', methods=['GET'])
def get_config():
    """获取配置"""
    try:
        config = load_config()
        return jsonify({
            'status': 'success',
            'data': config
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        })

def run_monitor(config):
    """运行监控"""
    global auto_reply
    try:
        # 从新格式的配置中提取邮箱信息
        recv_monitor_mail = config.get('recv_monitor_mail', [])
        if not recv_monitor_mail:
            raise ValueError("未配置邮箱账号")
            
        auto_reply = AutoReplyMail(
            recv_monitor_mail=recv_monitor_mail,
            send_copy_mail=config.get('send_copy_mail', []),
            auto_reply_text=config.get('auto_reply_text', """您好！

我们已收到您发送的入场资料，我们会尽快处理。
如有任何问题，请随时与我们联系。

此邮件为自动回复，请勿直接回复。

祝好！"""),
            auth_enter_mail=config.get('auth_enter_mail', []),
            keywords=config.get('keywords', '')
        )
        
        if config.get('monitor_mode', 'realtime') == 'realtime':
            auto_reply.start_monitoring(mode='realtime')
        else:
            auto_reply.start_monitoring(
                mode='scheduled', 
                schedule_time=config.get('schedule_time', '09:00')
            )
    except Exception as e:
        print(f"监控运行出错: {str(e)}")
        print(f"错误详情: {traceback.format_exc()}")

def update_monitor_status(status):
    """更新监控状态到配置文件"""
    try:
        config = load_config()
        config['monitor_status'] = status
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"更新监控状态失败: {str(e)}")

@app.route('/api/monitor/start', methods=['POST'])
def start_monitor():
    """启动监控"""
    global monitor_thread
    try:
        # 停止现有监控
        if monitor_thread and monitor_thread.is_alive():
            if auto_reply:
                auto_reply.stop_monitoring()
            monitor_thread.join()
            
        # 启动新监控
        config = load_config()
        monitor_thread = threading.Thread(target=run_monitor, args=(config,))
        monitor_thread.daemon = True
        monitor_thread.start()
        update_monitor_status('running')  # 更新状态
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/monitor/stop', methods=['POST'])
def stop_monitor():
    """停止监控"""
    global monitor_thread, auto_reply
    try:
        if auto_reply:
            auto_reply.stop_monitoring()
        if monitor_thread and monitor_thread.is_alive():
            monitor_thread.join()
        update_monitor_status('stopped')  # 更新状态
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/monitor/status', methods=['GET'])
def get_monitor_status():
    """获取监控状态"""
    try:
        config = load_config()
        status = config.get('monitor_status', 'stopped')
        if monitor_thread and monitor_thread.is_alive() and auto_reply:
            status = 'running'
        return jsonify({'status': status})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/mail_monitor/download_report', methods=['GET'])
def download_report():
    """下载邮件监控报表"""
    try:
        # 获取请求参数中的年月
        year_month = request.args.get('year_month')
        if not year_month:
            # 如果没有指定年月，使用当前年月
            year_month = datetime.datetime.now().strftime('%Y%m')
            
        # 构建Excel文件路径
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        excel_file = os.path.join(base_dir, 'download', 'mail_attach', f'mail_report_{year_month}.xlsx')
        
        if not os.path.exists(excel_file):
            return jsonify({
                'status': 'error',
                'code': 404,
                'message': f'未找到 {year_month} 的报表文件'
            }), 404
        
        # 返回Excel文件
        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'mail_report_{year_month}.xlsx'
        )
        
    except Exception as e:
        return jsonify({
            'status': 'error',
            'code': 500,
            'message': f'下载报表失败: {str(e)}'
        }), 500

if __name__ == '__main__':
    try:
        # 禁用自动重载和调试模式
        app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
        app.config['TEMPLATES_AUTO_RELOAD'] = False
        
        print("正在启动服务器...")
        # 使用 waitress 作为生产级 WSGI 服务器
        from waitress import serve
        print("服务器已启动，访问 http://localhost:5000")
        serve(app, host='localhost', port=5000, threads=4)
        
    except Exception as e:
        print(f"服务器启动失败: {str(e)}")
        traceback.print_exc() 