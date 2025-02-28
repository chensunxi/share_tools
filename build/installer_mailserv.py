import os
import subprocess
import sys
import shutil
import glob
import platform
import time
from typing import List, Tuple

def get_project_root():
    """获取项目根目录"""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.dirname(current_dir)

def get_python_files(directory):
    """获取指定目录下所有的Python文件"""
    py_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.py'):
                rel_path = os.path.relpath(root, directory)
                module_path = os.path.join(rel_path, file[:-3]).replace(os.sep, '.')
                if module_path.startswith('.'):
                    module_path = module_path[1:]
                if module_path:
                    py_files.append(module_path)
    return py_files

def build_windows(root_dir: str, build_dir: str, main_script: str, 
                 hidden_imports: List[str], data_files: List[Tuple[str, str]], 
                 icon_file: str = None) -> bool:
    """构建Windows版本"""
    try:
        platform_dist_dir = os.path.join(build_dir, 'dist', 'mailserv','windows')
        os.makedirs(platform_dist_dir, exist_ok=True)

        # 构建基本命令
        cmd = [
            'pyinstaller',
            '--onefile',
            '--console',
            '--noupx',
            '--clean',
            '--name', 'mail_monitor_service',
            '--paths', root_dir,
            '--distpath', platform_dist_dir,
            '--workpath', os.path.join(build_dir, 'build', 'windows'),
            '--specpath', os.path.join(build_dir, 'spec', 'windows'),
        ]

        # 添加图标
        if icon_file and os.path.exists(icon_file):
            cmd.extend(['--icon', icon_file])

        # 添加所有隐藏导入
        for hidden_import in hidden_imports:
            cmd.extend(['--hidden-import', hidden_import])

        # 添加数据文件
        for src, dst in data_files:
            if os.path.exists(src):
                cmd.extend(['--add-data', f'{src};{dst}'])

        # 添加主脚本
        cmd.append(main_script)

        # 执行打包命令
        print("\n开始打包 Windows 版本...")
        print(f"执行命令: {' '.join(cmd)}")
        subprocess.run(cmd, check=True, shell=True)

        # 复制配置文件到输出目录
        config_src = os.path.join(root_dir, 'resources', 'mail_server_cfg.json')
        config_dst = os.path.join(platform_dist_dir, 'resources', 'mail_server_cfg.json')
        if os.path.exists(config_src):
            os.makedirs(os.path.dirname(config_dst), exist_ok=True)
            shutil.copy2(config_src, config_dst)

        print("Windows 版本打包完成!")
        print(f"可执行文件位置: {os.path.join(platform_dist_dir, 'mail_monitor_serv.exe')}")
        return True

    except Exception as e:
        print(f"Windows 版本打包失败: {str(e)}")
        return False

def build_linux_docker(root_dir: str, build_dir: str) -> bool:
    """使用Docker构建Linux版本"""
    try:
        # 检查是否安装了Docker
        try:
            subprocess.run(['docker', '--version'], check=True, capture_output=True)
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("错误: 未安装Docker或Docker未运行。请安装并启动Docker后重试。")
            return False

        # 创建Dockerfile
        dockerfile_content = '''
FROM ccr.ccs.tencentyun.com/library/centos:7

# 设置环境变量
ENV LANG=en_US.UTF-8
ENV LC_ALL=en_US.UTF-8

# 配置阿里云 yum 源
RUN curl -o /etc/yum.repos.d/CentOS-Base.repo http://mirrors.aliyun.com/repo/Centos-7.repo && \
    curl -o /etc/yum.repos.d/epel.repo http://mirrors.aliyun.com/repo/epel-7.repo && \
    yum clean all && \
    yum makecache

# 安装基础工具
RUN yum install -y gcc gcc-c++ make wget curl && \
    yum clean all

# 下载并安装 Python 3.8.12
RUN cd /tmp && \
    wget https://www.python.org/ftp/python/3.8.12/Python-3.8.12.tgz && \
    tar -xzf Python-3.8.12.tgz && \
    cd Python-3.8.12 && \
    ./configure --prefix=/usr/local/python3 && \
    make && \
    make install && \
    ln -sf /usr/local/python3/bin/python3 /usr/bin/python3 && \
    ln -sf /usr/local/python3/bin/pip3 /usr/bin/pip3 && \
    cd /tmp && \
    rm -rf Python-3.8.12*

# 配置 pip 源为 USTC
RUN mkdir -p ~/.pip && \
    echo '[global]' > ~/.pip/pip.conf && \
    echo 'index-url = https://mirrors.ustc.edu.cn/pypi/web/simple' >> ~/.pip/pip.conf && \
    echo 'trusted-host = mirrors.ustc.edu.cn' >> ~/.pip/pip.conf

# 安装 Python 依赖
RUN pip3 install --no-cache-dir pyinstaller==4.5.1 psutil==5.9.0 requests==2.27.1

# 设置工作目录
WORKDIR /app

# 复制源代码
COPY . .

# 构建参数
ENV BUILD_PARAMS="-F --clean -y"
'''
        dockerfile_path = os.path.join(root_dir, 'Dockerfile')
        with open(dockerfile_path, 'w', encoding='utf-8') as f:
            f.write(dockerfile_content)

        try:
            # 设置环境变量来配置Docker
            env = os.environ.copy()
            env['DOCKER_BUILDKIT'] = '1'
            env['DOCKER_CLI_EXPERIMENTAL'] = 'enabled'
            
            # 构建Docker镜像，添加重试机制
            max_retries = 3
            retry_count = 0
            while retry_count < max_retries:
                try:
                    print(f"\n开始构建Linux版本... (尝试 {retry_count + 1}/{max_retries})")
                    
                    # 构建应用镜像
                    print("构建应用镜像...")
                    subprocess.run(
                        [
                            'docker', 'build',
                            '--platform', 'linux/amd64',
                            '--no-cache',
                            '--pull',
                            '-t', 'mail-monitor-builder',
                            '.'
                        ],
                        cwd=root_dir,
                        check=True,
                        env=env,
                        timeout=1200  # 增加超时时间到20分钟
                    )
                    break
                    
                except subprocess.TimeoutExpired:
                    print(f"构建超时，正在重试... ({retry_count + 1}/{max_retries})")
                    retry_count += 1
                    if retry_count >= max_retries:
                        raise Exception("构建多次超时，请检查网络连接")
                    time.sleep(5)
                    
                except subprocess.CalledProcessError as e:
                    print(f"构建失败，正在重试... ({retry_count + 1}/{max_retries})")
                    print(f"错误信息: {e}")
                    retry_count += 1
                    if retry_count >= max_retries:
                        raise
                    time.sleep(5)

            # 创建临时容器并复制构建结果
            container_id = subprocess.run(
                ['docker', 'create', 'mail-monitor-builder'],
                capture_output=True,
                text=True,
                check=True,
                env=env
            ).stdout.strip()

            # 创建输出目录
            linux_dist_dir = os.path.join(build_dir, 'dist', 'mailsrv', 'linux')
            os.makedirs(linux_dist_dir, exist_ok=True)

            # 从容器复制文件
            subprocess.run(
                ['docker', 'cp',
                 f'{container_id}:/app/dist/mail_monitor_service',
                 os.path.join(linux_dist_dir, 'mail_monitor_service')],
                check=True,
                env=env
            )

            # 复制配置文件
            config_src = os.path.join(root_dir, 'resources', 'mail_server_cfg.json')
            config_dst = os.path.join(linux_dist_dir, 'resources', 'mail_server_cfg.json')
            if os.path.exists(config_src):
                os.makedirs(os.path.dirname(config_dst), exist_ok=True)
                shutil.copy2(config_src, config_dst)

            # 清理
            subprocess.run(['docker', 'rm', container_id], check=True, env=env)
            subprocess.run(['docker', 'rmi', 'mail-monitor-builder'], check=True, env=env)

            print("Linux 版本打包完成!")
            print(f"可执行文件位置: {os.path.join(linux_dist_dir, 'mail_monitor_service')}")
            return True

        finally:
            # 清理Dockerfile
            if os.path.exists(dockerfile_path):
                os.remove(dockerfile_path)

    except Exception as e:
        print(f"Linux 版本打包失败: {str(e)}")
        return False

def main():
    try:
        # 检查是否在Windows上运行
        if not platform.system().lower() == 'windows':
            raise EnvironmentError("此打包脚本需要在 Windows 平台上运行")

        # 获取项目根目录和构建目录
        root_dir = get_project_root()
        build_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 设置文件路径
        main_script = os.path.join(root_dir, 'mail_monitor', 'run_mail_serv.py')
        icon_file = os.path.join(root_dir, 'resources', 'icon.ico')
        
        # 清理旧的构建文件
        dir_path = os.path.join(build_dir, 'dist', 'mailserv')
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path)
        
        # 检查主脚本是否存在
        if not os.path.exists(main_script):
            raise FileNotFoundError(f"主脚本文件不存在: {main_script}")

        # 设置隐藏导入
        hidden_imports = [
            'mail_monitor',
            'utils',
            'colorama',
            'waitress',
            'flask',
            'flask_cors',
            'pandas',
            'openpyxl',
            'cryptography',
            'bs4',
            'schedule'
        ]
        
        # 添加模块导入
        mail_monitor_dir = os.path.join(root_dir, 'mail_monitor')
        utils_dir = os.path.join(root_dir, 'utils')
        hidden_imports.extend([f'mail_monitor.{module}' 
                             for module in get_python_files(mail_monitor_dir)])
        hidden_imports.extend([f'utils.{module}' 
                             for module in get_python_files(utils_dir)])

        # 设置数据文件
        data_files = [
            (os.path.join(root_dir, "mail_monitor"), "mail_monitor"),
            (os.path.join(root_dir, "utils"), "utils"),
            (os.path.join(root_dir, "resources"), "resources"),
        ]

        # 构建Windows版本，临时屏蔽
        # windows_success = build_windows(
        #     root_dir=root_dir,
        #     build_dir=build_dir,
        #     main_script=main_script,
        #     hidden_imports=hidden_imports,
        #     data_files=data_files,
        #     icon_file=icon_file
        # )

        # 构建Linux版本
        linux_success = build_linux_docker(root_dir, build_dir)

        # 打印最终结果
        print("\n打包结果汇总:")
        print("=" * 50)
        # print(f"Windows: {'成功' if windows_success else '失败'}")
        print(f"Linux: {'成功' if linux_success else '失败'}")
        print("=" * 50)

    except Exception as e:
        print(f"打包过程出错: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main() 