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
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 返回上一级目录（项目根目录）
    return os.path.dirname(current_dir)

def get_python_files(directory):
    """获取指定目录下所有的Python文件"""
    py_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.py'):
                # 获取相对于目录的模块路径
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
        platform_dist_dir = os.path.join(build_dir, 'dist', 'share_tools', 'windows')
        os.makedirs(platform_dist_dir, exist_ok=True)

        # 构建基本命令
        cmd = [
            'pyinstaller',
            '--onefile',
            '--noconsole',
            '--noupx',
            '--clean',
            '--name', 'share_tools',
            '--paths', root_dir,
            '--distpath', platform_dist_dir,
            '--workpath', os.path.join(build_dir, 'dist','share_tools', 'windows'),
            '--specpath', os.path.join(build_dir, 'dist', 'share_tools', 'windows'),
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

        print("Windows 版本打包完成!")
        print(f"可执行文件位置: {os.path.join(platform_dist_dir, 'share_tools.exe')}")
        return True

    except Exception as e:
        print(f"Windows 版本打包失败: {str(e)}")
        return False

def build_macos_docker(root_dir: str, build_dir: str) -> bool:
    """使用Docker构建macOS版本"""
    try:
        # 检查是否安装了Docker
        try:
            subprocess.run(['docker', '--version'], check=True, capture_output=True)
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("错误: 未安装Docker或Docker未运行。请安装并启动Docker后重试。")
            return False

        # 创建Dockerfile
        dockerfile_content = """FROM --platform=linux/amd64 joseluisq/rust-linux-darwin-builder:latest

# 设置环境变量
ENV PATH="/root/.cargo/bin:/usr/local/osxcross/bin:$PATH" \\
    CC=o64-clang \\
    CXX=o64-clang++ \\
    PYTHONIOENCODING=utf8 \\
    LANG=C.UTF-8 \\
    LC_ALL=C.UTF-8

# 安装 Python 和依赖
RUN apt-get update && apt-get install -y \\
    python3.8 \\
    python3-pip \\
    python3.8-dev \\
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . /app

# 安装 Python 依赖
RUN pip3 install -i https://pypi.tuna.tsinghua.edu.cn/simple --no-cache-dir \\
    pyinstaller \\
    pandas \\
    openpyxl \\
    cryptography

# 执行打包
RUN MACOSX_DEPLOYMENT_TARGET=10.15 pyinstaller --onefile --noconsole --noupx --clean \\
    --name share_tools \\
    --paths /app \\
    --hidden-import social_download \\
    --hidden-import utils \\
    --hidden-import gui \\
    --hidden-import bidding_docx \\
    --hidden-import social_desensitize \\
    --add-data "social_download:social_download" \\
    --add-data "utils:utils" \\
    --add-data "gui:gui" \\
    --add-data "bidding_docx:bidding_docx" \\
    --add-data "social_desensitize:social_desensitize" \\
    --add-data "resources:resources" \\
    --target-arch universal2 \\
    share_main.py
"""
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
                    print(f"\n开始构建macOS版本... (尝试 {retry_count + 1}/{max_retries})")
                    
                    # 尝试拉取基础镜像
                    print("拉取基础镜像...")
                    subprocess.run(
                        ['docker', 'pull', '--platform=linux/amd64', 'joseluisq/rust-linux-darwin-builder:latest'],
                        check=True,
                        env=env,
                        timeout=300
                    )
                    
                    # 构建应用镜像
                    print("构建应用镜像...")
                    subprocess.run(
                        [
                            'docker', 'build',
                            '--platform=linux/amd64',
                            '--network', 'host',
                            '--no-cache',
                            '-t', 'share-tools-builder',
                            '.'
                        ],
                        cwd=root_dir,
                        check=True,
                        env=env,
                        timeout=600
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
                ['docker', 'create', 'share-tools-builder'],
                capture_output=True,
                text=True,
                check=True,
                env=env
            ).stdout.strip()

            # 创建输出目录
            macos_dist_dir = os.path.join(build_dir, 'dist', 'share_tools', 'macos')
            os.makedirs(macos_dist_dir, exist_ok=True)

            # 从容器复制文件
            subprocess.run(
                ['docker', 'cp',
                 f'{container_id}:/app/dist/share_tools',
                 os.path.join(macos_dist_dir, 'share_tools')],
                check=True,
                env=env
            )

            # 复制资源文件
            resources_dir = os.path.join(root_dir, 'resources')
            if os.path.exists(resources_dir):
                shutil.copytree(
                    resources_dir,
                    os.path.join(macos_dist_dir, 'resources'),
                    dirs_exist_ok=True
                )

            # 清理
            subprocess.run(['docker', 'rm', container_id], check=True, env=env)
            subprocess.run(['docker', 'rmi', 'share-tools-builder'], check=True, env=env)

            print("macOS 版本打包完成!")
            print(f"可执行文件位置: {os.path.join(macos_dist_dir, 'share_tools')}")
            return True

        finally:
            # 清理Dockerfile
            if os.path.exists(dockerfile_path):
                os.remove(dockerfile_path)

    except Exception as e:
        print(f"macOS 版本打包失败: {str(e)}")
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
        main_script = os.path.join(root_dir, 'share_main.py')
        icon_file = os.path.join(root_dir, 'resources', 'icon.ico')
        
        # 清理旧的构建文件
        dir_path = os.path.join(build_dir, 'dist', 'share_tools')
        if os.path.exists(dir_path):
            shutil.rmtree(dir_path)
        
        # 检查主脚本是否存在
        if not os.path.exists(main_script):
            raise FileNotFoundError(f"主脚本文件不存在: {main_script}")

        # 设置隐藏导入
        hidden_imports = [
            'social_download',
            'utils',
            'gui',
            'bidding_docx',
            'social_desensitize'
        ]
        
        # 添加各目录下的所有模块
        for dir_name in ['social_download', 'utils', 'gui', 'bidding_docx', 'social_desensitize']:
            dir_path = os.path.join(root_dir, dir_name)
            if os.path.exists(dir_path):
                hidden_imports.extend([f'{dir_name}.{module}' 
                                    for module in get_python_files(dir_path)])

        # 设置数据文件
        data_files = []
        
        # 添加代码目录
        for dir_name in ['social_download', 'utils', 'gui', 'bidding_docx', 'social_desensitize']:
            dir_path = os.path.join(root_dir, dir_name)
            if os.path.exists(dir_path):
                data_files.append((dir_path, dir_name))
        
        # 添加资源文件
        resources_dir = os.path.join(root_dir, 'resources')
        if os.path.exists(resources_dir):
            data_files.append((resources_dir, 'resources'))

        # 构建Windows版本，临时屏蔽
        windows_success = build_windows(
            root_dir=root_dir,
            build_dir=build_dir,
            main_script=main_script,
            hidden_imports=hidden_imports,
            data_files=data_files,
            icon_file=icon_file
        )

        # 构建macOS版本
        # macos_success = build_macos_docker(root_dir, build_dir)

        # 打印最终结果
        print("\n打包结果汇总:")
        print("=" * 50)
        # print(f"Windows: {'成功' if windows_success else '失败'}")
        print(f"macOS: {'成功' if macos_success else '失败'}")
        print("=" * 50)

    except Exception as e:
        print(f"打包过程出错: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
