import os
import re
import sys

def extract_chinese(input_text):
    match = re.match(r"([^\d]+)", input_text)
    if match:
        return match.group(1)
    return input_text

def find_key_in_string(input_string, keyword):
    # 使用正则表达式查找包含指定名称并符合特定模式的文件名
    pattern = rf".*{keyword}.*\.pdf"  # keyword 参数使得查找更通用
    if re.search(pattern, input_string):
        return input_string
    return None  # 如果不匹配则返回 None

# 查找指定目录下包含姓名的PDF文件
def find_pdf_files(pdf_dir, find_name):
    pdf_path = None
    for root, dirs, files in os.walk(pdf_dir):
        for file in files:
            if find_key_in_string(file, find_name):
                pdf_path = os.path.join(root, file)
                break  # 如果找到一个文件，可以提前退出内层循环
        if pdf_path:  # 如果找到匹配的文件，跳出外层循环
            break
    return pdf_path

def get_script_directory():
    if getattr(sys, 'frozen', False):
        # 如果是打包后的可执行文件
        return os.path.dirname(sys.executable)
    else:
        # 如果是源代码脚本
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # 返回上一级目录（项目根目录）
        return os.path.dirname(current_dir)

if __name__ == '__main__':
    test = find_pdf_files(r'D:/project/share_tools/bidding_docx/social', '曹立丹')
    print(f'找到：{test}')