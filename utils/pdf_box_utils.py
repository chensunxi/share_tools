import fitz
import re

def add_red_box_to_company_name(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    
    # 使用不同的文本提取方式
    text_dict = page.get_text("dict")
    blocks = text_dict["blocks"]
    
    # 存储当前处理的spans
    current_spans = []
    processed_texts = set()  # 使用集合来跟踪已处理的文本
    
    # 在每轮处理开始时重置processed_texts
    def reset_processed_texts():
        nonlocal processed_texts
        processed_texts = set()
    
    def check_and_draw_box(spans):
        """检查并绘制框"""
        if not spans:
            return False
            
        # 计算边界框
        x0 = min(s["bbox"][0] for s in spans)
        y0 = min(s["bbox"][1] for s in spans)
        x1 = max(s["bbox"][2] for s in spans)
        y1 = max(s["bbox"][3] for s in spans)
        
        # 获取完整文本
        full_text = "".join(s["text"].strip() for s in spans)
        
        print("\n=== 准备框选 ===")
        print(f"框选文本: {full_text}")
        print(f"框选区域: ({x0}, {y0}, {x1}, {y1})")
        
        # 绘制红色矩形框
        rect = fitz.Rect(x0-5, y0-5, x1+5, y1+5)
        page.draw_rect(rect, color=(1, 0, 0), width=2)
        
        return True
    
    def is_part_of_company_name(text):
        """检查文本是否是公司名称的一部分"""
        keywords = [
            "深圳", "市", "长亮", "科技",
            "股份", "有限", "公司",
            "分公", "分", "公司", "司"
        ]
        result = any(keyword in text for keyword in keywords)
        print(f"检查文本部分: {text} -> {result}")
        return result
    
    def is_complete_company_name_with_branch(text):
        """检查是否是完整的分公司名称"""
        result = ("深圳市长亮科技" in text and 
                "股份有限公司" in text and 
                "分公司" in text)
        print(f"检查分公司名称: {text} -> {result}")
        return result
    
    def is_simple_company_name(text):
        """检查是否是简单的公司名称（不含分公司）"""
        result = ("深圳市长亮科技" in text and 
                "股份有限公司" in text)
        print(f"检查简单公司名称: {text} -> {result}")
        return result
    
    def process_blocks(check_name_func):
        """处理文本块"""
        nonlocal current_spans
        
        print(f"\n=== 开始处理文本块 ===")
        print(f"检查函数: {check_name_func.__name__}")
        
        for block in blocks:
            if "lines" not in block:
                continue
            
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    
                    print(f"\n--- 处理新的span ---")
                    print(f"文本: {text}")
                    print(f"位置: {span['bbox']}")
                    
                    # 检查是否已处理过这个位置的文本
                    span_key = f"{text}_{span['bbox']}"
                    if span_key in processed_texts:
                        print(f"跳过重复文本")
                        continue
                    
                    # 先检查单个span是否包含完整的公司名称
                    print("检查是否是完整名称...")
                    if check_name_func(text):
                        print("找到完整名称！")
                        processed_texts.add(span_key)
                        if check_and_draw_box([span]):
                            return True
                    
                    # 如果不是完整名称，但包含公司名称的部分，则收集起来
                    print("检查是否是名称的一部分...")
                    if is_part_of_company_name(text):
                        print("是名称的一部分，添加到收集")
                        current_spans.append(span)
                        processed_texts.add(span_key)
                        
                        # 检查当前收集的spans是否构成完整名称
                        full_text = "".join(s["text"].strip() for s in current_spans)
                        print(f"当前累积文本: {full_text}")
                        
                        if check_name_func(full_text):
                            print("累积文本构成完整名称！")
                            if check_and_draw_box(current_spans):
                                return True
                    else:
                        print("不是名称的一部分，重置收集")
                        current_spans = []
        
        print("未找到匹配的名称")
        return False
    
    # 第一轮：查找带分公司的完整名称
    print("\n========== 第一轮：查找带分公司的完整名称 ==========")
    if process_blocks(is_complete_company_name_with_branch):
        output_path = pdf_path.replace('.pdf', '_with_box.pdf')
        doc.save(output_path, garbage=4, deflate=True)
        doc.close()
        print(f"\n已保存标注后的PDF到: {output_path}")
        return
    
    # 重置已处理文本的记录，开始第二轮查找
    reset_processed_texts()
    
    # 第二轮：查找不带分公司的公司名称
    print("\n========== 第二轮：查找不带分公司的公司名称 ==========")
    if process_blocks(is_simple_company_name):
        output_path = pdf_path.replace('.pdf', '_with_box.pdf')
        doc.save(output_path, garbage=4, deflate=True)
        doc.close()
        print(f"\n已保存标注后的PDF到: {output_path}")
        return
    
    # 如果没找到任何公司名称，关闭文档
    doc.close()
    print("\n未找到完整的公司名称")

if __name__ == "__main__":
    try:
        pdf_path = r"D:\project\share_tools\bidding_docx\test\社保证明--5.pdf"
        add_red_box_to_company_name(pdf_path)
    except Exception as e:
        print(f"处理PDF时出错: {str(e)}") 