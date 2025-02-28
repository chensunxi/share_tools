from paddleocr import PaddleOCR
import re

class OcrExtractorUtils:
    def __init__(self):
        # 初始化PaddleOCR
        self.ocr = PaddleOCR(use_angle_cls=True, lang="ch", use_gpu=False, show_log=False)
        # 定义要提取的字段
        self.fields = {
            "学校名称": "",
            "专业": ""
        }

    def extract_field_value(self, text, field):
        # 根据关键字提取学信网截图的文字
        pattern = fr'{field}[:：]\s*([^\n]+)'
        match = re.search(pattern, text)
        if match:
            value = match.group(1).strip()
            return value
        return None
    
    staticmethod
    def extract_info(self, image_path):
        try:
            # OCR识别
            result = self.ocr.ocr(image_path)

            # 合并所有识别的文本
            full_text = ""
            for line in result:
                for word_info in line:
                    text = word_info[1][0]  # 获取识别的文字
                    # confidence = word_info[1][1]  # 获取置信度
                    full_text += text + "\n"

            # 提取各个字段的值
            extracted_info = {}
            for field in self.fields.keys():
                value = self.extract_field_value(full_text, field)
                if value:
                    extracted_info[field] = value

            return {
                'success': True,
                'data': extracted_info,
                'full_text': full_text
            }

        except Exception as e:
            return {
                'success': False,
                'message': f'处理失败: {str(e)}'
            }

if __name__ == "__main__":
    # 创建提取器实例
    extractor = EducationInfoExtractor()

    # 图片路径
    image_path = r"D:\project\share_tools\bidding_assistant\files\39887\china_edu.png"

    # 提取信息
    result = extractor.extract_info(image_path)

    if result['success']:
        extracted_info = result['data']

