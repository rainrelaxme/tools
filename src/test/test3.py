from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER

document = Document()
paragraph = document.add_paragraph('\t制表符')
paragraph_format = paragraph.paragraph_format
tab_stops = paragraph_format.tab_stops
# tab_stop = tab_stops.add_tab_stop(Cm(5.0))  # 插入制表符
# print(tab_stop.position)  # 1800225
# print(tab_stop.position.cm)  # 5.000625
tab_stop = tab_stops.add_tab_stop(Cm(10.0), alignment=WD_TAB_ALIGNMENT.RIGHT, leader=WD_TAB_LEADER.DOTS)  # 右对齐，前导字符为点
# tab_stop = tab_stops.add_tab_stop(Cm(10.0), alignment=WD_TAB_ALIGNMENT.RIGHT, leader=WD_TAB_LEADER.DOTS)  # 右对齐，前导字符为点

document.save('test.docx')  # 保存

import json
import csv


class FileBasedTranslator:
    def __init__(self, glossary_file=None):
        self.glossary = {}
        if glossary_file:
            self.load_glossary(glossary_file)

    def load_glossary(self, file_path):
        """从文件加载词库"""
        if file_path.endswith('.json'):
            self._load_json_glossary(file_path)
        elif file_path.endswith('.csv'):
            self._load_csv_glossary(file_path)
        elif file_path.endswith('.txt'):
            self._load_txt_glossary(file_path)

    def _load_json_glossary(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            self.glossary = json.load(f)

    def _load_csv_glossary(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 2:
                    self.glossary[row[0].lower()] = row[1]

    def _load_txt_glossary(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                parts = line.strip().split('|')
                if len(parts) >= 2:
                    self.glossary[parts[0].lower()] = parts[1]

    def translate(self, text):
        text_lower = text.lower()
        if text_lower in self.glossary:
            return self.glossary[text_lower]

        # 调用翻译API
        return self._call_translation_api(text)

    def _call_translation_api(self, text):
        # 实际翻译API调用
        return f"翻译结果: {text}"

    def save_glossary(self, file_path):
        """保存词库到文件"""
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(self.glossary, f, ensure_ascii=False, indent=2)


# 使用示例
translator = FileBasedTranslator('glossary.json')