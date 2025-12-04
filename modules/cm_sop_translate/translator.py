#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@Project : tools
@File    : translator.py
@Author  : Shawn
@Date    : 2025/10/11 9:31
@Info    : Description of this file

"""

import json
import logging
import os

from openai import OpenAI

from modules.common.log import setup_logger
from modules.cm_sop_translate.conf.conf import LOG_PATH
from modules.cm_sop_translate.conf.conf import DS_KEY, GLOSSARY

logger = setup_logger(log_dir=LOG_PATH, name='logs', level=logging.INFO)


class Translator:
    def __init__(self):
        self.api_key = DS_KEY
        self.base_url = "https://api.deepseek.com"
        self.glossary_folder = GLOSSARY['dir']

    def translate(self, text: str, language, display=False):
        # 先过滤是否在词库中
        glossary_res = self.translate_filter(text, language)
        if glossary_res:
            return glossary_res

        # 构建DeepSeek请求
        client = OpenAI(api_key=self.api_key, base_url=self.base_url)

        messages = [
            {"role": "system",
             "content": f"你是一名工业领域翻译工作者，请翻译成{language}，保持专业术语的准确性,只需返回翻译后的文本，不要添加任何额外的解释或注释。"
                        f"待翻译内容如下：{text}"
             }]

        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=messages,
            temperature=0.5,
            stream=False,  # 非流式输出
        )
        resp_text = response.choices[0].message.content
        if display:
            print(text, "-->", resp_text)
        return resp_text

    def translate_filter(self, text: str, language):
        """
        如果是指定的专有名词，在词库中的，则不再提交deepseek翻译。
        """
        try:
            glossary_folder = self.glossary_folder
            # if not os.path.isdir(glossary_folder):
            #     print(f"词库路径不存在，请检查{glossary_folder}，采用AI翻译")
            #     return False

            # if language in GLOSSARY['languages']:
            #     glossary = os.path.join(glossary_folder, GLOSSARY['languages'][language])
            # else:
            #     print(f"{language}该语言词库未配置，采用AI翻译。")
            #     return False

            glossary = os.path.join(glossary_folder, GLOSSARY['languages'][language])

            # if not os.path.isfile(glossary):
            #     print(f"词库文件不存在，请检查{glossary}，采用AI翻译。")
            #     return False

            with open(glossary, "r", encoding="utf-8") as f:
                dicts = json.load(f)

            text_lite = text.replace(" ", "").replace("　", "")

            if text_lite in dicts:
                res = dicts[text_lite]
                print(f"{text} --> {res}")
                return res

        except Exception as e:
            logger.error(f"词库异常{e}，将采用AI翻译")
            # return False

# 使用示例
# if __name__ == "__main__":
#     # input_file = "../../file/1.docx"
#     # output_file = "../../file/QMS-Z000-A00_Quality_Management_Manual_Bilingual.docx"
#     text = "为保证供应商及材料的来源,同时保证供应商开发、选择、评价和再评价、引入、退出的客观性、公正性、科学性，加强对供应商的日常管理和考核，促使其推动质量（含HSF）、环境、职业健康安全、有害物质、社会责任等体系的改进，确保提供产品的质量（含环境、职业健康安全、有害物质、社会责任等）以及交付、服务符合我公司、法律法规及客户的要求，促进我司产品质量、有害物质稳定提高，确保任何HSF采购产品没有污染或混杂的可能，特制订本程序。"
#     # 开始翻译
#     try:
#         translator = Translator()
#         translator.translate(text, language="英文")
#     except Exception as e:
#         print(f"翻译过程发生严重错误: {e}")

# if __name__ == "__main__":
#     translator = Translator()
#     translated_text = translator.translate_filter("文件编号", "英语")
#     print(translated_text)
