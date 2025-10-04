
from openai import OpenAI

from config.private import API_KEY


class Translator:
    def __init__(self):
        self.api_key = API_KEY
        self.base_url = "https://api.deepseek.com"

    def translate(self, text: str, language, display=False) -> str:
        # 构建DeepSeek请求
        client = OpenAI(api_key=self.api_key, base_url=self.base_url)

        messages = [
            {"role": "system",
             "content": f"你是一名工业领域翻译工作者，请进行{language}翻译，保持专业术语的准确性,只需返回翻译后的文本，不要添加任何额外的解释或注释。"
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
