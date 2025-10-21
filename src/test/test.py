# from docx import Document
# import json
# from docx.oxml import OxmlElement
# from docx.oxml import OxmlElement
# from docx.oxml.ns import qn
#
# def table_line():
#     doc = Document()
#
#     # 创建3行1列的表格
#     table = doc.add_table(rows=3, cols=1)
#     table_rows = table.rows
#
#     # 设置各单元格样式
#     for row_id, row in enumerate(table.rows):
#         if row_id == 0:  # 首行
#             for cell in row.cells:
#                 set_cell_border(
#                     cell,
#                     top={"sz": 16, "val": "single", "color": "#000000"},  # 上边框加粗
#                     bottom={"sz": 12, "val": "none"},  # 底边无边框
#                     insideH={"sz": 12, "val": "single", "color": "#FFFFFF"}  # 隐藏内部线
#                 )
#         elif row_id == 1:  # 中间行
#             for cell in row.cells:
#                 set_cell_border(
#                     cell,
#                     top={"sz": 12, "val": "none"},
#                     bottom={"sz": 12, "val": "single"},
#                     insideH={"sz": 12, "val": "single", "color": "#FFFFFF"}
#                 )
#         else:  # 末行
#             for cell in row.cells:
#                 set_cell_border(
#                     cell,
#                     top={"sz": 12, "val": "none"},
#                     bottom={"sz": 16, "val": "single", "color": "#000000"},  # 下边框加粗
#                     insideH={"sz": 12, "val": "single", "color": "#FFFFFF"}
#                 )
#     doc.save('test.docx')
#
# def set_cell_border(cell, **kwargs):
#     tc = cell._tc
#     tcPr = tc.get_or_add_tcPr()
#     tcBorders = tcPr.first_child_found_in("w:tcBorders")
#
#     if tcBorders is None:
#         tcBorders = OxmlElement('w:tcBorders')
#         tcPr.append(tcBorders)
#
#     for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
#         edge_data = kwargs.get(edge)
#         if edge_data:
#             tag = f'w:{edge}'
#             element = tcBorders.find(qn(tag))
#             if element is None:
#                 element = OxmlElement(tag)
#                 tcBorders.append(element)
#
#             for key in ["sz", "val", "color", "space", "shadow"]:
#                 if key in edge_data:
#                     element.set(qn(f'w:{key}'), str(edge_data[key]))
#
#
# if __name__ == '__main__':
#     print("***********test************")
#     import win32com.client as win32
#     from win32com.client import constants
#     import os
#
#     doc_app = win32.gencache.EnsureDispatch('Word.Application')  # 打开word应用程序
#     doc_app.Visible = True
#
#     doc = doc_app.Documents.Add()
#     footer = doc.Sections(1).Footers(constants.wdHeaderFooterPrimary)
#     footer.Range.Text = ""
#     footer = doc.Sections(1).Footers(constants.wdHeaderFooterPrimary)
#     footer.LinkToPrevious = False
#     footer_rng = footer.Range
#     footer_rng.Text = "自动插入页码 "
#     footer.PageNumbers.Add(PageNumberAlignment=constants.wdAlignPageNumberRight, FirstPage=True)
#
#     # doc.Save('test.docx')
#
# def extract_header_footer_content(docx_path):
#     """
#     提取文档中所有页眉页脚的文本内容，返回结构化数据
#     """
#     doc = Document(docx_path)
#
#     result = {
#         'sections': [],
#         'all_headers': [],
#         'all_footers': []
#     }
#
#     for section_num, section in enumerate(doc.sections, 1):
#         section_data = {
#             'section_number': section_num,
#             'headers': {},
#             'footers': {}
#         }
#
#         # # 提取页眉
#         # header_types = [
#         #     ('default', section.header),
#         #     ('first_page', section.first_page_header),
#         #     ('even_page', section.even_page_header),
#         #     # ('odd_page', section.odd_page_header)
#         # ]
#         #
#         # for header_name, header in header_types:
#         #     if not header.is_linked_to_previous:
#         #         texts = [p.text.strip() for p in header.paragraphs if p.text.strip()]
#         #         if texts:
#         #             section_data['headers'][header_name] = texts
#         #             result['all_headers'].extend(texts)
#
#         # 提取页脚
#         footer_types = [
#             ('default', section.footer),
#             ('first_page', section.first_page_footer),
#             ('even_page', section.even_page_footer),
#             # ('odd_page', section.odd_page_footer)
#         ]
#
#         for footer_name, footer in footer_types:
#             if not footer.is_linked_to_previous:
#                 texts = [p.text.strip() for p in footer.paragraphs if p.text.strip()]
#                 if texts:
#                     section_data['footers'][footer_name] = texts
#                     result['all_footers'].extend(texts)
#
#         result['sections'].append(section_data)
#
#     return result
#
#
# def main1():
#     # 使用示例
#     file_path = r"D:\Code\Project\tools\data\13. 封面模板.docx"  # 替换为你的文件路径
#     content = extract_header_footer_content(file_path)
#
#     # print("提取到的页眉内容:")
#     # for i, header in enumerate(set(content['all_headers']), 1):
#     #     print(f"  {i}. {header}")
#
#     print("\n提取到的页脚内容:")
#     for i, footer in enumerate(set(content['all_footers']), 1):
#         print(f"  {i}. {footer}")
#
#     # 打印详细结构
#     print(f"\n详细结构:\n{json.dumps(content, indent=2, ensure_ascii=False)}")
#
#
# def get_header_content(doc):
#     """
#     获取页眉内容
#     """
#     pass
#
#
# def get_footer_content(doc):
#     """
#     获取页脚内容
#     """
#     footer_content = {
#         'type': 'paragraph',
#         'index': '',
#         'element_index': '',
#         'section': 'footer',
#         'flag': 'footer',
#         'text': '',
#         'para_format': '',
#         'runs': ''
#     }
#     for index, section in enumerate(doc.sections):
#         footer = section.footer
#         if footer:
#             footer_data = {
#                 'type': 'footer',
#                 'index': index,
#                 'element_index': index,
#                 'flag': 'footer',
#                 'content': [],
#             }
#             # 先处理所有段落，记录它们在文档中的位置
#             para_positions = {}
#             for para_index, para in enumerate(doc.paragraphs):
#                 # 获取段落在XML中的位置
#                 para_xml = para._element
#                 parent = para_xml.getparent()
#                 if parent is not None:
#                     position = list(parent).index(para_xml)
#                     para_positions[(position, para.text)] = para_index
#
#             # 处理所有表格，记录它们在文档中的位置
#             table_positions = {}
#             for table_index, table in enumerate(doc.tables):
#                 table_xml = table._element
#                 parent = table_xml.getparent()
#                 if parent is not None:
#                     position = list(parent).index(table_xml)
#                     table_positions[position] = table_index
#
#             # 按照文档顺序处理所有元素
#             all_elements = []
#
#             # 收集所有段落位置
#             for (pos, text), para_index in para_positions.items():
#                 all_elements.append(('paragraph', pos, para_index, text))
#
#             # 收集所有表格位置
#             for pos, table_index in table_positions.items():
#                 all_elements.append(('table', pos, table_index, None))
#
#             # 按位置排序
#             all_elements.sort(key=lambda x: x[1])
#
#             # 按排序后的顺序处理元素
#             for element_type, position, index, text in all_elements:
#                 if element_type == 'paragraph':
#                     para = doc.paragraphs[index]
#                     para_data = {
#                         'type': 'paragraph',
#                         'index': position,
#                         'element_index': index,
#                         'section': 'body',
#                         'flag': '',
#                         'text': para.text,
#                         'para_format': self.get_paragraph_format(para),
#                         'runs': self.get_run_format(para)
#                     }
#                     data.append(para_data)
#
#                 elif element_type == 'table':
#                     table = doc.tables[index]
#                     table_data = {
#                         'type': 'table',
#                         'index': position,
#                         'element_index': index,
#                         'section': 'body',
#                         'flag': '',
#                         'table_alignment': WD_TABLE_ALIGNMENT.CENTER,  # 表格居中，非内容居中
#                         'rows': self.get_table_content(table),
#                         'cols': len(table.columns),
#                     }
#                     data.append(table_data)
#
#
