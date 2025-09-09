from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import datetime
import os
import re


def get_content(input_path):
    """
    读取word文件的内容，返回位置、内容、格式，记录在dataform中
    """
    doc = Document(input_path)
    data = []

    # 处理段落
    for para_index, para in enumerate(doc.paragraphs):

        # 获取段落格式
        para_format = get_paragraph_format(para)

        # 获取运行(run)级别的格式
        runs_data = []
        for run in para.runs:
            run_format = get_run_format(run)
            runs_data.append(run_format)

        data.append({
            'type': 'paragraph',
            'index': para_index,
            'text': para.text,
            'para_format': para_format,
            'runs': runs_data
        })

    # 处理表格
    # for table_index, table in enumerate(doc.tables):
    #     table_data = {
    #         'type': 'table',
    #         'index': table_index,
    #         'rows': [],
    #         'cols': len(table.columns)
    #     }
    #
    #     for row_index, row in enumerate(table.rows):
    #         row_data = {'cells': []}
    #         for cell_index, cell in enumerate(row.cells):
    #             if cell.text.strip():
    #                 cell_content = {
    #                     'text': cell.text,
    #                     'paragraphs': []
    #                 }
    #
    #                 # 处理单元格中的段落
    #                 for cell_para_index, cell_para in enumerate(cell.paragraphs):
    #                     if cell_para.text.strip():
    #                         cell_para_format = get_paragraph_format(cell_para)
    #
    #                         cell_runs_data = []
    #                         for run in cell_para.runs:
    #                             if run.text.strip():
    #                                 run_format = get_run_format(run)
    #                                 cell_runs_data.append(run_format)
    #
    #                         cell_content['paragraphs'].append({
    #                             'para_index': cell_para_index,
    #                             'text': cell_para.text,
    #                             'para_format': cell_para_format,
    #                             'runs': cell_runs_data
    #                         })
    #
    #                 row_data['cells'].append(cell_content)
    #
    #         table_data['rows'].append(row_data)
    #
    #     data.append(table_data)
    return data


def get_paragraph_format(paragraph):
    """提取段落格式信息"""
    return {
        'style': paragraph.style.name if paragraph.style else 'Normal',
        'alignment': str(paragraph.alignment) if paragraph.alignment else None,
        'space_before': paragraph.paragraph_format.space_before,
        'space_after': paragraph.paragraph_format.space_after,
        'line_spacing': paragraph.paragraph_format.line_spacing,
        'first_line_indent': paragraph.paragraph_format.first_line_indent,
        'left_indent': paragraph.paragraph_format.left_indent
    }


def get_run_format(run):
    """提取运行文本格式信息"""
    return {
        'text': run.text,
        'bold': run.bold,
        'italic': run.italic,
        'underline': run.underline,
        'font_size': run.font.size.pt if run.font.size else None,
        'font_name': run.font.name,
        'font_color': get_rgb_color(run.font.color.rgb) if run.font.color and run.font.color.rgb else None,
        'font_color_theme': run.font.color.theme_color if run.font.color else None
    }


def get_rgb_color(rgb_value):
    """将RGB颜色值转换为十六进制字符串"""
    if rgb_value:
        return f"#{rgb_value:06x}"
    return None


def create_new_document(content_data, output_path, translations=None):
    """
    根据记录的内容和格式生成新的Word文档
    """
    doc = Document()

    for item in content_data:
        if item['type'] == 'paragraph':
            # 创建新段落
            paragraph = doc.add_paragraph()

            # 应用段落格式
            apply_paragraph_format(paragraph, item['para_format'])

            # 添加运行文本
            for run_data in item['runs']:
                run = paragraph.add_run(run_data['text'])
                apply_run_format(run, run_data)

            # 如果需要显示翻译
            # if translations and item['text'] in translations:
            #     add_translation_paragraph(doc, item['text'], translations[item['text']])

        # elif item['type'] == 'table':
        #     # 创建表格
        #     table = doc.add_table(rows=len(item['rows']), cols=item['cols'])
        #     table.style = 'Table Grid'
        #
        #     # 填充表格内容
        #     for row_idx, row_data in enumerate(item['rows']):
        #         for cell_idx, cell_data in enumerate(row_data['cells']):
        #             if cell_idx < len(table.rows[row_idx].cells):
        #                 cell = table.rows[row_idx].cells[cell_idx]
        #
        #                 # 清空默认段落
        #                 for paragraph in cell.paragraphs:
        #                     p = paragraph._element
        #                     p.getparent().remove(p)
        #
        #                 # 添加内容到单元格
        #                 if cell_data['paragraphs']:
        #                     for para_data in cell_data['paragraphs']:
        #                         cell_para = cell.add_paragraph()
        #                         apply_paragraph_format(cell_para, para_data['para_format'])
        #
        #                         for run_data in para_data['runs']:
        #                             run = cell_para.add_run(run_data['text'])
        #                             apply_run_format(run, run_data)
        #                 else:
        #                     # 如果没有详细的段落信息，只添加文本
        #                     cell.text = cell_data['text']

    # 保存文档
    doc.save(output_path)
    print(f"新文档已保存到: {output_path}")


def apply_paragraph_format(paragraph, format_info):
    """应用段落格式"""
    # 使用安全的get方法访问字典，避免KeyError
    if format_info.get('alignment'):
        alignment_map = {
            'CENTER': WD_ALIGN_PARAGRAPH.CENTER,
            'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT,
            'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY,
            'LEFT': WD_ALIGN_PARAGRAPH.LEFT
        }
        alignment = format_info['alignment'].split('.')[-1] if '.' in format_info['alignment'] else format_info[
            'alignment']
        paragraph.alignment = alignment_map.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)

    # 使用get方法安全访问，如果键不存在则返回None
    space_before = format_info.get('space_before')
    space_after = format_info.get('space_after')
    line_spacing = format_info.get('line_spacing')
    first_line_indent = format_info.get('first_line_indent')
    left_indent = format_info.get('left_indent')

    if space_before is not None:
        paragraph.paragraph_format.space_before = space_before
    if space_after is not None:
        paragraph.paragraph_format.space_after = space_after
    if line_spacing is not None:
        paragraph.paragraph_format.line_spacing = line_spacing
    if first_line_indent is not None:
        paragraph.paragraph_format.first_line_indent = first_line_indent
    if left_indent is not None:
        paragraph.paragraph_format.left_indent = left_indent


def apply_run_format(run, run_data):
    """应用运行文本格式"""
    # 使用get方法安全访问
    run.bold = run_data.get('bold', False)
    run.italic = run_data.get('italic', False)
    run.underline = run_data.get('underline', False)

    font_size = run_data.get('font_size')
    if font_size:
        run.font.size = Pt(font_size)

    font_name = run_data.get('font_name')
    if font_name:
        run.font.name = font_name

    font_color = run_data.get('font_color')
    if font_color:
        try:
            # 将十六进制颜色转换为RGB
            hex_color = font_color.lstrip('#')
            rgb = tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
            run.font.color.rgb = RGBColor(*rgb)
        except:
            pass


def add_translation_paragraph(doc, original_text, translations):
    """添加翻译段落"""
    if 'english' in translations:
        para = doc.add_paragraph()
        para.add_run("English: ").bold = True
        para.add_run(translations['english'])

    if 'vietnamese' in translations:
        para = doc.add_paragraph()
        para.add_run("Vietnamese: ").bold = True
        para.add_run(translations['vietnamese'])

    doc.add_paragraph()  # 添加空行分隔


# def create_simple_document(content_data, output_path):
#     """
#     创建简化版文档，只复制文本内容
#     """
#     doc = Document()
#
#     for item in content_data:
#         if item['type'] == 'paragraph':
#             para = doc.add_paragraph(item['text'])
#
#         elif item['type'] == 'table':
#             # 创建表格
#             table = doc.add_table(rows=len(item['rows']), cols=item['cols'])
#             table.style = 'Table Grid'
#
#             for row_idx, row_data in enumerate(item['rows']):
#                 for cell_idx, cell_data in enumerate(row_data['cells']):
#                     if cell_idx < len(table.rows[row_idx].cells):
#                         table.rows[row_idx].cells[cell_idx].text = cell_data['text']
#
#     doc.save(output_path)
#     print(f"简化文档已保存到: {output_path}")


# 使用示例
if __name__ == "__main__":
    datetime = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    print(f"********************start at {datetime}********************\n")
    input_file = "../../../file/1.C2LG-001-000-A08 供应商管理程序 - 副本.docx"
    output_file = f"../../../file/temp/output-{datetime}.docx"

    # 1. 读取原文档内容及格式
    content_data = get_content(input_file)
    print(f"总共提取了 {len(content_data)} 个内容块")

    # 2. 创建新文档
    create_new_document(content_data, output_file)

    # if os.path.exists(input_file):
    #     try:
    #         # 1. 读取原始文档内容
    #         content_data = get_content(input_file)
    #         print(f"总共提取了 {len(content_data)} 个内容块")
    #
    #         # 2. 尝试创建保持原格式的新文档
    #         # create_new_document(content_data, output_file)
    #
    #     except Exception as e:
    #         print(f"创建格式文档时出错: {e}")
    #         print("尝试创建简化版文档...")
    #         # 3. 如果出错，创建简化版文档
    #         # create_simple_document(content_data, simple_output)
    #
    # else:
    #     print(f"文件 {input_file} 不存在")
    print("********************END********************\n")