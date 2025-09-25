import os
from win32com import client as wc


def doc_to_docx(doc_path, docx_path=None):
    """
    将doc文件转换为docx格式

    Args:
        doc_path: 输入的doc文件路径
        docx_path: 输出的docx文件路径，默认为None（自动生成）
    """
    # 如果未指定输出路径，自动生成
    if docx_path is None:
        docx_path = doc_path.replace('.doc', '.docx')

    # 启动Word应用程序
    word = wc.Dispatch('Word.Application')
    word.Visible = False  # 不显示Word界面

    try:
        # 打开doc文档
        doc = word.Documents.Open(doc_path)
        # 另存为docx格式
        doc.SaveAs(docx_path, FileFormat=16)  # 16表示docx格式
        doc.Close()
        print(f"转换成功: {doc_path} -> {docx_path}")
        return True
    except Exception as e:
        print(f"转换失败: {e}")
        return False
    finally:
        word.Quit()


# 使用示例
old_doc = r'D:\Code\Project\tools\data\input\1.C2LG-001-000-A08 供应商管理程序.doc'
new_doc = r'D:\Code\Project\tools\data\input\translate_output\1.docx'
doc_to_docx(old_doc, new_doc)


# 批量转换示例
def batch_convert(folder_path):
    """批量转换文件夹中的所有doc文件"""
    for filename in os.listdir(folder_path):
        if filename.endswith('.doc'):
            doc_path = os.path.join(folder_path, filename)
            doc_to_docx(doc_path)

# 批量转换
# batch_convert('./doc_files/')