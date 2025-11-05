# 抽取 docx 中的图片
from typing import List
from docx import Document
from docx.enum.shape import WD_INLINE_SHAPE_TYPE
from docx.image.image import Image
from docx.oxml import CT_Inline, CT_GraphicalObject, CT_GraphicalObjectData, CT_Picture, CT_BlipFillProperties, CT_Blip, \
    CT_R


def extract_pics(file_path: str) -> None:
    doc = Document(file_path)
    for idx, shape in enumerate(doc.inline_shapes):
        # 获取图片的rid
        rid = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
        # 图片的名称
        name = shape._inline.docPr.name
        # 根据rid获取图片对象
        image_obj = doc.part.rels.get(rid).target_part.image
        print("image:", image_obj)
        print("file_name:", image_obj.filename)
        print("扩展名:", image_obj.ext)
        # print("图片二进制数据:", image_obj.blob)
        print("图片二进制数据hash:", image_obj.sha1)
        print("内容类型:", image_obj.content_type)

        print("像素宽高{}x{}:".format(image_obj.px_width, image_obj.px_height))
        print("分辨率{} x {}".format(image_obj.horz_dpi, image_obj.vert_dpi))
        # 解析图片信息
        width = shape.width,  # Lendth对象  可以继续调用.inches/.pt/.cm
        height = shape.height

        print("\n")


def extract_shapes(file_path: str) -> None:
    """ 抽取图形 """
    doc = Document(file_path)
    # 在段落中抽取图形
    for idx, para in enumerate(doc.element.body):
        for run in para:
            # 找到run
            if isinstance(run, CT_R):
                for inner_run in run:
                    tag_name = inner_run.tag.split("}")[1]
                    if tag_name == "AlternateContent":  # lxml.etree._Element对象
                        fallback = inner_run[1]
                        pict = fallback[0]
                        shape = pict[0]  # 内部图形
                        print("图形数据:", shape.items())
                        # textbox = list(shape) # shape内部是否有文本，需判断
                        # if textbox:
                        #     textbox = textbox[0]
                        #     # 图形内的段落文本
                        #     para_list = list(textbox[0])
                        #     print("图形内部文本:", [para.text for para in para_list])


if __name__ == '__main__':
    word_path = r"/data/test/13.docx"
    extract_pics(word_path)
    extract_shapes(word_path)

