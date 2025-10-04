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
tab_stop = tab_stops.add_tab_stop(Cm(5.0), alignment=WD_TAB_ALIGNMENT.RIGHT, leader=WD_TAB_LEADER.DOTS)  # 右对齐，前导字符为点
document.save('test.docx')  # 保存
