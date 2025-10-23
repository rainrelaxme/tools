from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Cm
import datetime

datetime = datetime.datetime.now().strftime('%y%m%d%H%M%S')
# input_file = f"../../../file/test.docx"
output_file = f"../../data/temp/output_{datetime}.docx"
# doc = Document(input_file)

new_doc = Document()

table = new_doc.add_table(rows=2, cols=3)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
# table.cell(0, 0).width = Cm(1)
# table.cell(0, 1).width = Cm(1)
# table.cell(0, 2).width = Cm(1)
# table.cell(1, 0).width = Cm(1)
# table.cell(1, 1).width = Cm(1)
# table.cell(1, 2).width = Cm(1)
for i in range(len(table.rows)):
    for j in range(len(table.columns)):
        table.cell(i, j).width = Cm(2)
        table.rows[i].height = Cm(2)
#
# table.cell(1,0).merge(table.cell(1,1))
# table.cell(1, 0).width = Inches(3)

# for i in range(len(table.rows)):
#     table.rows[i].height = Cm(3.0)

# new_doc.save(output_file)
# print(f"新文档已保存到: {output_file}")

main_text_cells = []
if len(main_text_cells) > 0 :
    print("********",main_text_cells)