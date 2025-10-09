from docx import Document
from docx.shared import Inches, Cm
import datetime

datetime = datetime.datetime.now().strftime('%y%m%d%H%M%S')
# input_file = f"../../../file/test.docx"
output_file = f"../../data/temp/output_{datetime}.docx"
# doc = Document(input_file)

new_doc = Document()

table = new_doc.add_table(rows=2, cols=3)
table.style = 'Table Grid'

# table.cell(0,0).width = Inches(3.6)
# table.cell(0,1).width = Inches(2.5)
# table.cell(0,1).width = Inches(2.5)
#
# table.cell(1,0).merge(table.cell(1,1))
# table.cell(1, 0).width = Inches(3)

for i in range(len(table.rows)):
    table.rows[i].height = Cm(3.0)

new_doc.save(output_file)
print(f"新文档已保存到: {output_file}")