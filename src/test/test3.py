from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Cm
import datetime

datetime = datetime.datetime.now().strftime('%y%m%d%H%M%S')
input_file = r"D:\Code\Project\tools\data\test\test.docx"
output_file = f"../../data/temp/output_{datetime}.docx"
doc = Document(input_file)

new_doc = Document()

for item in doc.tables:
    print(item.rows, item.columns)
    for row in item.rows:
        for cell in row.cells:
            print(cell)