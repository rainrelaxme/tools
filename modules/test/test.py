import datetime
from docx import Document

input_file = r"D:\Code\Project\tools\data\test\C3LG-Z03-002(A00) 备品备件与低值易耗品管理规定.docx"
doc = Document(input_file)

for table in doc.sections[0].header.tables:
    for row in table.rows:
        for cell in row.cells:
            print(cell.text)