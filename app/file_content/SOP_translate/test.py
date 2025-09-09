import datetime
from docx import Document
from docx.shared import Pt

if __name__ == '__main__':
    datetime = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    output_file = f"../../../file/temp/outputxxx-{datetime}.docx"
    doc = Document()

    para = doc.add_paragraph('test text content.')
    para.runs.font.name = '宋体'
    para.font.size = Pt(12)

    doc.save(output_file)