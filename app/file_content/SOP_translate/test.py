import datetime
from docx import Document
from docx.shared import Pt

if __name__ == '__main__':
    datetime = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    output_file = f"../../../file/temp/outputxxx-{datetime}.docx"
    doc = Document()

    para = doc.add_paragraph()

    run = para.add_run('你好啊，鲍老师')
    run2 = para.add_run('test text content.')

    # run.font.name = 'Times New Roman'

    run.font.name = u'宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    run.font.size = Pt(36)

    run2.font.name = '宋体'
    run2.font.size = Pt(12)

    doc.save(output_file)
