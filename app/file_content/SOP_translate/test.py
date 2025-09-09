import datetime
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

if __name__ == '__main__':
    datetime = datetime.datetime.now().strftime('%y%m%d%H%M%S')
    output_file = f"../../../file/temp/outputxxx.docx"
    doc = Document()

    para = doc.add_paragraph()

    run = para.add_run('testn\n你好')
    # run1 = para.add_run('你好，鲍老师')
    # run2 = para.add_run('test text content.')

    # run.font.name = 'Times New Roman'

    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体')
    run.font.size = Pt(36)

    # run1.font.name = '宋体'
    # run1.font.size = Pt(36)
    #
    # run2.font.name = '宋体'
    # run2.font.size = Pt(12)

    doc.save(output_file)
