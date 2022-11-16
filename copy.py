from docx import Document

document_save = Document()

for document_file in range(3):
    document = Document('test/test' + str(document_file) + '.docx')
    

document.save('test/save.docx')
