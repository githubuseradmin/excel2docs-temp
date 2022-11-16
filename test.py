from docx import Document
from docxcompose.composer import Composer

dir = "C:/Users/User/Desktop/New folder/test/test"

result = Document(dir + "1.docx")

composer = Composer(result)

for i in range(1, 15):
    document = Document(dir + str(i) + ".docx")
    composer.append(document)

composer.save(dir + "123.docx")
