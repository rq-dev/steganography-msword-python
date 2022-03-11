import docx
from docx.shared import RGBColor

document = docx.Document('./test.docx')

for para in document.paragraphs:
    text = para.text.split()
    print(text)
    para.text = ''
    print(len(text))
    c = 0
    for i in range(len(text)):
        para.add_run(text[i] + '\n').font.color.rgb = RGBColor(i+c, c, c)
        print(c)
        c += 30


document.save('./test.docx')

