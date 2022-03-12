import docx
from docx.shared import RGBColor
import string

document = docx.Document('./test.docx')

f = open('./text.txt', 'r')
to_hide_text = "".join(f.readlines())


def get_letters_to_hide(t):
    return list("".join(t.translate(str.maketrans('', '', string.punctuation)).split()))


# print(get_letters_to_hide(to_hide_text))

letters_to_hide = get_letters_to_hide(to_hide_text)

for para in document.paragraphs:
    text = para.text.split()
    para.text = ''
    for word in text:
        if letters_to_hide[0] in word:
            words_letters = list(word)
            print(words_letters)
            is_marked = False
            for l in words_letters:
                if not is_marked and l == letters_to_hide[0]:
                    del letters_to_hide[0]
                    print(letters_to_hide)
                    para.add_run(l).font.color.rgb = RGBColor(255, 255, 1)
                    is_marked = True
                else:
                    para.add_run(l)
            para.add_run(" ")
        else:
            para.add_run(word)


document.save('./test.docx')

