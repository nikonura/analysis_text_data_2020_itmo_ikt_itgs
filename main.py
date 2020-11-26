# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import docx
from docx.shared import Inches
import textract
import docx2txt
# from docx2python import docx2python
import urllib
from bs4 import BeautifulSoup

try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
import re


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

doc_path = "/home/nikon-cook/Documents/МИТМО/Analisys_TD/Med_karta_1_bez_personalnykh_dannykh.doc"
docx_path = "/home/nikon-cook/Documents/МИТМО/Analisys_TD/Med_karta_1_bez_personalnykh_dannykh.docx"

"""
text = textract.process(doc_path)
text = text.decode("utf-8")
print(type(text))
print(text)
"""
doc = docx.Document(docx_path)
# fullText = docx.getdocumenttext(doc)
styles = doc.styles
for i in styles:
    print(i)

print(type(doc))


def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)


file_context = getText(docx_path)
print(len(file_context))


def para2text(p):
    rs = p._element.xpath('.//w:t')
    return u" ".join([r.text for r in rs])


WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'


def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))

    return '\n\n'.join(paragraphs)


def get_doc_paragraphs(path):
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    print_next = 0
    for paragraph in tree.iter(PARA):
        for node in paragraph.iter(TEXT):
            """
            if print_next > 0:
                print('print_next', node.text)
                print_next -= 1
            """

            if re.match(r'/?[0-9]{5,}/?', node.text):
                print('reg_num', node.text)
            # желательно распознавать что за номер по слову перед номером

            if re.match(r'\d{2}[.]\d{2}[.]\d{2,4}[ -г. ]+\d{2}[.]\d{2}[.]\d{2,4} \w+', node.text) or \
                    re.match(r'\d{2}[.]\d{2}[.]\d{2,4}[ г.]?\d{0,2}:?\d{0,2}[ -]+\d{0,2}:?\d{0,2} \w+', node.text):
                print('fact_of_treatment', node.text)
            elif re.match(r'\d{2}[.]\d{2}[.]\d{2,4} [A-zА-я]+', node.text):
                print('fact_of_treatment3', node.text)
            elif re.match(r'[\w -]+ /+ \d{2}[.]\d{2}[.]\d{2,4} \d{0,2}[:.]?\d{0,2}[ ]?/+ [\w -]+', node.text):
                print('2slash1', node.text)
            elif re.match(r'\d{2}[.]\d{2}[.]\d{2,4} \d{0,2}[:.]?\d{0,2}[ ][/][ \w]+[/][ \w]+', node.text):
                print('2slash2', node.text)
            elif re.match(r'\d{2}[.]\d{2}[.]\d{2,4} \d{0,2}[:.]?\d{0,2}[ ]?/+', node.text):
                print('1slash', node.text)
            elif re.match(r'\d{2}[.]\d{2}[.]\d{2,4}[ -]+\d{2}[.]\d{2}', node.text):
                print('period', node.text)
            elif re.match(r'\d{2}[.]\d{2}[.]\d{2,4}', node.text):
                print("date", node.text)
                print_next = 2
            elif re.match(r'\d{2}.\d{2}.\d{2-4}', node.text):
                print('time or missed date', node.text)

            if re.match(r'Этаж', node.text):
                print('word_matching', node.text)
                print_next = 1
            paragraphs.append([node.text])
    return paragraphs


text = get_docx_text(docx_path)
print(type(text))
print(len(text))
# print(text[0:1000])

paragraphs = get_doc_paragraphs(docx_path)
# print(paragraphs)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
