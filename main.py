import docx
try:
	from xml.etree.cElementTree import XML
except ImportError:
	from xml.etree.ElementTree import XML
import zipfile
import re


# doc_path = "/home/nikon-cook/Documents/МИТМО/Analisys_TD/Med_karta_1_bez_personalnykh_dannykh.doc"
# docx_path = "/home/nikon-cook/Documents/МИТМО/Analisys_TD/Med_karta_1_bez_personalnykh_dannykh.docx"
docx_path = "Med_karta_1_bez_personalnykh_dannykh.docx"


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
print()


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
	block = []
	print_next = 0

	word_matching_list = {'sex': r'[Пп]ол[: ]? \w+\b',
						'age': r'[Вв]озраст[: ]? \d{1,3}лет',
						'years_old': r'\d{1,3}[ ]?год[а]?[ ]?\d{0,2}[ ]?[месяц]?[ев]?[ ]?\d{0,2}',
						'birth_place': r'[Мм]есто рождения[: ]? .*',
						'person': r'личность[: ]? .*',
						'registration': r'[Рр]егистраци.[: ]? .*',
						'insurance': r'страхов.+[: ]? .*',
						'document': r'(СНИЛС|ИНН|[Пп]аспорт)[: ]? .*',
						'nationality': r'[Гг]ражданство[: ]? .*',
						'family_status': r'[Сс]емейное положение[: ]? (замужем|женат)',
						'income': r'[Дд]оход[: ]? .*',
						'living_place': r'[Мм]есто жительства[: ]? .*',
						'phone': r'[Тт]елефон[у]?[:\s]+\d{3,}\b',
						'email': r'[A-z\w]@[A-z].[a-z]',
						'floor1': r'(\d? Этаж \d?)',
						'floor2': r'(\d? этаж \d?.*$)',
						'floor3': r'(\d{1,2}эт)',
						'reg_num': r'/?[0-9]{5,}/?',
						'alphanumeric_code': r'[A-ZА-Я]\d{4,}',
						# 'alphanumeric_code2': r'[A-ZА-Я]?\d+[A-ZА-Я]+\d+[A-ZА-Я]+)',
						# 'period': r'\d{2}[.]\d{2}[.]\d{2,4}[ -]+\d{2}[.]\d{2}.\d{2}',
						# 'date_time': r'\d{2}[.]\d{2}[.]\d{2,4}[ г.]?[ ]?\d{0,2}?[:]?\d{0,2}?',  # это не хочет работать правильно почему-то (время не затирает), поправила

						'time_full': r'\d{2}[:]\d{2}[:]\d{2}',
						'date_time': r'\d{2}[.]\d{2}[.]\d{2,4}[ ]?\d{2}[:]\d{2}',
						'date': r'\d{2}[.]\d{2}[.]\d{2,4}',
						'time': r'\d{2}[:]\d{2}',
						'date_or_time_short': r'(с|от|на) \d{2}[.]\d{2}',
						'blood_pressure': r'\d{2,3}[\\]?[/]?\d{2,3} мм [рp]т[.]?[ ]?[сc]т[.]?',

						'FIO_short_double': r'[А-ЯЁ][а-яё]+[ ]+[А-ЯЁ]{2}[ \t\n\r.?!/$]+',
						'FIO_short_with_space': r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]{1}[.]?[ ]+[А-ЯЁ]{1}[.]?[ \t\n\r.?!$]+',
						'FIO_short_without_space': r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]{1}[.]?[А-ЯЁ]{1}[.]?[ \t\n\r.?!$]+',
						'FIO_full': r'[А-ЯЁ][а-яё\-]+[ ]+[А-ЯЁ][а-яё\-]+[ ]*[А-ЯЁ][а-яё\-]+'
						}
	block_separators = {
						'2slash1': r'[\w -]+ /+ \d{2}[.]\d{2}[.]\d{2,4} \d{0,2}[:.]?\d{0,2}[ ]?/+ [\w -]+',
						'2slash2': r'\d{2}[.]\d{2}[.]\d{2,4} \d{0,2}[:.]?\d{0,2}[ ][/][ \w]+[/][ \w]+',
						'1slash': r'\d{2}[.]\d{2}[.]\d{2,4} \d{0,2}[:.]?\d{0,2}[ ]?/+'
						}

	for paragraph in tree.iter(PARA):
		for node in paragraph.iter(TEXT):

			"""
			if print_next > 0:
				print('print_next', node.text)
				print_next -= 1
			"""
			"""
			if re.match(r'/?[0-9]{5,}/?', node.text):
				print('reg_num', node.text)
			"""
			# желательно распознавать что за номер по слову перед номером
			"""
			if re.findall(r'[А-ЯЁ][а-яё]+[ ]+[А-ЯЁ]{2}[ \t\n\r.?!/$]+', node.text):
				print('FIO_short_double', re.findall(r'[А-ЯЁ][а-яё]+[ ]+[А-ЯЁ]{2}[ \t\n\r.?!$]+', node.text), node.text)
			elif re.findall(r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]{1}[.]?[ ]+[А-ЯЁ]{1}[.]?[ \t\n\r.?!$]+', node.text):
				print('FIO_short_with_space', re.findall(r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]{1}[.]?[ ]+[А-ЯЁ]{1}[.]?[ \t\n\r.?!$]+', node.text), node.text)
			elif re.findall(r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]{1}[.]?[А-ЯЁ]{1}[.]?[ \t\n\r.?!$]+', node.text):
				print('FIO_short_without_space',  re.findall(r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]{1}[.]?[А-ЯЁ]{1}[.]?[ \t\n\r.?!$]+', node.text), node.text)
			elif re.findall(r'[А-ЯЁ][а-яё\-]+[ ]+[А-ЯЁ][а-яё\-]+[ ]*[А-ЯЁ][а-яё\-]+', node.text):
				print('FIO_full', re.findall(r'[А-ЯЁ][а-яё\-]+[ ]+[А-ЯЁ][а-яё\-]+[ ]*[А-ЯЁ][а-яё\-]+', node.text), node.text)
			elif re.findall(r'[А-ЯЁ]{2}[ ]+[А-ЯЁ][а-яё]+[ \t\n\r.?!S]', node.text):
				print('IOF_short_double', re.findall(r'[А-ЯЁ]{2}[ ]+[А-ЯЁ][а-яё]+[ \t\n\r.?!S]', node.text))
			elif re.findall(r'[А-ЯЁ][.][ ]+[А-ЯЁ][.][ ][А-ЯЁ][а-яё]+[ \t\n\r.?!$]', node.text):
				print('IOF_short_with_space', re.findall(r'[А-ЯЁ][.][ ]+[А-ЯЁ][.][ ][А-ЯЁ][а-яё]+[ \t\n\r.?!$]', node.text))
			elif re.findall(r'[А-ЯЁ][.][А-ЯЁ][.][ ][А-ЯЁ][а-яё]+[ \t\n\r.?!$]', node.text):
				print('IOF_short_without_space', re.findall(r'[А-ЯЁ][.][А-ЯЁ][.][ ][А-ЯЁ][а-яё]+[ \t\n\r.?!$]', node.text))
			"""
			"""
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
			elif re.match(r'\d{2}.\d{2}.\d{2,4}', node.text):
				print('time or missed date', node.text)
			"""
			"""
			if re.findall(r'Этаж|этаж|(\d{1,2}эт)', node.text):
				print('floor', node.text)
				print_next = 1
			"""
			"""
			if re.match(r'Эпидномер', node.text):
				print('word_matching', node.text)
				print_next = 1
			"""

			for i in block_separators:
				expression = block_separators[i]
				if re.findall(expression, node.text):
					paragraphs.append(block)
					block = []

			for i in word_matching_list.keys():
				# if i in ('', '', '', '', 'blood_pressure'):  # убрать в результате, это проверка
					expression = word_matching_list[i]
					new_node = re.sub(expression, '<'+i+'>', node.text)
					node_matches = re.findall(expression, node.text)
					node.text = new_node

					if len(node_matches) > 0:
						# if i in ('blood_pressure', ''):  # убрать в результате, это проверка
							print(i, node_matches, node.text)
							print(new_node)
							print()

			block.append(node.text)
	if len(block) > 0:
		paragraphs.append(block)
	return paragraphs


text = get_docx_text(docx_path)
print(type(text))
print(len(text))
# print(text[0:1000])

paragraphs = get_doc_paragraphs(docx_path)
print(len(paragraphs))

for i in paragraphs:
	print(i)
