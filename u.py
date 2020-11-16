from bs4 import BeautifulSoup as bs
import textract

doc_path = '/home/nikon-cook/Documents/МИТМО/Analisys_TD/Med_karta_1_bez_personalnykh_dannykh.doc'
text = textract.process(doc_path)
print(text)
#soup = bs(open(doc_path).read())
#[s.extract() for s in soup(['style', 'script'])]
#tmpText = soup.get_text()
#text = "".join("".join(tmpText.split('\t')).split('\n')).encode('utf-8').strip()
#print(text)