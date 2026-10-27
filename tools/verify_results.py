"""Independent numerical and structural checks of final lab deliverables."""
from pathlib import Path
from fractions import Fraction as F
import struct,json,re,sys
from zipfile import ZipFile
from lxml import etree
from docx import Document
import numpy as np
sys.stdout.reconfigure(encoding='utf-8')
R=Path(__file__).resolve().parents[1]

def number(s,b):
 a,_,f=s.replace(',','.').partition('.')
 return F(int(a,b))+sum((F(int(c,b),b**(i+1)) for i,c in enumerate(f)),F(0))

assert number('1110000,010110',2)==F(11234375,100000)
assert number('11000,101110',2)==F(2471875,100000)
for a,b,x,y in [('52','2A',8,16),('13,1','B,2',8,16),('71,45','39,94',8,16),('22','42',16,8),('5B,7','133,34',16,8),('E0,59','340,262',16,8)]:assert number(a,x)==number(b,y)
assert number('71,45',8)-number('22,06',8)==number('47,37',8)
assert number('981,23',16)+number('E0,59',16)==number('A61,7C',16)
assert F(13,10)-number('1,01',2)==F(1,20)
for h,v in [('45000000',2048),('C2F88000',-124.25),('44FF9000',2044.5),('40A0000000000000',2048),('C05F900000000000',-126.25),('409FF20000000000',2044.5)]:assert struct.unpack('>f' if len(h)==8 else '>d',bytes.fromhex(h))[0]==v
assert 105-46-88+22-(-55)==48
for name in ['work/Name1','work/run_1/Name1']:
 p=R/'04_Командная_строка'/name;s=(p/'Test1.txt').read_bytes()
 assert (p/'Test5.txt').read_bytes()==s+s
 for f in ['Name2/Test2.txt','Name2/Name3/Test3.txt','Name2/Name4/Test4.txt','Name5/Test1.txt']:assert (p/f).read_bytes()==s

A=np.array([[3,1,2],[0,4,1],[2,1,5]])
B=np.array([[1,0,1],[0,2,0],[1,0,2]])
C=np.array([[1,2,0],[0,1,1],[1,0,2]])
assert np.array_equal((A-B)@C@C,[[6,11,11],[5,4,10],[11,11,17]])

for p in R.glob('0*/*.docx'):
 if p.name.startswith('~$'):continue
 d=Document(p);text='\n'.join(q.text for q in d.paragraphs)
 assert 'Дубограй Артем Юрьевич' in text
 assert 'Элементы оглавления не найдены' not in text
 assert 'Контрольные вопросы из методички' not in text
 print(p.parent.name,'DOCX OK')

p=R/'05_Word/Дубограй_А_Ю_Word.docx'
with ZipFile(p) as z:
 root=etree.fromstring(z.read('word/document.xml'))
 ns={'m':'http://schemas.openxmlformats.org/officeDocument/2006/math','w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
 assert len(root.findall('.//m:oMath',ns))==4
 headings=[q for q in root.findall('.//w:body/w:p',ns) if re.match(r'Задание [1-7] ',''.join(q.xpath('.//w:t/text()',namespaces=ns))) and not q.findall('w:hyperlink',ns)]
 assert len(headings)==7
 assert any('TOC' in (q.text or '') for q in root.findall('.//w:instrText',ns))
 for g in range(5):
  table=root.findall('.//w:body/w:tbl',ns)[-1];row=table.findall('w:tr',ns)[3+3*g]
  assert ''.join(row.find('w:tc',ns).xpath('.//w:t/text()',namespaces=ns))==str(g+1)
for p in [R/'05_Word/Дубограй_А_Ю_Word.doc',R/'06_Excel/Дубограй_А_Ю_Excel_1.xls']:assert p.read_bytes()[:8]==bytes.fromhex('D0CF11E0A1B11AE1')
assert len(list(R.glob('0*/Для_защиты.md')))==7
html=(R/'07_HTML/index.html').read_text(encoding='utf-8')
for tag in ['<h1>','<h2>','<h3>','<pre ','<strong>','<em ']:assert tag in html
assert 'src="http' not in html and 'href="http' not in html
print('Conversions, arithmetic, IEEE754, BAT, matrices, equations, contents, files and private preparation: OK')
