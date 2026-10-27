from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree

p=Path(r'F:\study\inf-labs\05_Word\Дубограй_А_Ю_Word.docx')
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W='{'+ns['w']+'}'
with ZipFile(p) as z:
 parts={i.filename:(i,z.read(i.filename)) for i in z.infolist()}
root=etree.fromstring(parts['word/document.xml'][1])
table=root.findall('.//w:body/w:tbl',ns)[-1]
rows=table.findall('w:tr',ns)

def set_cell(cell,text):
 ps=cell.findall('w:p',ns)
 first=ps[0]
 for child in list(first):
  if child.tag!=W+'pPr':first.remove(child)
 r=etree.SubElement(first,W+'r');rp=etree.SubElement(r,W+'rPr')
 sz=etree.SubElement(rp,W+'sz');sz.set(W+'val','20')
 f=etree.SubElement(rp,W+'rFonts');f.set(W+'ascii','Times New Roman');f.set(W+'hAnsi','Times New Roman')
 etree.SubElement(r,W+'t').text=text
 for other in ps[1:]:cell.remove(other)

for g in range(5):set_cell(rows[3+3*g].find('w:tc',ns),str(g+1))
cs=rows[1].findall('w:tc',ns)
for j in [2,6]:set_cell(cs[j],'Энергия, ккал')
for j in [5,9]:set_cell(cs[j],'Углев., г')
patched=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
tmp=p.with_suffix('.patch.docx')
with ZipFile(tmp,'w',ZIP_DEFLATED) as out:
 for name,(info,data) in parts.items():out.writestr(info,patched if name=='word/document.xml' else data)
tmp.replace(p)
print('Table merged cells and headers patched')
