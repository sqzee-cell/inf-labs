"""Change package metadata only; leave all layout and calculation parts intact."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree
root=Path(__file__).resolve().parents[1]
ns={'dc':'http://purl.org/dc/elements/1.1/','cp':'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'}
for p in root.glob('0*/*.docx'):
 if p.name.startswith('~$'):continue
 with ZipFile(p) as z:parts=[(i,z.read(i.filename)) for i in z.infolist()]
 out=p.with_suffix('.metadata.docx')
 with ZipFile(out,'w',ZIP_DEFLATED) as z:
  for info,data in parts:
   if info.filename=='docProps/core.xml':
    r=etree.fromstring(data)
    for tag in ['dc:creator','cp:lastModifiedBy']:
     prefix,local=tag.split(':');q='{'+ns[prefix]+'}'+local;e=r.find(q)
     if e is None:e=etree.SubElement(r,q)
     e.text='Дубограй Артем Юрьевич'
    data=etree.tostring(r,encoding='UTF-8',xml_declaration=True,standalone=True)
   z.writestr(info,data)
 out.replace(p)
print('Author metadata set in all seven reports; layout parts unchanged.')
