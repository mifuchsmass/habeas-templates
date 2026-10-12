#!/usr/bin/env python3
"""Check actual DOCX package connections; this is not a substitute for Word acceptance."""
import argparse,posixpath,zipfile,xml.etree.ElementTree as E
from pathlib import Path
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
REL='{http://schemas.openxmlformats.org/package/2006/relationships}'
CT='{http://schemas.openxmlformats.org/package/2006/content-types}'
def check(path):
 with zipfile.ZipFile(path) as z:
  assert z.testzip() is None,'ZIP checksum failure'
  names=set(z.namelist());xml={n:E.fromstring(z.read(n)) for n in names if n.endswith(('.xml','.rels'))}
  types=xml['[Content_Types].xml'];defaults={e.get('Extension') for e in types.findall(CT+'Default')};overrides={e.get('PartName').lstrip('/') for e in types.findall(CT+'Override')}
  assert 'word/document.xml' in overrides,'Missing main document content type'
  for n in names:
   if n.endswith('/') or n=='[Content_Types].xml':continue
   assert n in overrides or n.rsplit('.',1)[-1] in defaults,f'Missing content type: {n}'
  for n,root in xml.items():
   if not n.endswith('.rels'):continue
   base=posixpath.dirname(posixpath.dirname(n));ids=set()
   for rel in root.findall(REL+'Relationship'):
    rid=rel.get('Id');assert rid not in ids,f'Duplicate relationship {n} {rid}';ids.add(rid)
    if rel.get('TargetMode')=='External':continue
    target=rel.get('Target').split('#')[0]
    target=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base,target))
    assert target in names,f'Dangling relationship {n}: {target}'
  styles={e.get(W+'styleId') for e in xml.get('word/styles.xml',E.Element('none')).findall(W+'style')}
  nums={e.get(W+'numId') for e in xml.get('word/numbering.xml',E.Element('none')).findall(W+'num')}
  refs=[]
  for n,root in xml.items():
   if not n.startswith('word/') or n.endswith('.rels'):continue
   if n in ('word/styles.xml','word/numbering.xml'):continue
   for tag in ['pStyle','rStyle','tblStyle']:
    for node in root.iter(W+tag):assert node.get(W+'val') in styles,f'Unknown style {n}: {node.get(W+"val")}'
   for node in root.iter(W+'numId'):assert node.get(W+'val')=='0' or node.get(W+'val') in nums,f'Unknown numbering {n}'
   refs.extend(node.get(W+'id') for node in root.iter(W+'footnoteReference'))
   relpath=posixpath.join(posixpath.dirname(n),'_rels',posixpath.basename(n)+'.rels')
   relationship_ids={e.get('Id') for e in xml.get(relpath,E.Element('none')).findall(REL+'Relationship')}
   for node in root.iter():
    for name,value in node.attrib.items():
     if name.startswith(R):assert value in relationship_ids,f'Dangling relationship attribute {n} {value}'
   for cell in root.iter(W+'tc'):assert cell.find(W+'p') is not None or cell.find(W+'tbl') is not None,f'Empty table cell {n}'
  notes=xml.get('word/footnotes.xml',E.Element('none'));ids=[node.get(W+'id') for node in notes.findall(W+'footnote')]
  assert len(ids)==len(set(ids)),'Duplicate footnote IDs'
  assert set(refs).issubset(ids),'Dangling footnote reference'
  for node in notes.findall(W+'footnote'):
   if node.get(W+'type') in ('separator','continuationSeparator','continuationNotice'):continue
   assert node.get(W+'id') in refs,'Unreferenced normal footnote'
   assert len(list(node.iter(W+'footnoteRef')))==1,'Footnote own marker missing or duplicated'
 print('PASS package connections, numbering, styles and footnote markers:',path.name)
if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('paths',nargs='+');args=parser.parse_args()
 for name in args.paths:
  path=Path(name)
  for file in sorted(path.glob('*.docx')) if path.is_dir() else [path]:check(file)
