#!/usr/bin/env python3
"""Verify original migration plus explicit approved October 10 and October 13 Word edits.
--self-test changes package copies only in memory; never writes the template.
"""
import argparse,copy,io,re,subprocess,zipfile,xml.etree.ElementTree as E
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
BASE='5cde55966da8641665849ff27ed8c3fd3ae3b3e2'
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14='{http://schemas.microsoft.com/office/word/2010/wordml}'
SPACE='{http://www.w3.org/XML/1998/namespace}space'

def package(raw):
 with zipfile.ZipFile(io.BytesIO(raw)) as z:
  assert z.testzip() is None,'ZIP integrity failure'
  assert len(z.namelist())==len(set(z.namelist())),'Duplicate package entries'
  return {n:z.read(n) for n in z.namelist()}

def sig(n):
 # Expanded names ignore namespace-prefix and attribute-order serialization.
 return (n.tag,tuple(sorted(n.attrib.items())),n.text or '',n.tail or '',tuple(sig(c) for c in n))

def convert(m):
 body=m[1].strip()
 substitutions={'He/She':'Petitioner','is/are':'is',"Petitioner's":"Petitioner's"}
 if body in substitutions:return substitutions[body]
 if body.startswith('NOTE TO PRACTITIONER:'):return m[0]
 if body in ('ELSE','END IF'):return '{'+body+'}'
 if body.startswith(('IF ','ELSE IF ')):
  c=re.fullmatch(r'(IF|ELSE IF)\s+(.+?)\s*(>=|<=|!=|=|<|>)\s*(.+)',body)
  assert c,'Unexpected original conditional'
  return '{'+c[1]+' <'+c[2].strip()+'> '+c[3]+' '+c[4]+'}'
 return '<'+body+'>'

def normalize_runs(p,original):
 # Combine adjacent text-only runs only when their formatting is identical.
 # Preserve real breaks, fields, drawings, bookmarks and all formatting.
 for r in p.findall(W+'r'):
  for attr in (W+'rsidR',W+'rsidRPr'):r.attrib.pop(attr,None)
  for t in r.findall(W+'t'):t.attrib.pop(SPACE,None) # retain actual whitespace
 groups=[]
 for n in p:
  if n.tag==W+'r' and any(c.tag==W+'t' for c in n) and all(c.tag in (W+'rPr',W+'t') for c in n):
   props=n.find(W+'rPr')
   key=(tuple(sorted(n.attrib.items())),sig(props) if props is not None else None)
   text=''.join(c.text or '' for c in n if c.tag==W+'t')
   if groups and groups[-1][0]=='text' and groups[-1][1]==key:groups[-1][2]+=text
   else:groups.append(['text',key,text,props,dict(n.attrib)])
  else:groups.append(['node',n])
 p[:]=[]
 for item in groups:
  if item[0]=='node':p.append(item[1])
  else:
   r=E.SubElement(p,W+'r',item[4])
   if item[3] is not None:r.append(item[3])
   E.SubElement(r,W+'t').text=re.sub(r'\[([^\]]+)\]',convert,item[2]) if original else item[2]

def document(raw,original=False):
 root=E.fromstring(raw);ps=list(root.iter(W+'p'))
 assert len(ps)==39,'Paragraph count changed'
 if original:
  h=ps[0]
  # Derive heading formatting from original, adding only approved All Caps.
  h.set(W14+'textId','26AF6C36');h.set(W+'rsidRPr','00064280');h.set(W+'rsidRDefault','00F66102')
  E.SubElement(h.find(W+'pPr/'+W+'rPr'),W+'caps')
  r=copy.deepcopy(h.find(W+'r'));E.SubElement(r.find(W+'rPr'),W+'caps')
  r.find(W+'t').text=('IN THE UNITED STATES DISTRICT COURT FOR THE DISTRICT OF '
   '<Jurisdiction>')
  for n in list(h):
   if n.tag==W+'r':h.remove(n)
  h.append(r)
  for text in ps[27].iter(W+'t'):
   if text.text and 'Detention Facility' in text.text:
    text.text=text.text.replace('Detention Facility','<Detention Facility>')
  ps[34].set(W14+'textId','1BE7CCFB')
  t=ps[34].find('.//'+W+'t')
  assert t.text.startswith('2. Petitioner entered'),'Original numbering anchor changed'
  t.text=t.text[3:] # remove extra literal number, retain numPr
 for i,p in enumerate(ps,1):
  # Observed Word proofing and cached pagination artifacts, location constrained.
  # Never ignore tracked changes or author-inserted page breaks.
  if i in (31,38):
   markers=p.findall(W+'proofErr')
   expected=['gramStart','gramEnd']*(1 if i==31 else 2)
   assert [n.get(W+'type') for n in markers]==([] if original else expected),'Unexpected proofing markers'
   for n in markers:
    assert set(n.attrib)=={W+'type'} and not len(n),'Unexpected proofing structure'
    p.remove(n)
  if i in (36,37):
   markers=list(p.iter(W+'lastRenderedPageBreak'))
   assert len(markers)==int(i==(36 if original else 37)),'Unexpected cached page break'
   for r in p.findall(W+'r'):
    for n in list(r):
     if n.tag==W+'lastRenderedPageBreak':
      assert not n.attrib and not len(n),'Unexpected cached-break structure'
      r.remove(n)
  normalize_runs(p,original)
 return sig(root)

def check(before,current):
 assert set(before)==set(current),'Package entries changed'
 exceptions={'word/document.xml','word/settings.xml','word/numbering.xml','docProps/core.xml','docProps/app.xml'}
 for n in before:
  if n not in exceptions:assert before[n]==current[n],'Non-document part changed: '+n
 old=E.fromstring(before['word/settings.xml']);new=E.fromstring(current['word/settings.xml'])
 allowed={W+'zoom',W+'proofState',W+'rsids'}
 for tag in allowed:assert len(new.findall(tag))==1,'Missing/duplicate allowed setting'
 zoom=new.find(W+'zoom');proof=new.find(W+'proofState')
 assert zoom.attrib=={W+'percent':'90'} and not len(zoom),'Unexpected zoom change'
 assert proof.attrib=={W+'spelling':'clean',W+'grammar':'clean'} and not len(proof),'Unexpected proofing setting'
 expected=copy.deepcopy(old.find(W+'rsids'))
 for value in ('00064280','00E0022B','00F17F8A','00F66102'):E.SubElement(expected,W+'rsid',{W+'val':value})
 actual=copy.deepcopy(new.find(W+'rsids'))
 for n in (expected,actual):n[:]=sorted(n,key=lambda c:(c.tag,c.get(W+'val')))
 assert sig(expected)==sig(actual),'Unexpected revision-session changes'
 for root in (old,new):
  for n in list(root):
   if n.tag in allowed:root.remove(n)
 assert sig(old)==sig(new),'Substantive settings changed'
 assert sig(E.fromstring(before['word/numbering.xml']))==sig(E.fromstring(current['word/numbering.xml'])),'Numbering definitions changed'
 properties={
 'docProps/core.xml':{'{http://purl.org/dc/terms/}modified':'2026-10-10T16:17:00Z','{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}revision':'26'},
 'docProps/app.xml':{'{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'+k:v for k,v in {'TotalTime':'88','Words':'300','Characters':'1716','Lines':'14','Paragraphs':'4','CharactersWithSpaces':'2012'}.items()}}
 for part,values in properties.items():
  expected=E.fromstring(before[part])
  for tag,value in values.items():
   n=expected.find(tag);assert n is not None,'Missing original metadata anchor';n.text=value
  assert sig(expected)==sig(E.fromstring(current[part])),'Unexpected metadata changed: '+part
 assert document(before['word/document.xml'],True)==document(current['word/document.xml']),'Unapproved document text or formatting'

def self_test(before,current):
 # A numbering-only XML reserialization must still pass, without a new baseline.
 serialized=dict(current)
 serialized['word/numbering.xml']=E.tostring(E.fromstring(current['word/numbering.xml']))
 check(before,serialized)
 print('PASS accepts: equivalent numbering XML serialization')
 def italic_body(root):
  run=list(root.iter(W+'p'))[26].find(W+'r')
  props=run.find(W+'rPr')
  if props is None:
   props=E.Element(W+'rPr');run.insert(0,props)
  E.SubElement(props,W+'i')
 def reject(label,part,mutate):
  changed=dict(current)
  if part:
   root=E.fromstring(changed[part]);mutate(root);changed[part]=E.tostring(root)
  else:mutate(changed)
  try:check(before,changed)
  except AssertionError:print('PASS rejects:',label)
  else:raise AssertionError('Accepted forbidden change: '+label)
 reject('substantive setting','word/settings.xml',lambda r:E.SubElement(r,W+'evenAndOddHeaders'))
 reject('numbering definition','word/numbering.xml',lambda r:next(r.iter(W+'numFmt')).set(W+'val','bullet'))
 reject('unrelated text','word/document.xml',lambda r:setattr(list(r.iter(W+'t'))[-1],'text','Unapproved text'))
 reject('unrelated italic formatting','word/document.xml',italic_body)
 reject('heading All Caps removed','word/document.xml',lambda r:next(r.iter(W+'caps')).set(W+'val','0'))
 reject('paragraph numbering changed','word/document.xml',lambda r:next(r.iter(W+'numId')).set(W+'val','999'))
 reject('extra literal paragraph number restored','word/document.xml',lambda r:setattr(list(r.iter(W+'p'))[34].find('.//'+W+'t'),'text','2. Petitioner entered'))
 reject('unexpected package part',None,lambda p:p.update({'word/unexpected.xml':b'<unexpected/>'}))
 reject('removed package part',None,lambda p:p.pop('word/styles.xml'))
 reject('relationship target','word/_rels/document.xml.rels',lambda r:r[0].set('Target','unexpected.xml'))
 reject('unapproved metadata','docProps/app.xml',lambda r:setattr(r[0],'text','Unexpected'))

if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--self-test',action='store_true');args=parser.parse_args()
 original=package(subprocess.check_output(['git','show',BASE+':templates/petition.docx'],cwd=ROOT))
 current=package((ROOT/'templates/petition.docx').read_bytes())
 check(original,current)
 print('PASS original migration and explicit approved Word revisions; package and formatting protections retained.')
 if args.self_test:self_test(original,current)
