#!/usr/bin/env python3
"""Offline browser regression runner. Uses existing pinned PizZip; installs nothing."""
import argparse,base64,json,re,subprocess,sys,zipfile,io,xml.etree.ElementTree as E
from html import unescape
from pathlib import Path
parser=argparse.ArgumentParser();parser.add_argument('--pizzip',default='/tmp/habeas-review-pizzip-3.1.7.js');parser.add_argument('--chrome',default='google-chrome');parser.add_argument('--output',default='/tmp/habeas-v3-implementation-tests');args=parser.parse_args()
root=Path(__file__).resolve().parent.parent;out=Path(args.output);out.mkdir(parents=True,exist_ok=True)
script_paths=[Path(args.pizzip),root/'js/template-logic.js',root/'js/docx-engine.js',root/'js/habeas-app.js',root/'js/drafting-studio.js']
# DOMContentLoaded app listeners run before the test adds controls.
html='<html><head><meta charset="UTF-8"></head><body><pre id="test-output">pending</pre>'+''.join('<script>'+p.read_text(encoding='utf-8').replace('</script','<\\/script')+'</script>' for p in script_paths)
html+='<script>const template64='+json.dumps(base64.b64encode((root/'templates/petition.docx').read_bytes()).decode())+';window.addEventListener("load",()=>{'+(root/'tests/browser.test.js').read_text(encoding='utf-8')+'});</script></body></html>'
page=out/'browser-tests.html';page.write_text(html,encoding='utf-8')
command=[args.chrome,'--headless','--no-sandbox','--disable-gpu','--disable-background-networking','--allow-file-access-from-files','--user-data-dir='+str(out/'chrome-profile'),'--virtual-time-budget=10000','--dump-dom',page.as_uri()]
run=subprocess.run(command,capture_output=True,text=True,encoding='utf-8',timeout=60);(out/'chrome-stderr.txt').write_text(run.stderr,encoding='utf-8')
if run.returncode:print(run.stderr);sys.exit(run.returncode)
m=re.search(r'<pre id="test-output">(.*?)</pre>',run.stdout,re.S)
if not m:raise RuntimeError('Missing browser result')
try:report=json.loads(unescape(m[1]))
except Exception:(out/'browser-dom.html').write_text(run.stdout,encoding='utf-8');raise
artifacts=report.pop('artifacts',{});(out/'results.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
for name,data in artifacts.items():
 raw=base64.b64decode(data);(out/(name+'.docx')).write_bytes(raw);z=zipfile.ZipFile(io.BytesIO(raw))
 for n in z.namelist():
  if n.endswith(('.xml','.rels')):E.fromstring(z.read(n))
 if z.testzip():raise RuntimeError('ZIP integrity failure')
 print('DOCX package/XML checked:',name)
failed=[r for r in report.get('results',[]) if not r['passed']]
for r in report.get('results',[]): print(('PASS ' if r['passed'] else 'FAIL ')+r['name']+(('\n'+r['error']) if not r['passed'] else ''))
if report.get('fatal'):print(report['fatal'])
print(f"{len(report.get('results',[]))} browser checks; {len(failed)} failures; {len(artifacts)} DOCX artifacts; output {out}")
sys.exit(1 if failed or report.get('fatal') else 0)
