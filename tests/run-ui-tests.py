#!/usr/bin/env python3
"""Serve test pages only on loopback; rewrite CDN paths in responses to pinned local copies."""
import argparse,json,re,subprocess,sys,threading
from html import unescape
from http.server import ThreadingHTTPServer,BaseHTTPRequestHandler
from pathlib import Path
parser=argparse.ArgumentParser();parser.add_argument('--vendor-dir',default='/tmp');parser.add_argument('--artifacts',default='/tmp/habeas-v3-implementation-tests');parser.add_argument('--chrome',default='google-chrome');args=parser.parse_args()
root=Path(__file__).resolve().parent.parent;out=Path(args.artifacts);vendor=Path(args.vendor_dir)
versions={'pizzip':'3.1.7','docxtemplater':'3.50.0','file-saver':'2.0.5','jszip':'3.10.1','docx-preview':'0.3.3'}
for name,version in versions.items():
 if not (vendor/f'habeas-review-{name}-{version}.js').exists():raise RuntimeError('Missing pinned local library: '+name)
class Handler(BaseHTTPRequestHandler):
 def log_message(self,*args):pass
 def do_GET(self):
  path=self.path.split('?')[0]
  if path=='/ui-tests.html':data=('<meta charset="UTF-8"><pre id="test-output">pending</pre><script>'+(root/'tests/ui.test.js').read_text()+'</script>').encode();mime='text/html'
  elif path.startswith('/vendor/'):
   name=path[8:].removesuffix('.js');version=versions.get(name)
   if not version:self.send_error(404);return
   data=(vendor/f'habeas-review-{name}-{version}.js').read_bytes();mime='text/javascript'
  else:
   directory=out if path.startswith('/artifacts/') else root
   relative=path.removeprefix('/artifacts/') if path.startswith('/artifacts/') else path.lstrip('/')
   target=(directory/relative).resolve()
   if not target.is_relative_to(directory.resolve()) or not target.is_file():self.send_error(404);return
   data=target.read_bytes();mime={'html':'text/html','js':'text/javascript','css':'text/css'}.get(target.suffix[1:],'application/octet-stream')
   if target.suffix=='.html':
    text=data.decode()
    text=re.sub(r'https://cdn.jsdelivr.net/npm/([^@]+)@[^"\s]+',lambda m:'/vendor/'+m[1]+'.js',text);data=text.encode()
  self.send_response(200);self.send_header('Content-Type',mime);self.send_header('Content-Length',str(len(data)));self.end_headers();self.wfile.write(data)
server=ThreadingHTTPServer(('127.0.0.1',0),Handler);threading.Thread(target=server.serve_forever,daemon=True).start()
try:
 cmd=[args.chrome,'--headless','--no-sandbox','--disable-gpu','--disable-background-networking','--user-data-dir='+str(out/'ui-chrome-profile'),'--virtual-time-budget=45000','--dump-dom',f'http://127.0.0.1:{server.server_port}/ui-tests.html']
 run=subprocess.run(cmd,capture_output=True,text=True,timeout=60);(out/'ui-chrome-stderr.txt').write_text(run.stderr)
 if run.returncode:print(run.stderr);sys.exit(run.returncode)
 match=re.search(r'<pre id="test-output"[^>]*>(.*?)</pre>',run.stdout,re.S)
 try:report=json.loads(unescape(match[1]))
 except Exception:(out/'ui-dom.html').write_text(run.stdout);raise
 (out/'ui-results.json').write_text(json.dumps(report,indent=2))
 for r in report.get('results',[]):print(('PASS ' if r['passed'] else 'FAIL ')+r['name']+(('\n'+r['error']) if not r['passed'] else ''))
 if report.get('fatal'):print(report['fatal'])
 sys.exit(1 if report.get('fatal') or any(not r['passed'] for r in report.get('results',[])) else 0)
finally:server.shutdown();server.server_close()
