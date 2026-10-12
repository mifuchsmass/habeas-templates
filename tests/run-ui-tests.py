#!/usr/bin/env python3
"""Run UI checks in disposable Chrome, using explicitly pinned local CDN copies."""
import argparse
import base64
from collections import deque
import hashlib
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import re
import shutil
import signal
import socket
import struct
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from urllib.parse import urlparse
from urllib.request import urlopen


def check(condition, message):
    if not condition:
        raise RuntimeError(message)


def remaining(deadline):
    seconds = deadline - time.monotonic()
    if seconds <= 0:
        raise TimeoutError('Execution deadline exceeded')
    return seconds


class WebSocket:
    """Small synchronous transport for the single CDP connection used here."""
    def __init__(self, endpoint, deadline):
        endpoint = urlparse(endpoint)
        self.sock = socket.create_connection((endpoint.hostname, endpoint.port), remaining(deadline))
        try:
            key = base64.b64encode(os.urandom(16)).decode()
            self.sock.settimeout(remaining(deadline))
            self.sock.sendall((f'GET {endpoint.path} HTTP/1.1\r\nHost: {endpoint.hostname}:{endpoint.port}\r\n'
                               f'Upgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Key: {key}\r\n'
                               'Sec-WebSocket-Version: 13\r\n\r\n').encode())
            header = b''
            while not header.endswith(b'\r\n\r\n'):
                header += self.read(1, deadline)
                check(len(header) <= 16384, 'Oversized WebSocket handshake')
            lines = header.decode('ascii').split('\r\n')
            headers = dict(line.split(':', 1) for line in lines[1:] if ':' in line)
            headers = {key.lower(): value.strip() for key, value in headers.items()}
            expected = base64.b64encode(hashlib.sha1((key + '258EAFA5-E914-47DA-95CA-C5AB0DC85B11').encode()).digest()).decode()
            check(lines[0].split()[1] == '101' and headers.get('upgrade', '').lower() == 'websocket'
                  and 'upgrade' in [token.strip() for token in headers.get('connection', '').lower().split(',')]
                  and headers.get('sec-websocket-accept') == expected, 'Invalid WebSocket handshake')
        except BaseException:
            self.sock.close()
            raise

    def read(self, count, deadline):
        data = bytearray()
        while len(data) < count:
            self.sock.settimeout(remaining(deadline))
            chunk = self.sock.recv(count - len(data))
            if not chunk:
                raise ConnectionError('Chrome debugging connection closed')
            data.extend(chunk)
        return bytes(data)

    def frame(self, opcode, data, deadline):
        mask = os.urandom(4)
        length = len(data)
        header = bytes([128 | opcode, 128 | length]) if length < 126 else (
            bytes([128 | opcode, 254]) + struct.pack('!H', length) if length < 65536 else
            bytes([128 | opcode, 255]) + struct.pack('!Q', length))
        self.sock.settimeout(remaining(deadline))
        self.sock.sendall(header + mask + bytes(value ^ mask[index % 4] for index, value in enumerate(data)))

    def send(self, message, deadline):
        self.frame(1, json.dumps(message).encode(), deadline)

    def receive(self, deadline):
        fragments = bytearray()
        fragmented = False
        while True:
            first, second = self.read(2, deadline)
            opcode, final = first & 15, bool(first & 128)
            check(not first & 112 and not second & 128, 'Unsupported WebSocket frame flags')
            length = second & 127
            if length == 126:
                length = struct.unpack('!H', self.read(2, deadline))[0]
            elif length == 127:
                length = struct.unpack('!Q', self.read(8, deadline))[0]
            check(length + len(fragments) <= 16 * 1024 * 1024, 'Oversized WebSocket message')
            if opcode >= 8:
                check(final and length <= 125, 'Invalid WebSocket control frame')
            data = self.read(length, deadline)
            if opcode == 8:
                check(length != 1, 'Invalid WebSocket close frame')
                self.frame(8, data, deadline)
                raise ConnectionError(f'Chrome closed WebSocket: {data!r}')
            if opcode == 9:
                self.frame(10, data, deadline)
                continue
            if opcode == 10:
                continue
            check(opcode in (0, 1) and (opcode == 0) == fragmented, 'Unexpected WebSocket data frame')
            fragments.extend(data)
            if final:
                return json.loads(fragments.decode('utf-8'))
            fragmented = True


def main():
 parser=argparse.ArgumentParser()
 parser.add_argument('--vendor-dir',default='/tmp')
 parser.add_argument('--artifacts',default='/tmp/habeas-v3-implementation-tests')
 parser.add_argument('--chrome',default='google-chrome')
 parser.add_argument('--interaction-origin',default='',help='Optional actual loopback server for interaction checks')
 parser.add_argument('--timeout',type=float,default=120,help='Overall execution limit in seconds')
 parser.add_argument('--command-timeout',type=float,default=15,help='Discovery and CDP communication limit in seconds')
 args=parser.parse_args()
 check(args.timeout>0 and args.command_timeout>0,'Timeouts must be positive')
 if args.interaction_origin:
  parsed=urlparse(args.interaction_origin)
  check(parsed.scheme in ('http','https') and parsed.hostname in ('127.0.0.1','localhost')
        and not parsed.username and not parsed.password and not parsed.query and not parsed.fragment
        and parsed.path in ('','/'),'Interaction origin must be a loopback HTTP(S) origin')
  args.interaction_origin=args.interaction_origin.rstrip('/')
 root=Path(__file__).resolve().parent.parent;out=Path(args.artifacts);vendor=Path(args.vendor_dir)
 out.mkdir(parents=True,exist_ok=True)
 versions={'pizzip':('3.1.7','dist/pizzip.min.js'),'docxtemplater':('3.50.0','build/docxtemplater.js'),
           'file-saver':('2.0.5','dist/FileSaver.min.js'),'jszip':('3.10.1','dist/jszip.min.js'),
           'docx-preview':('0.3.3','dist/docx-preview.min.js')}
 deadline=time.monotonic()+args.timeout
 report={'results':[]};events=deque(maxlen=60);substitutions=set();server_errors=[]
 stage='startup';server=proc=transport=profile=log=watchdog=None
 def save_report():
  report['runner']={'stage':stage,'substitutions':sorted(substitutions),'browserEvents':list(events)}
  (out/'ui-results.json').write_text(json.dumps(report,indent=2))
 def dependency(url):
  parsed=urlparse(url)
  match=re.fullmatch(r'/npm/([^/@]+)@([^/]+)/(.+)',parsed.path)
  check(parsed.scheme=='https' and parsed.netloc=='cdn.jsdelivr.net' and not parsed.query
        and not parsed.fragment and match is not None,'Unexpected CDN URL: '+url)
  name,version,path=match.groups()
  check(name in versions and versions[name]==(version,path),'Unpinned CDN dependency: '+url)
  if url not in substitutions:
   substitutions.add(url);print('SUBSTITUTE '+url+' -> pinned local copy',flush=True)
  return name,(vendor/f'habeas-review-{name}-{version}.js').read_bytes()
 class Handler(BaseHTTPRequestHandler):
  def log_message(self,*unused):pass
  def do_GET(self):
   try:
    path=urlparse(self.path).path
    if path=='/ui-tests.html':
     data=('<meta charset="UTF-8"><pre id="test-output">pending</pre><script>'+(root/'tests/ui.test.js').read_text()+'</script>').encode();mime='text/html'
    elif path.startswith('/vendor/'):
     name=path[8:].removesuffix('.js');check(name in versions,'Unknown local dependency')
     data=(vendor/f'habeas-review-{name}-{versions[name][0]}.js').read_bytes();mime='text/javascript'
    else:
     directory=out if path.startswith('/artifacts/') else root
     relative=path.removeprefix('/artifacts/') if path.startswith('/artifacts/') else path.lstrip('/')
     target=(directory/relative).resolve()
     if not target.is_relative_to(directory.resolve()) or not target.is_file():self.send_error(404);return
     data=target.read_bytes();mime={'html':'text/html','js':'text/javascript','css':'text/css'}.get(target.suffix[1:],'application/octet-stream')
     if target.suffix=='.html':
      data=re.sub(r'https://cdn\.jsdelivr\.net/[^"\s<>]+',lambda match:'/vendor/'+dependency(match[0])[0]+'.js',data.decode()).encode()
    self.send_response(200);self.send_header('Content-Type',mime);self.send_header('Content-Length',str(len(data)));self.end_headers();self.wfile.write(data)
   except Exception as error:
    server_errors.append(str(error));self.send_error(500,str(error))
 def stop_browser():
  if proc is not None:
   try:
    if os.name=='posix':os.killpg(proc.pid,signal.SIGKILL)
    elif proc.poll() is None:proc.kill()
   except ProcessLookupError:pass
   except OSError as error:events.append({'browserCleanupError':str(error)})
 def snapshot():
  try:
   dom=evaluate('document.documentElement.outerHTML',limit=min(deadline,time.monotonic()+2))
   (out/'ui-dom.html').write_text(dom or '')
  except Exception as error:events.append({'snapshotError':str(error)})
 try:
  (out/'ui-dom.html').unlink(missing_ok=True)
  for name,(version,_) in versions.items():
   check((vendor/f'habeas-review-{name}-{version}.js').is_file(),'Missing pinned local library: '+name)
  server=ThreadingHTTPServer(('127.0.0.1',0),Handler)
  threading.Thread(target=server.serve_forever,daemon=True).start()
  profile=Path(tempfile.mkdtemp(prefix='habeas-ui-'))
  log=(out/'ui-chrome-stderr.txt').open('w')
  proc=subprocess.Popen([args.chrome,'--headless','--no-sandbox','--disable-gpu','--disable-background-networking','--no-first-run',
       '--remote-debugging-port=0','--remote-debugging-address=127.0.0.1','--user-data-dir='+str(profile),
       '--proxy-server=http://127.0.0.1:9','--proxy-bypass-list=127.0.0.1;localhost','about:blank'],
       stdout=log,stderr=log,start_new_session=os.name=='posix')
  watchdog=threading.Timer(remaining(deadline),stop_browser);watchdog.daemon=True;watchdog.start()
  startup=min(deadline,time.monotonic()+args.command_timeout)
  while not (profile/'DevToolsActivePort').exists():
   remaining(startup);check(proc.poll() is None,'Chrome exited during startup; see ui-chrome-stderr.txt');time.sleep(.05)
  port=int((profile/'DevToolsActivePort').read_text().splitlines()[0])
  def discovery(path):
   with urlopen(f'http://127.0.0.1:{port}/{path}',timeout=remaining(startup)) as response:return json.load(response)
  report['browser']=discovery('json/version')['Browser']
  targets=discovery('json')
  transport=WebSocket(next(target['webSocketDebuggerUrl'] for target in targets if target['type']=='page'),startup)
  seq=0;pending=set()
  def call(method,params,limit=None):
   nonlocal seq
   limit=min(deadline,time.monotonic()+args.command_timeout,limit if limit is not None else deadline)
   seq+=1;request=seq;transport.send({'id':request,'method':method,'params':params},limit)
   while True:
    response=transport.receive(limit)
    if response.get('method')=='Fetch.requestPaused':
     paused=response['params'];_,body=dependency(paused['request']['url'])
     seq+=1;pending.add(seq)
     transport.send({'id':seq,'method':'Fetch.fulfillRequest','params':{'requestId':paused['requestId'],'responseCode':200,
        'responseHeaders':[{'name':'Content-Type','value':'text/javascript'}],'body':base64.b64encode(body).decode()}},limit)
    elif response.get('method') in ('Runtime.exceptionThrown','Runtime.consoleAPICalled','Log.entryAdded','Network.loadingFailed'):
     events.append(response)
    if response.get('id') in pending:
     pending.remove(response['id']);check('error' not in response,'Dependency interception failed: '+str(response))
    if response.get('id')==request:
     check('error' not in response,method+' failed: '+str(response));return response
  def evaluate(expression,limit=None):
   response=call('Runtime.evaluate',{'expression':expression,'returnByValue':True},limit)['result']
   check('exceptionDetails' not in response,'Browser evaluation failed: '+str(response));return response['result'].get('value')
  def until(expression):
   limit=min(deadline,time.monotonic()+30)
   while True:
    remaining(limit);check(not server_errors,'Local server failed: '+str(server_errors))
    value=evaluate(expression,limit)
    if value:return value
    time.sleep(min(.1,remaining(limit)))
  origin=f'http://127.0.0.1:{server.server_port}'
  for method in ('Runtime.enable','Log.enable','Network.enable'):call(method,{})
  call('Fetch.enable',{'patterns':[{'urlPattern':'https://cdn.jsdelivr.net/*'}]})
  stage='34 JavaScript UI checks'
  call('Page.navigate',{'url':origin+'/ui-tests.html'})
  raw=until("document.getElementById('test-output')?.textContent!=='pending' && document.getElementById('test-output')?.textContent")
  report.update(json.loads(raw));save_report();snapshot()
  check(len(report.get('results',[]))==34,'Expected all 34 JavaScript UI checks to finish')
  # Trusted keyboard events operate native select controls on the actual served pages.
  def keypress(key,code):
   for kind in ['keyDown','keyUp']:call('Input.dispatchKeyEvent',{'type':kind,'key':key,'code':key,'windowsVirtualKeyCode':code})
  def click_element(element):
   evaluate(f"document.getElementById('{element}').scrollIntoView({{block:'center'}})")
   point=json.loads(evaluate(f"JSON.stringify((()=>{{const r=document.getElementById('{element}').getBoundingClientRect();return {{x:r.x+r.width/2,y:r.y+r.height/2}};}})())"))
   for kind in ['mousePressed','mouseReleased']:call('Input.dispatchMouseEvent',{'type':kind,'x':point['x'],'y':point['y'],'button':'left','clickCount':1})
  def select_index(element,index):
   evaluate(f"document.getElementById('{element}').focus()")
   keypress('Home',36)
   for _ in range(index):keypress('ArrowDown',40)
   keypress('Tab',9)
  interaction_origin=args.interaction_origin or origin
  for page in ['index.html','drafting-studio.html']:
   call('Page.navigate',{'url':interaction_origin+'/'+page})
   if page=='index.html':until("location.pathname.endsWith('/index.html') && document.readyState==='complete' && document.getElementById('renderBtn')?.disabled===false")
   else:
    until("location.pathname.endsWith('/drafting-studio.html') && document.readyState==='complete' && !!document.getElementById('petitioner_first_name')")
    evaluate("document.getElementById('accessPin').focus()")
    call('Input.insertText',{'text':'Habeas123!'});click_element('unlockBtn')
    until("getComputedStyle(document.getElementById('auth-overlay')).display==='none'")
    # Pasted inspection activates the sandbox through normal keyboard interaction.
    evaluate("document.getElementById('snippetInput').focus()")
    call('Input.insertText',{'text':'Synthetic inspection <Review Field>'});click_element('inspectSnippetBtn')
    until("getComputedStyle(document.getElementById('sandbox')).display!=='none'")
   for index,state in enumerate(['Massachusetts','Connecticut','Maine','New Hampshire','Rhode Island','Vermont'],1):
    stage=f'{page}: {state}'; snapshot()
    name=f'trusted keyboard {page}: {state} Other visibility and selection'
    try:
     select_index('jurisdiction',index)
     if state not in ['Connecticut','Maine']:
      count=evaluate("document.getElementById('detention_facility').options.length")
      select_index('detention_facility',count-1)
     actual=evaluate("JSON.stringify({state:document.getElementById('jurisdiction').value,facility:document.getElementById('detention_facility').value,visible:document.getElementById('other_detention_location').getBoundingClientRect().height>0,placeholder:document.getElementById('other_detention_location').placeholder})")
     actual=json.loads(actual)
     check(actual=={'state':state,'facility':'__other__','visible':True,'placeholder':'Facility or ICE office name and/or location'},f'Unexpected selection or visibility: {actual}')
     expected={'Massachusetts':['','Plymouth','Burlington','Boston','__other__'],'Connecticut':['__other__'],'Maine':['__other__'],'New Hampshire':['','Berlin','Strafford','__other__'],'Rhode Island':['','Wyatt','__other__'],'Vermont':['','Northwest','__other__']}
     options=json.loads(evaluate("JSON.stringify([...document.getElementById('detention_facility').options].map(option=>option.value))"))
     check(options==expected[state],f'Unexpected options: {options}')
     if state in ['Connecticut','Maine']:
      select_index('detention_facility',0)
      check(evaluate("document.getElementById('detention_facility').value==='__other__' && document.getElementById('other_detention_location').getBoundingClientRect().height>0"),'CT/ME custom location is not visible')
     if state not in ['Connecticut','Maine']:
      select_index('detention_facility',1)
      check(evaluate("document.getElementById('other_detention_location').getBoundingClientRect().height===0"),'Named facility did not hide custom location')
     report.setdefault('results',[]).append({'name':name,'passed':True})
    except Exception as error:
     report.setdefault('results',[]).append({'name':name,'passed':False,'error':str(error)})
     snapshot()
    save_report()
  check(not server_errors,'Local server failed: '+str(server_errors))
  check(len(report.get('results',[]))==46,'Expected all 46 UI checks to finish')
 except Exception as error:
  report['fatal']=str(error)
  report['traceback']=traceback.format_exc()
  if transport is not None:snapshot()
 finally:
  if watchdog is not None:watchdog.cancel()
  if transport is not None:
   try:transport.sock.close()
   except OSError as error:events.append({'connectionCleanupError':str(error)})
  stop_browser()
  if proc is not None:
   try:proc.wait(timeout=5)
   except subprocess.TimeoutExpired:report['fatal']='Chrome did not exit after cleanup'
  if server is not None:
   try:server.shutdown();server.server_close()
   except OSError as error:report['fatal']='Test server cleanup failed: '+str(error)
  if log is not None:
   try:log.close()
   except OSError as error:report['fatal']='Browser log cleanup failed: '+str(error)
  if profile is not None:
   try:shutil.rmtree(profile)
   except OSError as error:report['fatal']='Temporary profile cleanup failed: '+str(error)
  save_report()
 for result in report.get('results',[]):
  print(('PASS ' if result['passed'] else 'FAIL ')+result['name']+(' '+result.get('error','') if not result['passed'] else ''))
 if report.get('fatal'):print('FATAL '+str(report['fatal']))
 passed=sum(bool(result['passed']) for result in report.get('results',[]))
 print(f"{passed} passed; {len(report.get('results',[]))-passed} failed")
 return int(bool(report.get('fatal') or any(not result['passed'] for result in report.get('results',[]))))


if __name__=='__main__':
 sys.exit(main())
