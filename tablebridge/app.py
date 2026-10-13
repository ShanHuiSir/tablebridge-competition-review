from __future__ import annotations
import base64, io, json, os, re, uuid, threading, hashlib
from pathlib import Path
from datetime import datetime, timezone
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from zipfile import ZipFile, ZIP_DEFLATED
from core import ContractError, read_table, suggest, SCHEMAS, OUTPUTS, LABELS, execute, export_template, csv_bytes, clean
from planner import create_plan, config
from universal_api import handle as universal_handle

ROOT=Path(__file__).resolve().parent
DATA=Path(os.getenv('BRIDGE_DATA_DIR',str(ROOT/'data')))
SESSIONS={}; LOCK=threading.Lock()

def public_table(table):
 return {k:table[k] for k in ('filename','sheet','headers','header_row','sha256')} | {'row_count':len(table['rows']),'preview':[{'row':r['row'],'values':{k:clean(v) for k,v in r['values'].items()},'formula':r['formula']} for r in table['rows'] if not r['blank']][:4]}

class Handler(BaseHTTPRequestHandler):
 def log_message(self,fmt,*args): pass
 def respond(self,status,obj):
  content=json.dumps(obj,ensure_ascii=False).encode(); self.send_response(status); self.send_header('Content-Type','application/json; charset=utf-8'); self.send_header('Content-Length',str(len(content))); self.send_header('Cache-Control','no-store'); self.end_headers(); self.wfile.write(content)
 def authorized(self):
  host=self.headers.get('Host',''); origin=self.headers.get('Origin')
  expected=f'127.0.0.1:{self.server.server_port}'
  if host not in (expected,f'localhost:{self.server.server_port}'): raise ContractError('仅接受本机 Host')
  if origin and origin not in ('http://'+expected,f'http://localhost:{self.server.server_port}'): raise ContractError('拒绝跨站请求')
 def do_GET(self):
  try:
   self.authorized()
   if self.path=='/api/config': return self.respond(200,{'model':config(),'schemas':SCHEMAS,'outputs':OUTPUTS,'labels':LABELS})
   if self.path.startswith('/download/'):
    m=re.fullmatch(r'/download/([a-f0-9]{32})/(import.xlsx|import.xls|import.csv|workflow.py|audit.json|exceptions.csv|ledger.csv|result.zip)',self.path)
    if not m: return self.respond(404,{'error':'文件不存在'})
    path=DATA/m[1]/m[2]
    if not path.is_file(): return self.respond(404,{'error':'文件不存在'})
    raw=path.read_bytes(); types={'xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','xls':'application/vnd.ms-excel','py':'text/plain; charset=utf-8','json':'application/json','csv':'application/octet-stream','zip':'application/zip'}
    self.send_response(200);self.send_header('Content-Type',types[path.suffix[1:]]);self.send_header('Content-Disposition',f'attachment; filename="{m[2]}"');self.send_header('Content-Length',str(len(raw)));self.end_headers();self.wfile.write(raw);return
   path={'/':'workbench.html','/legacy':'index.html','/app.js':'app.js','/style.css':'style.css','/workbench.js':'workbench.js','/workbench.css':'workbench.css'}.get(self.path)
   if not path: return self.respond(404,{'error':'页面不存在'})
   data=(ROOT/'static'/path).read_bytes(); self.send_response(200);self.send_header('Content-Type',{'html':'text/html; charset=utf-8','js':'text/javascript; charset=utf-8','css':'text/css; charset=utf-8'}[path.rsplit('.',1)[1]]);self.send_header('Content-Length',str(len(data)));self.send_header('X-Content-Type-Options','nosniff');self.send_header('Content-Security-Policy',"default-src 'self'; style-src 'self'; script-src 'self'; connect-src 'self'; img-src 'self' data:; frame-ancestors 'none'");self.end_headers();self.wfile.write(data)
  except ContractError as e: self.respond(403,{'error':str(e)})
 def do_POST(self):
  try:
   self.authorized(); n=int(self.headers.get('Content-Length','0'))
   if n<=0 or n>58*1024*1024: raise ContractError('请求须为 JSON，且不得超过 58 MB')
   body=json.loads(self.rfile.read(n))
   if self.path=='/api/session':
    sid=uuid.uuid4().hex
    with LOCK:
     if len(SESSIONS)>=30: raise ContractError('已达到 30 个会话上限，请重启本机服务')
     SESSIONS[sid]={'tables':{},'raw':{},'busy':False}
    return self.respond(200,{'session_id':sid})
   sid=body.get('session_id'); session=SESSIONS.get(sid)
   if not session: raise ContractError('会话不存在或服务已重启，请刷新页面')
   if self.path.startswith('/api/universal/'):
    with LOCK:
     if session['busy']: raise ContractError('当前会话正在处理，请稍后再试')
     session['busy']=True
    try:
     return self.respond(200,universal_handle(self.path[len('/api/universal/'):],body,session,DATA))
    finally:
     with LOCK: session['busy']=False
   if self.path in ('/api/upload','/api/sample'):
    with LOCK:
     if session['busy']: raise ContractError('当前会话正在计算，请稍后替换文件')
     if self.path=='/api/sample':
      mode=body.get('operation','leave')
      if mode not in OUTPUTS: raise ContractError('未知样例')
      files={'leave':('OA_请假_模拟.csv',1),'balance':('HR_年假余额_模拟.csv',1),'template':('HR_请假导入模板_模拟.xlsx',1)} if mode=='leave' else {'overtime':('ERP_加班_模拟.csv',1),'template':('HR_加班导入模板_模拟.xlsx',1)}
      newtables={};newraw={}
      for role,(fn,h) in files.items():
       raw=(ROOT/'samples'/fn).read_bytes();newtables[role]=read_table(raw,fn,h,role=='template');newraw[role]=raw
      session['tables']=newtables;session['raw']=newraw
     else:
      role=body.get('role')
      if role not in (*SCHEMAS,'template'): raise ContractError('未知文件角色')
      fn=Path(body.get('filename','')).name; raw=base64.b64decode(body.get('data',''),validate=True)
      table=read_table(raw,fn,body.get('header_row',1),role=='template')
      session['tables'][role]=table;session['raw'][role]=raw
    return self.respond(200,{'tables':{r:public_table(t) for r,t in session['tables'].items()},'mappings':{r:suggest(t['headers'],SCHEMAS[r])[0] for r,t in session['tables'].items() if r in SCHEMAS},'target_suggestions':self.target_suggestions(session,body.get('operation','leave'))})
   if self.path=='/api/plan':
    return self.respond(200,create_plan(str(body.get('prompt',''))[:4000],session['tables'],bool(body.get('use_model'))))
   if self.path=='/api/run':
    if body.get('confirmed') is not True: raise ContractError('请明确确认映射、业务规则和目标模板')
    with LOCK:
     if session['busy']: raise ContractError('已有运行进行中')
     session['busy']=True
     tables=dict(session['tables']);raw=dict(session['raw'])
    try:
     if 'template' not in tables: raise ContractError('缺少目标导入模板')
     plan=body.get('plan',{}); result=execute({r:t for r,t in tables.items() if r in SCHEMAS},body.get('mappings',{}),plan)
     template=export_template(raw['template'],tables['template'],body.get('target_mapping',{}),result)
     run_id=uuid.uuid4().hex; dest=DATA/run_id;dest.mkdir(parents=True)
     result.update(run_id=run_id,created_at=datetime.now(timezone.utc).isoformat(),template=public_table(tables['template']),target_mapping=body['target_mapping'],planning=body.get('planning',{'mode':'人工配置'}),import_sha256=hashlib.sha256(template).hexdigest())
     result['template'].pop('preview',None)
     fields=['source','sheet','row','employee_id','code','reason','values']
     files={'import.xlsx':template,'audit.json':json.dumps(result,ensure_ascii=False,indent=2).encode(),'exceptions.csv':csv_bytes(result['errors'],fields),'ledger.csv':csv_bytes(result['ledger'],list(result['ledger'][0]) if result['ledger'] else ['employee_id','record_id','source_row'])}
     for fn,value in files.items(): (dest/fn).write_bytes(value)
     buffer=io.BytesIO()
     with ZipFile(buffer,'w',ZIP_DEFLATED) as z:
      for fn,value in files.items(): z.writestr(fn,value)
     (dest/'result.zip').write_bytes(buffer.getvalue())
     result['downloads']={fn:f'/download/{run_id}/{fn}' for fn in (*files,'result.zip')}
     return self.respond(200,result)
    finally:
     with LOCK: session['busy']=False
   raise ContractError('接口不存在')
  except ContractError as e: self.respond(400,{'error':str(e)})
  except (ValueError,KeyError,TypeError,OverflowError) as e: self.respond(400,{'error':'请求格式或字段类型错误，请核对文件及配置'})
  except Exception as e:
   print('Request failed:',type(e).__name__,flush=True); self.respond(500,{'error':'内部处理错误，当前运行未标记成功；请查看服务终端'})
 def target_suggestions(self,session,mode):
  table=session['tables'].get('template')
  if not table or mode not in OUTPUTS:return {}
  field_map,_=suggest(table['headers'],OUTPUTS[mode]);return {h:{'field':f} for f,h in field_map.items()}

if __name__=='__main__':
 port=int(os.getenv('BRIDGE_PORT','8877')); server=ThreadingHTTPServer(('127.0.0.1',port),Handler)
 print(f'表桥已启动 http://127.0.0.1:{port}  按 Ctrl+C 停止',flush=True)
 try:server.serve_forever()
 except KeyboardInterrupt:server.server_close()
