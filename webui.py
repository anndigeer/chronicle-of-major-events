#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""银河麒麟/龙芯本地网页操作台。仅用Python标准库提供网页；采集仍调用main.py。"""
from __future__ import annotations

import argparse
import hmac
from kylin_check import diagnose
import json
import os
from pathlib import Path
import re
import secrets
import subprocess
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs
import webbrowser

ROOT = Path(__file__).resolve().parent
MONTH = re.compile(r'20\d{2}-(?:0[1-9]|1[0-2])\Z')
MODES = {'all': '', 'sync': '--sync', 'export': '--export',
         'verify': '--verify', 'check-excel': '--check-excel',
         'check-names': '--check-names', 'update-leaders': '--update-leaders',
         'diagnose': ''}

PAGE = r'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>市委大事记 · 新闻采集工作台</title>
<style>
:root{font-family:"Noto Sans CJK SC",system-ui,sans-serif;background:#f4f6fa;color:#203048}
*{box-sizing:border-box}main{max-width:1030px;margin:auto;padding:24px 15px}
h1{font-size:26px;margin:0}h2{font-size:19px}.muted{color:#5a6d80}
.panel{background:#fff;border:1px solid #dce3ee;border-radius:14px;padding:20px;margin:16px 0}
.actions{display:flex;flex-wrap:wrap;gap:10px;margin:16px 0}
button{font:inherit;padding:10px 15px;background:white;border:1px solid #aab9ca;
border-radius:9px;color:#193b63;cursor:pointer}button.primary{background:#1761a8;color:white;border-color:#1761a8}
button:disabled{opacity:.45;cursor:wait}button:focus-visible{outline:3px solid #80adf2}
input{font:inherit;padding:9px;border:1px solid #aab9ca;border-radius:8px}
pre{background:#16273b;color:#e9f3ff;padding:16px;white-space:pre-wrap;word-break:break-word;
overflow:auto;min-height:130px;max-height:420px;border-radius:9px;font-size:13px;line-height:1.7}
.note{font-size:13px;line-height:1.8;color:#566979}#status{font-weight:700}
@media(max-width:600px){.actions button{width:100%}}
</style></head><body><main><h1>市委大事记 · 新闻采集工作台</h1>
<p class="muted">银河麒麟 / 龙芯 3A4000 · 仅在本机浏览器访问</p>
<section class="panel"><h2>① 选择月份并处理新闻</h2>
<label>目标月份 <input type="month" id="month" min="2000-01" max="2099-12"></label>
<div class="actions">
<button class="primary" data-run="all">增量采集 + 更新 Word</button>
<button data-run="sync">仅增量同步 Excel</button>
<button data-run="export">按人工收录/排除重建 Word</button>
<button data-run="verify">核验成果</button></div>
<p class="note">先采集，再打开 Excel：E 列选「收录/排除」，H 列核实事件日期。
保存并关闭 Excel、Word 后，点击「重建 Word」。只同步 Excel 不会改 Word。</p></section>
<section class="panel"><h2>② 文件与姓名</h2><div class="actions">
<button data-open="excel">打开 Excel</button><button data-open="word">打开本月 Word</button>
<button data-open="folder">打开结果目录</button><button data-open="names">编辑 names.txt</button>
<button data-run="check-names">检查名单</button><button data-run="check-excel">检查 Excel</button>
<button data-run="update-leaders">联网更新领导名单</button><button data-run="diagnose">检测电脑环境</button></div>
<p class="note" id="files">正在检查结果文件……</p><p class="note" id="environment"></p></section>
<section class="panel"><h2>③ 执行日志　<span id="status">未运行</span></h2>
<pre id="log">尚未运行任务。</pre><button id="refresh">刷新日志</button>
<p class="note">服务仅监听 127.0.0.1，不会向局域网开放。
关闭启动服务的终端或按 Ctrl+C 可停止界面。</p></section></main>
<script>
const TOKEN='__TOKEN__';
const el=id=>document.getElementById(id);
const headers={'X-UI-Token':TOKEN,'Content-Type':'application/json'};
const d=new Date();d.setDate(1);d.setMonth(d.getMonth()-1);
el('month').value=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
async function request(url,method='GET',data){
 const r=await fetch(url,{method,headers,body:data?JSON.stringify(data):undefined,cache:'no-store'});
 const v=await r.json();if(!r.ok)throw Error(v.error||'请求失败');return v;
}
async function refresh(){
 try{
  const month=el('month').value;
  const s=await request('/api/status?month='+encodeURIComponent(month));
  el('status').textContent=s.running?'（正在运行）':s.code===null?'（尚未运行）':
   s.code===0?'（最近一次成功）':'（最近一次失败，请看日志）';
  el('log').textContent=s.log||'尚未运行任务。';
  el('environment').textContent=s.environment.ready?'运行环境：采集功能已就绪':'运行环境：需要处理 '+s.environment.missing.join('、')+'。可点击「检测电脑环境」查看原因。';
  el('files').textContent='Excel：'+(s.files.excel?'已生成':'暂无')+
   '｜本月 Word：'+(s.files.word?'已生成':'暂无')+
   '｜链接待复核：'+(s.files.review?'有，请处理':'无');
  document.querySelectorAll('[data-run],[data-open]').forEach(b=>{
   b.disabled=s.running || (b.dataset.run && b.dataset.run!=='diagnose' && !s.environment.ready);
  });
 }catch(e){el('status').textContent='（连接失败：'+e.message+'）';}
}
async function run(mode){
 try{await request('/api/run','POST',{month:el('month').value,mode});await refresh();}
 catch(e){alert(e.message);}
}
async function openFile(kind){
 try{await request('/api/open','POST',{month:el('month').value,kind});}
 catch(e){alert(e.message);}
}
document.querySelectorAll('[data-run]').forEach(b=>b.onclick=()=>run(b.dataset.run));
document.querySelectorAll('[data-open]').forEach(b=>b.onclick=()=>openFile(b.dataset.open));
el('refresh').onclick=refresh;setInterval(refresh,1600);refresh();
</script></body></html>'''


class Jobs:
    """Keep a single subprocess alive; never run concurrent writes to the same Excel."""
    def __init__(self, root: Path, python: str | None = None):
        self.root = Path(root)
        self.python = python or sys.executable
        self.lock = threading.Lock()
        self.running = False
        self.mode = ''
        self.code = None
        self.job_id = 0
        self.month = ''
        self.lines = []
        result = diagnose(self.root)
        missing = [p['name'] for p in result['packages'] if not p['ok']]
        if not result['python_ok']:
            missing.insert(0, 'Python 3.10+')
        self.environment = {'ready': result['core_ready'], 'missing': missing,
                            'architecture': result['arch_description'],
                            'python': result['python']}

    def start(self, month: str, mode: str):
        if not MONTH.fullmatch(month) or mode not in MODES:
            raise ValueError('请提供正确月份和操作模式')
        with self.lock:
            if self.running:
                raise RuntimeError('已有任务正在运行，请先等待其结束')
            if not self.environment['ready'] and mode != 'diagnose':
                raise RuntimeError('Python或采集依赖未满足，先点击「检测电脑环境」查看原因')
            if mode == 'diagnose':
                command = [self.python, '-u', str(Path(__file__).with_name('kylin_check.py')),
                           '--root', str(self.root)]
            else:
                command = [self.python, '-u', str(self.root / 'main.py')]
            if mode in {'all', 'sync', 'export', 'verify'}:
                command.append(month)
            if MODES[mode]:
                command.append(MODES[mode])
            self.job_id += 1
            self.running, self.mode, self.code, self.month = True, mode, None, month
            self.lines = ['执行：'+' '.join(command)]
        threading.Thread(target=self._worker, args=(command,), daemon=True).start()

    def _worker(self, command):
        code = 1
        try:
            env = os.environ.copy()
            env['PYTHONUNBUFFERED'] = '1'
            env['PYTHONIOENCODING'] = 'utf-8'
            p = subprocess.Popen(command, cwd=self.root, env=env,
                                 stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                 text=True, encoding='utf-8', errors='replace')
            assert p.stdout is not None
            for line in p.stdout:
                with self.lock:
                    self.lines.append(line.rstrip('\r\n'))
                    self.lines = self.lines[-400:]
            code = p.wait()
        except Exception as exc:
            with self.lock:
                self.lines.append('执行异常：'+str(exc))
        finally:
            with self.lock:
                self.code, self.running = code, False
                self.lines.append('进程结束，退出码：'+str(code))

    def status(self, month: str):
        if not MONTH.fullmatch(month):
            month = self.month if MONTH.fullmatch(self.month) else '2026-09'
        y, m = map(int, month.split('-'))
        folder = self.root/'results'
        word = folder/f'市委大事记（{y}年{m}月）.docx'
        with self.lock:
            return {'running':self.running,'job_id':self.job_id,
                    'mode':self.mode,'code':self.code,
                    'log':'\n'.join(self.lines),
                    'environment':self.environment,
                    'files':{'excel':(folder/'jiujiang_news.xlsx').is_file(),
                             'word':word.is_file(),
                             'review':(folder/'链接冲突待复核.csv').is_file()}}


class LocalHandler(BaseHTTPRequestHandler):
    jobs: Jobs
    token: str
    page: str

    def log_message(self, fmt, *args):
        pass

    def respond(self, status, obj):
        data=json.dumps(obj,ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type','application/json; charset=utf-8')
        self.send_header('Cache-Control','no-store')
        self.send_header('X-Content-Type-Options','nosniff')
        self.send_header('Content-Length',str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def authorized(self):
        return hmac.compare_digest(self.headers.get('X-UI-Token',''), self.token)

    def do_GET(self):
        path=urlparse(self.path)
        if path.path=='/':
            page=self.page.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type','text/html; charset=utf-8')
            self.send_header('Content-Security-Policy',
                             "default-src 'none'; script-src 'unsafe-inline'; "
                             "style-src 'unsafe-inline'; connect-src 'self'; base-uri 'none'")
            self.send_header('Cache-Control','no-store')
            self.send_header('Content-Length',str(len(page)))
            self.end_headers()
            self.wfile.write(page)
        elif not self.authorized():
            self.respond(403, {'error':'会话已过期，请刷新页面'})
        elif path.path=='/api/status':
            month=parse_qs(path.query).get('month',[''])[0]
            self.respond(200,self.jobs.status(month))
        else:
            self.respond(404,{'error':'未找到接口'})

    def do_POST(self):
        if not self.authorized():
            return self.respond(403,{'error':'会话已过期，请刷新页面'})
        if self.headers.get('Content-Type','').split(';')[0]!='application/json':
            return self.respond(415,{'error':'仅支持 JSON 请求'})
        try:
            length=int(self.headers.get('Content-Length','-1'))
            if not 0 <= length <= 4096:
                return self.respond(413,{'error':'提交内容超出限制'})
            body=json.loads(self.rfile.read(length).decode('utf-8'))
            if not isinstance(body,dict):
                raise ValueError('参数必须为对象')
            path=urlparse(self.path).path
            month=body.get('month','')
            if path=='/api/run':
                self.jobs.start(month,body.get('mode',''))
                return self.respond(202,{'ok':True})
            if path=='/api/open':
                kind=body.get('kind')
                if kind not in {'excel','word','folder','names'} or not MONTH.fullmatch(month):
                    raise ValueError('文件类型或月份不合法')
                y,m=map(int,month.split('-'))
                root=self.jobs.root
                mapping={'excel':root/'results'/'jiujiang_news.xlsx',
                         'word':root/'results'/f'市委大事记（{y}年{m}月）.docx',
                         'folder':root/'results','names':root/'names.txt'}
                target=mapping[kind]
                if not target.exists():
                    return self.respond(404,{'error':'文件尚未生成：'+str(target)})
                if sys.platform.startswith('linux'):
                    subprocess.Popen(['xdg-open',str(target)],
                                     stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
                elif sys.platform=='win32':
                    os.startfile(str(target))
                else:
                    subprocess.Popen(['open',str(target)])
                return self.respond(200,{'ok':True})
            self.respond(404,{'error':'未知操作'})
        except (ValueError,TypeError,UnicodeError) as exc:
            self.respond(400,{'error':str(exc)})
        except RuntimeError as exc:
            self.respond(409,{'error':str(exc)})
        except OSError as exc:
            self.respond(500,{'error':str(exc)})


def launch(argv=None, *, root=None, open_browser=True):
    parser=argparse.ArgumentParser(description='银河麒麟本地网页图形操作台')
    parser.add_argument('--port',type=int,default=0,help='本机监听端口；默认使用空闲端口')
    parser.add_argument('--no-browser',action='store_true',help='不自动打开浏览器')
    args=parser.parse_args(argv)
    if not 0 <= args.port <= 65535:
        parser.error('端口范围必须为 0～65535')
    token=secrets.token_urlsafe(24)
    jobs=Jobs(root or ROOT)
    platform_label = ('Windows · 仅在本机浏览器访问'
                      if sys.platform == 'win32'
                      else '银河麒麟 / 龙芯 3A4000 · 仅在本机浏览器访问')
    page = PAGE.replace('__TOKEN__', token).replace(
        '银河麒麟 / 龙芯 3A4000 · 仅在本机浏览器访问', platform_label)
    handler=type('ChronicleHandler',(LocalHandler,),
                 {'jobs':jobs,'token':token,'page':page})
    server=ThreadingHTTPServer(('127.0.0.1',args.port),handler)
    url='http://127.0.0.1:'+str(server.server_port)+'/'
    print('已启动市委大事记网页操作台：'+url,flush=True)
    print('仅本机可访问。按 Ctrl+C 关闭。',flush=True)
    if open_browser and not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n网页操作台已停止')
    finally:
        server.server_close()


if __name__=='__main__':
    launch()
