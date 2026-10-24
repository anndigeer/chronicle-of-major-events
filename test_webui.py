"""本地网页操作台离线回归测试：不需要第三方 Python 依赖。"""
import json
import sys
import tempfile
import threading
import time
import unittest
from http.server import ThreadingHTTPServer
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import webui


class TestKylinWebUI(unittest.TestCase):
    def test_months_and_modes_whitelisted(self):
        self.assertTrue(webui.MONTH.fullmatch('2026-09'))
        for value in ('2026-13', '2026-00', '2026-09/../../x', 'bad'):
            self.assertFalse(webui.MONTH.fullmatch(value))
        with tempfile.TemporaryDirectory() as d:
            worker=webui.Jobs(Path(d))
            with self.assertRaises(ValueError):
                worker.start('../2026-09','all')
            with self.assertRaises(ValueError):
                worker.start('2026-09','shell')

    def test_process_uses_exact_mode_and_runs_once(self):
        with tempfile.TemporaryDirectory() as d:
            root=Path(d)
            (root/'main.py').write_text(
                "import sys,time\nprint('参数='+str(sys.argv[1:]),flush=True)\n"
                "time.sleep(0.1)\n",encoding='utf-8')
            worker=webui.Jobs(root, sys.executable)
            worker.start('2026-09','export')
            with self.assertRaises(RuntimeError):
                worker.start('2026-09','sync')
            for _ in range(100):
                if not worker.status('2026-09')['running']:
                    break
                time.sleep(0.02)
            status=worker.status('2026-09')
            self.assertFalse(status['running'])
            self.assertEqual(status['code'],0)
            self.assertIn("2026-09",status['log'])
            self.assertIn("--export",status['log'])
            self.assertFalse((root/'results').exists())

    def test_local_http_session_guard_and_no_external_access(self):
        with tempfile.TemporaryDirectory() as d:
            manager=webui.Jobs(Path(d))
            token='test-local-secret'
            handler=type('TestHandler',(webui.LocalHandler,),
                         {'jobs':manager,'token':token,
                          'page':webui.PAGE.replace('__TOKEN__',token)})
            server=ThreadingHTTPServer(('127.0.0.1',0),handler)
            thread=threading.Thread(target=server.serve_forever,daemon=True)
            thread.start()
            try:
                self.assertEqual(server.server_address[0],'127.0.0.1')
                origin='http://127.0.0.1:'+str(server.server_port)
                with urlopen(origin+'/',timeout=4) as response:
                    html=response.read().decode('utf-8')
                self.assertIn('本机浏览器访问',html)
                self.assertIn(token,html)
                with self.assertRaises(HTTPError) as caught:
                    urlopen(origin+'/api/status',timeout=4)
                self.assertEqual(caught.exception.code,403)
                req=Request(origin+'/api/status?month=2026-09',
                            headers={'X-UI-Token':token})
                with urlopen(req,timeout=4) as response:
                    status=json.load(response)
                self.assertFalse(status['running'])
                self.assertFalse(status['files']['excel'])
                # Remote websites cannot fire commands without session token.
                req=Request(origin+'/api/run',data=json.dumps(
                    {'month':'2026-09','mode':'all'}).encode('utf-8'),
                    headers={'Content-Type':'application/json'},method='POST')
                with self.assertRaises(HTTPError) as caught:
                    urlopen(req,timeout=4)
                self.assertEqual(caught.exception.code,403)
                self.assertFalse((Path(d)/'results').exists())
            finally:
                server.shutdown()
                server.server_close()
                thread.join(timeout=4)

    def test_no_dynamic_filesystem_paths_from_requests(self):
        with tempfile.TemporaryDirectory() as d:
            manager=webui.Jobs(Path(d))
            token='test-secret'
            handler=type('TestHandler',(webui.LocalHandler,),
                         {'jobs':manager,'token':token,'page':webui.PAGE})
            server=ThreadingHTTPServer(('127.0.0.1',0),handler)
            thread=threading.Thread(target=server.serve_forever,daemon=True)
            thread.start()
            try:
                origin='http://127.0.0.1:'+str(server.server_port)
                for payload in ({'month':'2026-09','kind':'../../etc/passwd'},
                                {'month':'2026-09/../../etc','kind':'excel'}):
                    req=Request(origin+'/api/open',
                        data=json.dumps(payload).encode('utf-8'),
                        headers={'Content-Type':'application/json','X-UI-Token':token},
                        method='POST')
                    with self.assertRaises(HTTPError) as caught:
                        urlopen(req,timeout=4)
                    self.assertEqual(caught.exception.code,400)
            finally:
                server.shutdown()
                server.server_close()
                thread.join(timeout=4)


if __name__=='__main__':
    unittest.main()
