"""Python 3.8/3.9 smoke: run the GUI and diagnostics with NO scraper packages."""
import json
from pathlib import Path
import sys
import tempfile
import threading
import time
from http.server import ThreadingHTTPServer
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import webui


def main():
    assert (3, 8) <= sys.version_info[:2] < (3, 10), sys.version
    with tempfile.TemporaryDirectory(prefix='kylin-old-python-') as temp:
        folder=Path(temp)
        manager=webui.Jobs(folder)
        assert not manager.environment['ready'], manager.environment
        assert 'Python 3.10+' in manager.environment['missing']
        try:
            manager.start('2026-09','sync')
        except RuntimeError:
            pass
        else:
            raise AssertionError('Old Python must not run scraper core')
        token='local-ci-token'
        handler=type('KylinOldPythonHandler',(webui.LocalHandler,),{
            'jobs':manager,'token':token,
            'page':webui.PAGE.replace('__TOKEN__',token)})
        server=ThreadingHTTPServer(('127.0.0.1',0),handler)
        thread=threading.Thread(target=server.serve_forever,daemon=True)
        thread.start()
        try:
            origin='http://127.0.0.1:'+str(server.server_port)
            with urlopen(origin, timeout=5) as response:
                page=response.read().decode('utf-8')
            assert '检测电脑环境' in page, 'Diagnostic GUI button absent'
            request=Request(
                origin+'/api/run',
                data=json.dumps({'month':'2026-09','mode':'diagnose'}).encode('utf-8'),
                headers={'X-UI-Token':token,'Content-Type':'application/json'},
                method='POST'
            )
            with urlopen(request,timeout=5) as response:
                assert response.status==202
            deadline=time.monotonic()+15
            while time.monotonic()<deadline:
                state=manager.status('2026-09')
                if not state['running']:
                    assert state['code']==0, state['log']
                    assert '环境诊断' in state['log']
                    assert not list(folder.iterdir()), 'Diagnostics wrote unexpected files'
                    print('PASS: Python '+sys.version.split()[0]+' diagnostic-only GUI, no scraper dependencies')
                    return
                time.sleep(0.05)
            raise TimeoutError('GUI diagnostic did not finish')
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=5)


if __name__=='__main__':
    main()
