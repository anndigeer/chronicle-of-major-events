"""Windows CMD launcher E2E smoke: PowerShell-compatible entry, localhost HTTP.
No real news requests and no changes to user Excel/Word.
"""
import json
from pathlib import Path
import re
import socket
import subprocess
import sys
import time
import urllib.request

ROOT = Path(__file__).resolve().parents[1]
COMMAND = ROOT / '市委大事记.cmd'


def call(*args):
    return subprocess.run(
        ['cmd', '/d', '/c', str(COMMAND), *args],
        cwd=ROOT,
        capture_output=True, text=True, encoding='utf-8', errors='replace',
        timeout=25, check=False
    )


def check():
    assert sys.platform == 'win32', 'Windows smoke must run on a Windows runner'
    assert COMMAND.is_file()
    help_result = call('help')
    assert help_result.returncode == 0, (help_result.stdout, help_result.stderr)
    assert 'python webui.py' in help_result.stdout
    assert '.cmd' in help_result.stdout
    invalid = call('unknown-command')
    assert invalid.returncode != 0
    diagnostic = call('check')
    assert diagnostic.returncode == 0, (diagnostic.stdout, diagnostic.stderr)
    assert 'Python' in diagnostic.stdout
    assert '依赖检测' in diagnostic.stdout

    # Allocate a temporary localhost port and start the real webui process
    # through the same .cmd users invoke from PowerShell.
    with socket.socket() as sock:
        sock.bind(('127.0.0.1', 0))
        port = sock.getsockname()[1]
    proc = subprocess.Popen(
        ['cmd', '/d', '/c', str(COMMAND), 'start', '--no-browser', '--port', str(port)],
        cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )
    try:
        deadline = time.monotonic() + 25
        url = 'http://127.0.0.1:{}/'.format(port)
        html = None
        while time.monotonic() < deadline:
            if proc.poll() is not None:
                raise AssertionError('CMD GUI launcher exited early: {}'.format(proc.returncode))
            try:
                with urllib.request.urlopen(url, timeout=0.75) as response:
                    html = response.read().decode('utf-8')
                break
            except (OSError, ValueError):
                time.sleep(0.15)
        assert html, 'Windows GUI did not start at {}'.format(url)
        assert '市委大事记' in html
        assert '增量采集 + 更新 Word' in html
        print('PASS: Windows CMD help/check/start and real local browser HTTP')
    finally:
        # cmd /c starts a child Python process: terminate its entire tree.
        subprocess.run(
            ['taskkill', '/PID', str(proc.pid), '/T', '/F'],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            timeout=15, check=False,
        )
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()


if __name__ == '__main__':
    check()
