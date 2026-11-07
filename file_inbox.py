# -*- coding: utf-8 -*-
"""扫码收文件：在会务电脑上开一个局域网收件服务。

现场的人连同一个 WiFi，用手机扫码打开网页，就能把微信里的文件直接传上来；
文件按时间落到指定文件夹，重名自动加序号，全部走局域网、不经过公网。
"""

import os
import re
import socket
import sys
import tempfile
import threading
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

MAX_FILE_BYTES = 1024 * 1024 * 1024        # 单个文件上限 1GB
MAX_NAME_LEN = 120
CHUNK = 256 * 1024
_UNSAFE = re.compile(r'[\\/:*?"<>|\r\n\t]')


def local_ip():
    """取本机在局域网里的地址（只做一次 UDP connect，不会真的发包）。"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        sock.connect(('10.255.255.255', 1))
        return sock.getsockname()[0]
    except Exception:
        return '127.0.0.1'
    finally:
        try:
            sock.close()
        except Exception:
            pass


def all_local_ips():
    """本机所有局域网地址（多网卡时备用）。"""
    ips = []
    first = local_ip()
    if first and not first.startswith('127.'):
        ips.append(first)
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None,
                                       socket.AF_INET):
            ip = info[4][0]
            if ip not in ips and not ip.startswith('127.'):
                ips.append(ip)
    except Exception:
        pass
    return ips


def safe_filename(name, fallback='文件'):
    """清洗上传上来的文件名：去掉路径和非法字符，限制长度。"""
    name = urllib.parse.unquote(name or '')
    name = name.replace('\\', '/').split('/')[-1]
    name = _UNSAFE.sub('_', name).strip().strip('.')
    if not name:
        name = fallback
    if len(name) > MAX_NAME_LEN:
        stem, ext = os.path.splitext(name)
        name = stem[:MAX_NAME_LEN - len(ext)] + ext
    return name


def unique_path(folder, name):
    """重名时自动加 (1)(2)…，不覆盖已有文件。"""
    path = os.path.join(folder, name)
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(name)
    index = 1
    while True:
        path = os.path.join(folder, '%s(%d)%s' % (stem, index, ext))
        if not os.path.exists(path):
            return path
        index += 1


def ensure_folder(folder):
    """确保保存目录可用：不可写时退到程序目录、临时目录。"""
    base = os.path.dirname(os.path.abspath(
        sys.executable if getattr(sys, 'frozen', False) else __file__))
    candidates = [folder,
                  os.path.join(base, '会务收件'),
                  os.path.join(tempfile.gettempdir(), '会务收件')]
    for path in candidates:
        if not path:
            continue
        try:
            os.makedirs(path, exist_ok=True)
            probe = os.path.join(path, '.write_test')
            with open(probe, 'w'):
                pass
            os.remove(probe)
            return path
        except Exception:
            continue
    return folder


class Inbox(object):
    """收件服务：负责起 HTTP 服务、保存文件、记录收到的文件列表。"""

    def __init__(self, folder=None, port=8000, max_files=500,
                 on_received=None):
        self.folder = folder or os.path.join(
            os.path.expanduser('~'), 'Documents', '会务收件')
        self.port = int(port)
        self.max_files = max_files
        # 每收到一个文件就在 HTTP 线程上回调（item dict）；
        # 回调内部自己负责切回 UI 线程，异常不能影响上传响应。
        self.on_received = on_received
        self.files = []            # [{name, size, time, path, who}]
        self.lock = threading.Lock()
        self._httpd = None
        self._thread = None

    def notify_received(self, item):
        cb = self.on_received
        if callable(cb):
            try:
                cb(item)
            except Exception:
                pass

    # ---------- 启停 ----------
    @property
    def running(self):
        return self._httpd is not None

    def start(self):
        """启动服务；端口被占用时自动往后找。返回实际使用的端口。"""
        if self.running:
            return self.port
        self.folder = ensure_folder(self.folder)
        os.makedirs(self.folder, exist_ok=True)
        last_error = None
        for offset in range(0, 20):
            port = self.port + offset
            if port_in_use(port):
                last_error = OSError('端口 %d 已被其它程序占用' % port)
                continue
            try:
                httpd = _make_server(port, self)
            except OSError as exc:
                last_error = exc
                continue
            self.port = port
            self._httpd = httpd
            self._thread = threading.Thread(target=httpd.serve_forever,
                                            kwargs={'poll_interval': 0.3},
                                            daemon=True)
            self._thread.start()
            return port
        raise OSError('端口都被占用了，换一个端口再试（%s）' % last_error)

    def stop(self):
        httpd, self._httpd = self._httpd, None
        if httpd is not None:
            try:
                httpd.shutdown()
                httpd.server_close()
            except Exception:
                pass
        self._thread = None

    # ---------- 状态 ----------
    def url(self, ip=None):
        return 'http://%s:%d' % (ip or local_ip(), self.port)

    def add_file(self, path, size, who='', sender_ip=''):
        item = {'name': os.path.basename(path), 'size': size,
                'time': time.strftime('%H:%M:%S'), 'path': path,
                'who': who or '', 'ip': sender_ip or ''}
        with self.lock:
            self.files.insert(0, item)
            del self.files[self.max_files:]
        return item

    def snapshot(self):
        with self.lock:
            return list(self.files)

    def clear_list(self):
        with self.lock:
            self.files = []

    def total_bytes(self):
        with self.lock:
            return sum(item['size'] for item in self.files)


class _Handler(BaseHTTPRequestHandler):
    server_version = 'MeetingInbox/1.0'
    timeout = 300

    # ---------- 工具 ----------
    @property
    def inbox(self):
        return self.server.inbox

    def log_message(self, fmt, *args):
        pass                      # 不往控制台刷日志

    def _send(self, code, body, content_type='text/plain; charset=utf-8'):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _send_json(self, code, obj):
        import json
        self._send(code, json.dumps(obj, ensure_ascii=False),
                   'application/json; charset=utf-8')

    # ---------- 路由 ----------
    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path in ('/', '/index.html'):
            self._send(200, _page(self.inbox), 'text/html; charset=utf-8')
            return
        if path == '/api/status':
            files = self.inbox.snapshot()
            self._send_json(200, {
                'ok': True,
                'count': len(files),
                'total': self.inbox.total_bytes(),
                'folder': self.inbox.folder,
                'files': [{'name': f['name'], 'size': f['size'],
                           'time': f['time'], 'who': f.get('who', '')}
                          for f in files[:50]],
            })
            return
        if path == '/favicon.ico':
            self._send(204, b'')
            return
        self._send(404, '页面不存在')

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if path != '/upload':
            self._send_json(404, {'ok': False, 'error': '接口不存在'})
            return
        try:
            result = self._receive_file()
        except _UploadError as exc:
            self._send_json(400, {'ok': False, 'error': str(exc)})
            return
        except Exception as exc:                     # 兜底，别让服务挂掉
            self._send_json(500, {'ok': False, 'error': '服务器错误：%s' % exc})
            return
        self._send_json(200, {'ok': True, 'file': result['name'],
                              'size': result['size']})

    # ---------- 收文件 ----------
    def _receive_file(self):
        length = int(self.headers.get('Content-Length') or 0)
        if length <= 0:
            raise _UploadError('没有收到文件内容')
        if length > MAX_FILE_BYTES:
            raise _UploadError('文件太大（上限 %d MB）'
                               % (MAX_FILE_BYTES // 1048576))
        raw_name = self.headers.get('X-File-Name') or ''
        who = self.headers.get('X-Sender') or ''
        name = safe_filename(raw_name, '文件_%s' % time.strftime('%H%M%S'))
        who = safe_filename(who, '') if who else ''
        if who:
            name = '%s_%s' % (who[:20], name)

        folder = self.inbox.folder
        os.makedirs(folder, exist_ok=True)
        path = unique_path(folder, name)
        sender_ip = self.client_address[0] if self.client_address else ''
        remaining = length
        written = 0
        # 现场电脑的杀软/索引器偶尔会瞬时占用新建文件（EACCES），短重试 3 次
        handle = None
        for attempt in range(3):
            try:
                handle = open(path, 'wb')
                break
            except PermissionError:
                if attempt == 2:
                    raise
                time.sleep(0.15)
        with handle:
            while remaining > 0:
                chunk = self.rfile.read(min(CHUNK, remaining))
                if not chunk:
                    break
                handle.write(chunk)
                written += len(chunk)
                remaining -= len(chunk)
        if written == 0:
            try:
                os.remove(path)
            except OSError:
                pass
            raise _UploadError('上传中断，请重试')
        item = self.inbox.add_file(path, written, who=who,
                                   sender_ip=sender_ip)
        self.inbox.notify_received(item)
        return {'name': os.path.basename(path), 'size': written}


class _UploadError(Exception):
    pass


class _Server(ThreadingHTTPServer):
    daemon_threads = True
    # Windows 上 SO_REUSEADDR 表示“可以抢占已被占用的端口”，
    # 会让第二个实例以为端口是空闲的，把别人的连接抢过来，所以关掉
    allow_reuse_address = os.name != 'nt'

    def __init__(self, addr, inbox):
        self.inbox = inbox
        ThreadingHTTPServer.__init__(self, addr, _Handler)


def _make_server(port, inbox):
    return _Server(('0.0.0.0', port), inbox)


def port_in_use(port, host='127.0.0.1', timeout=0.3):
    """端口上已经有程序在监听时返回 True（Windows 抢端口问题靠这个兜底）。"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(timeout)
    try:
        return sock.connect_ex((host, port)) == 0
    except Exception:
        return False
    finally:
        try:
            sock.close()
        except Exception:
            pass


def _page(inbox):
    """手机上传页（内联样式和脚本，离线可用）。"""
    return PAGE.replace('__FOLDER__', os.path.basename(inbox.folder.rstrip('/\\')))


PAGE = u"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1,
      maximum-scale=1, viewport-fit=cover">
<title>上传文件到会务电脑</title>
<style>
  * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
  body {
    margin: 0; padding: 18px 16px 32px;
    font-family: -apple-system, BlinkMacSystemFont, "PingFang SC",
                 "Microsoft YaHei", sans-serif;
    background: #f4f6fb; color: #1f2937;
  }
  .card {
    background: #fff; border-radius: 16px; padding: 18px;
    box-shadow: 0 6px 20px rgba(37, 99, 235, .08); margin-bottom: 14px;
  }
  h1 { font-size: 19px; margin: 0 0 6px; }
  .sub { font-size: 13px; color: #6b7280; margin: 0; }
  .folder { color: #2563eb; font-weight: 600; }
  label { display: block; font-size: 13px; color: #6b7280;
          margin: 14px 0 6px; }
  input[type=text] {
    width: 100%; padding: 12px 14px; font-size: 16px;
    border: 1px solid #e5eaf3; border-radius: 12px; background: #fff;
    outline: none;
  }
  input[type=text]:focus { border-color: #2563eb; }
  .req { color: #dc2626; font-weight: 600; }
  #reqhint {
    display: none; margin-top: 8px; padding: 10px 12px;
    background: #fef2f2; border: 1px solid #fecaca; border-radius: 10px;
    font-size: 13px; color: #b91c1c;
  }
  /* 文件输入框直接铺在按钮上：点哪儿都是原生触发，
     避免用 JS 去点 display:none 的输入框（微信里会被拦截） */
  .pickwrap { position: relative; margin-top: 14px; }
  .pickwrap input[type=file] {
    position: absolute; left: 0; top: 0; width: 100%; height: 100%;
    opacity: 0; filter: alpha(opacity=0); cursor: pointer;
    font-size: 100px; z-index: 2; margin: 0; padding: 0; border: 0;
  }
  .btnface {
    display: block; padding: 16px; font-size: 17px; font-weight: 600;
    text-align: center; color: #fff; background: #2563eb;
    border-radius: 14px;
  }
  #fallback {
    display: none; margin-top: 12px; padding: 12px 14px;
    background: #fff7ed; border: 1px solid #fed7aa; border-radius: 12px;
    font-size: 13px; color: #b45309; line-height: 1.7;
  }
  .hint { font-size: 12px; color: #9aa3b2; margin-top: 10px;
          line-height: 1.6; }
  ul { list-style: none; margin: 0; padding: 0; }
  li { padding: 12px 0; border-bottom: 1px solid #f0f3f9; font-size: 14px; }
  li:last-child { border-bottom: 0; }
  .row { display: flex; align-items: center; gap: 10px; }
  .name { flex: 1; word-break: break-all; }
  .size { color: #9aa3b2; font-size: 12px; white-space: nowrap; }
  .state { font-size: 12px; white-space: nowrap; }
  .st-wait { color: #9aa3b2; }
  .st-run { color: #2563eb; }
  .st-ok { color: #16a34a; }
  .st-err { color: #dc2626; }
  .bar { display: none; height: 4px; background: #eef2fa; border-radius: 2px;
         overflow: hidden; margin-top: 6px; }
  .bar.on { display: block; }
  .bar > i { display: block; height: 100%; width: 0;
             background: #2563eb; transition: width .15s; }
  #done { display: none; text-align: center; padding: 8px 0 4px; }
  #done .big { font-size: 15px; font-weight: 600; color: #16a34a; }
  .retry { display: none; margin: 10px auto 0; padding: 10px 18px;
           font-size: 14px; color: #2563eb; background: #e8f0fe;
           border: 0; border-radius: 10px; }
  .err { display: none; margin-top: 12px; padding: 12px 14px;
         background: #fef2f2; border: 1px solid #fecaca; border-radius: 12px;
         font-size: 13px; color: #b91c1c; word-break: break-all; }
</style>
</head>
<body>

<div class="card">
  <h1>上传文件到会务电脑</h1>
  <p class="sub">保存到：<span class="folder">__FOLDER__</span></p>
  <label>你的名字 / 单位 <span class="req">*</span>（必填，方便会务组认领）</label>
  <input type="text" id="who" placeholder="例如：张三 / 市场部" required>
  <div id="reqhint">请先填写你的名字 / 单位，再选择文件</div>
  <div class="pickwrap">
    <input type="file" id="file" multiple>
    <span class="btnface" id="pick">选择文件上传</span>
  </div>
  <div id="fallback">
    没有弹出选择框？点右上角「···」→「在浏览器打开」，再试一次。<br>
    微信里的文件也可以先“保存到手机”或在电脑微信里直接转发。
  </div>
  <p class="hint">可以一次选多个文件；手机和电脑要在同一个 WiFi 下。</p>
</div>

<div class="card" id="listcard" style="display:none">
  <ul id="list"></ul>
</div>

<div class="card" id="done">
  <div class="big" id="donebig">全部上传完成</div>
  <p class="hint" id="donesub">文件已保存到会务电脑</p>
  <button id="retry" class="retry">重试失败的文件</button>
</div>

<div id="errbox" class="err"></div>

<script>
var input = document.getElementById('file');
var list = document.getElementById('list');
var fallback = document.getElementById('fallback');
var who = document.getElementById('who');
var reqhint = document.getElementById('reqhint');
var queue = [];
var running = false;
var hintTimer = null;

function fmt(n) {
  if (n < 1024) return n + ' B';
  if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
  if (n < 1073741824) return (n / 1048576).toFixed(1) + ' MB';
  return (n / 1073741824).toFixed(2) + ' GB';
}

/* 兜底：万一有浏览器不吃“铺满式”输入框，也可用 JS 触发（此时输入框是可见的） */
document.getElementById('pick').onclick = function (event) {
  event.preventDefault();
  try { input.click(); } catch (err) {}
};

/* 点了按钮却没弹出选择框时，给出“用浏览器打开”的提示 */
input.addEventListener('click', function (event) {
  /* 名字/单位必填：没填就不让打开选择框，并给出红字提示 */
  if (!senderFilled()) {
    event.preventDefault();
    reqhint.style.display = 'block';
    try { who.focus(); } catch (err) {}
    return;
  }
  reqhint.style.display = 'none';
  if (hintTimer) clearTimeout(hintTimer);
  hintTimer = setTimeout(function () {
    if (!input.files || input.files.length === 0) {
      fallback.style.display = 'block';
    }
  }, 2500);
});

function senderFilled() {
  return (who.value || '').replace(/\s/g, '').length > 0;
}

who.addEventListener('input', function () {
  if (senderFilled()) reqhint.style.display = 'none';
});

/* 失败重试：把失败项状态清掉，重新排队 */
document.getElementById('retry').onclick = function () {
  for (var i = 0; i < queue.length; i++) {
    if (queue[i].state === 'err') {
      queue[i].state = '';
      queue[i].pct = 0;
      queue[i].error = '';
    }
  }
  document.getElementById('done').style.display = 'none';
  render();
  if (!running) uploadNext();
};

input.onchange = function () {
  if (hintTimer) clearTimeout(hintTimer);
  fallback.style.display = 'none';
  if (!senderFilled()) {
    reqhint.style.display = 'block';
    try { input.value = ''; } catch (err) {}
    return;
  }
  var picked = input.files || [];
  for (var i = 0; i < picked.length; i++) {
    /* 存成对象：file 才是文件本体，state/pct/error 记录进度 */
    queue.push({ file: picked[i], state: '', pct: 0, error: '' });
  }
  document.getElementById('listcard').style.display = 'block';
  document.getElementById('done').style.display = 'none';
  render();
  if (!running) uploadNext();
};

/* 页面里任何脚本错误都直接显示出来，不再“点了没反应” */
window.onerror = function (msg, src, line) {
  var box = document.getElementById('errbox');
  if (box) {
    box.style.display = 'block';
    box.textContent = '页面出错：' + msg + '（第 ' + line + ' 行）';
  }
};

function render() {
  list.innerHTML = '';
  queue.forEach(function (item) {
    var li = document.createElement('li');
    var cls = item.state === 'ok' ? 'st-ok'
            : item.state === 'err' ? 'st-err'
            : item.state === 'run' ? 'st-run' : 'st-wait';
    var txt = item.state === 'ok' ? '已上传'
            : item.state === 'err' ? (item.error || '失败')
            : item.state === 'run'
              ? '上传中 ' + Math.round(item.pct || 0) + '%' : '等待';
    li.innerHTML =
      '<div class="row">' +
        '<span class="name">' + escapeHtml(item.file.name) + '</span>' +
        '<span class="size">' + fmt(item.file.size) + '</span>' +
        '<span class="state ' + cls + '">' + txt + '</span>' +
      '</div>' +
      '<div class="bar' + (item.state === 'run' ? ' on' : '') + '">' +
        '<i style="width:' + (item.pct || 0) + '%"></i></div>';
    list.appendChild(li);
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;',
             '"': '&quot;', "'": '&#39;' }[c];
  });
}

function uploadNext() {
  var item = null;
  for (var i = 0; i < queue.length; i++) {
    if (!queue[i].state) { item = queue[i]; break; }
  }
  if (!item) {
    running = false;
    try { input.value = ''; } catch (err) {}
    var bad = queue.filter(function (q) { return q.state === 'err'; });
    document.getElementById('retry').style.display =
      bad.length ? 'inline-block' : 'none';
    document.getElementById('done').style.display = 'block';
    document.getElementById('donebig').textContent =
      bad.length ? '有 ' + bad.length + ' 个文件上传失败' : '全部上传完成';
    document.getElementById('donesub').textContent =
      '共 ' + queue.length + ' 个文件，已保存到会务电脑';
    return;
  }
  running = true;
  item.state = 'run';
  item.pct = 0;
  render();

  var xhr = new XMLHttpRequest();
  xhr.open('POST', '/upload', true);
  xhr.setRequestHeader('X-File-Name', encodeURIComponent(item.file.name));
  xhr.setRequestHeader('X-Sender', encodeURIComponent(
    (who.value || '').trim()));
  xhr.setRequestHeader('Content-Type', 'application/octet-stream');
  xhr.upload.onprogress = function (e) {
    if (e.lengthComputable) {
      item.pct = e.loaded * 100 / e.total;
      render();
    }
  };
  xhr.onload = function () {
    var ok = false;
    try { ok = JSON.parse(xhr.responseText).ok; } catch (err) { ok = false; }
    if (xhr.status === 200 && ok) {
      item.state = 'ok';
      item.pct = 100;
    } else {
      item.state = 'err';
      item.error = '失败';
      try { item.error = JSON.parse(xhr.responseText).error || '失败'; }
      catch (err) {}
    }
    render();
    uploadNext();
  };
  xhr.onerror = function () {
    item.state = 'err';
    item.error = '网络中断';
    render();
    uploadNext();
  };
  try {
    xhr.send(item.file);
  } catch (err) {
    item.state = 'err';
    item.error = '无法读取';
    render();
    uploadNext();
  }
}
</script>
</body>
</html>
"""
