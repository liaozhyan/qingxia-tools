# -*- coding: utf-8 -*-
"""局域网互传：两台同局域网电脑直接互传文件，替代 U 盘拷贝。

- 接收端：复用 file_inbox 的 HTTP 上传服务，对端一发即自动存入指定文件夹；
- 发送端：UDP 广播自动发现局域网在线电脑（+ 手动填 IP 兜底），
          多文件后台逐文件上传并显示进度，失败可重试。
一切走局域网内网，不经过公网。
"""

import json
import os
import socket
import struct
import sys
import tempfile
import threading
import time
import urllib.parse
import urllib.request
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import ui_theme
from file_inbox import (Inbox, local_ip, all_local_ips, port_in_use,
                        ensure_folder)

DISCOVER_PORT = 29109        # UDP 自动发现端口
DISCOVER_ADDR = ('255.255.255.255', DISCOVER_PORT)
HEATBEAT = 2.0               # 广播间隔（秒）
PEER_TTL = 6.0               # 超过该时长未广播视为下线
MAGIC = b'LANSHARE'

CHUNK = 256 * 1024


def machine_name():
    try:
        name = socket.gethostname()
    except Exception:
        name = '电脑'
    return (name or '电脑').strip() or '电脑'


def _pad_str(s, length=40):
    return (s.encode('utf-8', 'ignore') + b'\x00' * length)[:length]


def _unpad_str(b):
    return b.split(b'\x00')[0].decode('utf-8', 'ignore').strip()


class PeerTable(object):
    """在线对端表：{ip: {'name','port','seen'}}，供自动发现写入、UI 读取。"""

    def __init__(self):
        self._peers = {}
        self.lock = threading.Lock()

    def announce_in(self, ip, name, port):
        with self.lock:
            self._peers[ip] = {'name': name, 'port': port,
                               'seen': time.time()}

    def purge(self):
        now = time.time()
        with self.lock:
            for ip in [k for k, v in self._peers.items()
                       if now - v['seen'] > PEER_TTL]:
                del self._peers[ip]

    def snapshot(self):
        with self.lock:
            return list(self._peers.items())


class _Discover(threading.Thread):
    """UDP 广播 + 监听，维护在线对端表。daemon，随窗口关闭。

    recv_port 可传 int，也可传 callable——接收服务是开窗后才点“开始接收”
    启动的，端口还可能因占用顺延，用 callable 每次广播取最新值，
    就不用每次起停发现线程了（开窗即能看到在线机器）。
    """

    def __init__(self, table, recv_port):
        super().__init__(daemon=True)
        self.table = table
        if callable(recv_port):
            self._port_fn = recv_port
        else:
            self._port_fn = lambda: int(recv_port)
        self._stop = threading.Event()

    def stop(self):
        self._stop.set()

    def run(self):
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            sock.bind(('', DISCOVER_PORT))
            sock.settimeout(0.5)
        except OSError:
            # 发现端口被占：仍发广播，放弃监听也能手动传
            if not self._stop.wait(HEATBEAT):
                try:
                    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                    sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
                    sock.settimeout(0.5)
                except OSError:
                    return

        # 自己发的广播自己也可能收到（多网卡/回环），按本机 IP 过滤掉
        local_ips = set(all_local_ips())
        name = machine_name()
        while not self._stop.is_set():
            try:
                payload = MAGIC + b'\x00' + _pad_str(name) + \
                    struct.pack('!H', int(self._port_fn() or 0))
                sock.sendto(payload, DISCOVER_ADDR)
            except OSError:
                pass
            # 收集这段时间收到的广播
            deadline = time.time() + HEATBEAT
            while time.time() < deadline:
                try:
                    data, addr = sock.recvfrom(1024)
                except socket.timeout:
                    continue
                except OSError:
                    break
                if data[:len(MAGIC)] != MAGIC:
                    continue
                ip = addr[0]
                if ip.startswith('127.') or ip in local_ips:
                    continue
                body = data[len(MAGIC) + 1:]
                pname = _unpad_str(body[:40])
                port = struct.unpack('!H', body[40:42])[0] if len(body) >= 42 else 0
                self.table.announce_in(ip, pname, port)
            self.table.purge()
        try:
            sock.close()
        except OSError:
            pass


def _field(name, port, pct=0.0, state='wait', error=''):
    return {'name': name, 'port': port, 'pct': pct, 'state': state,
            'error': error}


class _CountingReader(object):
    """按需读文件并累计已读字节，供发送进度回调。"""

    def __init__(self, handle, total, on_progress):
        self._h = handle
        self.total = total
        self.sent = 0
        self._on_progress = on_progress

    def read(self, n=-1):
        data = self._h.read(n if n else CHUNK)
        if data:
            self.sent += len(data)
            try:
                self._on_progress(self.sent)
            except Exception:
                pass
        return data

    def __getattr__(self, attr):
        return getattr(self._h, attr)


def send_file_to(target_ip, port, path, on_progress):
    """把单个文件 POST 到接收端 /upload。阻塞返回。抛异常即失败。"""
    name = os.path.basename(path)
    size = os.path.getsize(path)
    headers = {
        'Content-Type': 'application/octet-stream',
        'Content-Length': str(size),
        'X-File-Name': urllib.parse.quote(name),
        'X-Sender': urllib.parse.quote(machine_name()),
    }
    url = 'http://%s:%s/upload' % (target_ip, port)
    with open(path, 'rb') as handle:
        reader = _CountingReader(handle, size, on_progress)
        req = urllib.request.Request(url, data=reader, headers=headers,
                                     method='POST')
        with urllib.request.urlopen(req, timeout=600) as resp:
            body = resp.read()
    try:
        obj = json.loads(body.decode('utf-8', 'ignore'))
    except Exception:
        obj = {}
    if not obj.get('ok'):
        raise RuntimeError(obj.get('error') or '接收端返回失败')


class _Sender(threading.Thread):
    """逐文件发送队列。daemon。每个字段 state: wait/run/ok/err。"""

    def __init__(self, peer, fields, on_item):
        """peer=(ip,port)；fields=[_field(...)]；on_item(全字段列表) 刷 UI。"""
        super().__init__(daemon=True)
        self.peer = peer
        self.fields = fields
        self.on_item = on_item
        self._stop = threading.Event()

    def run(self):
        for f in self.fields:
            if f['state'] == 'ok' or f['state'] == 'err':
                # 手动重试时只重新发 status 非 ok 的
                continue
            f['state'] = 'run'
            f['pct'] = 0.0
            self.on_item(self.fields)
            try:
                ip, port = self.peer
                send_file_to(ip, int(port), f['name'],
                             lambda got, f=f: self._tick(f, got))
                f['state'] = 'ok'
                f['pct'] = 100.0
            except Exception as exc:
                f['state'] = 'err'
                f['error'] = str(exc)
            self.on_item(self.fields)

    def _tick(self, f, got):
        f['pct'] = min(100.0, got / max(1, f['port']) * 100.0)
        self.on_item(self.fields)


class _ToastBubble(object):
    """右下角收到文件气泡：置顶无边框、淡入、6 秒后淡出，点开文件夹。

    纯 Tk 实现（零第三方依赖）；位置由 manager 统一排版，多个向上叠。
    """

    W, H = 380, 86

    def __init__(self, manager, title, message, sub, on_open):
        self._mgr = manager
        self._on_open = on_open
        self._dying = False
        p = ui_theme.PALETTE
        top = tk.Toplevel(manager.master)
        self.top = top
        top.overrideredirect(True)
        try:
            top.attributes('-topmost', True)
        except tk.TclError:
            pass
        try:
            top.attributes('-alpha', 0.0)
        except tk.TclError:
            pass
        card = tk.Frame(top, bg=p['surface'], highlightthickness=1,
                        highlightbackground=p['border'], bd=0)
        card.pack(fill=tk.BOTH, expand=True)
        tk.Frame(card, bg=ui_theme.ACCENTS.get('lan', '#0891b2'),
                 width=5).pack(side=tk.LEFT, fill=tk.Y)
        body = tk.Frame(card, bg=p['surface'])
        body.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=12, pady=9)
        head = tk.Frame(body, bg=p['surface'])
        head.pack(fill=tk.X)
        tk.Label(head, text=title, font=ui_theme.font(10, 'bold'),
                 bg=p['surface'], fg=ui_theme.ACCENTS.get('lan', '#0891b2'),
                 anchor='w').pack(side=tk.LEFT)
        tk.Label(head, text='✕', font=ui_theme.font(10),
                 bg=p['surface'], fg=p['faint'], cursor='arrow'
                 ).pack(side=tk.RIGHT)
        head.winfo_children()[-1].bind('<Button-1>', lambda e: self.close())
        tk.Label(body, text=message, font=ui_theme.font(10, 'bold'),
                 bg=p['surface'], fg=p['text'], anchor='w',
                 wraplength=300, justify=tk.LEFT).pack(fill=tk.X,
                                                       pady=(3, 0))
        tk.Label(body, text=sub, font=ui_theme.font(8),
                 bg=p['surface'], fg=p['muted'], anchor='w',
                 wraplength=300, justify=tk.LEFT).pack(fill=tk.X)
        # 点气泡任意空白处都打开文件夹
        for wgt in (card, body, head) + tuple(head.winfo_children()):
            try:
                wgt.configure(cursor='hand2')
                wgt.bind('<Button-1>', self._click_open)
            except tk.TclError:
                pass
        # 关闭按钮不触发打开
        head.winfo_children()[-1].configure(cursor='arrow')
        head.winfo_children()[-1].bind('<Button-1>', lambda e: self.close())
        self._relocate(0)
        top.after(20, self._fade_in)
        self._job = top.after(6000, self.close)

    def _relocate(self, index):
        try:
            sw, sh = self.top.winfo_screenwidth(), self.top.winfo_screenheight()
            x = sw - self.W - 14
            y = sh - 48 - (index + 1) * (self.H + 10)
            self.top.geometry('%dx%d+%d+%d' % (self.W, self.H, x, y))
        except tk.TclError:
            pass

    def reposition(self, index):
        self._relocate(index)

    def _fade_in(self, alpha=0.0):
        try:
            alpha = min(1.0, alpha + 0.18)
            self.top.attributes('-alpha', alpha)
            if alpha < 1.0:
                self.top.after(18, lambda: self._fade_in(alpha))
        except tk.TclError:
            pass

    def _click_open(self, _event=None):
        try:
            if callable(self._on_open):
                self._on_open()
        except Exception:
            pass
        self.close()

    def close(self):
        if self._dying:
            return
        self._dying = True
        try:
            self.top.after_cancel(self._job)
        except Exception:
            pass
        self._mgr.drop_toast(self)
        self._fade_out()

    def _fade_out(self, alpha=1.0):
        try:
            alpha -= 0.2
            if alpha <= 0:
                self.top.destroy()
                return
            self.top.attributes('-alpha', alpha)
            self.top.after(18, lambda: self._fade_out(alpha))
        except tk.TclError:
            try:
                self.top.destroy()
            except tk.TclError:
                pass


class LanShareApp(object):
    """局域网互传窗口。master 为主程序根窗口。"""

    def __init__(self, master):
        self.master = master
        self.win = None
        self.peers = PeerTable()
        self._discover = None
        self._inbox = None
        self._sender = None
        self._send_file = None          # ttk.Progressbar 引用
        self._after = None
        self._toasts = []               # 右下角收到文件气泡
        self._recv_items = {}           # 树行 iid → item
        self._recv_seq = 0
        self._recv_seen = set()         # 去重：(path,size,time)
        self._target_ip = None
        self._target_port = None
        self._target_name = ''
        self._build()
        # 窗口一打开就开始发现在线电脑（不必等点“开始接收”）
        self._start_discovery()

    # ---------------- 窗口 ----------------
    def _build(self):
        win = tk.Toplevel(self.master)
        win.title('局域网互传')
        win.configure(bg=ui_theme.PALETTE['bg'])
        self.win = win
        try:
            ui_theme.style_window(win)
        except Exception:
            pass
        win.resizable(True, True)
        win.geometry('860x620')
        ui_theme._center(win, 860, 620) if hasattr(ui_theme, '_center') \
            else _center(win, 860, 620)
        win.protocol('WM_DELETE_WINDOW', self.close)

        acc = ui_theme.ACCENTS.get('lan', '#0891b2')
        top = tk.Label(win, text='局域网互传', font=ui_theme.font(15, 'bold'),
                       bg=ui_theme.PALETTE['bg'],
                       fg=ui_theme.PALETTE['text'],
                       anchor='w')
        top.pack(fill=tk.X, padx=18, pady=(14, 6))
        sub = tk.Label(win, text='同一局域网内，自动发现对端，互传文件（不经过公网）。'
                                '机器名：%s  本机IP：%s'
                       % (machine_name(), local_ip()),
                       font=ui_theme.font(9),
                       bg=ui_theme.PALETTE['bg'],
                       fg=ui_theme.PALETTE['muted'], anchor='w')
        sub.pack(fill=tk.X, padx=18)

        nb = ttk.Notebook(win)
        nb.pack(fill=tk.BOTH, expand=True, padx=18, pady=(10, 12))
        self._recv_tab = self._build_recv(nb)
        self._send_tab = self._build_send(nb)

        self._after = win.after(400, self._tick_ui)

    # ---------- 接收 ----------
    def _build_recv(self, nb):
        page = tk.Frame(nb, bg=ui_theme.PALETTE['bg'])
        nb.add(page, text='  接收  ')
        p = ui_theme.PALETTE

        row = tk.Frame(page, bg=p['bg'])
        row.pack(fill=tk.X, pady=(10, 0))

        tk.Button(row, text='开始接收', relief='flat', bd=0, padx=14, pady=5,
                  bg=p['primary'], fg='#ffffff',
                  activebackground=p['primary_dark'],
                  command=self._start_recv).pack(side=tk.LEFT)
        tk.Button(row, text='停止', relief='flat', bd=0, padx=10, pady=5,
                  bg=p['surface_alt'], fg=p['text'],
                  activebackground=p['hover'],
                  command=self._stop_recv).pack(side=tk.LEFT, padx=(8, 0))
        tk.Button(row, text='选保存文件夹', relief='flat', bd=0, padx=10,
                  pady=5, bg=p['surface_alt'], fg=p['text'],
                  activebackground=p['hover'],
                  command=self._choose_folder).pack(side=tk.LEFT, padx=(8, 0))
        tk.Button(row, text='打开文件夹', relief='flat', bd=0, padx=10, pady=5,
                  bg=p['surface_alt'], fg=p['text'],
                  activebackground=p['hover'],
                  command=self._open_folder).pack(side=tk.LEFT, padx=(8, 0))

        self._recv_status = tk.Label(
            page, text='未接收 · 已自动开始检测局域网在线电脑（接收文件请点“开始接收”）',
            font=ui_theme.font(9), bg=p['bg'], fg=p['muted'], anchor='w')
        self._recv_status.pack(fill=tk.X, padx=2, pady=(6, 0))

        cols = ('name', 'size', 'time')
        self._recv_tree = ttk.Treeview(page, columns=cols, show='headings',
                                       height=14)
        for cid, title, w in (('name', '文件', 380), ('size', '大小', 90),
                              ('time', '时间', 90)):
            self._recv_tree.heading(cid, text=title)
            self._recv_tree.column(cid, width=w, anchor='w')
        self._recv_tree.pack(fill=tk.BOTH, expand=True, pady=(4, 0))
        tip = tk.Label(page, text='收到文件会在屏幕右下角即时提示；双击记录可打开该文件。',
                       font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                       anchor='w')
        tip.pack(fill=tk.X, pady=(4, 0))
        self._recv_tree.bind('<Double-1>', self._open_received_file)
        return page

    def _start_recv(self):
        if self._inbox is not None:
            return
        try:
            inbox = Inbox(folder=self._recv_folder(), max_files=500,
                          on_received=self._on_file_received)
            self._inbox = inbox
            inbox.start()
        except OSError as exc:
            self._inbox = None
            messagebox.showerror('局域网互传', '启动接收失败：%s' % exc, parent=self.win)
            return
        self._refresh_status()
        # 把之前会话里残留的记录补进列表
        for item in inbox.snapshot():
            self._handle_received_item(item, toast=False)

    def _stop_recv(self):
        if self._inbox is not None:
            self._inbox.stop()
            self._inbox = None
        # 注意：停止接收不停发现，在线检测一直保持
        self._recv_status.configure(
            text='未接收 · 在线电脑检测持续运行中（需收文件请重新点“开始接收”）',
            fg=ui_theme.PALETTE['muted'])

    def _recv_folder(self):
        override = getattr(self, '_recv_folder_override', None)
        if override:
            return ensure_folder(override)
        base = os.path.join(os.path.expanduser('~'), 'Documents', '会务互传')
        return ensure_folder(base)

    def _choose_folder(self):
        folder = filedialog.askdirectory(parent=self.win, title='选择保存到文件夹',
                                         mustexist=False)
        if not folder:
            return
        self._recv_folder_override = folder
        self._inbox and self._inbox.folder and setattr(self._inbox, 'folder', folder)
        self._refresh_status()

    def _open_folder(self):
        folder = getattr(self, '_recv_folder_override', None) or self._recv_folder()
        try:
            os.startfile(folder)
        except Exception:
            pass

    def _refresh_status(self):
        if self._inbox is not None:
            self._recv_status.configure(
                text='接收中  http://%s:%d   保存到：%s'
                     % (local_ip(), self._inbox.port, self._inbox.folder),
                fg=ui_theme.PALETTE['primary'])

    # ---------- 发送 ----------
    def _build_send(self, nb):
        page = tk.Frame(nb, bg=ui_theme.PALETTE['bg'])
        nb.add(page, text='  发送  ')
        p = ui_theme.PALETTE

        head = tk.Label(page, text='在线电脑', font=ui_theme.font(10, 'bold'),
                        bg=p['bg'], fg=p['text'], anchor='w')
        head.pack(fill=tk.X, pady=(10, 2))
        self._peer_head = head

        cols = ('host', 'ip', 'port', 'state')
        self._peer_tree = ttk.Treeview(page, columns=cols, show='headings',
                                       height=5)
        for cid, title, w in (('host', '机器名', 260), ('ip', 'IP', 150),
                              ('port', '端口', 70), ('state', '状态', 110)):
            self._peer_tree.heading(cid, text=title)
            self._peer_tree.column(cid, width=w, anchor='w')
        self._peer_tree.pack(fill=tk.X, pady=(0, 2))
        self._peer_tree.tag_configure('ready', foreground=p['success'])
        self._peer_tree.tag_configure('idle', foreground=p['muted'])
        self._peer_tree.tag_configure('gone', foreground=p['danger'])
        self._peer_tree.bind('<<TreeviewSelect>>', self._on_peer_select)
        self._peer_empty = tk.Label(
            page, text='未发现在线电脑：请确认对方电脑已打开本功能，且与本机在同一局域网（同 WiFi/同网段）。',
            font=ui_theme.font(9), bg=p['bg'], fg=p['warning'], anchor='w',
            wraplength=760, justify=tk.LEFT)
        self._peer_empty.pack(fill=tk.X, pady=(0, 6))

        manual = tk.Frame(page, bg=p['bg'])
        manual.pack(fill=tk.X)
        tk.Label(manual, text='或手动目标 IP：', font=ui_theme.font(9),
                 bg=p['bg'], fg=p['muted']).pack(side=tk.LEFT)
        self._ip_var = tk.StringVar()
        ip_e = ttk.Entry(manual, textvariable=self._ip_var, width=16)
        ip_e.pack(side=tk.LEFT, padx=(4, 8))
        tk.Label(manual, text='端口：', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        self._port_var = tk.StringVar(value='8000')
        ttk.Entry(manual, textvariable=self._port_var, width=7)\
            .pack(side=tk.LEFT, padx=(4, 8))
        tk.Button(manual, text='刷新', relief='flat', bd=0, padx=10, pady=2,
                  bg=p['surface_alt'], fg=p['text'], activebackground=p['hover'],
                  command=self._refresh_peers).pack(side=tk.LEFT)

        self._target_label = tk.Label(page, text='目标：未选择', font=ui_theme.font(9),
                                      bg=p['bg'], fg=p['muted'], anchor='w')
        self._target_label.pack(fill=tk.X, pady=(6, 0))

        sel = tk.Frame(page, bg=p['bg'])
        sel.pack(fill=tk.X, pady=(6, 0))
        tk.Button(sel, text='选择文件…', relief='flat', bd=0, padx=12, pady=5,
                  bg=p['surface_alt'], fg=p['text'], activebackground=p['hover'],
                  command=self._pick_files).pack(side=tk.LEFT)
        tk.Button(sel, text='发送', relief='flat', bd=0, padx=20, pady=5,
                  bg=p['primary'], fg='#ffffff', activebackground=p['primary_dark'],
                  command=self._send).pack(side=tk.LEFT, padx=(10, 0))
        tk.Button(sel, text='重试失败', relief='flat', bd=0, padx=10, pady=5,
                  bg=p['surface_alt'], fg=p['text'], activebackground=p['hover'],
                  command=self._retry_failed).pack(side=tk.LEFT, padx=(10, 0))

        self._send_tree = ttk.Treeview(page, columns=('name', 'size', 'pct'),
                                       show='headings', height=7)
        for cid, title, w in (('name', '文件', 320), ('size', '大小', 90),
                              ('pct', '进度', 160)):
            self._send_tree.heading(cid, text=title)
            self._send_tree.column(cid, width=w, anchor='w')
        self._send_tree.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        return page

    def _on_peer_select(self, event):
        sel = self._peer_tree.selection()
        if not sel:
            return
        vals = self._peer_tree.item(sel[0], 'values')
        if not vals or len(vals) < 3:
            return
        self._target_ip = vals[1]
        try:
            self._target_port = int(vals[2])
        except (TypeError, ValueError):
            self._target_port = 0
        self._target_name = vals[0]
        if self._target_port:
            self._target_label.configure(
                text='目标：%s  %s:%s（在线，可接收）'
                     % (self._target_name, self._target_ip, self._target_port),
                fg=ui_theme.PALETTE['success'])
        else:
            self._target_label.configure(
                text='目标：%s  %s（在线但对方尚未点“开始接收”，暂不能发送）'
                     % (self._target_name, self._target_ip),
                fg=ui_theme.PALETTE['warning'])

    def _refresh_peers(self):
        self.peers.purge()
        self._redraw_peers()

    def _redraw_peers(self):
        tree = self._peer_tree
        # 全量重绘会丢选中，按 IP 记住，重绘后恢复
        prev_sel = set(tree.selection())
        prev_ip = self._target_ip
        tree.delete(*tree.get_children())
        snap = self.peers.snapshot()
        ready = 0
        restore_iid = None
        for ip, info in snap:
            port = int(info.get('port') or 0)
            if port:
                ready += 1
                state_text, tag = '在线 · 可接收', 'ready'
            else:
                state_text, tag = '在线 · 未开接收', 'idle'
            iid = tree.insert('', 'end',
                              values=(info['name'], ip, port or '-',
                                      state_text),
                              tags=(tag,))
            if ip in prev_sel or ip == prev_ip:
                restore_iid = iid
        n = len(snap)
        try:
            self._peer_head.configure(
                text='在线电脑（实时检测：%d 台在线，其中 %d 台可接收）'
                     % (n, ready))
        except Exception:
            pass
        if n == 0:
            self._peer_empty.pack(fill=tk.X, pady=(0, 6))
        else:
            self._peer_empty.pack_forget()
        if restore_iid:
            try:
                tree.selection_set(restore_iid)
                tree.see(restore_iid)
            except tk.TclError:
                pass
        # 选中的目标已下线：提示用户
        if prev_ip and not any(ip == prev_ip for ip, _ in snap):
            self._target_label.configure(
                text='目标：%s（%s）已下线，请重新选择在线电脑'
                     % (self._target_name or '', prev_ip),
                fg=ui_theme.PALETTE['danger'])

    def _pick_files(self):
        paths = filedialog.askopenfilenames(parent=self.win, title='选择要发送的文件')
        if not paths:
            return
        items = []
        for pth in paths:
            try:
                size = os.path.getsize(pth)
            except OSError:
                size = 0
            items.append(_field(pth, size))
        self._send_fields = items
        self._redraw_send()

    def _send(self):
        if not self._target_ready():
            return
        fields = getattr(self, '_send_fields', None)
        if not fields:
            messagebox.showinfo('局域网互传', '请先选择要发送的文件', parent=self.win)
            return
        for f in fields:
            f['state'] = 'wait'
            f['pct'] = 0.0
        self._sender = _Sender((self._target_ip, self._target_port), fields,
                               self._on_send_item)
        self._sender.start()
        self._redraw_send()

    def _retry_failed(self):
        for f in getattr(self, '_send_fields', []):
            if f['state'] == 'err':
                f['state'] = 'wait'
                f['pct'] = 0.0
                f['error'] = ''
        if self._send_fields:
            self._redraw_send()

    def _on_send_item(self, fields):
        try:
            if self.win is not None and self.win.winfo_exists():
                self.win.after(0, self._redraw_send)
        except Exception:
            pass

    def _redraw_send(self):
        tree = self._send_tree
        tree.delete(*tree.get_children())
        for f in getattr(self, '_send_fields', []):
            st = f['state']
            tag = 'ok' if st == 'ok' else 'err' if st == 'err' \
                else 'run' if st == 'run' else 'wait'
            tree.insert('', 'end',
                        values=(os.path.basename(f['name']), f['port'],
                                fmt(f['pct'])),
                        tags=(tag,))

    def _target_ready(self):
        manual = (self._ip_var.get() or '').strip()
        if manual:
            ip = manual
            port_raw = (self._port_var.get() or '').strip()
            if not port_raw:
                messagebox.showinfo('局域网互传', '请填写目标端口', parent=self.win)
                return False
            try:
                port = int(port_raw)
            except (TypeError, ValueError):
                messagebox.showinfo('局域网互传', '目标端口不正确', parent=self.win)
                return False
        else:
            ip = getattr(self, '_target_ip', None)
            port = getattr(self, '_target_port', None)
            if not ip or port is None:
                messagebox.showinfo(
                    '局域网互传', '请先在列表点选在线电脑，或填写目标 IP/端口',
                    parent=self.win)
                return False
            try:
                port = int(port)
            except (TypeError, ValueError):
                messagebox.showinfo('局域网互传', '目标端口不正确', parent=self.win)
                return False
        self._target_ip = ip
        self._target_port = port
        if port <= 0:
            messagebox.showinfo(
                '局域网互传',
                '对方电脑已在线，但还没有点“开始接收”，暂时无法发送。\n'
                '请通知对方切换到“接收”页点【开始接收】后再发。',
                parent=self.win)
            return False
        return True

    # ---------- 发现与刷新 ----------
    def _current_port(self):
        """发现广播里携带的接收端口：没开始接收时为 0（仍可被看到“在线”）。"""
        return self._inbox.port if self._inbox is not None else 0

    def _start_discovery(self):
        if self._discover is not None:
            return
        self._discover = _Discover(self.peers, self._current_port)
        self._discover.start()

    def _stop_discover(self):
        d, self._discover = self._discover, None
        if d is not None:
            try:
                d.stop()
            except Exception:
                pass

    def _tick_ui(self):
        try:
            self._redraw_peers()
            # 兜底：即使 on_received 漏掉（如线程时序），也能从收件箱补记录
            if self._inbox is not None:
                for item in self._inbox.snapshot():
                    key = (item.get('path'), item.get('size'),
                           item.get('time'))
                    if key not in self._recv_seen:
                        self._handle_received_item(item, toast=True)
        except Exception:
            pass
        if self.win is not None and self.win.winfo_exists():
            self._after = self.win.after(2000, self._tick_ui)

    # ---------- 收到文件：列表 + 右下角提示 ----------
    def _on_file_received(self, item):
        """HTTP 工作线程回调：切回 Tk 主线程。"""
        try:
            if self.win is not None and self.win.winfo_exists():
                self.win.after(0,
                    lambda: self._handle_received_item(item, toast=True))
        except Exception:
            pass

    def _handle_received_item(self, item, toast=True):
        try:
            if self.win is None or not self.win.winfo_exists():
                return
        except Exception:
            return
        key = (item.get('path'), item.get('size'), item.get('time'))
        if key in self._recv_seen:
            return
        self._recv_seen.add(key)
        iid = str(self._recv_seq)
        self._recv_seq += 1
        self._recv_items[iid] = item
        self._recv_tree.insert('', 0, iid=iid,
                               values=(item.get('name', ''),
                                       fmt(item.get('size', 0)),
                                       item.get('time', '')))
        if toast:
            self._show_received_toast(item)

    def _short(self, text, limit=34):
        text = text or ''
        return text if len(text) <= limit else text[:limit - 1] + '…'

    def _show_received_toast(self, item):
        who = item.get('who') or item.get('ip') or '未知电脑'
        name = self._short(item.get('name', '文件'), 40)
        sub = '来自：%s · %s · %s' % (self._short(who, 20),
                                     fmt(item.get('size', 0)),
                                     item.get('time', ''))
        path = item.get('path', '')

        def open_folder():
            folder = os.path.dirname(path) or self._recv_folder()
            try:
                os.startfile(folder)
            except Exception:
                pass

        bubble = _ToastBubble(self, '收到文件', name, sub, open_folder)
        self._toasts.append(bubble)
        self._relayout_toasts()
        # 气泡最多同时叠 4 个，最老的提前消失
        while len(self._toasts) > 4:
            old = self._toasts[0]
            try:
                old.close()
            except Exception:
                pass

    def drop_toast(self, bubble):
        try:
            self._toasts.remove(bubble)
        except ValueError:
            pass
        self._relayout_toasts()

    def _relayout_toasts(self):
        for i, bubble in enumerate(reversed(self._toasts)):
            try:
                bubble.reposition(i)
            except Exception:
                pass

    def _open_received_file(self, _event=None):
        sel = self._recv_tree.selection()
        if not sel:
            return
        try:
            item = self._recv_items[sel[0]]
        except (KeyError, TypeError):
            return
        path = item.get('path', '')
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as exc:
                messagebox.showwarning('局域网互传', '无法打开文件：%s' % exc,
                                       parent=self.win)

    def close(self, force=False):
        self._stop_discover()
        if self._inbox is not None:
            try:
                self._inbox.stop()
            except Exception:
                pass
            self._inbox = None
        if self._after:
            try:
                self.win.after_cancel(self._after)
            except Exception:
                pass
            self._after = None
        for bubble in list(self._toasts):
            try:
                bubble.top.destroy()
            except Exception:
                pass
        self._toasts = []
        try:
            self.win.destroy()
        except Exception:
            pass
        self.win = None


def fmt(n):
    n = float(n or 0)
    if n < 1024:
        return '%d B' % n
    if n < 1048576:
        return '%.1f KB' % (n / 1024)
    if n < 1073741824:
        return '%.1f MB' % (n / 1048576)
    return '%.2f GB' % (n / 1073741824)


def _center(win, w, h):
    try:
        win.update_idletasks()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        win.geometry('+%d+%d' % (x, y))
    except Exception:
        pass