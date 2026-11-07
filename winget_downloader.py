# -*- coding: utf-8 -*-
"""软件下载器：winget 清单源，分类推荐 + 搜索，真实下载进度（无广告）。"""

import os
import re
import subprocess
import threading
import urllib.request
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import ui_theme

DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# 分类推荐（名称, winget ID）
RECO = [
    ("浏览器", [
        ("Chrome", "Google.Chrome"),
        ("Edge", "Microsoft.Edge"),
        ("Firefox", "Mozilla.Firefox"),
        ("Brave", "BraveSoftware.BraveBrowser"),
        ("Opera", "Opera.Opera"),
    ]),
    ("通讯", [
        ("微信", "Tencent.WeChat"),
        ("QQ", "Tencent.QQ"),
        ("TIM", "Tencent.TIM"),
        ("钉钉", "Alibaba.DingTalk.Mainland"),
        ("飞书", "ByteDance.Feishu"),
        ("Telegram", "Telegram.TelegramDesktop"),
    ]),
    ("会议", [
        ("腾讯会议", "Tencent.TencentMeeting"),
        ("Zoom", "Zoom.Zoom"),
    ]),
    ("办公", [
        ("WPS Office", "Kingsoft.WPSOffice"),
        ("LibreOffice", "TheDocumentFoundation.LibreOffice"),
        ("Notepad++", "Notepad++.Notepad++"),
        ("Adobe Acrobat Reader", "Adobe.Acrobat.Reader.64-bit"),
        ("calibre 电子书", "calibre.calibre"),
    ]),
    ("开发", [
        ("VS Code", "Microsoft.VisualStudioCode"),
        ("Git", "Git.Git"),
        ("Python 3.13", "Python.Python.3.13"),
        ("Node.js LTS", "OpenJS.NodeJS.LTS"),
        ("Windows Terminal", "Microsoft.WindowsTerminal"),
        ("PowerToys", "Microsoft.PowerToys"),
        ("Docker Desktop", "Docker.DockerDesktop"),
        ("GitHub CLI", "GitHub.cli"),
        ("PuTTY", "PuTTY.PuTTY"),
    ]),
    ("影音", [
        ("VLC 播放器", "VideoLAN.VLC"),
        ("OBS Studio 录屏", "OBSProject.OBSStudio"),
        ("Spotify", "Spotify.Spotify"),
    ]),
    ("图形", [
        ("GIMP", "GIMP.GIMP"),
        ("IrfanView", "IrfanSkiljan.IrfanView"),
        ("Inkscape", "Inkscape.Inkscape"),
        ("Krita", "KDE.Krita"),
        ("Blender", "BlenderFoundation.Blender"),
        ("Paint.NET", "dotPDN.PaintDotNet"),
    ]),
    ("工具", [
        ("7-Zip", "7zip.7zip"),
        ("Everything", "voidtools.Everything"),
        ("Snipaste", "liule.Snipaste"),
        ("qBittorrent", "qBittorrent.qBittorrent"),
        ("Rufus 启动U盘", "Rufus.Rufus"),
        ("AutoHotkey", "AutoHotkey.AutoHotkey"),
        ("VirtualBox", "Oracle.VirtualBox"),
        ("Bitwarden 密码", "Bitwarden.Bitwarden"),
    ]),
]


def run(cmd):
    """执行命令，返回 (返回码, 文本)。"""
    try:
        p = subprocess.run(cmd, capture_output=True, shell=True)
        raw = p.stdout.decode("utf-8", errors="replace")
        if "\uFFFD" in raw:  # utf-8 乱码回退 gbk
            try:
                raw = p.stdout.decode("gbk", errors="replace")
            except Exception:
                pass
        return p.returncode, raw.rstrip()
    except Exception as e:
        return -1, str(e)


def get_url(pid):
    """用 winget show 解析安装包下载 URL。"""
    code, out = run(
        f'winget show --id "{pid}" --accept-source-agreements --disable-interactivity'
    )
    for line in out.splitlines():
        if "安装程序 URL" in line:
            m = re.search(r"https?://\S+", line)
            if m:
                return m.group(0).rstrip(".")
    return None


class DownloaderPane(ttk.Frame):
    """下载器界面主体：放进任意父容器即可使用。

    会务助手把本 Frame 嵌在 Toplevel 子窗口里；直接运行本脚本时
    由下方 App(tk.Tk) 承载，两种方式共用同一套界面与逻辑。
    """

    def __init__(self, master):
        super().__init__(master, padding=(12, 10))
        self.dl = {"busy": False}
        self._build()
        self._load_reco(0)

    # ---------- 通用小组件 ----------
    def _make_listbox(self, parent, **kw):
        """白底细边框扁平列表，风格和主程序其他列表一致。"""
        p = ui_theme.PALETTE
        wrap = tk.Frame(parent, bg=p['border'], highlightthickness=0)
        lb = tk.Listbox(
            wrap, bg=p['surface'], fg=p['text'],
            selectbackground=p['primary'], selectforeground='#ffffff',
            selectborderwidth=0, activestyle='none',
            relief='flat', bd=0, highlightthickness=0,
            font=ui_theme.font(10), cursor='arrow', **kw)
        return wrap, lb

    # ---------- 界面 ----------
    def _build(self):
        p = ui_theme.PALETTE

        # 顶部搜索条（不切页签，始终可用）
        top = ttk.Frame(self)
        top.pack(fill='x', padx=6, pady=(0, 8))
        self.search_var = tk.StringVar()
        en = ttk.Entry(top, textvariable=self.search_var)
        en.pack(side='left', fill='x', expand=True, padx=(0, 6))
        en.bind('<Return>', lambda e: self.do_search())
        ttk.Button(top, text='搜索', style='Accent.TButton',
                   command=self.do_search).pack(side='right')

        # 主区域：左分类 + 右侧推荐列表/搜索结果（同位置切换显示）
        main = ttk.Frame(self)
        main.pack(fill='both', expand=True, padx=6)

        # 左侧分类
        lwrap, self.cat = self._make_listbox(main, width=12,
                                             selectmode='single', height=14)
        lwrap.pack(side='left', fill='y')
        self.cat.pack(side='left', fill='y', expand=True, padx=1, pady=1)
        lsb = ttk.Scrollbar(lwrap, orient='vertical', command=self.cat.yview)
        self.cat.configure(yscrollcommand=lsb.set)
        lsb.pack(side='right', fill='y')
        self.cat.bind('<<ListboxSelect>>', self._on_cat)

        right = ttk.Frame(main)
        right.pack(side='left', fill='both', expand=True, padx=(10, 0))

        # 右侧 A：推荐软件列表（默认显示）
        self.reco_wrap, self.soft = self._make_listbox(
            right, selectmode='single')
        self.soft.pack(side='left', fill='both', expand=True,
                       padx=1, pady=1)
        rsb = ttk.Scrollbar(self.reco_wrap, orient='vertical',
                            command=self.soft.yview)
        self.soft.configure(yscrollcommand=rsb.set)
        rsb.pack(side='right', fill='y')
        self.soft.bind('<Double-1>',
                       lambda e: self.start_download(self._active_pid()))

        # 右侧 B：搜索结果表（默认隐藏，搜索后替换推荐列表）
        self._pid_map = {}
        self.result_wrap = tk.Frame(right, bg=p['border'])
        self.tree = ttk.Treeview(self.result_wrap,
                                 columns=('name', 'pid', 'ver', 'src'),
                                 show='headings', height=10)
        self.tree.heading('name', text='软件')
        self.tree.heading('pid', text='软件 ID')
        self.tree.heading('ver', text='版本')
        self.tree.heading('src', text='来源')
        self.tree.column('name', width=200, anchor='w', stretch=False)
        self.tree.column('pid', width=240, anchor='w', stretch=True)
        self.tree.column('ver', width=110, anchor='center', stretch=False)
        self.tree.column('src', width=90, anchor='center', stretch=False)
        vsb = ttk.Scrollbar(self.result_wrap, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True,
                       padx=1, pady=1)
        vsb.pack(side='right', fill='y')
        self.tree.bind('<Double-1>', lambda e: self._dl_searched())
        self.reco_wrap.pack(fill='both', expand=True)

        # 底部下载区
        bar = ttk.Frame(self)
        bar.pack(fill='x', padx=6, pady=(10, 2))
        tk.Label(bar, text='保存到', font=ui_theme.font(10),
                 fg=p['text'], bg=p['bg']).pack(side='left')
        self.dir_var = tk.StringVar(value=DOWNLOAD_DIR)
        ttk.Entry(bar, textvariable=self.dir_var).pack(
            side='left', fill='x', expand=True, padx=(8, 6))
        ttk.Button(bar, text='选择目录', style='Ghost.TButton',
                   command=self.pick_dir).pack(side='left')
        self.dl_btn = ttk.Button(bar, text='下载', style='Accent.TButton',
                                 command=self._dl_current)
        self.dl_btn.pack(side='left', padx=(8, 0))

        pf = ttk.Frame(self)
        pf.pack(fill='x', padx=6)
        self.prog = ttk.Progressbar(pf, length=1, mode='determinate')
        self.prog.pack(side='left', fill='x', expand=True)
        self.percent = tk.Label(pf, text='', width=8, anchor='e',
                                font=ui_theme.font(9), fg=p['muted'],
                                bg=p['bg'])
        self.percent.pack(side='right', padx=(8, 0))

        self.status = tk.StringVar(
            value='就绪：顶部可直接搜索，或左侧选分类，双击软件名即可下载')
        tk.Label(self, textvariable=self.status, anchor='w',
                 font=ui_theme.font(9), fg=p['muted'],
                 bg=p['bg']).pack(fill='x', padx=6, pady=(4, 0))

    # ---------- 右侧视图切换 ----------
    def _show_reco_view(self):
        if self.result_wrap.winfo_manager():
            self.result_wrap.pack_forget()
        if not self.reco_wrap.winfo_manager():
            self.reco_wrap.pack(fill='both', expand=True)

    def _show_result_view(self):
        if self.reco_wrap.winfo_manager():
            self.reco_wrap.pack_forget()
        if not self.result_wrap.winfo_manager():
            self.result_wrap.pack(fill='both', expand=True)

    # ---------- 推荐 ----------
    def _load_reco(self, idx):
        self.cat.delete(0, 'end')
        self._cats = [c[0] for c in RECO]
        for i, c in enumerate(self._cats):
            self.cat.insert('end', c)
        self.cat.selection_set(idx)
        self._show_soft(idx)

    def _on_cat(self, _=None):
        sel = self.cat.curselection()
        if sel:
            self._show_reco_view()
            self._show_soft(sel[0])
            self.status.set(f'{self._cats[sel[0]]}：双击软件名即可下载')

    def _show_soft(self, idx):
        self._softs = RECO[idx][1]
        self.soft.delete(0, 'end')
        for name, _ in self._softs:
            self.soft.insert('end', name)
        self.soft.selection_clear(0, 'end')

    def _active_pid(self):
        if self.soft.curselection():
            return self._softs[self.soft.curselection()[0]][1]
        return None

    def _dl_current(self):
        if self.result_wrap.winfo_manager():
            self._dl_searched()
        else:
            self.start_download(self._active_pid())

    # ---------- 搜索 ----------
    def do_search(self):
        q = self.search_var.get().strip()
        if not q:
            return
        self.tree.delete(*self.tree.get_children())
        self._pid_map.clear()
        self.cat.selection_clear(0, 'end')
        self._show_result_view()
        self.status.set(f'搜索中：{q} ...')
        threading.Thread(target=self._search_worker, args=(q,),
                         daemon=True).start()

    def _search_worker(self, q):
        code, out = run(
            f'winget search -q "{q}" --accept-source-agreements --disable-interactivity'
        )
        rows = []
        for line in out.splitlines():
            parts = line.split()
            if len(parts) < 4 or parts[0] == '名称':
                continue
            for i in range(1, len(parts) - 1):
                if '.' in parts[i] and parts[i + 1][0].isdigit():
                    rows.append((' '.join(parts[:i]), parts[i + 1],
                                 parts[i], parts[-1]))
                    break
        self.after(0, lambda: self._show_rows(rows))

    def _show_rows(self, rows):
        for name, ver, pid, src in rows:
            iid = self.tree.insert(
                '', 'end', values=(name, pid, ver, src))
            self._pid_map[iid] = pid
        self.status.set(f'共 {len(rows)} 个结果，双击一行即可下载'
                        if rows else '未找到结果')

    def _dl_searched(self):
        sel = self.tree.selection()
        if not sel:
            return
        pid = self._pid_map.get(sel[0])
        if pid and '.' in pid:
            self.start_download(pid)

    # ---------- 通用下载 ----------
    def pick_dir(self):
        d = filedialog.askdirectory(initialdir=self.dir_var.get())
        if d:
            self.dir_var.set(d)

    def start_download(self, pid):
        if not pid or self.dl.get('busy'):
            return
        self.dl = {'busy': True, 'file': ''}
        self.dl_btn.configure(state='disabled')
        self.status.set(f'获取下载地址 {pid} ...')
        threading.Thread(target=self._dl_worker, args=(pid,),
                         daemon=True).start()

    def _dl_worker(self, pid):
        url = get_url(pid)
        if not url:
            self.after(0, lambda: self._dl_done(False,
                                                f'未找到 {pid} 的安装地址'))
            return
        name = os.path.basename(url.split('?')[0]) or f'{pid}.exe'
        dst = os.path.join(self.dir_var.get().strip() or DOWNLOAD_DIR, name)
        self.after(0, lambda: self.status.set(f'下载 {name}'))
        self.after(0, self._prog_reset)
        try:
            req = urllib.request.Request(
                url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=60) as r, \
                    open(dst, 'wb') as f:
                total = int(r.headers.get('Content-Length', 0))
                done, n = 0, 0
                while True:
                    chunk = r.read(65536)
                    if not chunk:
                        break
                    f.write(chunk)
                    done += len(chunk)
                    n += 1
                    if n % 16 == 0:
                        self.after(0,
                                   lambda d=done, t=total: self._prog_set(d, t))
                self.after(0, lambda: self._prog_set(done, total))
            self.after(0, lambda: self._dl_done(True, f'完成：{name}'))
        except Exception as e:
            self.after(0, lambda: self._dl_done(False, f'下载失败：{e}'))

    def _prog_reset(self):
        self.prog.configure(value=0, maximum=100)
        self.percent.configure(text='0%')

    def _prog_set(self, done, total):
        pct = int(done * 100 / total) if total else 0
        self.prog.configure(maximum=total or 1, value=done)
        self.percent.configure(text=f'{pct}%')

    def _dl_done(self, ok, msg):
        self.dl['busy'] = False
        self.dl_btn.configure(state='normal')
        self.status.set(msg)
        if ok:
            messagebox.showinfo('完成', msg)


class App(tk.Tk):
    """独立运行入口：python winget_downloader.py。"""

    def __init__(self):
        super().__init__()
        self.title('软件下载')
        self.geometry('800x600')
        self.configure(bg=ui_theme.PALETTE['bg'])
        ui_theme.apply_theme(self)
        pane = DownloaderPane(self)
        pane.pack(fill='both', expand=True)


if __name__ == '__main__':
    App().mainloop()
