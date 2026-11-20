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

# 内置精选软件库（名称, winget ID）：覆盖会务现场常用场景，并支持 winget 全库搜索
RECO = [
    ("会务必备", [
        ("腾讯会议", "Tencent.TencentMeeting"),
        ("Zoom", "Zoom.Zoom"),
        ("OBS Studio 录屏", "OBSProject.OBSStudio"),
        ("VLC 播放器", "VideoLAN.VLC"),
        ("Snipaste 截图", "liule.Snipaste"),
        ("微信", "Tencent.WeChat"),
        ("钉钉", "Alibaba.DingTalk.Mainland"),
        ("飞书", "ByteDance.Feishu"),
        ("WPS Office", "Kingsoft.WPSOffice"),
        ("Adobe Acrobat Reader", "Adobe.Acrobat.Reader.64-bit"),
        ("7-Zip", "7zip.7zip"),
        ("Everything 搜索", "voidtools.Everything"),
    ]),
    ("浏览器", [
        ("Chrome", "Google.Chrome"),
        ("Edge", "Microsoft.Edge"),
        ("Firefox", "Mozilla.Firefox"),
        ("Brave", "BraveSoftware.BraveBrowser"),
        ("Opera", "Opera.Opera"),
        ("Vivaldi", "Vivaldi.Vivaldi"),
        ("Chromium", "Chromium.Chromium"),
    ]),
    ("通讯协作", [
        ("微信", "Tencent.WeChat"),
        ("QQ", "Tencent.QQ"),
        ("TIM", "Tencent.TIM"),
        ("钉钉", "Alibaba.DingTalk.Mainland"),
        ("飞书", "ByteDance.Feishu"),
        ("Telegram", "Telegram.TelegramDesktop"),
        ("Discord", "Discord.Discord"),
        ("Microsoft Teams", "Microsoft.Teams"),
        ("Slack", "Slack.Slack"),
    ]),
    ("会议直播", [
        ("腾讯会议", "Tencent.TencentMeeting"),
        ("Zoom", "Zoom.Zoom"),
        ("OBS Studio 录屏", "OBSProject.OBSStudio"),
        ("Streamlabs OBS", "Streamlabs.StreamlabsOBS"),
        ("Microsoft Teams", "Microsoft.Teams"),
    ]),
    ("办公文档", [
        ("WPS Office", "Kingsoft.WPSOffice"),
        ("LibreOffice", "TheDocumentFoundation.LibreOffice"),
        ("Notepad++", "Notepad++.Notepad++"),
        ("Adobe Acrobat Reader", "Adobe.Acrobat.Reader.64-bit"),
        ("Foxit PDF Reader", "Foxit.FoxitReader"),
        ("calibre 电子书", "calibre.calibre"),
        ("ONLYOFFICE 桌面版", "ONLYOFFICE.DesktopEditors"),
        ("Draw.io 流程图", "jgraph.Draw"),
        ("Typora", "Typora.Typora"),
        ("Obsidian", "Obsidian.Obsidian"),
    ]),
    ("开发编程", [
        ("VS Code", "Microsoft.VisualStudioCode"),
        ("Git", "Git.Git"),
        ("Python 3.13", "Python.Python.3.13"),
        ("Node.js LTS", "OpenJS.NodeJS.LTS"),
        ("Windows Terminal", "Microsoft.WindowsTerminal"),
        ("PowerToys", "Microsoft.PowerToys"),
        ("Docker Desktop", "Docker.DockerDesktop"),
        ("GitHub CLI", "GitHub.cli"),
        ("PuTTY", "PuTTY.PuTTY"),
        ("Postman", "Postman.Postman"),
        ("DBeaver Community", "DBeaver.DBeaverCommunity"),
        ("IntelliJ IDEA Community", "JetBrains.IntelliJIDEA.Community"),
        ("Visual Studio 2022 Community", "Microsoft.VisualStudio.2022.Community"),
        ("Go", "GoLang.Go"),
        ("CMake", "Kitware.CMake"),
    ]),
    ("影音媒体", [
        ("VLC 播放器", "VideoLAN.VLC"),
        ("Spotify", "Spotify.Spotify"),
        ("Audacity 音频剪辑", "Audacity.Audacity"),
        ("FFmpeg", "FFmpeg.FFmpeg"),
        ("HandBrake 转码", "HandBrake.HandBrake"),
        ("MKVToolNix", "MKVToolNix.MKVToolNix"),
        ("foobar2000", "foobar2000.foobar2000"),
        ("Apple iTunes", "Apple.iTunes"),
    ]),
    ("图像设计", [
        ("GIMP", "GIMP.GIMP"),
        ("IrfanView", "IrfanSkiljan.IrfanView"),
        ("Inkscape", "Inkscape.Inkscape"),
        ("Krita", "KDE.Krita"),
        ("Blender", "BlenderFoundation.Blender"),
        ("Paint.NET", "dotPDN.PaintDotNet"),
        ("ShareX 截图录屏", "ShareX.ShareX"),
        ("ImageMagick", "ImageMagick.ImageMagick"),
        ("Figma", "Figma.Figma"),
    ]),
    ("系统工具", [
        ("7-Zip", "7zip.7zip"),
        ("Everything 搜索", "voidtools.Everything"),
        ("Snipaste 截图", "liule.Snipaste"),
        ("Rufus 启动U盘", "Rufus.Rufus"),
        ("AutoHotkey", "AutoHotkey.AutoHotkey"),
        ("VirtualBox", "Oracle.VirtualBox"),
        ("Bitwarden 密码", "Bitwarden.Bitwarden"),
        ("Microsoft PowerToys", "Microsoft.PowerToys"),
        ("Process Explorer", "Microsoft.Sysinternals.ProcessExplorer"),
        ("Process Monitor", "Microsoft.Sysinternals.ProcessMonitor"),
        ("CrystalDiskInfo", "CrystalDewWorld.CrystalDiskInfo"),
        ("WinDirStat", "WinDirStat.WinDirStat"),
        ("TeamViewer", "TeamViewer.TeamViewer"),
    ]),
    ("传输下载", [
        ("qBittorrent", "qBittorrent.qBittorrent"),
        ("FileZilla", "FileZilla.FileZilla"),
        ("WinSCP", "WinSCP.WinSCP"),
        ("aria2 下载器", "aria2.aria2"),
        ("Free Download Manager", "FreeDownloadManager.FreeDownloadManager"),
        ("JDownloader", "JDownloader.JDownloader"),
    ]),
    ("阅读学习", [
        ("calibre 电子书", "calibre.calibre"),
        ("SumatraPDF", "SumatraPDF.SumatraPDF"),
        ("Foxit PDF Reader", "Foxit.FoxitReader"),
        ("Zotero 文献管理", "Zotero.Zotero"),
        ("Anki 记忆卡", "Anki.Anki"),
    ]),
]

# 精选条目的用途说明；其余条目自动用「分类 + winget ID」生成说明
SOFT_DESC = {
    "Tencent.TencentMeeting": "会务远程会议与线上协作，屏幕共享稳定。",
    "Zoom.Zoom": "国际会议常用客户端，适合外部嘉宾接入。",
    "OBSProject.OBSStudio": "现场录制与直播推流，支持多场景切换。",
    "VideoLAN.VLC": "本地音视频播放，会务视频循环播放更稳。",
    "liule.Snipaste": "截图贴图与标注，现场讲解很好用。",
    "Kingsoft.WPSOffice": "文档、表格、演示一体化办公套件。",
    "Adobe.Acrobat.Reader.64-bit": "PDF 阅读与批注，兼容性好。",
    "7zip.7zip": "高压缩比解压缩工具，资料打包传输必备。",
    "voidtools.Everything": "按文件名秒级搜索本机文件。",
    "Microsoft.VisualStudioCode": "轻量代码编辑器，支持多种扩展。",
    "Git.Git": "版本控制客户端，开发环境基础组件。",
    "Python.Python.3.13": "Python 运行环境，脚本与工具开发常用。",
    "Microsoft.WindowsTerminal": "现代化命令行终端，多标签管理。",
    "Microsoft.PowerToys": "Windows 效率增强工具集。",
    "ShareX.ShareX": "开源截图录屏与快速分享工具。",
    "Audacity.Audacity": "音频剪辑、降噪与格式转换。",
    "FFmpeg.FFmpeg": "音视频转码与处理的命令行工具。",
    "GIMP.GIMP": "开源图像编辑，可替代部分 Photoshop 场景。",
    "Inkscape.Inkscape": "矢量图设计，适合制作会场标识。",
    "BlenderFoundation.Blender": "三维建模与动画制作。",
    "Bitwarden.Bitwarden": "跨平台密码管理器。",
    "qBittorrent.qBittorrent": "无广告 BT 下载客户端。",
    "FileZilla.FileZilla": "FTP/SFTP 文件传输客户端。",
    "WinSCP.WinSCP": "Windows 下常用的 SFTP/SCP 工具。",
    "calibre.calibre": "电子书管理与格式转换。",
    "Zotero.Zotero": "文献收集、整理与引用管理。",
    "Anki.Anki": "间隔重复记忆卡片工具。",
}


def describe(name, pid, category):
    """给软件生成一句话说明，没有精选说明时用分类兜底。"""
    return SOFT_DESC.get(pid) or f"{category}类工具 · winget ID：{pid}"


def all_software():
    """展开内置库，供本地搜索与详情展示使用。"""
    return [(category, name, pid)
            for category, items in RECO
            for name, pid in items]


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
        self.searching = False
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
        self.online_btn = ttk.Button(top, text='全库',
                                     style='Ghost.TButton',
                                     command=self.do_online_search)
        self.online_btn.pack(side='right', padx=(6, 0))
        self.search_btn = ttk.Button(top, text='搜索',
                                     style='Accent.TButton',
                                     command=self.do_search)
        self.search_btn.pack(side='right')

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

        # 右侧上半区：推荐列表 / 搜索结果同位置切换
        list_host = ttk.Frame(right)
        list_host.pack(fill='both', expand=True)
        self._list_host = list_host

        # 右侧 A：推荐软件列表（默认显示）
        self.reco_wrap, self.soft = self._make_listbox(
            list_host, selectmode='single')
        self.soft.pack(side='left', fill='both', expand=True,
                       padx=1, pady=1)
        rsb = ttk.Scrollbar(self.reco_wrap, orient='vertical',
                            command=self.soft.yview)
        self.soft.configure(yscrollcommand=rsb.set)
        rsb.pack(side='right', fill='y')
        self.soft.bind('<Double-1>',
                       lambda e: self.start_download(self._active_pid()))
        self.soft.bind('<<ListboxSelect>>', self._on_soft_select)

        # 右侧 B：搜索结果表（默认隐藏，搜索后替换推荐列表）
        self._pid_map = {}
        self._row_info = {}
        self.result_wrap = tk.Frame(list_host, bg=p['border'])
        self.tree = ttk.Treeview(self.result_wrap,
                                 columns=('name', 'pid', 'ver', 'src'),
                                 show='headings', height=10)
        self.tree.heading('name', text='软件')
        self.tree.heading('pid', text='软件 ID')
        self.tree.heading('ver', text='版本')
        self.tree.heading('src', text='来源')
        self.tree.column('name', width=200, anchor='w', stretch=False)
        self.tree.column('pid', width=240, anchor='w', stretch=True)
        self.tree.column('ver', width=100, anchor='center', stretch=False)
        self.tree.column('src', width=100, anchor='center', stretch=False)
        vsb = ttk.Scrollbar(self.result_wrap, orient='vertical',
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True,
                       padx=1, pady=1)
        vsb.pack(side='right', fill='y')
        self.tree.bind('<Double-1>', lambda e: self._dl_searched())
        self.tree.bind('<<TreeviewSelect>>', self._on_tree_select)
        self.reco_wrap.pack(fill='both', expand=True)

        # 右侧下半区：软件详情，让内置库不只是“一个名字列表”
        detail = ttk.LabelFrame(right, text='软件详情', padding=(10, 8))
        detail.pack(fill='x', pady=(8, 0))
        self.detail_name = tk.Label(
            detail, text='请选择一个软件', anchor='w',
            font=ui_theme.font(11, 'bold'), fg=p['text'], bg=p['surface'])
        self.detail_name.pack(fill='x')
        self.detail_desc = tk.Label(
            detail, text='左侧选择分类，或在顶部搜索；双击软件可直接下载安装包。',
            anchor='w', justify='left', wraplength=620,
            font=ui_theme.font(9), fg=p['muted'], bg=p['surface'])
        self.detail_desc.pack(fill='x', pady=(3, 0))
        self.detail_meta = tk.Label(
            detail, text='', anchor='w', font=ui_theme.font(8),
            fg=p['faint'], bg=p['surface'])
        self.detail_meta.pack(fill='x', pady=(2, 0))
        detail_actions = ttk.Frame(detail)
        detail_actions.pack(fill='x', pady=(6, 0))
        self.install_btn = ttk.Button(
            detail_actions, text='静默安装', style='Ghost.TButton',
            command=self.install_selected)
        self.install_btn.pack(side='left')
        tk.Label(detail_actions,
                 text='下载只保存安装包；静默安装由 winget 执行',
                 font=ui_theme.font(8), fg=p['faint'],
                 bg=p['surface']).pack(side='left', padx=(10, 0))

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
            value=(f'内置库 {sum(len(items) for _, items in RECO)} 项 · '
                   '顶部“搜索”先查精选库，“全库”走 winget 在线目录'))
        tk.Label(self, textvariable=self.status, anchor='w',
                 font=ui_theme.font(9), fg=p['muted'],
                 bg=p['bg']).pack(fill='x', padx=6, pady=(4, 0))

    # ---------- 搜索状态 ----------
    def _set_searching(self, busy, text=None):
        """搜索时给出明确反馈：按钮禁用 + 文字 + 不确定进度条。"""
        self.searching = busy
        state = 'disabled' if busy else 'normal'
        self.search_btn.configure(state=state,
                                  text='搜索中…' if busy else '搜索')
        self.online_btn.configure(state=state,
                                  text='全库中…' if busy else '全库')
        if busy:
            self.prog.configure(mode='indeterminate')
            self.prog.start(12)
            self.percent.configure(text='')
        else:
            try:
                self.prog.stop()
            except Exception:
                pass
            self.prog.configure(mode='determinate', value=0, maximum=100)
            self.percent.configure(text='')
        if text is not None:
            self.status.set(text)

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
        self._update_detail('', '', '', '')

    def _on_soft_select(self, _=None):
        sel = self.soft.curselection()
        if not sel or not getattr(self, '_softs', None):
            return
        name, pid = self._softs[sel[0]]
        cat_sel = self.cat.curselection()
        category = self._cats[cat_sel[0]] if cat_sel else '内置精选'
        self._update_detail(name, pid, category, '内置精选')

    def _on_tree_select(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        name, pid, category, source = self._row_info.get(
            sel[0], ('', '', '', ''))
        self._update_detail(name, pid, category, source)

    def _update_detail(self, name, pid, category, source):
        if not name:
            self.detail_name.config(text='请选择一个软件')
            self.detail_desc.config(
                text='左侧选择分类，或在顶部搜索；双击软件可直接下载安装包。')
            self.detail_meta.config(text='')
            return
        self.detail_name.config(text=name)
        self.detail_desc.config(text=describe(name, pid, category))
        bits = [f'分类：{category}'] if category else []
        if source:
            bits.append(f'来源：{source}')
        if pid:
            bits.append(f'ID：{pid}')
        self.detail_meta.config(text='   ·   '.join(bits))

    def _active_info(self):
        if self.result_wrap.winfo_manager():
            sel = self.tree.selection()
            if not sel:
                return None
            info = self._row_info.get(sel[0])
            if info:
                return info
            pid = self._pid_map.get(sel[0])
            return ('搜索结果', pid, 'winget 全库', 'winget') if pid else None
        sel = self.soft.curselection()
        if not sel or not getattr(self, '_softs', None):
            return None
        name, pid = self._softs[sel[0]]
        cat_sel = self.cat.curselection()
        category = self._cats[cat_sel[0]] if cat_sel else '内置精选'
        return (name, pid, category, '内置精选')

    def _active_pid(self):
        info = self._active_info()
        return info[1] if info else None

    def _dl_current(self):
        if self.result_wrap.winfo_manager():
            self._dl_searched()
        else:
            self.start_download(self._active_pid())

    # ---------- 搜索 ----------
    def do_search(self):
        """先搜内置精选库，查不到再自动回落到 winget 全库。"""
        q = self.search_var.get().strip()
        if not q or self.searching:
            return
        self._set_searching(True, '正在搜索内置库...')
        needle = q.lower()
        local = []
        for category, name, pid in all_software():
            text = f'{name} {pid} {describe(name, pid, category)}'.lower()
            if needle in text:
                local.append((name, '推荐', pid, '内置库', category))
        self.tree.delete(*self.tree.get_children())
        self._pid_map.clear()
        self._row_info.clear()
        self.cat.selection_clear(0, 'end')
        self._show_result_view()
        if local:
            rows = [(name, ver, pid, src)
                    for name, ver, pid, src, _cat in local]
            self._show_rows(rows)
            for iid, (_n, _v, _p, _s, cat) in zip(
                    self.tree.get_children(), local):
                self._row_info[iid] = (_n, _p, cat, '内置库')
            self.status.set(
                f'内置库找到 {len(local)} 个结果；点“全库”可搜索 winget 在线目录')
            return
        self.status.set(f'内置库无结果，正在搜索 winget 全库：{q} ...')
        self._online_search(q)

    def do_online_search(self):
        """直接使用 winget 搜索在线全库。"""
        q = self.search_var.get().strip()
        if not q:
            self.status.set('请先输入要搜索的软件名或 ID')
            return
        if self.searching:
            return
        self._set_searching(True, f'正在搜索 winget 全库：{q} ...')
        self.cat.selection_clear(0, 'end')
        self._show_result_view()
        self._online_search(q)

    def _online_search(self, q):
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
        self.after(0, lambda: self._show_rows(rows, online=True))

    def _show_rows(self, rows, online=False):
        self._set_searching(False)
        self.tree.delete(*self.tree.get_children())
        self._pid_map.clear()
        self._row_info.clear()
        for name, ver, pid, src in rows:
            iid = self.tree.insert(
                '', 'end', values=(name, pid, ver, src))
            self._pid_map[iid] = pid
            self._row_info[iid] = (
                name, pid, 'winget 全库' if online else src,
                'winget' if online else src)
        self.status.set(f'共 {len(rows)} 个结果，双击一行即可下载'
                        if rows else '未找到结果')

    def _dl_searched(self):
        sel = self.tree.selection()
        if not sel:
            return
        pid = self._pid_map.get(sel[0])
        if pid and '.' in pid:
            self.start_download(pid)

    # ---------- 通用下载 / 安装 ----------
    def pick_dir(self):
        d = filedialog.askdirectory(initialdir=self.dir_var.get())
        if d:
            self.dir_var.set(d)

    def _set_busy(self, busy):
        self.dl['busy'] = busy
        self.dl_btn.configure(state='disabled' if busy else 'normal')
        self.install_btn.configure(state='disabled' if busy else 'normal')

    def start_download(self, pid):
        if not pid or self.dl.get('busy'):
            return
        self._set_busy(True)
        self._prog_reset()
        self.status.set(f'获取下载地址 {pid} ...')
        threading.Thread(target=self._dl_worker, args=(pid,),
                         daemon=True).start()

    def install_selected(self):
        """用 winget 静默安装当前选中的软件；必须先选中并二次确认。"""
        info = self._active_info()
        if not info:
            messagebox.showinfo('静默安装', '请先选中一个软件。', parent=self)
            return
        name, pid = info[0], info[1]
        if not pid or '.' not in pid:
            messagebox.showinfo('静默安装', '当前条目没有可用的 winget ID。',
                                parent=self)
            return
        if self.dl.get('busy'):
            return
        if not messagebox.askyesno(
                '静默安装',
                f'将使用 winget 安装：\n{name}\n{pid}\n\n确定继续吗？',
                parent=self):
            return
        self._set_busy(True)
        self.status.set(f'正在静默安装 {name} ...')
        threading.Thread(target=self._install_worker, args=(name, pid),
                         daemon=True).start()

    def _install_worker(self, name, pid):
        code, out = run(
            f'winget install --id "{pid}" --silent '
            '--accept-package-agreements --accept-source-agreements '
            '--disable-interactivity'
        )
        tail = ' '.join(out.splitlines()[-3:])[:160]
        ok = code == 0
        msg = f'已安装：{name}' if ok else f'安装失败：{name}\n{tail}'
        self.after(0, lambda: self._install_done(ok, msg))

    def _install_done(self, ok, msg):
        self._set_busy(False)
        self.status.set(msg)
        if ok:
            messagebox.showinfo('安装完成', msg, parent=self)
        else:
            messagebox.showerror('安装失败', msg, parent=self)

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
        try:
            self.prog.stop()
        except Exception:
            pass
        self.prog.configure(mode='determinate', value=0, maximum=100)
        self.percent.configure(text='0%')

    def _prog_set(self, done, total):
        pct = int(done * 100 / total) if total else 0
        self.prog.configure(maximum=total or 1, value=done)
        self.percent.configure(text=f'{pct}%')

    def _dl_done(self, ok, msg):
        self._set_busy(False)
        self.status.set(msg)
        if ok:
            messagebox.showinfo('完成', msg, parent=self)


class App(tk.Tk):
    """独立运行入口：python winget_downloader.py。"""

    def __init__(self):
        super().__init__()
        self.title('软件库')
        self.geometry('900x640')
        self.configure(bg=ui_theme.PALETTE['bg'])
        ui_theme.apply_theme(self)
        pane = DownloaderPane(self)
        pane.pack(fill='both', expand=True)


if __name__ == '__main__':
    App().mainloop()
