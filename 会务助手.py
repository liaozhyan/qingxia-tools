# -*- coding: utf-8 -*-
"""会务助手主界面：聚合会务相关工具入口。

- 会务PPT转换：PPT 转图片 / 转 16:9 全屏 PPT；
- 生成欢迎语：输入内容 + 纯色/上传背景图，直接生成一页 PPT；
- PPT倒计时：浮动透明倒计时，与演讲 PPT 放映联动，
  右键菜单设置字体/颜色/时长/投屏显示器（默认扩展屏），
  主屏同时显示一份可操作的副本，设置面板也在主屏弹出；
- 打印：在主界面的子窗口里打开远程打印功能。
  打印窗口依赖主窗口：同一时间只能打开一个；
  关闭主界面时，打印子窗口会一并关闭。
"""

import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import webbrowser
import winsound
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont

import safe_guard
import ui_theme
import file_inbox
import qrcode_lite
import video_player
import lottery
import strip_screen
import winget_downloader
import prompter
import lan_share
import av_tools
import image_tools
import focus_guard

try:
    import screenshot as _screenshot
except Exception:
    _screenshot = None

# native_countdown 可选导入——只有在 Win32 + 失败的 fallback 救回时需要；
# import 失败 / 跨平台就置 None，逻辑里按 is None 处理
try:
    import native_countdown as _native_countdown
except Exception:
    _native_countdown = None

# tray_icon 可选导入——把倒计时藏到任务栏右下角图标用；
# 任何 import 失败 / 非 Windows 就置 None，逻辑里按 is None 降级
try:
    import tray_icon as _tray_icon
except Exception:
    _tray_icon = None


WINDOW_TITLE = '会务助手'
WINDOW_BG = ui_theme.PALETTE['bg']
CARD_BG = ui_theme.PALETTE['surface']
BORDER_COLOR = ui_theme.PALETTE['border']
PRIMARY = ui_theme.PALETTE['primary']
ACCENT = ui_theme.PALETTE['primary_soft']

# 在线音乐下载上限：超过直接拒绝，避免把盘/内存写爆
MAX_MUSIC_BYTES = 100 * 1024 * 1024


def has_audio_output():
    """系统当前有没有可用的音频输出设备（Windows waveOut）。

    现场最常见的“点了播放没声音”其实是音箱/耳机没插、声卡被禁用：
    这时候任何播放器都放不出声，播放器状态却仍显示“播放中”，
    所以界面要单独给一句提示，不能让人以为是软件坏了。
    非 Windows 或探测失败一律当作“有”，避免误报。
    """
    if not sys.platform.startswith('win'):
        return True
    try:
        import ctypes
        return int(ctypes.windll.winmm.waveOutGetNumDevs()) > 0
    except Exception:
        return True

PURE_COLORS = [
    ('中国红', '#C62828'),
    ('枣红', '#8C1D40'),
    ('深蓝', '#1E3A8A'),
    ('青绿', '#0E7490'),
    ('墨绿', '#14532D'),
    ('深紫', '#4C1D95'),
    ('暖灰', '#4B5563'),
    ('墨黑', '#1F2937'),
]

if getattr(sys, 'frozen', False):
    # 打包后内置资源在临时解压目录里
    _BASE_DIR = sys._MEIPASS
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _welcome_default_name(content):
    """按输入文字拼存盘文件名：清理非法字符、截断超长、加时间戳防重名。"""
    s = (content or '').replace('\r', ' ')
    s = s.replace('\x0b', ' ').replace('\n', ' ')
    s = ' '.join(s.split())
    for ch in '<>:"/\\|?*':
        s = s.replace(ch, '')
    s = s.rstrip(' ._')
    if len(s) > 26:
        s = s[:26].rstrip(' ._')
    if not s:
        s = '欢迎语'
    return '%s_%s.pptx' % (s, time.strftime('%Y%m%d_%H%M%S'))


def _center_window(win, width, height):
    """把子窗口摆到屏幕中间（跟随主界面所在那块显示器），不要落在左上角。"""
    rect = None
    try:
        master = win.master
        rx = master.winfo_rootx() + master.winfo_width() // 2
        ry = master.winfo_rooty() + master.winfo_height() // 2
        for item in video_player.list_monitors():
            left, top, right, bottom = item['rect']
            if left <= rx < right and top <= ry < bottom:
                rect = item['rect']
                break
    except Exception:
        rect = None
    if rect is None:
        try:
            for item in video_player.list_monitors():
                if item['primary']:
                    rect = item['rect']
                    break
        except Exception:
            rect = None
    if rect is None:
        rect = (0, 0, win.winfo_screenwidth(), win.winfo_screenheight())
    left, top, right, bottom = rect
    x = max(left, left + ((right - left) - int(width)) // 2)
    y = max(top, top + ((bottom - top) - int(height)) // 2)
    win.geometry('%dx%d+%d+%d' % (int(width), int(height), x, y))


def guard_color_dialog_visible(anchor, title='选择颜色'):
    """把系统"选择颜色"对话框拉回 anchor 所在屏幕（后台线程兜底）。

    为什么需要：askcolor 是阻塞的原生模态框，它弹出期间 Tk 的 after
    定时器不会跑，只能从旁边的线程收拾它：
      * 双屏摆放可能有"死区"（两块屏之间的空隙，不属任何显示器），
        以父窗口居中的对话框万一落进死区 → 弹了也看不见；
      * 个别情况对话框会被压到 Z 序底下，用户以为没弹出来。
    做法：延迟几拍后，在本进程可见的 #32770（通用对话框）里找标题
    匹配的那个，检查其矩形是否完整落在 anchor 所在显示器内；不在就
    挪进去，并抬到 Z 序前面。找不到 / 非 Windows / 任何异常都静默。

    **注意**：这里绝不给 user32 的共享函数对象设 argtypes/restype——
    那是进程级全局（video_player.list_monitors 等模块也在用同一批
    函数，设了私有结构体会把它们全部毒化）。统一用"函数原型 + 地址"
    生成只属于本函数的私有副本。
    """
    try:
        import ctypes
        import threading
        from ctypes import wintypes

        _u32 = ctypes.windll.user32

        def _fn(name, restype, *argtypes):
            """按地址拿一份带类型签名的私有函数副本，不碰共享对象。"""
            proto = ctypes.WINFUNCTYPE(restype, *argtypes)
            addr = ctypes.cast(getattr(_u32, name), ctypes.c_void_p).value
            return proto(addr)

        class _MONITORINFO(ctypes.Structure):
            _fields_ = [('cbSize', wintypes.DWORD),
                        ('rcMonitor', wintypes.RECT),
                        ('rcWork', wintypes.RECT),
                        ('dwFlags', wintypes.DWORD)]

        WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, ctypes.c_void_p,
                                         wintypes.LPARAM)
        _enum_windows = _fn('EnumWindows', wintypes.BOOL, WNDENUMPROC,
                            wintypes.LPARAM)
        _get_class = _fn('GetClassNameW', ctypes.c_int, ctypes.c_void_p,
                         wintypes.LPWSTR, ctypes.c_int)
        _get_text = _fn('GetWindowTextW', ctypes.c_int, ctypes.c_void_p,
                        wintypes.LPWSTR, ctypes.c_int)
        _get_text_len = _fn('GetWindowTextLengthW', ctypes.c_int,
                            ctypes.c_void_p)
        _is_visible = _fn('IsWindowVisible', wintypes.BOOL, ctypes.c_void_p)
        _get_rect = _fn('GetWindowRect', wintypes.BOOL, ctypes.c_void_p,
                        ctypes.POINTER(wintypes.RECT))
        _get_pid = _fn('GetWindowThreadProcessId', wintypes.DWORD,
                       ctypes.c_void_p, ctypes.POINTER(wintypes.DWORD))
        _get_ancestor = _fn('GetAncestor', ctypes.c_void_p, ctypes.c_void_p,
                            wintypes.UINT)
        _monitor_from = _fn('MonitorFromWindow', ctypes.c_void_p,
                            ctypes.c_void_p, wintypes.DWORD)
        _monitor_info = _fn('GetMonitorInfoW', wintypes.BOOL, ctypes.c_void_p,
                            ctypes.POINTER(_MONITORINFO))
        _set_pos = _fn('SetWindowPos', wintypes.BOOL, ctypes.c_void_p,
                       ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
                       ctypes.c_int, ctypes.c_int, wintypes.UINT)

        GA_ROOT = 2
        MONITOR_DEFAULTTONEAREST = 2
        HWND_TOP = 0
        SWP_NOSIZE = 0x0001
        SWP_NOMOVE = 0x0002
        SWP_NOACTIVATE = 0x0010
        pid = os.getpid()

        def _attempt():
            try:
                found = []

                @WNDENUMPROC
                def _cb(hwnd, _lp):
                    try:
                        cls = ctypes.create_unicode_buffer(64)
                        _get_class(hwnd, cls, 64)
                        if cls.value != '#32770':
                            return True
                        if not _is_visible(hwnd):
                            return True
                        dw = wintypes.DWORD()
                        _get_pid(hwnd, ctypes.byref(dw))
                        if dw.value != pid:
                            return True
                        n = _get_text_len(hwnd)
                        buf = ctypes.create_unicode_buffer(n + 1)
                        _get_text(hwnd, buf, n + 1)
                        if title and title not in buf.value:
                            return True
                        found.append(hwnd)
                    except Exception:
                        pass
                    return True

                _enum_windows(_cb, 0)
                if not found:
                    return
                # anchor 所在显示器；拿不到就退回主屏
                mon_rect = None
                try:
                    ah = _get_ancestor(anchor.winfo_id(), GA_ROOT)
                    mon = _monitor_from(ah, MONITOR_DEFAULTTONEAREST) \
                        if ah else None
                    if mon:
                        mi = _MONITORINFO()
                        mi.cbSize = ctypes.sizeof(_MONITORINFO)
                        if _monitor_info(mon, ctypes.byref(mi)):
                            r = mi.rcMonitor
                            mon_rect = (r.left, r.top, r.right, r.bottom)
                except Exception:
                    mon_rect = None
                if mon_rect is None:
                    mon_rect = (0, 0, _u32.GetSystemMetrics(16),
                                _u32.GetSystemMetrics(17))
                ml, mt, mr, mb = mon_rect
                for hwnd in found:
                    r = wintypes.RECT()
                    if not _get_rect(hwnd, ctypes.byref(r)):
                        continue
                    inside = (r.left >= ml and r.top >= mt
                              and r.right <= mr and r.bottom <= mb)
                    if inside:
                        _set_pos(hwnd, HWND_TOP, 0, 0, 0, 0,
                                 SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE)
                    else:
                        w, h = r.right - r.left, r.bottom - r.top
                        x = min(max(r.left, ml), max(mr - w, ml))
                        y = min(max(r.top, mt), max(mb - h, mt))
                        _set_pos(hwnd, HWND_TOP, x, y, 0, 0,
                                 SWP_NOSIZE | SWP_NOACTIVATE)
            except Exception:
                pass

        for delay in (0.05, 0.2, 0.5, 1.0):
            t = threading.Timer(delay, _attempt)
            t.daemon = True
            t.start()
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root = root
        self._print_win = None
        self._print_app = None
        self._ppt_win = None
        self._welcome_win = None
        self._countdown_win = None
        self._music_win = None
        self._inbox_win = None
        self._inbox = None
        self._video_win = None
        self._video_player = None
        self._lottery_win = None
        self._lottery_app = None
        self._strip_win = None
        self._strip_screen = None
        self._qrcode_win = None
        self._downloader_win = None
        self._screenshot_win = None
        self._guard = None
        self._guard_win = None
        root.title(WINDOW_TITLE)
        width, height = 980, 680
        try:
            screen_w = root.winfo_screenwidth()
            screen_h = root.winfo_screenheight()
            left = max(0, (screen_w - width) // 2)
            top = max(0, (screen_h - height) // 2)
        except tk.TclError:
            left = top = 0
        root.geometry('%dx%d+%d+%d' % (width, height, left, top))
        root.resizable(False, False)
        root.configure(bg=WINDOW_BG)
        ui_theme.apply_theme(root)
        root.protocol('WM_DELETE_WINDOW', self._close_main_requested)
        # 先隐藏：卡片高度按字体真实行高计算、随系统 DPI 变化，
        # 等 _build 完成后按卡片实际需求补足窗口高度再显示，避免闪动
        root.withdraw()
        self._load_icon()
        self._build()
        self._start_focus_guard()
        self._poll_inbox()
        self._ensure_vlc_silent()
        root.after(20, self._show_root)

    def _ensure_vlc_silent(self):
        """启动即保证 VLC 可用：机器上没装就后台静默安装。

        优先跑程序目录 / soft 子目录里的 VLC 安装包（/S 静默），
        没有就用系统 winget（--silent）。全程无弹窗，结果写运行日志。
        """
        def work():
            try:
                if video_player.find_vlc() is not None:
                    return
                print('[VLC] 未检测到 VLC，开始静默安装…')
                ok, note = video_player.install_vlc()
                if ok:
                    video_player.reset_vlc_cache()
                print('[VLC] 自动安装结果: %s %s' % (ok, note))
            except Exception as exc:
                print('[VLC] 自动安装异常: %s' % exc)
        threading.Thread(target=work, daemon=True).start()

    def _show_root(self):
        """窗口高度自适应完成后再把主界面显示出来。"""
        try:
            self._fit_card_window()
            self.root.deiconify()
        except tk.TclError:
            try:
                self.root.deiconify()
            except tk.TclError:
                pass

    def _fit_card_window(self):
        """卡片总高（随 DPI 变化）超出可视区时，把窗口加高到一屏全显；
        屏幕实在不够高时维持原高，由现有滚动条兜底，绝不裁切文字。"""
        inner = getattr(self, '_card_inner', None)
        canvas = getattr(self, '_card_canvas', None)
        if inner is None or canvas is None:
            return
        try:
            self.root.update_idletasks()
            overflow = inner.winfo_reqheight() - canvas.winfo_height()
            if overflow <= 0:
                return
            w = self.root.winfo_width()
            h = self.root.winfo_height()
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            # 给任务栏/边缘留出余量
            new_h = min(h + overflow + 8, screen_h - 48)
            if new_h <= h:
                return
            left = max(0, (screen_w - w) // 2)
            top = max(0, (screen_h - new_h) // 2)
            self.root.geometry('%dx%d+%d+%d' % (w, new_h, left, top))
        except tk.TclError:
            pass

    def _poll_inbox(self):
        """主界面定时刷新“扫码收文件”卡片上的实时状态。"""
        card = getattr(self, '_inbox_card', None)
        inbox = self._inbox
        try:
            if card is not None and inbox is not None:
                files = inbox.snapshot()
                if inbox.running:
                    text = ('正在接收 · 已收到 %d 个文件\n'
                            '手机扫码上传，点开可查看/停止' % len(files))
                else:
                    text = self._inbox_desc
                card.set_description(text)
            self.root.after(1500, self._poll_inbox)
        except tk.TclError:
            pass

    def _load_icon(self):
        for base in (_BASE_DIR, os.path.dirname(
                os.path.abspath(__file__))):
            icon = os.path.join(base, '远程打印.ico')
            if os.path.exists(icon):
                try:
                    self.root.iconbitmap(icon)
                except Exception:
                    pass
                return

    def _build(self):
        p = ui_theme.PALETTE
        # 顶部渐变标题栏（右侧放大当前时间）
        topbar = tk.Frame(self.root, bg=p['primary'])
        band = ui_theme.GradientBand(
            topbar, title=WINDOW_TITLE, subtitle='会务工作，一站完成',
            height=60)
        band.pack(side=tk.LEFT, fill=tk.X, expand=True)
        # 现场防护状态（屏幕常亮 + 弹窗拦截），点击打开防护面板
        self._guard_label = tk.Label(
            topbar, text='● 现场防护中', font=ui_theme.font(9),
            fg='#dbeafe', bg=p['primary'], cursor='hand2',
            padx=10, pady=3)
        self._guard_label.pack(side=tk.RIGHT, padx=(0, 14))
        self._guard_label.bind('<Button-1>',
                               lambda _e: self._open_guard_panel())
        self._clock_label = tk.Label(topbar, font=ui_theme.font(24, 'bold'),
                                     fg='#ffffff', bg=p['primary'])
        self._clock_label.pack(side=tk.RIGHT, padx=(0, 24))
        topbar.pack(fill=tk.X)
        self._update_clock()

        body = tk.Frame(self.root, bg=p['bg'])
        body.pack(fill=tk.BOTH, expand=True)

        # 合并提醒条：离线只显示温馨提示；检测到外网时整条变红并入断网提示
        self._warn_ybg = ui_theme.mix(p['warning'], '#ffffff', 0.86)
        self._warn_rbg = ui_theme.mix(p['danger'], '#ffffff', 0.92)
        self._warn_bar = tk.Frame(body, bg=self._warn_ybg, padx=14, pady=4)
        self._warn_title = tk.Label(self._warn_bar, text='温馨提示：',
                                    font=ui_theme.font(9, 'bold'),
                                    fg=p['warning'], bg=self._warn_ybg)
        self._warn_title.pack(anchor='w')
        self._warn_body = tk.Label(
            self._warn_bar,
            text=('会场音响、话筒以及翻页笔等设备是否已准备好'
                  '（话筒、翻页笔等是否有电且准备妥当，请仔细检查）；'
                  '预祝会务圆满成功！'),
            font=ui_theme.font(9), fg=p['text'], bg=self._warn_ybg,
            justify=tk.LEFT, anchor='w')
        self._warn_body.pack(anchor='w')
        self._warn_bar.pack(fill=tk.X, padx=28, pady=(6, 2))
        self._start_net_check()

        # 卡片区在剩余空间里垂直居中
        holder = tk.Frame(body, bg=p['bg'])
        holder.pack(fill=tk.BOTH, expand=True)
        card_canvas = tk.Canvas(holder, bg=p['bg'], highlightthickness=0,
                                bd=0)
        card_scroll = ttk.Scrollbar(holder, orient=tk.VERTICAL,
                                    command=card_canvas.yview)
        card_canvas.configure(yscrollcommand=card_scroll.set)
        inner = tk.Frame(card_canvas, bg=p['bg'])
        self._card_canvas = card_canvas
        self._card_inner = inner
        inner_id = card_canvas.create_window((0, 0), window=inner,
                                             anchor='nw')
        card_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def _sync_card_width(_event=None):
            """把卡片容器宽度对齐到画布可视宽——否则 2 列会撑出屏幕。

            坑：画布的第一个 <Configure> 常常在画布还没有真实尺寸时就发
            过来（宽度是 1），之后窗口尺寸不再变化就不会再有事件，卡片
            区会一直停在“自然宽度”（2 列 × 440 撑到 900px 左右），
            看起来像“界面显示不正常、点一下才好”。
            所以这里要跳过无效宽度，并在窗口显示后主动补几次同步。
            """
            try:
                w = card_canvas.winfo_width()
                if w > 1:
                    card_canvas.itemconfigure(inner_id, width=w)
                    card_canvas.configure(scrollregion=card_canvas.bbox('all'))
            except tk.TclError:
                pass

        inner.bind('<Configure>', lambda e: card_canvas.configure(
            scrollregion=card_canvas.bbox('all')))
        card_canvas.bind('<Configure>', _sync_card_width)
        # 窗口真正显示出来之后再补几次（首个 Configure 可能来得太早）
        for delay in (30, 120, 400, 900):
            self.root.after(delay, _sync_card_width)
        inner.bind('<MouseWheel>', lambda e: card_canvas.yview_scroll(
            -int(e.delta / 120), 'units'))
        card_canvas.bind('<MouseWheel>', lambda e: card_canvas.yview_scroll(
            -int(e.delta / 120), 'units'))

        # 功能分三组展示：
        # ① 会务现场：现场环节用的投屏/播放/抽签，关闭窗口都会二次确认，
        #    放最前面，开场时一眼就能找到、也不怕误触关掉；
        # ② 文件收发：跨设备传文件；③ 制作工具：会前准备与素材处理。
        groups = (
            ('会务现场', '现场环节使用 · 关闭窗口前会先确认，防止误触', (
                ('PPT倒计时',
                 '扩展屏投屏 · 与PPT联动',
                 self._open_countdown, '时', ui_theme.ACCENTS['countdown']),
                ('会务音乐',
                 '本地音频暖场 / 循环',
                 self._open_music, '乐', ui_theme.ACCENTS['music']),
                ('会务视频',
                 '多视频依次全屏播放',
                 self._open_video, '影', ui_theme.ACCENTS['video']),
                ('现场抽签排序',
                 '名单洗牌 · 大屏投屏',
                 self._open_lottery, '签', ui_theme.ACCENTS['lottery']),
                ('长条屏合成投屏',
                 '两侧图 + 中间放PPT',
                 self._open_strip, '屏', ui_theme.ACCENTS['strip']),
                ('提词器',
                 '大屏词幕 · 自动滚动调速',
                 self._open_prompter, '词', ui_theme.ACCENTS['prompter']),
            )),
            ('文件收发', '跨设备收发文件', (
                ('微信传文件',
                 '打开微信网页版 · 扫码互传',
                 self._open_wechat_helper, '微', ui_theme.ACCENTS['wechat']),
                ('局域网互传',
                 '两台电脑局域网互传文件',
                 self._open_lan_share, '传', ui_theme.ACCENTS['lan']),
                ('扫码收文件',
                 '手机扫码即可上传文件',
                 self._open_file_inbox, '收', ui_theme.ACCENTS['inbox']),
            )),
            ('制作工具', '会前准备 · 素材处理', (
                ('会务PPT转换',
                 'PPT 转图片 / 转16:9全屏',
                 self._open_ppt, '演', ui_theme.ACCENTS['ppt']),
                ('生成欢迎语',
                 '模板改字或自定义背景',
                 self._open_welcome_choice, '语',
                 ui_theme.ACCENTS['welcome']),
                ('打印',
                 '座位牌与文件打印',
                 self._open_print, '印', ui_theme.ACCENTS['print']),
                ('二维码生成',
                 '链接/文本转二维码',
                 self._open_qrcode, '码', ui_theme.ACCENTS['qrcode']),
                ('快捷邮箱',
                 '一键直达常用邮箱登录页',
                 self._open_mail_menu, '邮', ui_theme.ACCENTS['mail']),
                ('软件下载',
                 '官方源 · 一键下载',
                 self._open_downloader, '载',
                 ui_theme.ACCENTS['downloader']),
                ('截图（仿 ShareX）',
                 '区域/全屏/窗口/滚动',
                 self._open_screenshot, '截',
                 ui_theme.ACCENTS['screenshot']),
                ('图片工具箱',
                 '裁剪缩放旋转·加水印压缩',
                 self._open_image_tools, '图', ui_theme.ACCENTS['photo']),
                ('音视频剪辑',
                 '剪切拼接·提音频压缩',
                 self._open_av_tools, '剪', ui_theme.ACCENTS['video']),
            )),
        )
        # 横条式小卡片：左边图标 + 右边标题说明，三列排布，组间有标题
        columns = 3
        grid_row = 0
        for gi, (group_title, group_hint, group_items) in enumerate(groups):
            hdr = tk.Frame(inner, bg=p['bg'])
            hdr.grid(row=grid_row, column=0, columnspan=columns,
                     sticky='ew', padx=7,
                     pady=((2 if gi == 0 else 14), 3))
            tk.Frame(hdr, bg=p['primary'], width=4,
                     height=15).pack(side=tk.LEFT, padx=(2, 8))
            tk.Label(hdr, text=group_title,
                     font=ui_theme.font(10, 'bold'),
                     fg=p['text'], bg=p['bg']).pack(side=tk.LEFT)
            tk.Label(hdr, text=group_hint,
                     font=ui_theme.font(8),
                     fg=p['faint'], bg=p['bg']).pack(side=tk.LEFT,
                                                      padx=(10, 0))
            grid_row += 1
            for index, (title, desc, cmd, icon, accent) in enumerate(
                    group_items):
                card = ui_theme.ModernCard(inner, title, desc, command=cmd,
                                           icon=icon, accent=accent,
                                           layout='row')
                card.grid(row=grid_row + index // columns,
                          column=index % columns,
                          sticky='nsew', padx=7, pady=(3, 9))
                card.bind('<MouseWheel>', lambda e: card_canvas.yview_scroll(
                    -int(e.delta / 120), 'units'))
                if title == '扫码收文件':
                    self._inbox_card = card
                    self._inbox_desc = desc
            grid_row += (len(group_items) + columns - 1) // columns

        for col in range(columns):
            inner.grid_columnconfigure(col, weight=1, uniform='card')

        # 卡片下方留出底部边距，最后一行不至于贴着窗口底边
        tk.Frame(inner, bg=p['bg'], height=14).grid(
            row=grid_row, column=0, columnspan=columns, sticky='ew')

        # 内容多到放不下时才显示右侧滚动条
        def _sync_scrollbar(_event=None):
            try:
                need = inner.winfo_reqheight() > card_canvas.winfo_height() + 2
                if need and not card_scroll.winfo_ismapped():
                    card_scroll.pack(side=tk.RIGHT, fill=tk.Y)
                elif not need and card_scroll.winfo_ismapped():
                    card_scroll.pack_forget()
            except tk.TclError:
                pass

        inner.bind('<Configure>', _sync_scrollbar, add='+')
        card_canvas.bind('<Configure>', _sync_scrollbar, add='+')

        # 底部说明
        tk.Label(body,
                 text='各功能都在主界面内打开子窗口，关闭主界面会一并关闭',
                 font=ui_theme.font(8), fg=p['faint'],
                 bg=p['bg']).pack(side=tk.BOTTOM, pady=(0, 6))

    def _update_clock(self):
        """标题栏右侧当前时间，每秒刷新。"""
        try:
            self._clock_label.config(text=time.strftime('%H:%M:%S'))
        except Exception:
            return
        try:
            self.root.after(1000, self._update_clock)
        except Exception:
            pass

    def _start_net_check(self):
        """后台探测能否访问外网，不阻塞主界面。"""
        threading.Thread(target=self._net_worker, daemon=True).start()

    def _net_worker(self):
        ok = False
        hosts = [('114.114.114.114', 53), ('223.5.5.5', 53), ('8.8.8.8', 53)]
        for host in hosts:
            try:
                with socket.create_connection(host, timeout=2.5):
                    ok = True
                    break
            except Exception:
                continue
        if ok:
            try:
                self.root.after(0, self._net_online)
            except Exception:
                pass

    def _net_online(self):
        """在线：整条提醒条变红，并入断网提示。"""
        try:
            p = ui_theme.PALETTE
            rb = self._warn_rbg
            self._warn_bar.config(bg=rb)
            self._warn_title.config(text='提醒：', fg=p['danger'], bg=rb)
            self._warn_body.config(
                text=('会务期间请断网，关闭一切屏保程序，防止广告弹窗。'
                      '会场音响、话筒以及翻页笔等设备请确认有电并准备妥当；'
                      '预祝会务圆满成功！'),
                fg=p['danger'], bg=rb)
        except Exception:
            pass

    def _open_ppt(self):
        """打开 PPT 转换子窗口：选择任意 PPT，转图片或转 16:9。"""
        if self._activate_or_close_others('ppt'):
            return
        win = self._ppt_win
        if win is not None:
            try:
                alive = bool(win.winfo_exists())
            except tk.TclError:
                alive = False
            if alive:
                win.deiconify()
                win.lift()
                win.focus_force()
                return
            self._ppt_win = None

        win = tk.Toplevel(self.root)
        win.title('会务PPT转换')
        _center_window(win, 640, 480)
        win.resizable(False, False)
        ui_theme.style_window(win)
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW', self._close_ppt_win)
        self._ppt_win = win

        path_var = tk.StringVar()
        mode_var = tk.StringVar(value='images')
        out_var = tk.StringVar(value='')
        progress_var = tk.IntVar(value=0)
        status_var = tk.StringVar(
            value='选择 PPT 文件、转换方式和输出位置后，'
                  '点击“开始转换”。')

        tk.Label(win, text='PPT 文件',
                 font=('微软雅黑', 10)).pack(anchor='w', padx=20,
                                            pady=(18, 4))
        row = ttk.Frame(win)
        row.pack(fill=tk.X, padx=20)
        entry = ttk.Entry(row, textvariable=path_var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        browse_btn = ttk.Button(row, text='选择文件…', width=10,
                                style='Ghost.TButton',
                                command=lambda: _browse())
        browse_btn.pack(
            side=tk.LEFT, padx=(8, 0))

        # 转换方式
        mode_box = ttk.LabelFrame(win, text='转换方式', padding=8)
        mode_box.pack(fill=tk.X, padx=20, pady=(14, 6))
        radio_images = ttk.Radiobutton(
            mode_box, text='转换为图片（每页一张）',
            variable=mode_var, value='images',
            command=lambda: _mode_changed())
        radio_images.pack(anchor='w')
        radio_169 = ttk.Radiobutton(
            mode_box, text='生成 16:9 全屏 PPT'
                           '（非常规比例自动拉伸铺满）',
            variable=mode_var, value='169',
            command=lambda: _mode_changed())
        radio_169.pack(anchor='w', pady=(4, 0))

        # 输出位置
        out_frame = tk.Frame(win)
        out_frame.pack(fill=tk.X, padx=20, pady=(6, 0))
        out_label = tk.Label(out_frame, text='输出文件夹：',
                             font=('微软雅黑', 10))
        out_label.pack(side=tk.LEFT)
        out_entry = ttk.Entry(out_frame, textvariable=out_var,
                              state='readonly')
        out_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        out_btn = ttk.Button(out_frame, text='选择文件夹…', width=12,
                             style='Ghost.TButton',
                             command=lambda: _browse_output())
        out_btn.pack(side=tk.LEFT, padx=(8, 0))

        # 开始转换
        start_btn = ttk.Button(
            win, text='开始转换', style='Accent.TButton',
            command=lambda: _start())
        start_btn.pack(pady=(14, 8))

        # 进度条
        bar = ttk.Progressbar(win, orient=tk.HORIZONTAL,
                              maximum=100, variable=progress_var)
        bar.pack(fill=tk.X, padx=20)

        ttk.Label(win, textvariable=status_var,
                  foreground='#555f6b', font=('微软雅黑', 9),
                  wraplength=580, justify=tk.LEFT).pack(
            anchor='w', padx=20, pady=(10, 16))

        def _browse():
            current = path_var.get().strip()
            initial = os.path.dirname(current) if current else ''
            path = filedialog.askopenfilename(
                parent=win, title='选择要转换的 PPT',
                initialdir=initial,
                filetypes=[('PPT 文件', '*.ppt *.pptx'),
                           ('所有文件', '*.*')])
            if path:
                path_var.set(path)
                progress_var.set(0)
                _refresh_status()

        def _mode_changed():
            out_var.set('')
            progress_var.set(0)
            images_mode = mode_var.get() == 'images'
            out_label.configure(
                text='输出文件夹：' if images_mode else '输出文件：')
            out_btn.configure(
                text='选择文件夹…' if images_mode else '选择保存位置…')
            _refresh_status()

        def _refresh_status():
            status_var.set(
                '已选择：%s' % path_var.get().strip()
                if path_var.get().strip() else
                '请选择要转换的 .ppt / .pptx 文件。')

        def _browse_output():
            src = path_var.get().strip()
            if not src or not os.path.isfile(src):
                messagebox.showwarning(
                    '提示', '请先选择要转换的 PPT 文件。', parent=win)
                return
            if mode_var.get() == 'images':
                folder = filedialog.askdirectory(
                    parent=win, title='选择图片保存文件夹',
                    initialdir=os.path.dirname(src))
                if folder:
                    out_var.set(folder)
                    progress_var.set(0)
                    status_var.set('输出文件夹已选择，点击“开始转换”。')
            else:
                stem = os.path.splitext(os.path.basename(src))[0]
                path = filedialog.asksaveasfilename(
                    parent=win, title='保存 16:9 PPT',
                    initialdir=os.path.dirname(src),
                    initialfile=stem + '_16x9.pptx',
                    defaultextension='.pptx',
                    filetypes=[('PPTX 文件', '*.pptx'),
                               ('PPT 文件', '*.ppt')])
                if path:
                    out_var.set(path)
                    progress_var.set(0)
                    status_var.set('输出文件已选择，点击“开始转换”。')

        def _set_busy(busy):
            state = 'disabled' if busy else 'normal'
            browse_btn.configure(state=state)
            radio_images.configure(state=state)
            radio_169.configure(state=state)
            out_btn.configure(state=state)
            start_btn.configure(state=state)
            win.configure(cursor='watch' if busy else '')

        def _progress(current, total, text):
            progress_var.set(
                100 if total <= 0 else int(current * 100 // total))
            status_var.set(text)
            win.update_idletasks()

        def _start():
            src = path_var.get().strip()
            if not src or not os.path.isfile(src):
                messagebox.showwarning(
                    '提示', '请先选择要转换的 .ppt / .pptx 文件。',
                    parent=win)
                return
            ext = os.path.splitext(src)[1].lower()
            if ext not in ('.ppt', '.pptx'):
                messagebox.showwarning(
                    '提示', '仅支持 .ppt / .pptx 文件。', parent=win)
                return
            out = out_var.get().strip()
            if not out:
                messagebox.showwarning(
                    '提示',
                    '请先选择输出文件夹或输出文件。', parent=win)
                return

            _set_busy(True)
            progress_var.set(0)
            try:
                import ppt_converter
            except Exception as e:
                messagebox.showerror(
                    '错误', '加载转换模块失败：\n%s' % e, parent=win)
                _set_busy(False)
                return

            try:
                if mode_var.get() == 'images':
                    out_dir, count, _ = \
                        ppt_converter.convert_ppt_to_images(
                            src, out, on_progress=_progress)
                    progress_var.set(100)
                    status_var.set('完成：已导出 %d 张图片。' % count)
                    messagebox.showinfo(
                        '完成', '已导出 %d 张图片（每页一张）：\n%s'
                        % (count, out_dir), parent=win)
                else:
                    out_path, count = \
                        ppt_converter.convert_ppt_to_16x9(
                            src, out, on_progress=_progress)
                    progress_var.set(100)
                    status_var.set(
                        '完成：已生成 16:9 PPT（%d 页）。' % count)
                    messagebox.showinfo(
                        '完成', '已生成拉伸铺满 16:9 的全屏 PPT'
                        '（共 %d 页）：\n%s' % (count, out_path),
                        parent=win)
            except ppt_converter.PptConvertError as e:
                messagebox.showerror('转换失败', str(e), parent=win)
            except Exception as e:
                messagebox.showerror('转换失败', '发生未知错误：\n%s' % e,
                                     parent=win)
            finally:
                _set_busy(False)

    def _close_ppt_win(self):
        """用户单独关闭 PPT 转换子窗口。"""
        win = self._ppt_win
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._ppt_win = None

    def _open_prompter(self):
        """打开内置提词器编辑窗（同一时间只保留一个）。"""
        c = getattr(self, '_prompter', None)
        if c is not None and c.win is not None:
            try:
                alive = bool(c.win.winfo_exists())
            except tk.TclError:
                alive = False
            if alive:
                c.win.deiconify()
                c.win.lift()
                c.win.focus_force()
                return
        self._prompter = prompter.PrompterApp(self.root)

    def _close_prompter_win(self):
        """关闭提词器（编辑窗 + 若在演示则一并退出）。"""
        c = getattr(self, '_prompter', None)
        if c is not None:
            try:
                c.close(force=True)
            except Exception:
                pass
            self._prompter = None

    def _open_lan_share(self):
        """打开局域网互传窗口（同一时间只保留一个）。"""
        c = getattr(self, '_lan_share', None)
        if c is not None and c.win is not None:
            try:
                alive = bool(c.win.winfo_exists())
            except tk.TclError:
                alive = False
            if alive:
                c.win.deiconify()
                c.win.lift()
                c.win.focus_force()
                return
        self._lan_share = lan_share.LanShareApp(self.root)

    def _close_lan_share_win(self):
        """关闭局域网互传窗口（含接收服务 / 发现线程）。"""
        c = getattr(self, '_lan_share', None)
        if c is not None:
            try:
                c.close(force=True)
            except Exception:
                pass
            self._lan_share = None

    def _open_image_tools(self):
        """打开图片工具箱窗口（同一时间只保留一个）。"""
        c = getattr(self, '_image_tools', None)
        if c is not None and c.win is not None:
            try:
                alive = bool(c.win.winfo_exists())
            except tk.TclError:
                alive = False
            if alive:
                c.win.deiconify()
                c.win.lift()
                c.win.focus_force()
                return
        self._image_tools = image_tools.ImageToolApp(self.root)

    def _close_image_tools_win(self):
        """关闭图片工具箱窗口。"""
        c = getattr(self, '_image_tools', None)
        if c is not None:
            try:
                c.close(force=True)
            except Exception:
                pass
            self._image_tools = None

    def _open_av_tools(self):
        """打开音视频剪辑窗口（同一时间只保留一个）。"""
        c = getattr(self, '_av_tools', None)
        if c is not None and c.win is not None:
            try:
                alive = bool(c.win.winfo_exists())
            except tk.TclError:
                alive = False
            if alive:
                c.win.deiconify()
                c.win.lift()
                c.win.focus_force()
                return
        self._av_tools = av_tools.AvToolApp(self.root)

    def _close_av_tools_win(self):
        """关闭音视频剪辑窗口。"""
        c = getattr(self, '_av_tools', None)
        if c is not None:
            try:
                c.close(force=True)
            except Exception:
                pass
            self._av_tools = None

    def _open_welcome_legacy(self):
        """旧版内置模板界面（已停用，保留参考）。"""
        win = self._welcome_win
        if win is not None:
            try:
                alive = bool(win.winfo_exists())
            except tk.TclError:
                alive = False
            if alive:
                win.deiconify()
                win.lift()
                win.focus_force()
                return
            self._welcome_win = None

        try:
            import ppt_converter
        except Exception as e:
            messagebox.showerror(
                '错误', '加载转换模块失败：\n%s' % e, parent=self.root)
            return
        bg_items = ppt_converter.list_welcome_backgrounds()

        win = tk.Toplevel(self.root)
        win.title('生成欢迎语')
        _center_window(win, 760, 640)
        win.resizable(False, False)
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW', self._close_welcome_win)
        self._welcome_win = win

        template_var = tk.StringVar(value='subject')
        bg_kind_var = tk.StringVar(value='pure')
        color_name_var = tk.StringVar(value='中国红')
        pattern_name_var = tk.StringVar(
            value=bg_items[0]['name'] if bg_items else '')
        auto_var = tk.BooleanVar(value=True)
        size_var = tk.IntVar(value=24)
        status_var = tk.StringVar(
            value='选择模板、输入中间内容后点“生成 PPT”。')

        # 模板选择
        tpl_row = tk.Frame(win)
        tpl_row.pack(anchor='w', padx=20, pady=(14, 4))
        ttk.Label(tpl_row, text='模板：',
                  font=('微软雅黑', 10)).pack(side=tk.LEFT)
        radio_subject = ttk.Radiobutton(
            tpl_row, text='有主体（模板1）', value='subject',
            variable=template_var,
            command=lambda: _toggle_template())
        radio_subject.pack(side=tk.LEFT)
        radio_plain = ttk.Radiobutton(
            tpl_row, text='无主体（模板2）', value='plain',
            variable=template_var,
            command=lambda: _toggle_template())
        radio_plain.pack(side=tk.LEFT, padx=(14, 0))

        # 背景设置（仅模板2）：纯色可选颜色 / 花纹图库
        bg_frame = ttk.LabelFrame(
            win, text='背景（模板2可用）', padding=6)
        bg_frame.pack(fill=tk.X, padx=20, pady=(6, 2))

        type_row = tk.Frame(bg_frame)
        type_row.pack(fill=tk.X)
        ttk.Label(type_row, text='方式：').pack(side=tk.LEFT)
        radio_pure = ttk.Radiobutton(
            type_row, text='纯色', value='pure',
            variable=bg_kind_var,
            command=lambda: _sync_bg_kind())
        radio_pure.pack(side=tk.LEFT)
        radio_pattern = ttk.Radiobutton(
            type_row, text='花纹背景', value='pattern',
            variable=bg_kind_var,
            command=lambda: _sync_bg_kind())
        radio_pattern.pack(side=tk.LEFT, padx=(14, 0))

        # 纯色行：颜色下拉 + 预览 + 自定义
        pure_row = tk.Frame(bg_frame)
        pure_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(pure_row, text='颜色：').pack(side=tk.LEFT)
        color_combo = ttk.Combobox(
            pure_row, textvariable=color_name_var, state='readonly',
            width=10, values=[name for name, _ in PURE_COLORS])
        color_combo.pack(side=tk.LEFT, padx=(4, 8))
        color_combo.bind(
            '<<ComboboxSelected>>',
            lambda e: _refresh_pure_preview())
        pure_preview = tk.Label(
            pure_row, width=8, height=2,
            bg=PURE_COLORS[0][1],
            font=('微软雅黑', 9))
        pure_preview.pack(side=tk.LEFT, padx=(0, 10))
        custom_color_btn = ttk.Button(
            pure_row, text='自定义…',
            command=lambda: _pick_custom_color())
        custom_color_btn.pack(side=tk.LEFT)

        # 花纹行：图库下拉 + 缩略图
        pattern_row = tk.Frame(bg_frame)
        pattern_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(pattern_row, text='花纹：').pack(side=tk.LEFT)
        pattern_combo = ttk.Combobox(
            pattern_row, textvariable=pattern_name_var,
            state='readonly', width=12,
            values=[item['name'] for item in bg_items])
        pattern_combo.pack(side=tk.LEFT, padx=(4, 8))
        pattern_combo.bind(
            '<<ComboboxSelected>>',
            lambda e: _refresh_pattern_preview())
        pattern_preview = tk.Label(
            pattern_row, width=12, height=2, bg='#eeeeee')
        pattern_preview.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(
            pattern_row, text='图库共 %d 种' % len(bg_items),
            foreground='#7a8ca0', font=('微软雅黑', 9)).pack(
            side=tk.LEFT)
        pattern_preview_photos = []

        content_label = tk.Label(
            win, text='中间内容（会替换模板中间文字，换行会保留）',
            font=('微软雅黑', 10))
        content_label.pack(anchor='w', padx=20, pady=(8, 4))
        content_text = tk.Text(
            win, height=6, wrap=tk.WORD, font=('微软雅黑', 14),
            relief=tk.GROOVE, bd=1, padx=10, pady=8,
            undo=True)
        content_text.pack(fill=tk.X, padx=20)

        # 字号设置
        size_row = tk.Frame(win)
        size_row.pack(anchor='w', padx=20, pady=(8, 0))
        auto_check = ttk.Checkbutton(
            size_row, text='自动调整字号（按内容长短自动缩小）',
            variable=auto_var, command=lambda: _toggle_auto())
        auto_check.pack(side=tk.LEFT)
        ttk.Label(size_row, text='字号(pt)：').pack(
            side=tk.LEFT, padx=(16, 4))
        size_spin = ttk.Spinbox(
            size_row, from_=8, to=120, textvariable=size_var,
            width=6, state='disabled')
        size_spin.pack(side=tk.LEFT)

        # 生成按钮
        gen_btn = ttk.Button(
            win, text='生成 PPT', command=lambda: _generate())
        gen_btn.pack(pady=(10, 6))

        ttk.Label(win, textvariable=status_var,
                  foreground='#7a8ca0', font=('微软雅黑', 9),
                  wraplength=700, justify=tk.LEFT).pack(
            anchor='w', padx=20, pady=(0, 12))

        def _bg_enabled():
            return template_var.get() == 'plain'

        def _sync_bg_kind():
            enabled = _bg_enabled()
            pure = bg_kind_var.get() == 'pure'
            radio_pure.configure(
                state='normal' if enabled else 'disabled')
            radio_pattern.configure(
                state='normal' if enabled else 'disabled')
            color_combo.configure(
                state='readonly' if enabled and pure else 'disabled')
            custom_color_btn.configure(
                state='normal' if enabled and pure else 'disabled')
            pattern_combo.configure(
                state='readonly' if enabled and not pure
                else 'disabled')

        def _toggle_template():
            if template_var.get() == 'plain':
                if not bg_frame.winfo_ismapped():
                    bg_frame.pack(
                        fill=tk.X, padx=20, pady=(6, 2),
                        before=content_label)
                _sync_bg_kind()
                status_var.set(
                    '模板2可设纯色或花纹背景，'
                    '输入内容后点“生成 PPT”。')
            else:
                bg_frame.pack_forget()
                status_var.set(
                    '模板1有固定主体，只替换中间文字，'
                    '点“生成 PPT”。')

        def _refresh_pure_preview():
            name = color_name_var.get()
            color = _color_hex(name)
            pure_preview.configure(bg=color, image='', text='')

        def _refresh_pattern_preview():
            name = pattern_name_var.get()
            pattern_preview_photos.clear()
            item = next((x for x in bg_items
                         if x['name'] == name), None)
            if item is None:
                pattern_preview.configure(bg='#eeeeee', text='')
                return
            try:
                from PIL import Image, ImageTk
                with Image.open(item['thumb']) as raw:
                    thumb = raw.convert('RGB')
                thumb = thumb.resize((150, 84), Image.LANCZOS)
                photo = ImageTk.PhotoImage(thumb)
                thumb.close()
                pattern_preview_photos.append(photo)
                pattern_preview.configure(image=photo, text='')
            except Exception as e:
                pattern_preview.configure(
                    text='预览失败', image='', bg='#999999')

        def _color_hex(name):
            for n, color in PURE_COLORS:
                if n == name:
                    return color
            if len(name) == 7 and name.startswith('#'):
                try:
                    int(name[1:], 16)
                    return name
                except ValueError:
                    pass
            return PURE_COLORS[0][1]

        def _pick_custom_color():
            from tkinter import colorchooser
            current = _color_hex(color_name_var.get())
            rgb, hex_color = colorchooser.askcolor(
                color=current, title='选择纯色背景', parent=win)
            if hex_color is None:
                return
            existing = next(
                (n for n, c in PURE_COLORS if c.lower() ==
                 hex_color.lower()), None)
            if existing is None:
                PURE_COLORS.append((hex_color, hex_color))
                color_combo.configure(
                    values=[n for n, _ in PURE_COLORS])
            color_name_var.set(existing or hex_color)
            _refresh_pure_preview()

        def _toggle_auto():
            size_spin.configure(
                state='disabled' if auto_var.get() else 'normal')

        def _set_busy(busy):
            radio_subject.configure(
                state='disabled' if busy else 'normal')
            radio_plain.configure(
                state='disabled' if busy else 'normal')
            gen_btn.configure(
                state='disabled' if busy else 'normal')
            content_text.configure(
                state='disabled' if busy else 'normal')
            auto_check.configure(
                state='disabled' if busy else 'normal')
            if busy:
                size_spin.configure(state='disabled')
                radio_pure.configure(state='disabled')
                radio_pattern.configure(state='disabled')
                color_combo.configure(state='disabled')
                custom_color_btn.configure(state='disabled')
                pattern_combo.configure(state='disabled')
            else:
                size_spin.configure(
                    state='normal' if not auto_var.get()
                    else 'disabled')
                _sync_bg_kind()
            win.configure(cursor='watch' if busy else '')

        def _generate():
            content = content_text.get('1.0', 'end-1c')
            if not content.strip():
                messagebox.showwarning(
                    '提示', '请输入欢迎语内容。', parent=win)
                return

            try:
                import ppt_converter
            except Exception as e:
                messagebox.showerror(
                    '错误', '加载转换模块失败：\n%s' % e, parent=win)
                return

            default_name = _welcome_default_name(content)
            save_path = filedialog.asksaveasfilename(
                parent=win, title='保存欢迎语 PPT',
                initialfile=default_name,
                defaultextension='.pptx',
                filetypes=[('PPTX 文件', '*.pptx'),
                           ('PPT 文件', '*.ppt')])
            if not save_path:
                return

            font_size = None if auto_var.get() else size_var.get()
            template_key = template_var.get()
            background_path = None
            background_color = None
            if template_key == 'plain':
                if bg_kind_var.get() == 'pattern':
                    bg_name = pattern_name_var.get()
                    item = next(
                        (x for x in bg_items if x['name'] == bg_name),
                        None)
                    if item:
                        background_path = item['image']
                else:
                    background_color = _color_hex(color_name_var.get())
            _set_busy(True)
            try:
                status_var.set('正在生成欢迎语 PPT…')
                win.update_idletasks()
                out_path, used_size = \
                    ppt_converter.generate_welcome_ppt(
                        content, save_path, font_size=font_size,
                        template_key=template_key,
                        background_path=background_path,
                        background_color=background_color)
                tpl_name = '有主体' if template_key == 'subject' \
                    else '无主体'
                bg_name = ''
                if background_path:
                    bg_name = '，背景：' + pattern_name_var.get()
                elif background_color:
                    bg_name = '，纯色：' + color_name_var.get()
                if used_size:
                    status_var.set(
                        '已生成（模板：%s%s），字号 %d 磅。' %
                        (tpl_name, bg_name, used_size))
                else:
                    status_var.set(
                        '已生成（模板：%s%s）：%s' %
                        (tpl_name, bg_name, out_path))
            except ppt_converter.PptConvertError as e:
                messagebox.showerror('生成失败', str(e), parent=win)
                return
            except Exception as e:
                messagebox.showerror('生成失败', '发生未知错误：\n%s' % e,
                                     parent=win)
                return
            finally:
                _set_busy(False)

            if messagebox.askyesno(
                    '生成完成', '欢迎语 PPT 已生成：\n%s\n\n'
                    '是否立即打开查看？' % out_path, parent=win):
                try:
                    os.startfile(out_path)
                except Exception as e:
                    messagebox.showerror(
                        '提示', '打开文件失败：\n%s' % e, parent=win)

        _toggle_template()
        _refresh_pure_preview()
        _refresh_pattern_preview()

    def _close_welcome_win(self):
        """用户单独关闭欢迎语子窗口。"""
        win = self._welcome_win
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._welcome_win = None

    def _open_print(self):
        """打开远程打印子窗口；已打开时只聚焦，不重复创建。"""
        if self._activate_or_close_others('print'):
            return
        win = self._print_win
        if win is not None:
            try:
                alive = bool(win.winfo_exists())
            except tk.TclError:
                alive = False
            if alive:
                win.deiconify()
                win.lift()
                win.focus_force()
                return
            # 子窗口已被销毁，清掉旧引用
            self._print_win = None
            self._print_app = None

        try:
            import seat_card
        except Exception as e:
            messagebox.showerror(
                '加载失败', '远程打印功能加载失败：\n%s' % e,
                parent=self.root)
            return

        win = tk.Toplevel(self.root)
        # 先藏起来：内容建完、摆到中间再显示，避免窗口先在左上角闪一下
        win.withdraw()
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW', self._close_print_win)
        try:
            self._print_app = seat_card.App(win)
        except Exception as e:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._print_win = None
            self._print_app = None
            messagebox.showerror(
                '加载失败', '远程打印功能打开失败：\n%s' % e,
                parent=self.root)
            return
        self._print_win = win
        # 打印窗口由子程序自己定尺寸，建好后摆到屏幕中间
        try:
            win.update_idletasks()
            _center_window(win, max(560, win.winfo_reqwidth()),
                           max(420, win.winfo_reqheight()))
        except tk.TclError:
            pass
        win.deiconify()
        win.lift()
        win.focus_force()

    def _close_print_win(self):
        """用户单独关闭打印子窗口。"""
        if self._print_win is not None:
            try:
                self._print_win.destroy()
            except tk.TclError:
                pass
            self._print_win = None
            self._print_app = None

    def _activate_or_close_others(self, keep):
        """只允许一个子窗口：同名窗口存在就激活，否则关闭其他子窗口。"""
        targets = {
            'ppt': self._ppt_win,
            'welcome': self._welcome_win,
            'print': self._print_win,
            'countdown': self._countdown_win,
            'inbox': self._inbox_win,
            'video': self._video_win,
            'lottery': self._lottery_win,
            'strip': self._strip_win,
            'qrcode': self._qrcode_win,
            'downloader': self._downloader_win,
            'screenshot': self._screenshot_win,
            'lan': self._lan_share.win if getattr(self, '_lan_share', None) else None,
            'img': self._image_tools.win if getattr(self, '_image_tools', None) else None,
            'av': self._av_tools.win if getattr(self, '_av_tools', None) else None,
        }
        attrs = {
            'ppt': ('_ppt_win', None),
            'welcome': ('_welcome_win', None),
            'print': ('_print_win', '_print_app'),
            'countdown': ('_countdown_win', None),
            'inbox': ('_inbox_win', None),
            'video': ('_video_win', None),
            'lottery': ('_lottery_win', '_lottery_app'),
            'strip': ('_strip_win', None),
            'qrcode': ('_qrcode_win', None),
            'downloader': ('_downloader_win', None),
            'screenshot': ('_screenshot_win', None),
            'lan': ('_lan_share_win', '_lan_share'),
            'img': ('_image_tools_win', '_image_tools'),
            'av': ('_av_tools_win', '_av_tools'),
        }
        target = targets[keep]
        if target is not None:
            alive = self._is_subtarget_alive(keep, target)
            if alive:
                # 倒计时处于"隐藏到托盘"状态时，从主程序点按钮也要能
                # 把它恢复出来——同时把托盘图标收掉，避免两个入口。
                if keep == 'countdown':
                    state = getattr(target, '_countdown_state', None)
                    if state is not None and state.get('_hidden'):
                        self._restore_countdown_from_tray(target, state)
                    else:
                        target.deiconify()
                        target.lift()
                        target.focus_force()
                else:
                    target.deiconify()
                    target.lift()
                    target.focus_force()
                return True
            attr, app_attr = attrs[keep]
            setattr(self, attr, None)
            if app_attr:
                setattr(self, app_attr, None)

        if keep != 'welcome':
            self._close_welcome_win()
        if keep != 'ppt':
            self._close_ppt_win()
        if keep != 'print':
            self._close_print_win()
        if keep != 'countdown':
            self._close_countdown_win()
        if keep != 'inbox':
            # 打开其它功能时只关窗口，收件服务继续跑（手机还能继续上传）
            self._close_inbox_win(keep_running=True)
        if keep != 'video':
            self._close_video_win()
        if keep != 'lottery':
            self._close_lottery_win()
        if keep != 'strip':
            self._close_strip_win()
        if keep != 'downloader':
            self._close_downloader_win()
        if keep != 'screenshot':
            self._close_screenshot_win()
        if keep != 'lan':
            self._close_lan_share_win()
        if keep != 'img':
            self._close_image_tools_win()
        if keep != 'av':
            self._close_av_tools_win()
        return False

    def _is_subtarget_alive(self, key, target):
        """各子功能的"还活着"判定（全部是 Toplevel，winfo_exists 即可）。"""
        try:
            return bool(target.winfo_exists())
        except tk.TclError:
            return False
        except Exception:
            return False

    def _open_welcome_template_library_legacy(self):
        """旧版整页模板库方案（已停用，保留参考）。"""
        if self._activate_or_close_others('welcome'):
            return
        try:
            import ppt_converter
        except Exception as e:
            messagebox.showerror(
                '错误', '加载转换模块失败：\n%s' % e, parent=self.root)
            return

        win = tk.Toplevel(self.root)
        win.title('生成欢迎语（模板库）')
        _center_window(win, 800, 720)
        win.resizable(True, True)
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW', self._close_welcome_win)
        self._welcome_win = win

        state = {'items': []}
        auto_var = tk.BooleanVar(value=True)
        size_var = tk.IntVar(value=24)
        status_var = tk.StringVar(
            value='请先导入你的欢迎语模板，再输入内容生成。')

        # 模板库
        tpl_frame = ttk.LabelFrame(
            win, text='当前模板', padding=6)
        tpl_frame.pack(fill=tk.X, padx=20, pady=(16, 4))
        row1 = tk.Frame(tpl_frame)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text='当前模板：').pack(side=tk.LEFT)
        tpl_combo = ttk.Combobox(
            row1, state='readonly', width=32)
        tpl_combo.pack(side=tk.LEFT, padx=(4, 10))
        tpl_combo.bind('<<ComboboxSelected>>',
                       lambda e: (_update_tpl_hint(),
                                  _update_tpl_preview()))
        pick_btn = ttk.Button(
            row1, text='指定文字框…',
            command=lambda: _repick_template())
        pick_btn.pack(side=tk.LEFT)
        tpl_hint_var = tk.StringVar(value='')
        ttk.Label(
            tpl_frame, textvariable=tpl_hint_var,
            foreground='#7a8ca0', font=('微软雅黑', 9)).pack(
            anchor='w', pady=(4, 0))
        ttk.Label(
            tpl_frame,
            text='模板自动加载，无需导入或删除。',
            foreground='#7a8ca0', font=('微软雅黑', 9)).pack(
            anchor='w', pady=(2, 0))

        # 模板预览：选择模板后自动显示其首页缩略图（无边框浅底，避免黑线）
        preview_frame = tk.Frame(win, bg='#f5f6f8',
                                 bd=0, highlightthickness=0)
        preview_frame.pack(fill=tk.X, padx=20, pady=(8, 4))
        tk.Label(preview_frame, text='模板预览', bg='#f5f6f8',
                 fg='#33466e', font=('微软雅黑', 9)).pack(
            anchor='w', padx=4, pady=(4, 0))
        preview_text = tk.StringVar(value='选择模板后自动显示首页预览')
        _preview_photo = [None]
        preview_label = tk.Label(
            preview_frame, textvariable=preview_text,
            anchor=tk.CENTER, foreground='#7a8ca0',
            background='#f5f6f8', bd=0, highlightthickness=0)
        preview_label.pack(fill=tk.X, padx=4, pady=(2, 4))

        # 内容与字号
        tk.Label(
            win, text='要替换的文字内容（换行会保留）',
            font=('微软雅黑', 10)).pack(
            anchor='w', padx=20, pady=(10, 4))
        content_text = tk.Text(
            win, height=4, wrap=tk.WORD, font=('微软雅黑', 14),
            relief=tk.GROOVE, bd=1, padx=10, pady=8, undo=True)
        content_text.pack(fill=tk.X, padx=20)

        size_row = tk.Frame(win)
        size_row.pack(anchor='w', padx=20, pady=(10, 0))
        auto_check = ttk.Checkbutton(
            size_row, text='自动调整字号（过长自动缩小）',
            variable=auto_var, command=lambda: _toggle_auto())
        auto_check.pack(side=tk.LEFT)
        ttk.Label(size_row, text='字号(pt)：').pack(
            side=tk.LEFT, padx=(16, 4))
        size_spin = ttk.Spinbox(
            size_row, from_=8, to=200, textvariable=size_var,
            width=6, state='disabled')
        size_spin.pack(side=tk.LEFT)

        act_row = tk.Frame(win)
        act_row.pack(pady=(12, 6))
        refresh_btn = ttk.Button(
            act_row, text='刷新预览',
            command=lambda: _preview_result())
        refresh_btn.pack(side=tk.LEFT, padx=6)
        gen_btn = ttk.Button(
            act_row, text='保存并生成 PPT', style='Accent.TButton',
            command=lambda: _generate())
        gen_btn.pack(side=tk.LEFT, padx=6)

        ttk.Label(win, textvariable=status_var,
                  foreground='#7a8ca0', font=('微软雅黑', 9),
                  wraplength=740, justify=tk.LEFT).pack(
            anchor='w', padx=20, pady=(0, 12))

        def _selected_template():
            name = tpl_combo.get()
            for item in state['items']:
                if item['name'] == name:
                    return item
            return None

        def _reload(select_id=None):
            state['items'] = ppt_converter.list_welcome_templates()
            names = [x['name'] for x in state['items']]
            tpl_combo.configure(values=names)
            if select_id:
                hit = next((x for x in state['items']
                            if x['id'] == select_id), None)
                if hit:
                    tpl_combo.set(hit['name'])
                    return
            if names:
                tpl_combo.set(names[0])
            else:
                tpl_combo.set('')
                status_var.set('模板库为空：点“导入模板…”添加你的模板。')

        def _choose_shape(shapes):
            if len(shapes) == 1:
                return shapes[0]
            pick = tk.Toplevel(win)
            pick.title('选择要替换的文字框')
            _center_window(pick, 680, 440)
            pick.transient(win)
            pick.resizable(False, False)
            tk.Label(
                pick, text='模板里有多处文字，请选择“生成欢迎语时”'
                           '要替换的那个：', font=('微软雅黑', 10)
            ).pack(anchor='w', padx=18, pady=(16, 6))
            lb = tk.Listbox(pick, font=('微软雅黑', 10),
                            activestyle='dotbox')
            sb = ttk.Scrollbar(pick, orient=tk.VERTICAL,
                               command=lb.yview)
            lb.configure(yscrollcommand=sb.set)
            for s in shapes:
                preview = (s['text'] or '（空文本）').replace(
                    '\n', ' ').replace('\r', ' ')
                lb.insert(
                    tk.END,
                    '第%d页 · %s：%s' % (s['slide'], s['name'],
                                         preview))
            lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                    padx=(18, 0), pady=(6, 12))
            sb.pack(side=tk.LEFT, fill=tk.Y, pady=(6, 12))
            if shapes:
                lb.selection_set(0)
            result = []

            def ok():
                sel = lb.curselection()
                if not sel:
                    messagebox.showwarning(
                        '提示', '请先选择一个文字框。', parent=pick)
                    return
                result.append(shapes[sel[0]])
                pick.destroy()

            btns = tk.Frame(pick)
            btns.pack(side=tk.BOTTOM, anchor='e', padx=18,
                      pady=(0, 14))
            ttk.Button(btns, text='选择这个',
                       command=ok).pack(side=tk.LEFT)
            ttk.Button(btns, text='取消',
                       command=pick.destroy).pack(
                side=tk.LEFT, padx=(8, 0))
            pick.grab_set()
            win.wait_window(pick)
            return result[0] if result else None

        def _update_tpl_hint():
            """在提示行显示当前模板将替换哪个文字框。"""
            item = _selected_template()
            if item is None:
                tpl_hint_var.set('')
                return
            try:
                shapes = ppt_converter.template_text_shapes(item['path'])
                hit = next((s for s in shapes
                            if s.get('index') == item.get('shape_index')),
                           None)
                txt = (hit or {}).get('text') or ''
                tpl_hint_var.set('将替换文字框：%s' %
                                 (txt[:40] if txt else '（空文字框）'))
            except Exception:
                tpl_hint_var.set('')

        def _update_tpl_preview():
            """把当前模板的首页导出成缩略图显示在预览区。"""
            item = _selected_template()
            if item is None:
                preview_text.set('（请先导入模板）')
                preview_label.configure(image='')
                return
            try:
                preview_text.set('正在加载首页预览…')
                win.update_idletasks()
                png = ppt_converter.export_slide_thumbnail(
                    item['path'], 1, width=900)
                from PIL import Image, ImageTk
                img = Image.open(png).convert('RGB')
                box_w, box_h = 560, 300
                iw, ih = img.size
                r = min(box_w / iw, box_h / ih)
                nw, nh = max(1, int(iw * r)), max(1, int(ih * r))
                _preview_photo[0] = ImageTk.PhotoImage(img.resize((nw, nh)))
                preview_label.configure(
                    image=_preview_photo[0], text='')
                preview_text.set('模板首页预览')
            except ppt_converter.PptConvertError as e:
                preview_text.set('预览失败：%s' % e)
            except Exception as e:
                preview_text.set('预览暂不可用：%s' % e)

        def _repick_template():
            """重新指定当前模板要替换哪个文字框（例如模板中间的红色文字框）。"""
            item = _selected_template()
            if item is None:
                messagebox.showinfo('提示', '模板库为空。', parent=win)
                return
            _set_busy(True)
            try:
                status_var.set('正在读取当前模板文字框…')
                win.update_idletasks()
                shapes = ppt_converter.template_text_shapes(item['path'])
                if not shapes:
                    messagebox.showwarning(
                        '提示', '模板里没有可替换的文本框。', parent=win)
                    return
                shape = _choose_shape(shapes)
                if shape is None:
                    return
                new = ppt_converter.add_welcome_template(
                    item['path'], item['name'],
                    shape['slide'], shape['index'])
                ppt_converter.delete_welcome_template(item['id'])
                _reload(new['id'])
                _update_tpl_hint()
                _update_tpl_preview()
                status_var.set(
                    '已把该模板的文字框改为：%s' %
                    (shape.get('text') or '（空）'))
            except ppt_converter.PptConvertError as e:
                messagebox.showerror('失败', str(e), parent=win)
            except Exception as e:
                messagebox.showerror(
                    '失败', '发生未知错误：\n%s' % e, parent=win)
            finally:
                _set_busy(False)

        def _toggle_auto():
            size_spin.configure(
                state='disabled' if auto_var.get() else 'normal')

        def _set_busy(busy):
            widgets = (tpl_combo, pick_btn, refresh_btn, gen_btn,
                       auto_check, size_spin)
            for w in widgets:
                w.configure(state='disabled' if busy else 'normal')
            content_text.configure(
                state='disabled' if busy else 'normal')
            if not busy and auto_var.get():
                size_spin.configure(state='disabled')
            win.configure(cursor='watch' if busy else '')

        def _generate():
            item = _selected_template()
            if item is None:
                messagebox.showwarning(
                    '提示', '请先导入模板。', parent=win)
                return
            content = content_text.get('1.0', 'end-1c')
            if not content.strip():
                messagebox.showwarning(
                    '提示', '请输入要替换的文字内容。', parent=win)
                return
            default_name = _welcome_default_name(content)
            save_path = filedialog.asksaveasfilename(
                parent=win, title='保存欢迎语 PPT',
                initialfile=default_name,
                defaultextension='.pptx',
                filetypes=[('PPTX 文件', '*.pptx'),
                           ('PPT 文件', '*.ppt')])
            if not save_path:
                return
            font_size = None if auto_var.get() else size_var.get()
            _set_busy(True)
            try:
                status_var.set('正在生成欢迎语 PPT…')
                win.update_idletasks()
                out_path, used_size = \
                    ppt_converter.generate_welcome_ppt(
                        content, save_path, font_size=font_size,
                        template_id=item['id'])
                if used_size:
                    status_var.set(
                        '已生成（模板：%s），字号 %d 磅。' %
                        (item['name'], used_size))
                else:
                    status_var.set(
                        '已生成（模板：%s）：%s' %
                        (item['name'], out_path))
            except ppt_converter.PptConvertError as e:
                messagebox.showerror('生成失败', str(e), parent=win)
                return
            except Exception as e:
                messagebox.showerror('生成失败', '发生未知错误：\n%s' % e,
                                     parent=win)
                return
            finally:
                _set_busy(False)

            if messagebox.askyesno(
                    '生成完成', '欢迎语 PPT 已生成：\n%s\n\n'
                    '是否立即打开查看？' % out_path, parent=win):
                try:
                    os.startfile(out_path)
                except Exception as e:
                    messagebox.showerror(
                        '提示', '打开文件失败：\n%s' % e, parent=win)

        def _preview_result():
            """刷新预览：真实生成到临时文件并渲染结果首页（替换文字后）。"""
            item = _selected_template()
            content = content_text.get('1.0', 'end-1c')
            if item is None or not content.strip():
                _update_tpl_preview()
                return
            font_size = None if auto_var.get() else size_var.get()
            _set_busy(True)
            tmp = None
            try:
                preview_text.set('正在渲染替换后的预览…')
                win.update_idletasks()
                fd, tmp = tempfile.mkstemp(suffix='.pptx')
                os.close(fd)
                out_path, _used = ppt_converter.generate_welcome_ppt(
                    content, tmp, font_size=font_size,
                    template_id=item['id'])
                png = ppt_converter.export_slide_thumbnail(
                    out_path, 1, width=900)
                from PIL import Image, ImageTk
                img = Image.open(png).convert('RGB')
                box_w, box_h = 560, 300
                iw, ih = img.size
                r = min(box_w / iw, box_h / ih)
                nw, nh = max(1, int(iw * r)), max(1, int(ih * r))
                _preview_photo[0] = ImageTk.PhotoImage(img.resize((nw, nh)))
                preview_label.configure(image=_preview_photo[0], text='')
                preview_text.set('已替换中间文字后的预览')
            except ppt_converter.PptConvertError as e:
                preview_text.set('预览失败：%s' % e)
            except Exception as e:
                preview_text.set('预览失败：%s' % e)
            finally:
                if tmp:
                    try:
                        os.remove(tmp)
                    except Exception:
                        pass
                _set_busy(False)

        # 模板库为空且自带 templet\欢迎语.pptx 时自动登记为默认模板，
        # 首页点“生成欢迎语”即可直接输文字改模板中间文案，无需手动导入。
        if not state['items']:
            builtin = os.path.join(_BASE_DIR, 'templet', '欢迎语.pptx')
            if os.path.isfile(builtin) and \
                    not ppt_converter.list_welcome_templates():
                try:
                    shapes = ppt_converter.template_text_shapes(builtin)
                    # 取第一个“非空”文字框当作中间文字（模板里空框排在
                    # 前面，直接取 shapes[0] 会选到空白框），见 #欢迎语模板
                    s = next((x for x in shapes if x.get('text')),
                             shapes[0] if shapes else None)
                    if s:
                        ppt_converter.add_welcome_template(
                            builtin, '欢迎语', s['slide'], s['index'])
                        status_var.set(
                            '已自动加载默认模板「欢迎语」（文字框：%s）：'
                            '输入文字后点“生成 PPT”。' %
                            (s.get('text') or '（空）'))
                        _reload()
                        _update_tpl_hint()
                        _update_tpl_preview()
                except Exception:
                    pass
        _reload()
        _update_tpl_preview()

    def _open_welcome_choice(self):
        """生成欢迎语 · 方式选择：①模板改中间字 ②自定义背景。"""
        if self._activate_or_close_others('welcome'):
            return
        win = tk.Toplevel(self.root)
        win.title('生成欢迎语 - 选择方式')
        _center_window(win, 460, 300)
        win.transient(self.root)
        win.resizable(False, False)
        tk.Label(win, text='选择欢迎语生成方式：',
                 font=('微软雅黑', 13)).pack(pady=(24, 16))

        def _go(mode):
            win.destroy()
            if mode == 'tpl':
                self._open_welcome_template_library_legacy()
            else:
                self._open_welcome()

        ttk.Button(win, text='①  用模板背景 · 上下固定，只改中间文字',
                   style='Accent.TButton',
                   command=lambda: _go('tpl')).pack(
            fill=tk.X, padx=40, pady=8)
        ttk.Button(win, text='②  自定义背景 · 选图片或纯色，中间写文字',
                   command=lambda: _go('bg')).pack(
            fill=tk.X, padx=40, pady=8)
        ttk.Button(win, text='取消',
                   command=win.destroy).pack(pady=(18, 8))

    def _open_welcome(self):
        """生成欢迎语：内容 + 上传背景图/纯色，自动生成一页 PPT。"""
        if self._activate_or_close_others('welcome'):
            return
        try:
            import ppt_converter
        except Exception as e:
            messagebox.showerror(
                '错误', '加载转换模块失败：\n%s' % e, parent=self.root)
            return

        win = tk.Toplevel(self.root)
        win.title('生成欢迎语')
        _center_window(win, 880, 880)
        win.resizable(False, False)
        ui_theme.style_window(win)
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW', self._close_welcome_win)
        self._welcome_win = win

        bg_kind_var = tk.StringVar(value='pure')
        color_name_var = tk.StringVar(value='中国红')
        bg_name_var = tk.StringVar(value='')
        auto_var = tk.BooleanVar(value=True)
        size_var = tk.IntVar(value=48)
        status_var = tk.StringVar(
            value='输入欢迎语内容，可选用纯色或上传的背景图片。')
        bg_state = {'items': []}
        bg_photos = []
        preview_photo_refs = []
        preview_info_var = tk.StringVar(value='')

        # 背景方式
        bg_frame = ttk.LabelFrame(
            win, text='背景（纯色或自己上传的图片）', padding=6)
        bg_frame.pack(fill=tk.X, padx=20, pady=(16, 6))
        kind_row = tk.Frame(bg_frame)
        kind_row.pack(fill=tk.X)
        ttk.Label(kind_row, text='方式：').pack(side=tk.LEFT)
        radio_pure = ttk.Radiobutton(
            kind_row, text='纯色', value='pure',
            variable=bg_kind_var,
            command=lambda: (_sync_kind(), _refresh_preview()))
        radio_pure.pack(side=tk.LEFT)
        radio_image = ttk.Radiobutton(
            kind_row, text='上传的背景图片', value='image',
            variable=bg_kind_var,
            command=lambda: (_sync_kind(), _refresh_preview()))
        radio_image.pack(side=tk.LEFT, padx=(14, 0))

        # 纯色行
        color_row = tk.Frame(bg_frame)
        color_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(color_row, text='颜色：').pack(side=tk.LEFT)
        color_combo = ttk.Combobox(
            color_row, textvariable=color_name_var,
            state='readonly', width=12,
            values=[n for n, _ in PURE_COLORS])
        color_combo.pack(side=tk.LEFT, padx=(4, 8))
        color_combo.bind(
            '<<ComboboxSelected>>',
            lambda e: (_refresh_color(), _refresh_preview()))
        color_preview = tk.Label(
            color_row, width=8, height=2, bg=PURE_COLORS[0][1])
        color_preview.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(
            color_row, text='自定义…', style='Ghost.TButton',
            command=lambda: (_pick_custom_color(),
                             _refresh_preview())).pack(side=tk.LEFT)

        # 上传背景图行
        image_row = tk.Frame(bg_frame)
        image_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(image_row, text='图片：').pack(side=tk.LEFT)
        bg_combo = ttk.Combobox(
            image_row, textvariable=bg_name_var,
            state='readonly', width=16)
        bg_combo.pack(side=tk.LEFT, padx=(4, 8))
        bg_combo.bind(
            '<<ComboboxSelected>>',
            lambda e: (_refresh_image(), _refresh_preview()))
        bg_preview = tk.Label(
            image_row, width=12, height=2, bg='#eeeeee')
        bg_preview.pack(side=tk.LEFT, padx=(0, 10))
        upload_btn = ttk.Button(
            image_row, text='上传背景…', style='Ghost.TButton',
            command=lambda: _upload_background())
        upload_btn.pack(side=tk.LEFT, padx=(0, 6))
        delete_bg_btn = ttk.Button(
            image_row, text='删除选中', style='Danger.TButton',
            command=lambda: _delete_background())
        delete_bg_btn.pack(side=tk.LEFT)
        ttk.Button(image_row, text='重命名…', style='Ghost.TButton',
                   command=lambda: _rename_background()).pack(side=tk.LEFT,
                                                             padx=(6, 0))

        # 内容
        tk.Label(
            win, text='欢迎语内容（换行会保留，自动居中）',
            font=('微软雅黑', 10)).pack(
            anchor='w', padx=20, pady=(8, 4))
        content_text = tk.Text(
            win, width=62, height=5, wrap=tk.WORD,
            font=('微软雅黑', 14),
            relief=tk.FLAT, bd=0, padx=10, pady=8, undo=True,
            bg=ui_theme.PALETTE['surface'], fg=ui_theme.PALETTE['text'],
            insertbackground=ui_theme.PALETTE['primary'],
            highlightthickness=1,
            highlightbackground=ui_theme.PALETTE['border'],
            highlightcolor=ui_theme.PALETTE['primary'])
        content_text.pack(fill=tk.X, padx=20)
        content_text.bind(
            '<<Modified>>', lambda e: _refresh_preview())

        size_row = tk.Frame(win)
        size_row.pack(anchor='w', padx=20, pady=(8, 0))
        auto_check = ttk.Checkbutton(
            size_row, text='自动调整字号（过长自动缩小）',
            variable=auto_var, command=lambda: _toggle_auto())
        auto_check.pack(side=tk.LEFT)
        ttk.Label(size_row, text='字号(pt)：').pack(
            side=tk.LEFT, padx=(16, 4))
        size_spin = ttk.Spinbox(
            size_row, from_=8, to=200, textvariable=size_var,
            width=6, state='disabled')
        size_spin.pack(side=tk.LEFT)

        gen_btn = ttk.Button(
            win, text='生成 PPT', style='Accent.TButton',
            command=lambda: _generate())
        gen_btn.pack(pady=(8, 4))
        ttk.Label(win, textvariable=status_var,
                  foreground='#7a8ca0', font=('微软雅黑', 9),
                  wraplength=760, justify=tk.LEFT).pack(
            anchor='w', padx=20, pady=(0, 12))

        # 输出预览（与生成结果使用同一渲染逻辑）
        preview_frame = ttk.LabelFrame(
            win, text='输出预览（与实际生成一致）', padding=6)
        preview_frame.pack(fill=tk.X, padx=20, pady=(2, 0))
        preview_label = tk.Label(
            preview_frame, bg='#f5f6f8', text='输入内容后自动刷新预览',
            font=('微软雅黑', 9))
        preview_label.pack()
        ttk.Label(preview_frame, textvariable=preview_info_var,
                  foreground='#7a8ca0', font=('微软雅黑', 9)).pack(
            anchor='e', pady=(3, 0))

        def _color_hex(name):
            for n, c in PURE_COLORS:
                if n == name:
                    return c
            if len(name) == 7 and name.startswith('#'):
                try:
                    int(name[1:], 16)
                    return name
                except ValueError:
                    pass
            return PURE_COLORS[0][1]

        def _refresh_color():
            color_preview.configure(bg=_color_hex(color_name_var.get()))

        def _pick_custom_color():
            from tkinter import colorchooser
            rgb, hex_color = colorchooser.askcolor(
                color=_color_hex(color_name_var.get()),
                title='选择纯色背景', parent=win)
            if hex_color is None:
                return
            hit = next(
                (n for n, c in PURE_COLORS
                 if c.lower() == hex_color.lower()), None)
            if hit is None:
                PURE_COLORS.append((hex_color, hex_color))
                color_combo.configure(
                    values=[n for n, _ in PURE_COLORS])
            color_name_var.set(hit or hex_color)
            _refresh_color()

        def _reload_bg(select_name=None):
            bg_state['items'] = ppt_converter.list_uploaded_backgrounds()
            names = [x['name'] for x in bg_state['items']]
            bg_combo.configure(values=names)
            if select_name and select_name in names:
                bg_name_var.set(select_name)
            elif names:
                bg_name_var.set(names[0])
            else:
                bg_name_var.set('')
            _refresh_image()

        def _selected_bg():
            name = bg_name_var.get()
            for item in bg_state['items']:
                if item['name'] == name:
                    return item
            return None

        def _refresh_image():
            bg_photos.clear()
            item = _selected_bg()
            if item is None:
                bg_preview.configure(
                    text='未选背景', image='', bg='#eeeeee')
                return
            try:
                from PIL import Image, ImageTk
                with Image.open(item['image']) as raw:
                    # draft 让大 JPEG 在解码阶段就缩小，省内存也更快
                    try:
                        raw.draft('RGB', (440, 248))
                    except Exception:
                        pass
                    im = raw.convert('RGB')
                try:
                    im.thumbnail((220, 124))
                    photo = ImageTk.PhotoImage(im)
                finally:
                    im.close()
                bg_photos.append(photo)
                bg_preview.configure(image=photo, text='')
            except Exception:
                bg_preview.configure(
                    text='预览失败', image='', bg='#eeeeee')

        def _sync_kind():
            is_pure = bg_kind_var.get() == 'pure'
            color_combo.configure(
                state='readonly' if is_pure else 'disabled')
            bg_combo.configure(
                state='readonly' if not is_pure else 'disabled')
            upload_btn.configure(
                state='normal' if not is_pure else 'disabled')
            delete_bg_btn.configure(
                state='normal' if not is_pure else 'disabled')

        def _upload_background():
            src = filedialog.askopenfilename(
                parent=win, title='选择背景图片',
                filetypes=[('图片文件', '*.png *.jpg *.jpeg *.bmp *.webp'),
                           ('所有文件', '*.*')])
            if not src:
                return
            stem = os.path.splitext(os.path.basename(src))[0]
            name = simpledialog.askstring(
                '背景名称', '给背景起个名字：',
                initialvalue=stem, parent=win)
            if not name:
                return
            try:
                item = ppt_converter.add_uploaded_background(src, name)
                _reload_bg(item['name'])
                _refresh_preview()
                status_var.set('背景已保存：%s' % name)
            except ppt_converter.PptConvertError as e:
                messagebox.showerror('上传失败', str(e), parent=win)
            except Exception as e:
                messagebox.showerror(
                    '上传失败', '发生未知错误：\n%s' % e, parent=win)

        def _delete_background():
            item = _selected_bg()
            if item is None:
                messagebox.showinfo('提示', '没有可删除的背景。', parent=win)
                return
            if not messagebox.askyesno(
                    '删除背景', '确定删除“%s”吗？' % item['name'],
                    parent=win):
                return
            ppt_converter.delete_uploaded_background(item['id'])
            _reload_bg()
            _refresh_preview()
            status_var.set('已删除背景：%s' % item['name'])

        def _rename_background():
            item = _selected_bg()
            if item is None:
                messagebox.showinfo('提示', '请先在下拉框里选一张背景图。',
                                    parent=win)
                return
            new_name = simpledialog.askstring(
                '重命名背景', '新的名称：', initialvalue=item['name'],
                parent=win)
            if not new_name or new_name.strip() == item['name']:
                return
            new_name = new_name.strip()
            if any(x['name'] == new_name
                   for x in ppt_converter.list_uploaded_backgrounds()):
                messagebox.showwarning('重名了', '已经有一张叫“%s”的图。'
                                       % new_name, parent=win)
                return
            if not ppt_converter.rename_uploaded_background(item['id'],
                                                            new_name):
                messagebox.showerror('重命名失败', '这张图没能改名。',
                                     parent=win)
                return
            _reload_bg(new_name)
            _refresh_preview()
            status_var.set('已重命名：%s' % new_name)

        def _toggle_auto():
            size_spin.configure(
                state='disabled' if auto_var.get() else 'normal')

        # 输入时合并刷新：连续打字只在停手 250ms 后渲染一次，
        # 避免每敲一个字都重新解码背景图、把界面卡住
        _preview_job = {'id': None}

        def _refresh_preview_later():
            if _preview_job['id'] is not None:
                try:
                    win.after_cancel(_preview_job['id'])
                except tk.TclError:
                    pass
                _preview_job['id'] = None
            try:
                _preview_job['id'] = win.after(250, _refresh_preview)
            except tk.TclError:
                pass

        def _refresh_preview():
            _preview_job['id'] = None
            try:
                if not win.winfo_exists():
                    return
            except tk.TclError:
                return
            try:
                content = content_text.get('1.0', 'end-1c')
                font_size = None if auto_var.get() else size_var.get()
                background_image = None
                background_color = None
                if bg_kind_var.get() == 'pure':
                    background_color = _color_hex(color_name_var.get())
                else:
                    item = _selected_bg()
                    if item:
                        background_image = item['image']
                img, used_size = \
                    ppt_converter.render_welcome_preview(
                        content,
                        background_image=background_image,
                        background_color=background_color,
                        font_size=font_size, width=480, height=270)
                from PIL import ImageTk
                try:
                    photo = ImageTk.PhotoImage(img)
                finally:
                    img.close()
                preview_photo_refs.clear()
                preview_photo_refs.append(photo)
                preview_label.configure(image=photo, text='')
                preview_info_var.set('预览字号：%d 磅' % used_size)
            except Exception:
                try:
                    preview_label.configure(image='', text='预览生成失败')
                except tk.TclError:
                    pass
            finally:
                try:
                    content_text.edit_modified(False)
                except tk.TclError:
                    pass

        def _set_busy(busy):
            for w in (radio_pure, radio_image, color_combo,
                      upload_btn, delete_bg_btn, bg_combo,
                      gen_btn, auto_check, size_spin):
                w.configure(state='disabled' if busy else 'normal')
            content_text.configure(
                state='disabled' if busy else 'normal')
            if not busy:
                _sync_kind()
                if auto_var.get():
                    size_spin.configure(state='disabled')
            win.configure(cursor='watch' if busy else '')

        def _generate():
            content = content_text.get('1.0', 'end-1c')
            if not content.strip():
                messagebox.showwarning(
                    '提示', '请输入欢迎语内容。', parent=win)
                return
            background_color = None
            background_image = None
            if bg_kind_var.get() == 'pure':
                background_color = _color_hex(color_name_var.get())
            else:
                item = _selected_bg()
                if item is None:
                    messagebox.showwarning(
                        '提示', '请先上传或选择一张背景图片。',
                        parent=win)
                    return
                background_image = item['image']

            default_name = _welcome_default_name(content)
            save_path = filedialog.asksaveasfilename(
                parent=win, title='保存欢迎语 PPT',
                initialfile=default_name, defaultextension='.pptx',
                filetypes=[('PPTX 文件', '*.pptx'),
                           ('PPT 文件', '*.ppt')])
            if not save_path:
                return
            font_size = None if auto_var.get() else size_var.get()
            _set_busy(True)
            try:
                status_var.set('正在生成欢迎语 PPT…')
                win.update_idletasks()
                out_path, used_size = \
                    ppt_converter.generate_welcome_slide(
                        content, save_path, font_size=font_size,
                        background_image=background_image,
                        background_color=background_color)
                status_var.set(
                    '已生成欢迎语 PPT，字号 %d 磅。' % used_size)
            except ppt_converter.PptConvertError as e:
                messagebox.showerror('生成失败', str(e), parent=win)
                return
            except Exception as e:
                messagebox.showerror('生成失败', '发生未知错误：\n%s' % e,
                                     parent=win)
                return
            finally:
                _set_busy(False)
            if messagebox.askyesno(
                    '生成完成', '欢迎语 PPT 已生成：\n%s\n\n'
                    '是否立即打开查看？' % out_path, parent=win):
                try:
                    os.startfile(out_path)
                except Exception as e:
                    messagebox.showerror(
                        '提示', '打开文件失败：\n%s' % e, parent=win)

        _reload_bg()
        _refresh_color()
        _sync_kind()
        _refresh_preview()

    # ---------- 扫码收文件 ----------
    def _open_file_inbox(self):
        """扫码收文件：手机扫码上传到这台电脑（局域网，不走公网）。"""
        if self._activate_or_close_others('inbox'):
            return

        p = ui_theme.PALETTE
        if self._inbox is None:
            self._inbox = file_inbox.Inbox()
        inbox = self._inbox

        win = tk.Toplevel(self.root)
        win.title('扫码收文件')
        _center_window(win, 800, 620)
        win.resizable(False, False)
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW', self._close_inbox_win)
        ui_theme.style_window(win)
        self._inbox_win = win

        body = tk.Frame(win, bg=p['bg'], padx=16, pady=14)
        body.pack(fill=tk.BOTH, expand=True)

        head = tk.Frame(body, bg=p['bg'])
        head.pack(fill=tk.X)
        tk.Label(head, text='扫码收文件', font=ui_theme.font(13, 'bold'),
                 bg=p['bg'], fg=p['text']).pack(side=tk.LEFT)
        dot = tk.Label(head, text='●', font=ui_theme.font(11),
                       bg=p['bg'], fg=p['faint'])
        dot.pack(side=tk.LEFT, padx=(12, 4))
        state_label = tk.Label(head, text='未启动', font=ui_theme.font(9),
                               bg=p['bg'], fg=p['muted'])
        state_label.pack(side=tk.LEFT)
        tk.Label(head, text='手机与电脑连同一个 WiFi，用微信扫一扫',
                 font=ui_theme.font(9), bg=p['bg'],
                 fg=p['faint']).pack(side=tk.RIGHT)

        top = tk.Frame(body, bg=p['bg'])
        top.pack(fill=tk.X, pady=(12, 10))

        qr_box = tk.Frame(top, bg=p['surface'], highlightthickness=1,
                          highlightbackground=p['border'])
        qr_box.pack(side=tk.LEFT, anchor='n')
        qr_label = tk.Label(qr_box, text='未启动', bg=p['surface'],
                            fg=p['faint'], font=ui_theme.font(9),
                            width=26, height=11)
        qr_label.pack(padx=10, pady=10)
        qr_btns = tk.Frame(qr_box, bg=p['surface'])
        qr_btns.pack(padx=10, pady=(0, 10))
        ttk.Button(qr_btns, text='保存二维码图片',
                   command=lambda: _save_qr()).pack(side=tk.LEFT,
                                                    padx=(0, 6))
        print_btn = ttk.Button(qr_btns, text='打印二维码',
                               style='Accent.TButton',
                               command=lambda: _print_qr())
        print_btn.pack(side=tk.LEFT)

        right = tk.Frame(top, bg=p['bg'])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(14, 0))

        def _row(label):
            row = tk.Frame(right, bg=p['bg'])
            row.pack(fill=tk.X, pady=3)
            tk.Label(row, text=label, font=ui_theme.font(9), bg=p['bg'],
                     fg=p['muted'], width=8, anchor='e').pack(
                side=tk.LEFT, padx=(0, 8))
            return row

        row = _row('保存到')
        folder_var = tk.StringVar(value=inbox.folder)
        tk.Entry(row, textvariable=folder_var, font=ui_theme.font(9),
                 state='readonly', readonlybackground=p['surface'],
                 relief=tk.FLAT, bd=0, highlightthickness=1,
                 highlightbackground=p['border']).pack(
            side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        ttk.Button(row, text='换位置', width=6,
                   command=lambda: _choose_folder()).pack(
            side=tk.LEFT, padx=(6, 0))
        ttk.Button(row, text='打开', width=5,
                   command=lambda: _open_folder()).pack(
            side=tk.LEFT, padx=(6, 0))

        row = _row('手机网址')
        url_var = tk.StringVar(value='—')
        tk.Entry(row, textvariable=url_var, font=ui_theme.font(10),
                 state='readonly', readonlybackground=p['surface'],
                 relief=tk.FLAT, bd=0, highlightthickness=1,
                 highlightbackground=p['border']).pack(
            side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        ttk.Button(row, text='复制', width=5,
                   command=lambda: _copy_url()).pack(side=tk.LEFT, padx=(6, 0))

        row = _row('端口')
        port_var = tk.IntVar(value=inbox.port)
        ttk.Spinbox(row, from_=1024, to=65535, width=7,
                    textvariable=port_var).pack(side=tk.LEFT)
        toggle_btn = ttk.Button(row, text='开始接收', style='Accent.TButton',
                                command=lambda: _toggle())
        toggle_btn.pack(side=tk.LEFT, padx=(10, 0))

        tk.Label(right,
                 text='支持 PPT / PDF / 图片 / Word / 压缩包等任意格式，'
                      '单个文件最大 1GB；\n手机打不开网页时，请确认在同一个 '
                      'WiFi，并在 Windows 防火墙弹窗里选“允许访问”。',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 justify='left', anchor='w').pack(fill=tk.X, pady=(10, 0))

        # 收到的文件列表
        list_head = tk.Frame(body, bg=p['bg'])
        list_head.pack(fill=tk.X, pady=(4, 6))
        tk.Label(list_head, text='收到的文件', font=ui_theme.font(10, 'bold'),
                 bg=p['bg'], fg=p['text']).pack(side=tk.LEFT)
        count_label = tk.Label(list_head, text='共 0 个',
                               font=ui_theme.font(9), bg=p['bg'],
                               fg=p['faint'])
        count_label.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(list_head, text='清空记录', style='Ghost.TButton',
                   command=lambda: _clear()).pack(side=tk.RIGHT)

        wrap = tk.Frame(body, bg=p['bg'])
        wrap.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(wrap, columns=('name', 'size', 'time'),
                            show='headings', height=8)
        tree.heading('name', text='文件名')
        tree.heading('size', text='大小')
        tree.heading('time', text='时间')
        tree.column('name', width=440, anchor='w', stretch=True)
        tree.column('size', width=90, anchor='e', stretch=False)
        tree.column('time', width=90, anchor='center', stretch=False)
        scroll = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        last = {'key': None}

        def _file_at(index):
            """按表格行号取回收到的文件（越界返回 None）。"""
            files = inbox.snapshot()
            if 0 <= index < len(files):
                return files[index]
            return None

        def _open_item(index):
            """双击/右键：用系统默认程序打开这个文件。"""
            item = _file_at(index)
            if item is None:
                return
            path = item.get('path')
            if not path or not os.path.exists(path):
                messagebox.showwarning('提示', '文件已经不在原位置了：\n%s'
                                       % (path or item.get('name', '')),
                                       parent=win)
                return
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showwarning('提示', '打不开这个文件：\n%s' % e,
                                       parent=win)

        def _reveal_item(index):
            """在资源管理器里定位这个文件。"""
            item = _file_at(index)
            if item is None or not item.get('path'):
                return
            try:
                subprocess.Popen(['explorer', '/select,',
                                  os.path.normpath(item['path'])])
            except Exception:
                _open_folder()

        def _on_double_click(_event=None):
            sel = tree.selection()
            if sel:
                _open_item(int(sel[0]))

        def _on_right_click(event):
            row = tree.identify_row(event.y)
            if row:
                tree.selection_set(row)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                try:
                    menu.grab_release()
                except tk.TclError:
                    pass

        menu = tk.Menu(win, tearoff=0, bd=0, bg=p['surface'], fg=p['text'],
                       activebackground=p['primary_soft'],
                       activeforeground=p['primary_dark'],
                       font=ui_theme.font(9))
        menu.add_command(label='打开文件',
                         command=lambda: _open_item(_selected_row()))
        menu.add_command(label='打开所在文件夹',
                         command=lambda: _reveal_item(_selected_row()))

        def _selected_row():
            sel = tree.selection()
            return int(sel[0]) if sel else -1

        tree.bind('<Double-Button-1>', _on_double_click)
        tree.bind('<Button-3>', _on_right_click)
        tk.Label(body, text='提示：双击列表里的文件即可打开，右键可定位到文件夹',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 anchor='w').pack(fill=tk.X, pady=(6, 0))

        def _alive():
            try:
                return bool(win.winfo_exists())
            except tk.TclError:
                return False

        def _fmt_size(size):
            if size < 1024:
                return '%d B' % size
            if size < 1048576:
                return '%.1f KB' % (size / 1024.0)
            if size < 1073741824:
                return '%.1f MB' % (size / 1048576.0)
            return '%.2f GB' % (size / 1073741824.0)

        def _copy_url():
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(url_var.get())
                state_label.config(text='网址已复制')
            except tk.TclError:
                pass

        def _choose_folder():
            path = filedialog.askdirectory(
                parent=win, title='选择收件保存位置',
                initialdir=inbox.folder)
            if not path:
                return
            inbox.folder = os.path.abspath(path)
            folder_var.set(inbox.folder)
            inbox.clear_list()
            _refresh(force=True)

        def _open_folder():
            try:
                os.makedirs(inbox.folder, exist_ok=True)
                os.startfile(inbox.folder)
            except Exception as e:
                messagebox.showwarning('提示', '打不开文件夹：\n%s' % e,
                                       parent=win)

        def _clear():
            inbox.clear_list()
            _refresh(force=True)

        def _make_qr():
            from PIL import ImageTk
            try:
                url = inbox.url()
                matrix = qrcode_lite.make_matrix(url)
                scale = max(3, int(230 // (len(matrix) + 12)))
                img = qrcode_lite.to_image(matrix, scale=scale)
                photo = ImageTk.PhotoImage(img)
                qr_label.configure(image=photo, width=0, height=0, text='')
                qr_label.image = photo
                url_var.set(url)
            except Exception as e:
                qr_label.configure(image='', text='二维码生成失败',
                                   width=26, height=11)
                safe_guard.write_log('二维码生成失败：%r' % (e,))

        def _build_poster(url):
            """生成高清二维码海报（标题 + 二维码 + 网址 + 提示）。

            保存图片和直接打印共用这一张，保证屏幕预览、打印件一致。
            """
            from PIL import Image, ImageDraw, ImageFont

            def _font(sz, bold=False):
                names = ('msyhbd.ttc', 'msyh.ttc') if bold \
                    else ('msyh.ttc', 'simhei.ttf')
                for name in names:
                    try:
                        return ImageFont.truetype(
                            os.path.join(r'C:\Windows\Fonts', name), sz)
                    except Exception:
                        continue
                return ImageFont.load_default()

            matrix = qrcode_lite.make_matrix(url)
            scale = max(8, 1240 // (len(matrix) + 12))
            qr = qrcode_lite.to_image(matrix, scale=scale)
            pad, gap = 90, 50
            t_f, u_f, n_f = _font(64, True), _font(40), _font(32)
            tmp = Image.new('RGB', (10, 10))
            td = ImageDraw.Draw(tmp)
            t_w = td.textbbox((0, 0), '扫码上传文件', font=t_f)[2]
            u_w = td.textbbox((0, 0), url, font=u_f)[2]
            n_txt = '请用微信扫一扫（与电脑连接同一个 WiFi）'
            n_w = td.textbbox((0, 0), n_txt, font=n_f)[2]
            W = max(qr.width, t_w, u_w, n_w) + pad * 2
            img = Image.new('RGB', (W, pad + 64 + gap + qr.height
                                    + gap + 40 + gap + 32 + pad), '#ffffff')
            d = ImageDraw.Draw(img)
            y = pad
            d.text(((W - t_w) // 2, y), '扫码上传文件', font=t_f,
                   fill='#111111')
            y += 64 + gap
            img.paste(qr, ((W - qr.width) // 2, y))
            y += qr.height + gap
            d.text(((W - u_w) // 2, y), url, font=u_f, fill='#444444')
            y += 40 + gap
            d.text(((W - n_w) // 2, y), n_txt, font=n_f, fill='#999999')
            return img

        def _save_qr():
            """把当前二维码导出成高清海报 PNG（可打印 / 贴进 PPT）。"""
            url = url_var.get()
            if not url or url == '—':
                messagebox.showinfo('扫码收文件', '请先点击“开始接收”',
                                    parent=win)
                return
            try:
                img = _build_poster(url)
                desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
                path = filedialog.asksaveasfilename(
                    parent=win, title='保存二维码图片',
                    initialdir=desktop if os.path.isdir(desktop) else None,
                    initialfile='扫码上传二维码.png', defaultextension='.png',
                    filetypes=[('PNG 图片', '*.png')])
                if not path:
                    return
                img.save(path)
            except Exception as e:
                safe_guard.write_log('保存二维码失败：%r' % (e,))
                messagebox.showerror('扫码收文件', '保存失败：\n%s' % e,
                                     parent=win)
                return
            if messagebox.askyesno(
                    '扫码收文件', '已保存：\n%s\n\n打开所在文件夹？' % path,
                    parent=win):
                try:
                    os.startfile(os.path.dirname(path))
                except Exception:
                    pass

        def _print_qr():
            """已连实体打印机时，不弹对话框直接把二维码海报打出来。

            没有打印机或默认是“打印成 PDF”等虚拟设备时，引导先保存图片。
            """
            import seat_card
            url = url_var.get()
            if not url or url == '—':
                messagebox.showinfo('扫码收文件', '请先点击“开始接收”',
                                    parent=win)
                return
            btn = print_btn
            try:
                img = _build_poster(url)
            except Exception as e:
                safe_guard.write_log('打印二维码失败：%r' % (e,))
                messagebox.showerror('扫码收文件', '生成二维码失败：\n%s' % e,
                                     parent=win)
                return
            if btn is not None:
                btn.config(state=tk.DISABLED, text='打印中…')
                win.update_idletasks()
            try:
                ok, info = seat_card.print_image_silent(
                    img, doc_title='扫码上传二维码')
            finally:
                if btn is not None:
                    btn.config(state=tk.NORMAL, text='打印二维码')
            if ok:
                safe_guard.write_log('二维码已发送到打印机：%s' % info)
                messagebox.showinfo(
                    '扫码收文件',
                    '二维码已发送到打印机：\n%s\n请到打印机取纸。' % info,
                    parent=win)
            else:
                safe_guard.write_log('二维码未打印：%s' % info)
                if messagebox.askyesno(
                        '扫码收文件', info + '\n\n是否先把二维码保存为图片？',
                        parent=win):
                    _save_qr()

        def _toggle():
            if inbox.running:
                inbox.stop()
                state_label.config(text='已停止', fg=p['muted'])
                # 停止后旧网址已失效：清掉二维码和网址，
                # 避免保存/打印出扫不开的码
                url_var.set('—')
                qr_label.configure(image='', text='未启动',
                                   width=26, height=11)
                qr_label.image = None
            else:
                try:
                    inbox.port = max(1024, min(int(port_var.get()), 65535))
                except (tk.TclError, ValueError):
                    port_var.set(inbox.port)
                try:
                    port_var.set(inbox.start())
                except Exception as e:
                    messagebox.showerror(
                        '启动失败', '收件服务启动失败：\n%s' % e, parent=win)
                    return
                _make_qr()
            _refresh(force=True)

        def _refresh(force=False):
            if not _alive():
                return
            running = inbox.running
            dot.config(fg=p['success'] if running else p['faint'])
            if running:
                state_label.config(text='正在接收', fg=p['muted'])
            elif state_label.cget('text') == '正在接收':
                state_label.config(text='未启动', fg=p['muted'])
            toggle_btn.configure(
                text='停止接收' if running else '开始接收',
                style='Danger.TButton' if running else 'Accent.TButton')
            files = inbox.snapshot()
            count_label.config(
                text='共 %d 个 · %s'
                     % (len(files), _fmt_size(inbox.total_bytes())))
            key = tuple((f['name'], f['size'], f['time']) for f in files)
            if force or key != last['key']:
                last['key'] = key
                tree.delete(*tree.get_children())
                for i, item in enumerate(files):
                    tree.insert('', 'end', iid=str(i),
                                values=(item['name'],
                                        _fmt_size(item['size']),
                                        item['time']))
            try:
                win.after(1000, _refresh)
            except tk.TclError:
                pass

        # 打开窗口就自动开始接收
        try:
            if not inbox.running:
                port_var.set(inbox.start())
            _make_qr()
        except Exception as e:
            state_label.config(text='启动失败：%s' % e, fg=p['danger'])
        _refresh(force=True)

    def _close_inbox_win(self, keep_running=False):
        """关闭收件窗口：默认连服务一起停掉。

        keep_running=True 时只关窗口、服务继续跑，这样打开别的功能窗口时
        手机还能继续上传。
        """
        if self._inbox is not None and not keep_running:
            try:
                self._inbox.stop()
            except Exception:
                pass
        win = self._inbox_win
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._inbox_win = None

    # ---------- 微信传文件 ----------
    def _open_wechat_helper(self):
        """微信传文件：直接在浏览器打开微信“文件传输助手”网页版。

        微信官方不允许“扫码 → 跳转到文件传输助手”，所以不走二维码，
        点击卡片直接打开网页，由网页自己显示登录二维码。
        """
        url = 'https://filehelper.weixin.qq.com/'
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showwarning('提示', '打不开网页：\n%s' % e,
                                   parent=self.root)

    # ---------- 快捷邮箱 ----------
    # 只做“快捷打开邮箱网页”，不做收发：直接用默认浏览器打开登录页，
    # 免去浏览器里搜索的步骤，也不依赖任何第三方邮箱 API。
    _MAIL_SITES = (
        ('QQ 邮箱', 'https://mail.qq.com/'),
        ('腾讯企业邮箱', 'https://exmail.qq.com/'),
        ('163 邮箱', 'https://mail.163.com/'),
        ('126 邮箱', 'https://mail.126.com/'),
        ('Outlook / Hotmail', 'https://outlook.live.com/mail/'),
        ('Gmail', 'https://mail.google.com/'),
        ('阿里邮箱', 'https://mail.aliyun.com/'),
        ('新浪邮箱', 'https://mail.sina.com.cn/'),
    )

    def _make_mail_menu(self):
        """常用邮箱选择菜单（单独成方法，便于测试逐项打开）。"""
        menu = tk.Menu(self.root, tearoff=0)
        for name, url in self._MAIL_SITES:
            menu.add_command(
                label=name,
                command=lambda u=url, n=name: self._open_mail_site(u, n))
        return menu

    def _open_mail_menu(self):
        """点“快捷邮箱”卡片：在点击位置弹出邮箱选择菜单。"""
        menu = self._make_mail_menu()
        x, y = self.root.winfo_pointerxy()
        # 菜单项多时防止弹出位置超出屏幕
        menu_h = 28 * len(self._MAIL_SITES) + 12
        if y + menu_h > self.root.winfo_screenheight():
            y = max(0, self.root.winfo_screenheight() - menu_h - 10)
        if x + 180 > self.root.winfo_screenwidth():
            x = max(0, self.root.winfo_screenwidth() - 190)
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _open_mail_site(self, url, name):
        """用默认浏览器直达邮箱登录页。"""
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showwarning('快捷邮箱',
                                   '打不开 %s：\n%s' % (name, e),
                                   parent=self.root)

    # ---------- 会务视频 ----------
    def _open_video(self):
        """会务视频：选多个视频、排好顺序，依次全屏播放到指定显示器。"""
        if self._activate_or_close_others('video'):
            return

        p = ui_theme.PALETTE
        if self._video_player is None:
            self._video_player = video_player.VideoPlayer()
        player = self._video_player
        preview_ctl = video_player.EmbeddedPreview()

        win = tk.Toplevel(self.root)
        win.title('会务视频播放')
        _center_window(win, 1060, 880)
        win.minsize(900, 620)
        win.resizable(True, True)
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW',
                     lambda w=win: self._request_close(
                         '会务视频播放', self._close_video_win, w))
        ui_theme.style_window(win)
        self._video_win = win
        win._preview_ctl = preview_ctl      # 关窗口时要把预览一起停掉

        state = {'list': [], 'idx': -1, 'playing': False, 'topmost': True,
                 'audio_device': None}   # None=跟随系统默认播放设备
        state.update({'layout': 'full',              # full=整屏 / compose=合成
                      'video_box': [0.20, 0.08, 0.60, 0.84],
                      'video_bg': '',             # 背景图名（空=不用图）
                      'video_left': '',           # 左侧图名（空=不用图）
                      'video_right': '',          # 右侧图名（空=不用图）
                      'box_drag': None, 'box_origin': None,
                      'bg_screen': None,          # 合成投屏的背景窗口
                      'compose_screen': None,     # 大屏合成窗口（背景+视频框）
                      'frame_pic': None,          # 预览用的视频截图
                      'frame_for': None,          # 这张截图是哪个视频的
                      'pick_job': None,           # 稍后取帧的定时任务
                      'grab_win': None,           # 抓帧用的屏幕外小窗口
                      'paused': False,            # 是否暂停中
                      'duration': 0.0,            # 当前视频总长（秒）
                      'seek_drag': False,         # 正在拖进度条
                      'start_at': 0.0,            # 下次从这个位置开始播
                      'start_at_for': '',         # 这个位置是哪个视频的
                      'pending_seek': None,       # 起播后要跳到的位置
                      'seek_tries': 0})

        # 播放时两种退出方式：ESC 键 + 鼠标右键菜单。
        # VLC 全屏是一个独立进程窗口，焦点在它自己那边，tkinter 的 bind 收不到
        # 键盘和右键事件；RegisterHotKey 注册纯 ESC 在不少机器上会被系统直接拒绝
        # （实测返回失败），所以统一用 GetAsyncKeyState 轮询——它读的是全局物理
        # 键鼠状态，跟谁持有焦点无关，VLC 全屏时同样有效。
        esc_watch = {'after': None, 'esc_held': False, 'rbtn_held': False}

        def _keys_state():
            """一次读两个键：ESC 是否按下、鼠标右键是否按下。"""
            try:
                import ctypes
                u = ctypes.windll.user32
                return (bool(u.GetAsyncKeyState(0x1B) & 0x8000),    # VK_ESCAPE
                        bool(u.GetAsyncKeyState(0x02) & 0x8000))    # VK_RBUTTON
            except Exception:
                return False, False

        menu_box = {'win': None}

        def _close_play_menu():
            m = menu_box['win']
            menu_box['win'] = None
            if m is not None:
                try:
                    m.destroy()
                except Exception:
                    pass

        def _show_play_menu(x, y):
            """弹一个朴素小菜单，只有「退出播放」一项。
            只置顶一次就够了：VLC 的窗口仅在起播时摆一次位置，之后不会
            再来抢层级。之前每 200ms 反复 lift 反而把菜单刷得一闪一闪。"""
            _close_play_menu()
            m = tk.Toplevel(win)
            menu_box['win'] = m
            m.overrideredirect(True)
            try:
                m.attributes('-topmost', True)
            except Exception:
                pass
            line = tk.Frame(m, bg=p['border'])
            line.pack(fill=tk.BOTH, expand=True)
            item = tk.Label(line, text='退出播放', bg=p['surface'], fg=p['text'],
                            font=('Microsoft YaHei', 9), padx=14, pady=5,
                            cursor='hand2')
            item.pack(padx=1, pady=1)

            def _click(_e):
                _close_play_menu()
                _stop()

            def _enter(_e):
                item.configure(bg=p['hover'])

            def _leave(_e):
                item.configure(bg=p['surface'])

            item.bind('<Button-1>', _click)
            item.bind('<Enter>', _enter)
            item.bind('<Leave>', _leave)
            m.bind('<FocusOut>', lambda e: _close_play_menu())
            m.geometry('+%d+%d' % (x, y))
            m.lift()
            try:
                m.focus_force()
            except Exception:
                pass

        def _watch_keys():
            esc_watch['after'] = None
            if not _alive() or not state['playing']:
                _close_play_menu()
                return
            esc_down, rbtn_down = _keys_state()
            if esc_down and not esc_watch['esc_held']:
                esc_watch['esc_held'] = True     # 只在“刚按下”那一下停一次
                _close_play_menu()
                _stop()
                return
            esc_watch['esc_held'] = esc_down
            # 右键当开关：没菜单就弹，有菜单就收
            if rbtn_down and not esc_watch['rbtn_held']:
                esc_watch['rbtn_held'] = True
                if menu_box['win'] is not None:
                    _close_play_menu()
                else:
                    try:
                        x, y = win.winfo_pointerxy()
                    except Exception:
                        x, y = 0, 0
                    _show_play_menu(x, y)
            esc_watch['rbtn_held'] = rbtn_down
            esc_watch['after'] = win.after(60, _watch_keys)

        def _start_esc_watch():
            esc_watch['esc_held'] = False
            esc_watch['rbtn_held'] = False
            if esc_watch['after'] is None and state['playing']:
                esc_watch['after'] = win.after(60, _watch_keys)

        def _stop_esc_watch():
            aid = esc_watch['after']
            esc_watch['after'] = None
            esc_watch['esc_held'] = False
            esc_watch['rbtn_held'] = False
            _close_play_menu()
            if aid is not None:
                try:
                    win.after_cancel(aid)
                except Exception:
                    pass

        win._disarm_esc = _stop_esc_watch    # 关窗口时把轮询停掉

        body = tk.Frame(win, bg=p['bg'], padx=16, pady=14)
        body.pack(fill=tk.BOTH, expand=True)

        head = tk.Frame(body, bg=p['bg'])
        head.pack(fill=tk.X)
        tk.Label(head, text='会务视频播放', font=ui_theme.font(13, 'bold'),
                 bg=p['bg'], fg=p['text']).pack(side=tk.LEFT)
        dot = tk.Label(head, text='●', font=ui_theme.font(11), bg=p['bg'],
                       fg=p['faint'])
        dot.pack(side=tk.LEFT, padx=(12, 4))
        state_label = tk.Label(head, text='未播放', font=ui_theme.font(9),
                               bg=p['bg'], fg=p['muted'])
        state_label.pack(side=tk.LEFT)
        tk.Label(head, text='双击列表项从该视频开始播放；播放中按 ESC 退出播放',
                 font=ui_theme.font(9), bg=p['bg'],
                 fg=p['faint']).pack(side=tk.RIGHT)

        mid = tk.Frame(body, bg=p['bg'])

        left = tk.Frame(mid, bg=p['bg'])
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 预览区：播放时把播放器的静音画面嵌进这个黑框里（就在软件内显示）
        preview_host = tk.Canvas(left, bg='#000000', width=620, height=349,
                                 highlightthickness=0, bd=0)
        preview_host.pack(fill=tk.BOTH, expand=True)
        preview_hint = tk.Label(preview_host,
                                text=('选好视频后这里显示预览画面（静音）'
                                      if preview_ctl.available else
                                      '软件内预览需要 VLC（未检测到）'),
                                font=ui_theme.font(9), bg='#000000',
                                fg='#7c8797')
        preview_hint.place(relx=0.5, rely=0.5, anchor='center')
        # 播放时用来“装”静音画面的小窗口：合成模式下它只压在黄框那一块，
        # 外面一圈留给背景图，这样软件里看到的和扩展屏上投出来的是一个样式。
        preview_video_host = tk.Frame(preview_host, bg='#000000',
                                      highlightthickness=0, bd=0)

        list_box = tk.Frame(left, bg=p['bg'])
        list_box.pack(fill=tk.X, pady=(8, 0))
        tree = ttk.Treeview(list_box, columns=('no', 'name'), show='headings',
                            height=6)
        tree.heading('no', text='顺序')
        tree.heading('name', text='视频文件')
        tree.column('no', width=54, anchor='center', stretch=False)
        tree.column('name', width=520, anchor='w', stretch=True)
        scroll = ttk.Scrollbar(list_box, orient=tk.VERTICAL,
                               command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 投屏版式：整屏播放 / 合成投屏（背景图 + 视频框任意摆放）
        g_layout = ttk.LabelFrame(left, text=' 投屏版式 ', padding=(10, 8))
        g_layout.pack(fill=tk.X, pady=(10, 0))
        row = tk.Frame(g_layout, bg=p['bg'])
        row.pack(fill=tk.X)
        video_layout_var = tk.StringVar(value='full')
        ttk.Radiobutton(row, text='整屏播放', value='full',
                        variable=video_layout_var,
                        command=lambda: _sync_video_layout()).pack(side=tk.LEFT)
        ttk.Radiobutton(row, text='合成（背景+视频框）', value='compose',
                        variable=video_layout_var,
                        command=lambda: _sync_video_layout()).pack(
            side=tk.LEFT, padx=(10, 0))
        # 背景图和视频框参数都只属于“合成投屏”：整屏播放时整组藏起来
        compose_tools = tk.Frame(g_layout, bg=p['bg'])
        row = tk.Frame(compose_tools, bg=p['bg'])
        row.pack(fill=tk.X, pady=(6, 0))
        tk.Label(row, text='背景图', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        video_bg_var = tk.StringVar(value='（不用图）')
        video_bg_pick = ttk.Combobox(row, state='readonly', width=16,
                                     textvariable=video_bg_var,
                                     values=['（不用图）'])
        video_bg_pick.pack(side=tk.LEFT, padx=(6, 0))
        video_bg_pick.bind('<<ComboboxSelected>>',
                           lambda e: _sync_video_layout())
        video_bg_upload_btn = ttk.Button(row, text='上传…', style='Ghost.TButton',
                                         command=lambda: _upload_video_bg())
        video_bg_upload_btn.pack(side=tk.LEFT, padx=(6, 0))
        video_bg_mgr_btn = ttk.Button(
            row, text='管理…', style='Ghost.TButton',
            command=lambda: self._open_bg_manager(
                win, on_change=lambda: _sync_video_bg_list()))
        video_bg_mgr_btn.pack(side=tk.LEFT, padx=(6, 0))
        # 左中右版式：两侧各铺一张图，视频在中间那块
        row = tk.Frame(compose_tools, bg=p['bg'])
        row.pack(fill=tk.X, pady=(6, 0))
        tk.Label(row, text='两侧图', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        tk.Label(row, text='左', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT, padx=(6, 0))
        video_left_var = tk.StringVar(value='（不用图）')
        video_left_pick = ttk.Combobox(row, state='readonly', width=11,
                                       textvariable=video_left_var,
                                       values=['（不用图）'])
        video_left_pick.pack(side=tk.LEFT, padx=(4, 0))
        tk.Label(row, text='右', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT, padx=(6, 0))
        video_right_var = tk.StringVar(value='（不用图）')
        video_right_pick = ttk.Combobox(row, state='readonly', width=11,
                                        textvariable=video_right_var,
                                        values=['（不用图）'])
        video_right_pick.pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(row, text='上传…', style='Ghost.TButton',
                   command=lambda: _upload_video_side()).pack(side=tk.LEFT,
                                                              padx=(6, 0))
        ttk.Button(row, text='管理…', style='Ghost.TButton',
                   command=lambda: self._open_bg_manager(
                       win, on_change=lambda: _sync_video_bg_list())).pack(
            side=tk.LEFT, padx=(6, 0))
        row = tk.Frame(compose_tools, bg=p['bg'])
        row.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(row, text='视频框 16:9 居中', style='Ghost.TButton',
                   command=lambda: _video_box_center()).pack(side=tk.LEFT)
        ttk.Button(row, text='左右对称', style='Ghost.TButton',
                   command=lambda: _video_box_symmetric()).pack(
            side=tk.LEFT, padx=(6, 0))
        ttk.Button(row, text='铺满', style='Ghost.TButton',
                   command=lambda: _video_box_full()).pack(side=tk.LEFT,
                                                           padx=(6, 0))
        compose_tools.pack(fill=tk.X)
        video_layout_hint = tk.Label(
            g_layout, text='合成模式：背景图铺满大屏，视频放在黄框里；'
                           '播放时不会出现播放器标题栏/操作栏',
            font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
            justify='left', wraplength=560)
        video_layout_hint.pack(anchor='w', pady=(6, 0))
        frame_hint = tk.Label(g_layout, text='选中视频会自动取一帧贴到预览里',
                              font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                              justify='left', wraplength=560)
        frame_hint.pack(anchor='w', pady=(2, 0))

        right = tk.Frame(mid, bg=p['bg'])
        right.pack(side=tk.LEFT, fill=tk.Y, padx=(14, 0))
        ttk.Label(right, text='播放位置', font=ui_theme.font(9),
                  foreground=p['muted']).pack(anchor='w')
        monitors = video_player.list_monitors()
        default_index = 0
        for item in monitors:
            if not item['primary']:
                default_index = item['index']
                break
        screen_var = tk.StringVar(
            value=monitors[default_index]['name'] if monitors else '显示器 1')
        screen_box = ttk.Combobox(
            right, state='readonly', width=22, textvariable=screen_var,
            values=[m['name'] for m in monitors] or ['显示器 1'])
        screen_box.pack(anchor='w', pady=(4, 2))
        tk.Label(right, text='播放时会把画面全屏到这台显示器；\n'
                             '只有一个屏幕时会播在主屏上。',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 justify='left').pack(anchor='w', pady=(0, 10))

        # 播放方式放在面板靠上的位置：下拉列表往下展开，不会翻到上面去
        tk.Label(right, text='播放方式', font=ui_theme.font(9),
                 fg=p['muted']).pack(anchor='w')
        backend_names = {'自动选择（推荐）': 'auto',
                         'VLC 内嵌（画面固定，推荐）': 'libvlc',
                         'VLC（解码最全）': 'vlc'}
        backend_var = tk.StringVar(value='自动选择（推荐）')
        backend_box = ttk.Combobox(right, state='readonly', width=22,
                                   textvariable=backend_var,
                                   values=list(backend_names))
        backend_box.pack(anchor='w', pady=(4, 2))
        backend_label = tk.Label(right, text=player.describe(),
                                 font=ui_theme.font(8), bg=p['bg'],
                                 fg=p['faint'], anchor='w', justify='left',
                                 wraplength=190)
        backend_label.pack(anchor='w', pady=(0, 8))

        loop_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text='循环播放（一遍放完从头再来）',
                        variable=loop_var).pack(anchor='w')
        top_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text='播放时窗口置顶',
                        variable=top_var).pack(anchor='w', pady=(4, 10))

        preview_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text='软件内预览（静音）',
                        variable=preview_var).pack(anchor='w')
        tk.Label(right, text='画面投到扩展屏时，左边黑框里\n跟着放一份静音预览。',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 justify='left').pack(anchor='w', pady=(2, 0))

        tk.Label(right, text='音量', font=ui_theme.font(9),
                 fg=p['muted']).pack(anchor='w')
        volume = ui_theme.SeekBar(right, from_=0, to=100, value=100,
                                  width=180,
                                  command=lambda v: player.set_volume(int(v)))
        volume.pack(anchor='w', pady=(2, 0))

        # 声音输出：有些电脑的默认设备是没接音箱的 USB 声卡 / 显示器，
        # 画面在放却一点声音都没有，这里可以指定用哪一台设备出声。
        tk.Label(right, text='声音输出', font=ui_theme.font(9),
                 fg=p['muted']).pack(anchor='w', pady=(12, 0))
        audio_items = video_player.list_audio_devices()
        follow_label = '跟随系统默认（推荐）'
        audio_var = tk.StringVar(value=follow_label)
        audio_box = ttk.Combobox(
            right, state='readonly', width=22, cursor='arrow',
            textvariable=audio_var,
            values=[follow_label] + [item['name'] for item in audio_items])
        audio_box.pack(anchor='w', pady=(4, 2))

        def _audio_id():
            """下拉框当前选的是哪台设备；“跟随系统默认”返回 None。"""
            name = audio_var.get()
            for item in audio_items:
                if item['name'] == name:
                    return item['id']
            return None

        audio_hint = tk.Label(right, text='没声音就换一台设备，再点“试听”',
                              font=ui_theme.font(8), bg=p['bg'],
                              fg=p['faint'], anchor='w', justify='left',
                              wraplength=190)
        audio_hint.pack(anchor='w')

        def _on_audio(_event=None):
            state['audio_device'] = _audio_id()
            audio_hint.configure(text='将用：%s' % audio_var.get(),
                                 fg=p['muted'])

        audio_box.bind('<<ComboboxSelected>>', _on_audio)

        def _test_audio():
            ok, note = player.test_sound(_audio_id(), int(volume.get()))
            audio_hint.configure(text=note,
                                 fg=p['success'] if ok else p['danger'])

        ttk.Button(right, text='试听声音', style='Ghost.TButton',
                   command=_test_audio).pack(anchor='w', pady=(4, 0))

        # 系统当前默认是哪台设备：后台问一句，主线程更新提示（查不到就算了）
        audio_probe = {}

        def _ask_default_audio():
            try:
                audio_probe['id'] = video_player.default_audio_device_id()
            except Exception:
                audio_probe['id'] = None

        threading.Thread(target=_ask_default_audio, daemon=True).start()

        def _poll_audio_probe():
            if 'id' not in audio_probe:
                try:
                    win.after(400, _poll_audio_probe)
                except tk.TclError:
                    pass
                return
            dev_id = audio_probe.get('id')
            name = ''
            for item in audio_items:
                if dev_id and item['id'].lower() == dev_id.lower():
                    name = item['name']
                    break
            try:
                if name and state.get('audio_device') is None:
                    audio_hint.configure(
                        text='系统当前默认：%s\n没声音就换一台再点“试听”' % name)
            except tk.TclError:
                pass

        win.after(400, _poll_audio_probe)

        # VLC 没装/不可用时的提示 + 自动静默安装（失败时给重试按钮）
        vlc_hint = tk.Label(right, text='', font=ui_theme.font(8), bg=p['bg'],
                            fg=p['faint'], anchor='w', justify='left',
                            wraplength=190)
        vlc_install_btn = ttk.Button(right, text='重试安装 VLC',
                                     style='Ghost.TButton',
                                     command=lambda: _start_install())

        def _sync_vlc_ui():
            try:
                if video_player.find_vlc() is not None:
                    vlc_install_btn.pack_forget()
                    vlc_hint.pack_forget()
                    return
                ph = install_state.get('phase')
                if ph == 'running':
                    vlc_install_btn.pack_forget()
                    vlc_hint.pack(anchor='w', pady=(4, 0))
                    return
                if ph == 'done-fail':
                    vlc_hint.pack(anchor='w', pady=(4, 0))
                    vlc_install_btn.pack(anchor='w', pady=(4, 0))
                    return
                # 没装且没在装：自动开始静默安装（启动时也会装，双重入口有锁）
                vlc_install_btn.pack_forget()
                vlc_hint.config(text='视频完整功能需要 VLC，正在自动安装…')
                vlc_hint.pack(anchor='w', pady=(4, 0))
                _start_install()
            except tk.TclError:
                pass

        def _on_backend(_event=None):
            player.set_backend(backend_names.get(backend_var.get(), 'auto'))
            _sync_backend_label()
            _sync_seek_ui()

        def _sync_backend_label():
            """当前播放方式说明。"""
            try:
                backend_label.config(text=player.describe())
            except tk.TclError:
                pass

        backend_box.bind('<<ComboboxSelected>>', _on_backend)

        # ---------- VLC 自动静默安装 ----------
        install_state = {'phase': ''}      # '' 空闲 / running / done-ok / done-fail

        def _start_install():
            if install_state['phase'] in ('running', 'done-ok'):
                return
            install_state['phase'] = 'running'
            vlc_hint.config(text='正在自动安装 VLC（静默，约 1～3 分钟）…')
            try:
                vlc_install_btn.pack_forget()
                vlc_install_btn.configure(state='disabled')
            except tk.TclError:
                pass

            def work():
                install_state['result'] = video_player.install_vlc()

            threading.Thread(target=work, daemon=True).start()

            def poll():
                if not _alive():
                    return
                if 'result' not in install_state:
                    win.after(1000, poll)
                    return
                ok, note = install_state.pop('result')
                install_state['phase'] = 'done-ok' if ok else 'done-fail'
                try:
                    vlc_install_btn.configure(state='normal')
                except tk.TclError:
                    pass
                if ok:
                    video_player.reset_vlc_cache()
                    player.set_backend('auto')
                    _sync_backend_label()
                    _sync_seek_ui()
                    try:
                        preview_hint.configure(
                            text='选好视频后这里显示预览画面（静音）')
                    except tk.TclError:
                        pass
                    # 装好 VLC 后把已经选好的视频重新取一帧，预览马上有画面
                    if state['list'] and not state.get('playing'):
                        _grab_preview_frame(_pick_index(), force=True)
                vlc_hint.config(text='' if ok else note)
                _sync_vlc_ui()

            win.after(1000, poll)

        # 列表操作
        # 这一行紧跟在视频列表下面（见左栏的排版），高度先占好，
        # 窗口内容多了也不会被挤出去
        tools = tk.Frame(left, bg=p['bg'])
        ttk.Button(tools, text='添加视频…', style='Accent.TButton',
                   command=lambda: _add()).pack(side=tk.LEFT)
        ttk.Button(tools, text='上移', style='Ghost.TButton',
                   command=lambda: _move(-1)).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(tools, text='下移', style='Ghost.TButton',
                   command=lambda: _move(1)).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(tools, text='移除', style='Ghost.TButton',
                   command=lambda: _remove()).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(tools, text='清空', style='Ghost.TButton',
                   command=lambda: _clear()).pack(side=tk.LEFT, padx=(6, 0))

        # 播放控制
        ctrl = tk.Frame(body, bg=p['bg'])
        ctrl.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 10))
        # 进度条：拖动就能跳到视频任意位置，不用从头再放一遍
        seek_row = tk.Frame(ctrl, bg=p['bg'])
        seek_row.pack(anchor='center', pady=(0, 6))
        seek = ui_theme.SeekBar(seek_row, from_=0, to=1, value=0, width=420,
                                command=lambda v: _on_seek_drag(v),
                                on_seek_start=lambda: _on_seek_start(),
                                on_seek_end=lambda: _on_seek_end())
        seek.pack(side=tk.LEFT)
        # 固定宽度：时间长短变化时，居中的这一行不会左右乱晃
        seek_label = tk.Label(seek_row, text='--:-- / --:--',
                              font=ui_theme.font(9), bg=p['bg'],
                              fg=p['muted'], width=15, anchor='w')
        seek_label.pack(side=tk.LEFT, padx=(10, 0))
        # 这一行摆在窗口中间：按钮加进度提示合成一组，整组居中
        ctrl_row = tk.Frame(ctrl, bg=p['bg'])
        ctrl_row.pack(anchor='center')
        play_btn = ttk.Button(ctrl_row, text='开始播放', style='Accent.TButton',
                              command=lambda: _start())
        play_btn.pack(side=tk.LEFT)
        pause_btn = ttk.Button(ctrl_row, text='暂停', style='Ghost.TButton',
                               command=lambda: _pause_toggle())
        pause_btn.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(ctrl_row, text='停止', style='Danger.TButton',
                   command=lambda: _stop()).pack(side=tk.LEFT, padx=(6, 0))
        # 固定宽度：进度文字长短变化时，居中的这一行不会左右乱晃
        info_label = tk.Label(ctrl_row, text='', font=ui_theme.font(9),
                              bg=p['bg'], fg=p['muted'], width=38,
                              anchor='w')
        info_label.pack(side=tk.LEFT, padx=(14, 0))
        tk.Label(body,
                 text='提示：用电脑上装的 VLC 播放（各种格式都能解）；'
                      '双击列表项可从该视频开始，播放中可拖进度条跳位置。',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 anchor='w', justify='left').pack(side=tk.BOTTOM, fill=tk.X,
                                                  pady=(0, 8))

        # 左栏：先给“固定高度”的列表、列表按钮行和版式区从下往上占好位，
        # 预览区自适应剩余空间，这样窗口变小也不会把下面的按钮挤没。
        # 顺序（从上到下）：预览画面 → 视频列表 → 列表按钮 → 投屏版式
        for _widget in (preview_host, list_box, tools, g_layout):
            _widget.pack_forget()
        g_layout.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))
        tools.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        list_box.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        preview_host.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # ---------- 内部函数 ----------
        def _alive():
            try:
                return bool(win.winfo_exists())
            except tk.TclError:
                return False

        def _fmt(sec):
            sec = max(0, int(sec or 0))
            return '%02d:%02d' % (sec // 60, sec % 60)

        def _render(select=None):
            tree.delete(*tree.get_children())
            for i, path in enumerate(state['list']):
                tree.insert('', 'end', iid=str(i),
                            values=(i + 1, os.path.basename(path)))
            if select is not None and 0 <= select < len(state['list']):
                tree.selection_set(str(select))
                tree.see(str(select))

        def _selected():
            sel = tree.selection()
            return int(sel[0]) if sel else -1

        def _queue_preview_grab(index, delay=200):
            """挑好视频后过一下再取帧：连点几个只取最后选中的那个。"""
            job = state.get('pick_job')
            if job is not None:
                try:
                    win.after_cancel(job)
                except Exception:
                    pass
                state['pick_job'] = None
            if index is None or index < 0:
                return
            state['pick_job'] = win.after(
                delay, lambda idx=index: _grab_preview_frame(idx))

        def _on_pick(_event=None):
            """列表里点到哪个视频，就先取一帧放到预览里（还没开始播放时）。"""
            if state.get('playing'):
                return
            i = _selected()
            # 换了视频：进度条回起点，别把上一个视频的位置带过来
            if 0 <= i < len(state['list']) \
                    and state.get('start_at_for') != state['list'][i]:
                _reset_seek()
            _queue_preview_grab(i)

        def _add():
            paths = filedialog.askopenfilenames(
                parent=win, title='选择视频文件（可多选）',
                filetypes=[('视频文件', '*.mp4 *.wmv *.avi *.mov *.mkv '
                                        '*.flv *.m4v *.mpg *.mpeg *.ts '
                                        '*.webm'),
                           ('所有文件', '*.*')])
            if not paths:
                return
            added = 0
            for path in paths:
                full = os.path.abspath(path)
                if full not in state['list']:
                    state['list'].append(full)
                    added += 1
            _render(state['list'].index(state['list'][-1]) if added
                    else None)
            state_label.config(text='共 %d 个视频' % len(state['list']))
            # 加了视频就顺手取一帧：还没点“开始播放”也能在预览里看到画面
            if added:
                state['frame_pic'] = None
                state['frame_for'] = None
                _queue_preview_grab(len(state['list']) - 1, delay=30)

        def _move(delta):
            i = _selected()
            if i < 0:
                return
            j = i + delta
            if not (0 <= j < len(state['list'])):
                return
            state['list'][i], state['list'][j] = state['list'][j], state['list'][i]
            if state['idx'] == i:
                state['idx'] = j
            elif state['idx'] == j:
                state['idx'] = i
            _render(j)

        def _remove():
            i = _selected()
            if i < 0:
                return
            if i == state['idx']:
                _stop()
            gone = state['list'][i]
            del state['list'][i]
            if state['idx'] > i:
                state['idx'] -= 1
            _render(min(i, len(state['list']) - 1) if state['list'] else None)
            state_label.config(text='共 %d 个视频' % len(state['list']))
            if state.get('frame_for') == gone:      # 删掉的正好是预览那一个
                state['frame_pic'] = None
                state['frame_for'] = None
                _queue_preview_grab(min(i, len(state['list']) - 1)
                                    if state['list'] else None)
                _draw_video_box()

        def _clear():
            _stop()
            state['list'] = []
            state['frame_pic'] = None
            state['frame_for'] = None
            _reset_seek()
            _render()
            state_label.config(text='未播放')

        def _monitor_index():
            """下拉框选的显示器序号（插拔投影仪后按“主屏/扩展屏”重新对）。"""
            return video_player.resolve_monitor_index(screen_var.get())

        def _refresh_screen_box(_event=None):
            """投影仪是后插的也没关系：重新对一遍，别默默投回主屏。"""
            items = video_player.list_monitors()
            names = [m['name'] for m in items] or ['显示器 1']
            if list(screen_box.cget('values')) != names:
                index = video_player.resolve_monitor_index(screen_var.get(),
                                                           items)
                screen_box.configure(values=names)
                screen_var.set(items[index]['name'])

        screen_box.bind('<Button-1>', _refresh_screen_box, add='+')
        win.bind('<FocusIn>', _refresh_screen_box, add='+')

        def _want_preview():
            """要不要预览：勾了开关，而且正片不是投在主屏上。"""
            if not preview_var.get():
                return False
            monitors_now = video_player.list_monitors()
            index = max(0, min(_monitor_index(), len(monitors_now) - 1))
            return len(monitors_now) > 1 and not monitors_now[index]['primary']

        # ---------- 投屏版式：整屏 / 合成（背景图 + 视频框任意摆放） ----------
        def _target_item():
            items = video_player.list_monitors()
            index = max(0, min(_monitor_index(), len(items) - 1))
            return items[index]

        def _target_rect():
            return _target_item()['rect']

        def _video_bg_items():
            try:
                import ppt_converter
                return ppt_converter.list_uploaded_backgrounds()
            except Exception:
                return []

        def _sync_video_bg_list(select=None):
            names = ['（不用图）'] + [x['name'] for x in _video_bg_items()]
            try:
                for combo in (video_bg_pick, video_left_pick, video_right_pick):
                    combo.configure(values=names)
            except tk.TclError:
                return
            if select and select in names:
                video_bg_var.set(select)
            for var in (video_bg_var, video_left_var, video_right_var):
                if var.get() not in names:
                    var.set(names[0])

        def _video_bg_image():
            name = video_bg_var.get()
            if not name or name == '（不用图）':
                return None
            for item in _video_bg_items():
                if item['name'] == name:
                    return strip_screen.load_image(item['image'])
            return None

        def _video_side_image(which):
            """左侧图 / 右侧图；没选返回 None。"""
            name = (video_left_var if which == 'left' else video_right_var).get()
            if not name or name == '（不用图）':
                return None
            for item in _video_bg_items():
                if item['name'] == name:
                    return strip_screen.load_image(item['image'])
            return None

        def _video_box_local(width, height):
            """视频框在整屏里的像素矩形（左上角 + 宽高）。"""
            x, y, w, h = state['video_box']
            return (max(0, int(round(width * x))),
                    max(0, int(round(height * y))),
                    max(2, int(round(width * w))),
                    max(2, int(round(height * h))))

        def _video_base_image(width, height):
            """大屏底图：背景图铺满整屏 + 两侧图铺在视频框左右两边。

            左中右版式就是靠这块底图实现的：左边一张图、右边一张图，
            中间留出来给视频。两侧的宽度 = 视频框左右两边剩下的宽度，
            所以拖视频框的左右边（或点“左右对称”）就能调两侧的宽窄。
            """
            from PIL import Image
            bg = _video_bg_image()
            if bg is not None:
                base = strip_screen.cover_crop(bg, width, height)
            else:
                base = Image.new('RGB', (max(1, width), max(1, height)),
                                 '#0b1a33')
            x0, _y0, box_w, _box_h = _video_box_local(width, height)
            x1 = min(width, x0 + box_w)
            left_img = _video_side_image('left')
            right_img = _video_side_image('right')
            if left_img is not None and x0 > 2:
                base.paste(strip_screen.cover_crop(left_img, x0, height),
                           (0, 0))
            if right_img is not None and width - x1 > 2:
                base.paste(
                    strip_screen.cover_crop(right_img, width - x1, height),
                    (x1, 0))
            return base

        def _upload_video_bg():
            import ppt_converter
            src = filedialog.askopenfilename(
                parent=win, title='选择整屏背景图',
                filetypes=[('图片文件', '*.png *.jpg *.jpeg *.bmp *.webp'),
                           ('所有文件', '*.*')])
            if not src:
                return
            stem = os.path.splitext(os.path.basename(src))[0]
            name = simpledialog.askstring('背景名称', '给这张背景起个名字：',
                                          initialvalue=stem, parent=win)
            if not name:
                return
            try:
                item = ppt_converter.add_uploaded_background(src, name)
            except Exception as e:
                messagebox.showerror('上传失败', str(e), parent=win)
                return
            _sync_video_bg_list(item['name'])
            state_label.config(text='背景图已保存：%s' % name)
            _sync_video_layout()

        def _upload_video_side():
            """上传一张两侧用的图：空着的那一边先放，都有图就替换左边。"""
            import ppt_converter
            src = filedialog.askopenfilename(
                parent=win, title='选择两侧用的图片',
                filetypes=[('图片文件', '*.png *.jpg *.jpeg *.bmp *.webp'),
                           ('所有文件', '*.*')])
            if not src:
                return
            stem = os.path.splitext(os.path.basename(src))[0]
            name = simpledialog.askstring('图片名称', '给这张图起个名字：',
                                          initialvalue=stem, parent=win)
            if not name:
                return
            try:
                item = ppt_converter.add_uploaded_background(src, name)
            except Exception as e:
                messagebox.showerror('上传失败', str(e), parent=win)
                return
            _sync_video_bg_list()
            if video_left_var.get() in ('', '（不用图）'):
                video_left_var.set(item['name'])
                side = '左'
            elif video_right_var.get() in ('', '（不用图）'):
                video_right_var.set(item['name'])
                side = '右'
            else:
                video_left_var.set(item['name'])
                side = '左'
            state_label.config(text='已把「%s」放到%s侧图' % (name, side))
            _sync_video_layout()

        def _box_note(text):
            """把刚应用的视频框参数写在提示行上，点完能看出变化。"""
            try:
                frame_hint.config(text=text)
            except tk.TclError:
                pass

        def _video_box_center():
            """视频框：16:9 居中，占屏幕高度 90%。"""
            rect = _target_rect()
            width = max(2, rect[2] - rect[0])
            height = max(2, rect[3] - rect[1])
            box_h = 0.9
            box_w = min(1.0, (height * box_h) * 16.0 / 9.0 / float(width))
            state['video_box'] = [(1.0 - box_w) / 2.0, (1.0 - box_h) / 2.0,
                                  box_w, box_h]
            _sync_video_layout()
            _box_note('视频框：16:9 居中，占屏幕高度 90%')

        def _video_box_symmetric():
            """左右对称：视频框保持 16:9，左右留边一样宽。

            16:9 的屏幕和视频同比例，“铺满高度”就等于“铺满宽度”，
            看不出左右留边；所以这种情况下按 16:9 缩到 80% 居中，
            左右各留 10%，点一下就看得出效果。更宽的屏幕（长条屏、
            超宽 LED）则铺满高度，左右自动等宽留边。
            """
            rect = _target_rect()
            width = max(2, rect[2] - rect[0])
            height = max(2, rect[3] - rect[1])
            # 宽度最多占 80%（左右各留 10% 的边），高度按 16:9 算，不超出屏幕
            box_w = min(0.8, (height * 16.0 / 9.0) / float(width))
            box_h = box_w * width * 9.0 / 16.0 / float(height)
            state['video_box'] = [(1.0 - box_w) / 2.0, (1.0 - box_h) / 2.0,
                                  box_w, box_h]
            _sync_video_layout()
            _box_note('视频框：左右对称，左右各留 %d%% 的边（保持 16:9）'
                      % int(round((1.0 - box_w) * 50.0)))

        def _video_box_full():
            state['video_box'] = [0.0, 0.0, 1.0, 1.0]
            _sync_video_layout()
            _box_note('视频框：铺满整屏')

        def _video_box_pixels():
            """视频框在预览画布里的像素矩形。"""
            ox, oy, tw, th = _video_view()
            x, y, w, h = state['video_box']
            return (ox + tw * x, oy + th * y,
                    ox + tw * (x + w), oy + th * (y + h))

        def _video_view():
            """预览底图（合成效果图）在画布里的位置和大小。"""
            view = state.get('box_view')
            if view:
                return view
            return (0.0, 0.0, max(2, preview_host.winfo_width()),
                    max(2, preview_host.winfo_height()))

        def _preview_video_rect():
            """软件内预览的画面占预览区的哪一块：
            合成模式=黄框那一块（外面留背景图），整屏模式=整块预览区。"""
            if video_layout_var.get() == 'compose':
                return _video_box_pixels()
            return (0.0, 0.0, float(max(2, preview_host.winfo_width())),
                    float(max(2, preview_host.winfo_height())))

        def _place_preview_video(running=None):
            """把内嵌的静音画面挪到该在的位置；没在放就收起来。"""
            if running is None:
                running = bool(getattr(preview_ctl, 'running', False))
            try:
                if not running:
                    preview_video_host.place_forget()
                    return
                x0, y0, x1, y1 = _preview_video_rect()
                preview_video_host.place(
                    x=int(round(x0)), y=int(round(y0)),
                    width=max(2, int(round(x1 - x0))),
                    height=max(2, int(round(y1 - y0))))
                preview_video_host.lift()
            except tk.TclError:
                pass

        # ---------- 取一帧视频画面，当版式预览用 ----------
        grab_state = {}

        def _grab_window():
            """抓帧时给 VLC 用的窗口：放到屏幕外，用户看不到。"""
            w = state.get('grab_win')
            if w is None or not w.winfo_exists():
                w = tk.Toplevel(self.root)
                w.overrideredirect(True)
                w.geometry('32x32+-4000+-4000')
                w.configure(bg='#000000')
                state['grab_win'] = w
            return w

        def _pick_index():
            """该给哪个视频取预览：优先列表里正选中的那个。"""
            if not state['list']:
                return -1
            index = _selected()
            if index < 0:
                index = state['idx'] if state['idx'] >= 0 else 0
            return max(0, min(index, len(state['list']) - 1))

        def _grab_preview_frame(index=None, force=False):
            """静音放一下这个视频、抓一帧，贴到预览里当画面。

            整屏播放时贴满预览区，合成投屏时贴进黄框；这样还没点“开始播放”
            就能看到视频长什么样，不再是黑框。
            """
            if not state['list']:
                return
            if index is None:
                index = _pick_index()
            index = max(0, min(int(index), len(state['list']) - 1))
            src = state['list'][index]
            if not getattr(player, 'embedded', False):
                # 没装 VLC 就取不到帧，别硬等，直接在预览里说清楚
                try:
                    preview_hint.config(
                        text='没装 VLC，暂时取不到预览画面\n文件：%s'
                             % os.path.basename(src))
                except tk.TclError:
                    pass
                _set_preview_hint(True)
                return
            if grab_state.get('busy'):
                if grab_state.get('src') == src:
                    return                  # 正在取的就是它，别重复排队
                grab_state['next'] = index  # 记下最新一次选择，这次取完接着取
                return
            if not force and state.get('frame_for') == src \
                    and state.get('frame_pic') is not None:
                _draw_video_box()           # 还是这个视频，直接重画
                return
            out = os.path.join(tempfile.gettempdir(), 'meeting_frame.png')
            grab_state.update(busy=True, result=None, out=out, src=src)
            try:
                frame_hint.config(text='正在取一帧画面…')
            except tk.TclError:
                pass
            host = _grab_window().winfo_id()

            def work():
                grab_state['result'] = video_player.grab_frame(
                    src, out, host_hwnd=host)
                grab_state['busy'] = False

            threading.Thread(target=work, daemon=True).start()
            win.after(500, _poll_grab)

        def _poll_grab():
            if not _alive():
                return
            if grab_state.get('busy'):
                win.after(500, _poll_grab)
                return
            if grab_state.get('result'):
                try:
                    from PIL import Image
                    state['frame_pic'] = Image.open(
                        grab_state['out']).convert('RGB')
                    state['frame_for'] = grab_state.get('src')
                    if video_layout_var.get() == 'compose':
                        frame_hint.config(text='预览用的是这一帧画面；'
                                               '拖动黄框可改视频位置和大小')
                    else:
                        frame_hint.config(text='预览用的是这一帧画面；'
                                               '点“开始播放”就投到大屏')
                except Exception:
                    state['frame_pic'] = None
                    state['frame_for'] = None
                    frame_hint.config(text='取帧成功，但读图失败')
            else:
                state['frame_pic'] = None
                state['frame_for'] = None
                frame_hint.config(text='没能取到画面（这种格式取不到，'
                                       '播放本身没问题）')
                try:
                    preview_hint.config(
                        text='这个视频取不到预览画面\n文件：%s'
                             % os.path.basename(grab_state.get('src') or ''))
                except tk.TclError:
                    pass
            _draw_video_box()
            nxt = grab_state.pop('next', None)
            if nxt is not None and _alive():
                _grab_preview_frame(nxt)

        def _draw_video_box():
            """没在播放时：把预览画面画出来（整屏=视频画面本身，
            合成=背景图 + 视频位置），再叠上可拖动的黄框。"""
            _draw_preview_bg()
            _draw_box_overlay()
            _set_preview_hint()
            # 播放中：内嵌的静音画面要一直贴着黄框（窗口大小一变就得跟着挪）
            _place_preview_video()

        def _draw_preview_bg():
            """预览画面：整屏播放=视频画面本身；合成=背景图铺满 + 视频位置。"""
            try:
                preview_host.delete('vbg')
            except tk.TclError:
                return
            state['box_view'] = None
            state['layout_thumb'] = None
            pic = state.get('frame_pic')
            playing = bool(state.get('playing'))
            if video_layout_var.get() != 'compose':
                # 整屏播放：直接把取到的那一帧等比居中画在预览里
                if playing:
                    return          # 播放中整块预览区交给内嵌的播放器
                if pic is None:
                    return
                try:
                    from PIL import ImageTk
                    cw = max(80, preview_host.winfo_width())
                    ch = max(60, preview_host.winfo_height())
                    s = min(cw / float(max(1, pic.width)),
                            ch / float(max(1, pic.height)))
                    tw = max(1, int(pic.width * s))
                    th = max(1, int(pic.height * s))
                    photo = ImageTk.PhotoImage(pic.resize((tw, th)))
                    state['layout_thumb'] = photo      # 必须留引用
                    preview_host.create_image((cw - tw) / 2.0, (ch - th) / 2.0,
                                              image=photo, anchor='nw',
                                              tags=('vbg',))
                except Exception:
                    state['layout_thumb'] = None
                return
            # 合成模式：播放中也要把背景图铺在预览里（视频只盖住黄框那块），
            # 这样主屏上看到的和扩展屏上投出来的是同一个样式
            try:
                rect = _target_rect()
                width = max(2, rect[2] - rect[0])
                height = max(2, rect[3] - rect[1])
                frame = _video_base_image(width, height)
                from PIL import ImageDraw, ImageTk
                bx, by, bw, bh = _video_box_local(width, height)
                x1, y1 = bx + bw, by + bh
                # 视频框那块先留黑（真播放时 VLC 也是黑底 + 等比画面）
                ImageDraw.Draw(frame).rectangle([bx, by, x1, y1],
                                                fill='#000000')
                if pic is not None and not playing:
                    # 有截图就贴截图（等比留边，和真实播放一致）
                    try:
                        fitted = strip_screen.fit_inside(
                            pic.size, (bx, by, x1, y1))
                        small = pic.resize(fitted)
                        frame.paste(small, (bx + (bw - fitted[0]) // 2,
                                            by + (bh - fitted[1]) // 2))
                    except Exception:
                        pass
                cw = max(80, preview_host.winfo_width())
                ch = max(60, preview_host.winfo_height())
                scale = min(cw / float(frame.width), ch / float(frame.height))
                tw = max(1, int(frame.width * scale))
                th = max(1, int(frame.height * scale))
                photo = ImageTk.PhotoImage(frame.resize((tw, th)))
                state['layout_thumb'] = photo      # 必须留引用
                ox = (cw - tw) / 2.0
                oy = (ch - th) / 2.0
                preview_host.create_image(ox, oy, image=photo, anchor='nw',
                                          tags=('vbg',))
                state['box_view'] = (ox, oy, tw, th)
            except Exception:
                state['box_view'] = None

        def _draw_box_overlay():
            """可拖动的黄框（视频位置）。

            播放中也要保留：视频刚加载那一下还没出画面，黄框在着才看得清位置；
            视频出画面后（子窗口）自然会盖住它。
            """
            try:
                preview_host.delete('vbox')
            except tk.TclError:
                return
            if video_layout_var.get() != 'compose':
                return
            x0, y0, x1, y1 = _video_box_pixels()
            preview_host.create_rectangle(x0, y0, x1, y1, outline='#ffd54a',
                                          width=2, dash=(6, 4), tags=('vbox',))
            # 四个角 + 四条边的中点都能拖：边上的点用来单独调左/右/上/下
            for hx, hy in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)):
                preview_host.create_rectangle(hx - 6, hy - 6, hx + 6, hy + 6,
                                              fill='#ffd54a', outline='',
                                              tags=('vbox',))
            for hx, hy in (((x0 + x1) / 2.0, y0), ((x0 + x1) / 2.0, y1),
                           (x0, (y0 + y1) / 2.0), (x1, (y0 + y1) / 2.0)):
                preview_host.create_rectangle(hx - 4, hy - 4, hx + 4, hy + 4,
                                              fill='#ffd54a', outline='',
                                              tags=('vbox',))
            preview_host.create_text((x0 + x1) / 2, (y0 + y1) / 2,
                                     text='视频放这里', fill='#ffd54a',
                                     font=ui_theme.font(9), tags=('vbox',))
            # 立刻刷一次，别等到下一次空闲才显示（有时要“点一下”才出来的原因）
            try:
                preview_host.update_idletasks()
            except tk.TclError:
                pass

        box_job = {'job': None}

        def _schedule_box_refresh():
            """拖动时只让黄框跟手，效果图按 90ms 节流重算，避免拖起来卡。"""
            if box_job['job'] is not None:
                return

            def run():
                box_job['job'] = None
                _draw_preview_bg()
            box_job['job'] = win.after(90, run)

        def _sync_video_layout():
            """版式/背景/框变化后刷新界面状态。"""
            compose = video_layout_var.get() == 'compose'
            state['layout'] = video_layout_var.get()
            # 背景图 / 上传 / 视频框参数都只属于“合成投屏”，整屏播放时整组藏起来
            try:
                if compose:
                    compose_tools.pack(fill=tk.X, before=video_layout_hint)
                else:
                    compose_tools.pack_forget()
                for widget, ok_state in ((video_bg_pick, 'readonly'),
                                         (video_bg_upload_btn, 'normal')):
                    widget.configure(state=ok_state if compose else 'disabled')
                frame_hint.config(
                    text=('两侧图铺在黄框左右两边，拖黄框的左右边就能调两侧宽窄；'
                          '选中视频会自动取一帧贴到框里' if compose else
                          '选中视频会自动取一帧贴在预览里'))
            except tk.TclError:
                pass
            _draw_video_box()

        def _on_bg_pick(*_args):
            """选了背景图/两侧图就自动切到“合成”模式。

            不然还停在“整屏播放”，这些图根本用不上，看着像“没效果/没铺满”。
            """
            picked = [var.get() for var in (video_bg_var, video_left_var,
                                            video_right_var)]
            if any(n and n != '（不用图）' for n in picked) and \
                    video_layout_var.get() != 'compose':
                video_layout_var.set('compose')      # 会触发下面的 trace
            else:
                _sync_video_layout()

        # 下拉框选图 / 切版式，都立刻刷新预览（不依赖事件是否送达）
        video_bg_var.trace_add('write', _on_bg_pick)
        video_left_var.trace_add('write', _on_bg_pick)
        video_right_var.trace_add('write', _on_bg_pick)
        video_layout_var.trace_add('write', lambda *a: _sync_video_layout())

        def _video_hit(x, y):
            if video_layout_var.get() != 'compose' or state.get('playing'):
                return None
            x0, y0, x1, y1 = _video_box_pixels()
            for name, (hx, hy) in (('nw', (x0, y0)), ('ne', (x1, y0)),
                                   ('sw', (x0, y1)), ('se', (x1, y1)),
                                   ('n', ((x0 + x1) / 2.0, y0)),
                                   ('s', ((x0 + x1) / 2.0, y1)),
                                   ('w', (x0, (y0 + y1) / 2.0)),
                                   ('e', (x1, (y0 + y1) / 2.0))):
                if abs(x - hx) <= 12 and abs(y - hy) <= 12:
                    return ('resize', name)
            if x0 <= x <= x1 and y0 <= y <= y1:
                return ('move', None)
            return None

        def _video_motion(event):
            hit = _video_hit(event.x, event.y)
            cursor = 'arrow'
            if hit is not None:
                kind, part = hit
                if kind == 'move':
                    cursor = 'fleur'
                elif part in ('nw', 'se'):
                    cursor = 'size_nw_se'
                elif part in ('ne', 'sw'):
                    cursor = 'size_ne_sw'
                elif part in ('w', 'e'):
                    cursor = 'size_we'
                else:
                    cursor = 'size_ns'
            try:
                preview_host.configure(cursor=cursor)
            except tk.TclError:
                pass

        def _video_press(event):
            state['box_drag'] = _video_hit(event.x, event.y)
            state['box_moved'] = False
            if state['box_drag']:
                ox, oy, tw, th = _video_view()
                state['box_origin'] = ((event.x - ox) / float(max(1, tw)),
                                       (event.y - oy) / float(max(1, th)),
                                       list(state['video_box']))
            else:
                state['box_origin'] = None
            if state['box_drag']:
                _video_drag(event)

        def _video_drag(event):
            drag = state.get('box_drag')
            origin = state.get('box_origin')
            if not drag or not origin:
                return
            kind, part = drag
            ox, oy, tw, th = _video_view()
            rx = (event.x - ox) / float(max(1, tw))
            ry = (event.y - oy) / float(max(1, th))
            rx0, ry0, box0 = origin
            left, top, width, height = box0
            if kind == 'move':
                left = min(max(0.0, left + (rx - rx0)), max(0.0, 1.0 - width))
                top = min(max(0.0, top + (ry - ry0)), max(0.0, 1.0 - height))
            else:
                right = left + width
                bottom = top + height
                min_size = 0.04
                if 'w' in part:
                    left = min(max(0.0, rx), right - min_size)
                if 'e' in part:
                    right = max(min(1.0, rx), left + min_size)
                if 'n' in part:
                    top = min(max(0.0, ry), bottom - min_size)
                if 's' in part:
                    bottom = max(min(1.0, ry), top + min_size)
                width = right - left
                height = bottom - top
            box = [left, top, width, height]
            if box != list(box0):
                state['box_moved'] = True
            state['video_box'] = box
            _draw_box_overlay()          # 黄框立刻跟手
            _schedule_box_refresh()      # 效果图节流重算

        def _video_release(_event=None):
            if state.get('box_drag') and state.get('box_moved'):
                _box_note('视频框：已手动调整（四角、四边都能拖，框内拖动改位置）')
            state['box_drag'] = None
            state['box_origin'] = None

        def _video_window_rect():
            """视频框在大屏上的绝对像素矩形。"""
            rect = _target_rect()
            width = max(2, rect[2] - rect[0])
            height = max(2, rect[3] - rect[1])
            x, y, bw, bh = state['video_box']
            left = rect[0] + int(round(width * x))
            top = rect[1] + int(round(height * y))
            return (left, top, left + max(2, int(round(width * bw))),
                    top + max(2, int(round(height * bh))))

        def _show_bg_screen():
            """合成投屏：先在目标屏铺一张“背景 + 视频位置留黑”的底图。"""
            rect = _target_rect()
            width = max(2, rect[2] - rect[0])
            height = max(2, rect[3] - rect[1])
            frame = _video_base_image(width, height)
            from PIL import ImageDraw
            x0, y0, box_w, box_h = _video_box_local(width, height)
            ImageDraw.Draw(frame).rectangle(
                [x0, y0, x0 + box_w, y0 + box_h], fill='#000000')
            screen = state.get('bg_screen')
            if screen is None or not screen.is_open() or \
                    tuple(screen.rect) != tuple(rect):
                if screen is not None:
                    screen.close(notify=False)
                # 背景窗口也要置顶才会被系统画出来；视频窗口在它之后创建，
                # 并且播放期间每 1.5 秒抬一次，保证视频始终盖在背景上面。
                screen = strip_screen.StripScreen(self.root, rect,
                                                  topmost=True,
                                                  keep_lift=False)
                state['bg_screen'] = screen
            screen.show(frame)

        def _close_bg_screen():
            screen = state.get('bg_screen')
            if screen is not None:
                try:
                    screen.close(notify=False)
                except Exception:
                    pass
                state['bg_screen'] = None

        def _projection_rect():
            """播放窗口在大屏上的矩形：整屏=整块屏；合成=视频框。"""
            if video_layout_var.get() == 'compose':
                return _video_window_rect()
            return tuple(_target_rect())

        def _show_compose_screen():
            """大屏合成窗口：一块窗口里背景铺满 + 视频嵌在框里。

            视频是这块窗口的子窗口（VLC 直接画进去），所以 z 序天然正确，
            不需要在两个窗口之间抢层级，也不会出现"背景盖住视频"。
            """
            rect = tuple(_target_rect())
            screen = state.get('compose_screen')
            if screen is None or not screen.is_open() or \
                    tuple(screen.rect) != rect:
                if screen is not None:
                    try:
                        screen.close()
                    except Exception:
                        pass
                screen = strip_screen.ComposeScreen(self.root, rect)
                state['compose_screen'] = screen
            compose = video_layout_var.get() == 'compose'
            if compose:
                screen.show_background(_video_base_image(
                    max(2, rect[2] - rect[0]), max(2, rect[3] - rect[1])))
                screen.set_box(tuple(state['video_box']))
            else:
                screen.show_background(None)
                screen.set_box((0.0, 0.0, 1.0, 1.0))
            try:
                screen.win.attributes('-topmost', bool(top_var.get()))
            except tk.TclError:
                pass
            return screen

        def _close_compose_screen():
            screen = state.get('compose_screen')
            if screen is not None:
                try:
                    screen.close()
                except Exception:
                    pass
                state['compose_screen'] = None

        # ---------- 进度条：拖到哪就从哪播 ----------
        def _seek_seconds():
            """进度条现在指在第几秒。"""
            dur = float(state.get('duration') or 0.0)
            return max(0.0, min(1.0, float(seek.get()))) * dur

        def _seek_target_path():
            if 0 <= state['idx'] < len(state['list']):
                return state['list'][state['idx']]
            i = _selected()
            if 0 <= i < len(state['list']):
                return state['list'][i]
            return ''

        def _on_seek_start():
            state['seek_drag'] = True

        def _on_seek_drag(_value=None):
            """拖动过程中只更新时间文字，松手才真的跳。"""
            dur = float(state.get('duration') or 0.0)
            if dur > 0:
                seek_label.config(text='%s / %s'
                                  % (_fmt(_seek_seconds()), _fmt(dur)))

        def _on_seek_end():
            state['seek_drag'] = False
            dur = float(state.get('duration') or 0.0)
            if dur <= 0:
                return
            want = _seek_seconds()
            state['start_at'] = want                 # 下次也从这里开始
            state['start_at_for'] = _seek_target_path()
            if not state['playing']:
                # 还没播：记下来，点“开始播放”就从这里开始
                info_label.config(
                    text='开始播放将从 %s 继续（拖到最左＝从头播）'
                         % _fmt(want))
                return
            if not player.supports_seek:
                state['pending_seek'] = None
                info_label.config(
                    text='当前用 %s 播放，拖进度不可用；播放方式换成'
                         '“VLC 内嵌”就能拖' % player.backend.name)
                return
            ok = player.seek(want)
            info_label.config(text=('已跳到 %s' % _fmt(want)) if ok
                              else '这个位置跳不过去，换一个位置试试')

        def _reset_seek(clear_start=True):
            """把进度条收回起点（换视频、播完、清空时用）。"""
            state['duration'] = 0.0
            state['pending_seek'] = None
            state['seek_tries'] = 0
            if clear_start:
                state['start_at'] = 0.0
                state['start_at_for'] = ''
            try:
                seek.set(0.0)
                seek_label.config(text='--:-- / --:--')
            except tk.TclError:
                pass

        def _sync_seek_ui():
            """进度条能不能拖，取决于是哪种播放方式。"""
            ok = bool(player.supports_seek)
            try:
                seek.set_enabled(ok)
            except Exception:
                pass
            if state['playing']:
                return
            try:
                seek_label.config(text='--:-- / --:--' if ok
                                  else '换 VLC 内嵌可拖进度')
            except tk.TclError:
                pass

        def _do_play(start_index):
            """把整个列表交给播放器（顺序、循环、全屏都由后端负责）。"""
            if not state['list']:
                return
            if not player.available:
                messagebox.showerror(
                    '播放失败',
                    '没有可用的播放组件。\n\n请在这台电脑上装 VLC：'
                    '程序启动时会自动静默安装（程序目录或 soft 文件夹里放了 '
                    'VLC 安装包就直接装，否则走 winget），'
                    '装好重启会务助手即可。',
                    parent=win)
                return
            start_index = max(0, min(int(start_index), len(state['list']) - 1))
            # 进度条停在半路时，从这里接着播（没动过进度条就从头开始）
            start_at = 0.0
            if state.get('start_at_for') == state['list'][start_index]:
                start_at = float(state.get('start_at') or 0.0)
            state['duration'] = 0.0
            state['seek_tries'] = 0
            state['pending_seek'] = start_at if start_at > 0.5 else None
            # 软件内预览：勾了开关且画面投在扩展屏时，在黑框里画一份静音画面。
            # 稍微晚一点起（600ms）：视频加载那一下预览区不会立刻被黑窗盖住，
            # 编辑用的黄框还能看见，等视频真出画面再盖上去。
            preview_ctl.stop()
            _place_preview_video(False)
            if _want_preview() and preview_ctl.available:
                def _start_preview_later(i=start_index):
                    if not state.get('playing'):
                        return
                    # 先把画面窗口摆到该在的位置（合成模式=黄框），再让 VLC 画进去
                    _draw_video_box()
                    _place_preview_video(True)
                    preview_ctl.start(preview_video_host.winfo_id(),
                                      state['list'], i, bool(loop_var.get()))
                    _place_preview_video()   # 起失败就把那块窗口收回去
                try:
                    win.after(600, _start_preview_later)
                except tk.TclError:
                    pass
            # 大屏合成窗口：背景铺满 + 视频嵌在框里（内嵌后端时用）
            compose = video_layout_var.get() == 'compose'
            if compose and not getattr(player, 'embedded', False):
                # 合成投屏要靠 VLC 的 libvlc 把画面画进指定框里，
                # 没装 VLC（用独立进程的 VLC）时做不到，退回整屏播放
                video_layout_var.set('full')
                compose = False
                messagebox.showinfo(
                    '提示',
                    '“合成投屏（背景图 + 视频框）”需要这台电脑装 VLC'
                    '（用它的 libvlc 引擎把画面画进指定位置）。\n\n'
                    '现在用的是：%s，本次按“整屏播放”处理。\n'
                    '装好 VLC 后重启会务助手即可使用合成投屏。'
                    % player.describe(),
                    parent=win)
            host = None
            screen = None
            if getattr(player, 'embedded', False):
                try:
                    screen = _show_compose_screen()
                    host = screen.hwnd() if screen is not None else None
                except Exception:
                    host = None
            if compose and host is None:
                _show_bg_screen()          # 退回外部播放器时的老路子
            else:
                _close_bg_screen()
            ok, note = player.play(
                state['list'], start_index, _monitor_index(),
                loop_var.get(), int(volume.get()), bool(top_var.get()),
                audio_device=state.get('audio_device'),
                window_rect=_video_window_rect() if compose else None,
                host_hwnd=host)
            if not ok:
                messagebox.showerror('播放失败', note, parent=win)
                preview_ctl.stop()
                _close_bg_screen()
                _close_compose_screen()
                return
            state['playing'] = True
            state['idx'] = start_index
            state['paused'] = False
            pause_btn.configure(text='暂停')
            dot.config(fg=p['success'])
            play_btn.configure(text='重新播放')
            _sync_backend_label()
            state_label.config(
                text='正在播放 %d/%d：%s'
                     % (start_index + 1, len(state['list']),
                        os.path.basename(state['list'][start_index])))
            info_label.config(text=note)
            _render(start_index)
            _start_esc_watch()  # 播起来了才开始轮询 ESC

        def _start():
            if not state['list']:
                messagebox.showinfo('提示', '请先点“添加视频…”选择要播放的视频。',
                                    parent=win)
                return
            _do_play(_selected() if _selected() >= 0 else 0)

        def _pause_toggle():
            if not state['playing']:
                return
            if not player.supports_pause:
                messagebox.showinfo(
                    '提示',
                    '当前用 %s 播放，暂停请在全屏画面上操作：\n\n'
                    '空格 = 暂停 / 继续\n← → = 上一个 / 下一个'
                    % player.backend.name, parent=win)
                return
            # 注意：不能拿 state()==3 判断“在播”，暂停时它也返回 3，
            # 会导致点“继续”又被当成暂停，再也回不去。
            if not player.paused:
                player.pause()
                try:
                    preview_ctl.pause()      # 软件内预览跟着一起暂停
                except Exception:
                    pass
                state['paused'] = True
                pause_btn.configure(text='继续')
                state_label.config(text='已暂停')
            else:
                player.resume()
                try:
                    preview_ctl.resume()
                except Exception:
                    pass
                state['paused'] = False
                pause_btn.configure(text='暂停')
                state_label.config(text='播放中')

        def _stop():
            state['playing'] = False
            state['idx'] = -1
            state['paused'] = False
            _stop_esc_watch()
            preview_ctl.stop()
            _close_bg_screen()
            _close_compose_screen()
            try:
                player.stop()
            except Exception:
                pass
            if not _alive():
                return
            dot.config(fg=p['faint'])
            state_label.config(text='已停止' if state['list'] else '未播放')
            play_btn.configure(text='开始播放')
            pause_btn.configure(text='暂停')
            # 停在中途的话，下次“开始播放”从这儿接着放；进度条拖到最左就从头
            if float(state.get('start_at') or 0.0) > 0.5:
                info_label.config(
                    text='开始播放将从 %s 继续（拖到最左＝从头播）'
                         % _fmt(state['start_at']))
            else:
                info_label.config(text='')
            _render()
            _draw_video_box()      # 停止后把编辑框画回来（播放时被视频盖住了）

        def _on_double_click(_event=None):
            i = _selected()
            if i >= 0:
                _do_play(i)

        def _poll():
            if not _alive():
                return
            if state['playing']:
                player.tick()       # 先让后端处理“播完 → 接下一个”
                if player.state() == 1:          # 播放器已经退出/播完
                    state['playing'] = False
                    preview_ctl.stop()
                    _place_preview_video(False)
                    _close_bg_screen()
                    _close_compose_screen()
                    _stop_esc_watch()   # 播完了就别再轮询
                    dot.config(fg=p['faint'])
                    state_label.config(text='播放结束')
                    play_btn.configure(text='开始播放')
                    info_label.config(text='')
                    _reset_seek()       # 播完了：进度条回起点
                    _draw_video_box()   # 播完把编辑框画回来
                else:
                    dur = player.duration()
                    pos = player.position()
                    # 起播时进度条停在半路的话，等时长出来了再跳过去
                    pend = state.get('pending_seek')
                    if pend is not None and dur > 0:
                        state['seek_tries'] = int(state.get('seek_tries', 0)) + 1
                        if not player.supports_seek or player.seek(pend) \
                                or state['seek_tries'] >= 3:
                            state['pending_seek'] = None
                            state['seek_tries'] = 0
                    if dur > 0:
                        state['duration'] = dur
                        if not state.get('seek_drag'):
                            seek.set(max(0.0, min(1.0, pos / dur)))
                        seek_label.config(text='%s / %s'
                                          % (_fmt(pos), _fmt(dur)))
                        name = player.current_name()
                        info_label.config(
                            text='第 %d/%d 个：%s'
                                 % (state['idx'] + 1, len(state['list']),
                                    os.path.basename(name) if name else ''))
                    else:
                        name = player.current_name()
                        if name:
                            info_label.config(
                                text='正在播放：%s' % os.path.basename(name))
            try:
                win.after(500, _poll)
            except tk.TclError:
                pass

        # ---------- 软件内预览 ----------
        # 预览画面由 VLC 的库（libvlc）直接画进左边那块黑框里，静音、
        # 不额外开窗（软件外面不再出现第二个预览窗口）。

        def _preview_has_art():
            """预览区里现在有没有画面可看。"""
            if preview_ctl.running or state.get('playing'):
                return True
            if video_layout_var.get() == 'compose':
                return True             # 合成模式画的是排版效果图
            return state.get('frame_pic') is not None

        def _set_preview_hint(show=None):
            """预览区有画面时就把提示文字收起来。"""
            if show is None:
                show = not _preview_has_art()
            try:
                if show:
                    preview_hint.place(relx=0.5, rely=0.5, anchor='center')
                else:
                    preview_hint.place_forget()
            except tk.TclError:
                pass

        def _watch_preview():
            try:
                if not win.winfo_exists():
                    return
            except tk.TclError:
                return
            if preview_ctl.running:
                preview_ctl.tick()          # 播完一个自动接下个
            _set_preview_hint()
            # 编辑框必须一直在（播放刚加载、被视频盖掉、漏画都补回来）
            if video_layout_var.get() == 'compose':
                try:
                    if not preview_host.find_withtag('vbox'):
                        _draw_box_overlay()
                except tk.TclError:
                    pass
            _place_preview_video()      # 预览里那块静音画面一直贴着黄框
            try:
                win.after(400, _watch_preview)
            except tk.TclError:
                pass

        tree.bind('<Double-Button-1>', _on_double_click)
        tree.bind('<<TreeviewSelect>>', _on_pick)
        state_label.config(text='未播放')
        _render()
        # 底部按钮先占好位，中部内容再填剩下的空间
        mid.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(12, 8))
        # 预览上可以直接拖视频框（合成模式下）
        preview_host.bind('<Motion>', _video_motion)
        preview_host.bind('<Button-1>', _video_press)
        preview_host.bind('<B1-Motion>', _video_drag)
        preview_host.bind('<ButtonRelease-1>', _video_release)
        preview_host.bind('<Configure>', lambda e: _draw_video_box(), add='+')
        _sync_video_bg_list()
        _sync_video_layout()
        _sync_vlc_ui()
        _sync_seek_ui()
        _sync_backend_label()
        win._close_bg = _close_bg_screen
        win._close_proj = _close_compose_screen
        win._grab_win = state.get('grab_win')
        _poll()
        _watch_preview()

    def _close_video_win(self):
        """关闭视频窗口：停止播放并释放播放器。"""
        win = self._video_win
        ctl = getattr(win, '_preview_ctl', None) if win is not None else None
        if ctl is not None:
            try:
                ctl.stop()
            except Exception:
                pass
        esc_off = getattr(win, '_disarm_esc', None) if win is not None else None
        if esc_off is not None:
            try:
                esc_off()
            except Exception:
                pass
        closer = getattr(win, '_close_bg', None) if win is not None else None
        if closer is not None:
            try:
                closer()
            except Exception:
                pass
        proj_closer = getattr(win, '_close_proj', None) if win is not None else None
        if proj_closer is not None:
            try:
                proj_closer()
            except Exception:
                pass
        grab = getattr(win, '_grab_win', None) if win is not None else None
        if grab is not None:
            try:
                grab.destroy()
            except tk.TclError:
                pass
        if self._video_player is not None:
            try:
                self._video_player.stop()
            except Exception:
                pass
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._video_win = None

    # ---------- 背景图库管理（会务视频 / 长条屏 / 现场抽签 共用） ----------
    def _open_bg_manager(self, parent, on_change=None):
        """背景图库：看列表和预览，上传 / 重命名 / 删除。

        on_change 会在每次改动后回调，调用方用它刷新自己的下拉框。
        """
        import ppt_converter
        p = ui_theme.PALETTE
        win = tk.Toplevel(parent)
        win.title('背景图管理')
        _center_window(win, 800, 580)
        win.minsize(680, 480)
        win.transient(parent)
        ui_theme.style_window(win)

        body = tk.Frame(win, bg=p['bg'], padx=16, pady=14)
        body.pack(fill=tk.BOTH, expand=True)
        head = tk.Frame(body, bg=p['bg'])
        head.pack(fill=tk.X)
        tk.Label(head, text='背景图库', font=ui_theme.font(13, 'bold'),
                 bg=p['bg'], fg=p['text']).pack(side=tk.LEFT)
        tk.Label(head, text='会务视频 / 长条屏 / 现场抽签 共用这一套图',
                 font=ui_theme.font(9), bg=p['bg'],
                 fg=p['faint']).pack(side=tk.LEFT, padx=(10, 0))

        mid = tk.Frame(body, bg=p['bg'])
        mid.pack(fill=tk.BOTH, expand=True, pady=(10, 8))
        left = tk.Frame(mid, bg=p['bg'])
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(left, columns=('name', 'size', 'created'),
                            show='headings', height=12)
        tree.heading('name', text='名称')
        tree.heading('size', text='尺寸')
        tree.heading('created', text='添加时间')
        tree.column('name', width=220, anchor='w', stretch=True)
        tree.column('size', width=110, anchor='center', stretch=False)
        tree.column('created', width=150, anchor='center', stretch=False)
        scroll = ttk.Scrollbar(left, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        right = tk.Frame(mid, bg=p['bg'])
        right.pack(side=tk.LEFT, fill=tk.Y, padx=(14, 0))
        shot = tk.Canvas(right, width=300, height=180, bg='#0b1220',
                         highlightthickness=0, bd=0)
        shot.pack()
        shot_info = tk.Label(right, text='选一张图看预览', font=ui_theme.font(9),
                             bg=p['bg'], fg=p['faint'], wraplength=300,
                             justify='left')
        shot_info.pack(anchor='w', pady=(6, 0))

        bar = tk.Frame(body, bg=p['bg'])
        bar.pack(fill=tk.X)
        state = {'items': [], 'photo': None}

        def _notify():
            if callable(on_change):
                try:
                    on_change()
                except Exception:
                    pass

        def _selected():
            sel = tree.selection()
            if not sel:
                return None
            index = int(sel[0])
            if 0 <= index < len(state['items']):
                return state['items'][index]
            return None

        def _show_preview():
            item = _selected()
            shot.delete('all')
            state['photo'] = None
            if item is None:
                shot_info.config(text='选一张图看预览')
                return
            try:
                from PIL import Image, ImageTk
                with Image.open(item['image']) as raw:
                    img = raw.convert('RGB')
                w, h = img.size
                scale = min(300.0 / max(1, w), 180.0 / max(1, h))
                size = (max(1, int(w * scale)), max(1, int(h * scale)))
                photo = ImageTk.PhotoImage(img.resize(size))
                state['photo'] = photo          # 留引用
                shot.create_image(150, 90, image=photo)
                shot_info.config(text='%s\n%dx%d' % (item['name'], w, h))
            except Exception as exc:
                shot_info.config(text='这张图读不出来：%s' % exc)

        def _reload(keep=None):
            state['items'] = ppt_converter.list_uploaded_backgrounds()
            tree.delete(*tree.get_children())
            want = keep
            for i, item in enumerate(state['items']):
                try:
                    from PIL import Image
                    with Image.open(item['image']) as raw:
                        size = '%dx%d' % raw.size
                except Exception:
                    size = '读不出'
                tree.insert('', 'end', iid=str(i),
                            values=(item['name'], size,
                                    item.get('created', '')))
                if keep and item['name'] == keep:
                    want = i
            if state['items']:
                index = want if isinstance(want, int) else 0
                index = max(0, min(index, len(state['items']) - 1))
                tree.selection_set(str(index))
                tree.see(str(index))
            _show_preview()

        def _upload():
            paths = filedialog.askopenfilenames(
                parent=win, title='选择背景图（可多选）',
                filetypes=[('图片文件', '*.png *.jpg *.jpeg *.bmp *.webp'),
                           ('所有文件', '*.*')])
            if not paths:
                return
            last = ''
            for path in paths:
                stem = os.path.splitext(os.path.basename(path))[0]
                name = simpledialog.askstring(
                    '图片名称', '给这张图起个名字：',
                    initialvalue=stem, parent=win)
                if not name:
                    continue
                try:
                    item = ppt_converter.add_uploaded_background(path, name)
                    last = item['name']
                except Exception as exc:
                    messagebox.showerror('上传失败', str(exc), parent=win)
            if last:
                _reload(keep=last)
                _notify()

        def _rename():
            item = _selected()
            if item is None:
                messagebox.showinfo('提示', '先在列表里选一张图。', parent=win)
                return
            new_name = simpledialog.askstring(
                '重命名', '新的名称：', initialvalue=item['name'],
                parent=win)
            if not new_name or new_name.strip() == item['name']:
                return
            new_name = new_name.strip()
            if any(x['name'] == new_name for x in state['items']):
                messagebox.showwarning('重名了',
                                       '已经有一张叫“%s”的图，换个名字。'
                                       % new_name, parent=win)
                return
            if not ppt_converter.rename_uploaded_background(item['id'],
                                                            new_name):
                messagebox.showerror('重命名失败', '这张图没能改名。',
                                     parent=win)
                return
            _reload(keep=new_name)
            _notify()

        def _delete():
            item = _selected()
            if item is None:
                messagebox.showinfo('提示', '先在列表里选一张图。', parent=win)
                return
            if not messagebox.askyesno(
                    '删除背景图',
                    '确定删除“%s”吗？\n\n删掉后，正在用它做背景的窗口'
                    '要重新选一张。' % item['name'], parent=win):
                return
            ppt_converter.delete_uploaded_background(item['id'])
            _reload()
            _notify()

        ttk.Button(bar, text='上传背景图…', style='Accent.TButton',
                   command=lambda: _upload()).pack(side=tk.LEFT)
        ttk.Button(bar, text='重命名…', style='Ghost.TButton',
                   command=lambda: _rename()).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(bar, text='删除', style='Danger.TButton',
                   command=lambda: _delete()).pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(bar, text='双击列表项也能改名', font=ui_theme.font(8),
                 bg=p['bg'], fg=p['faint']).pack(side=tk.LEFT, padx=(12, 0))
        ttk.Button(bar, text='关闭', style='Accent.TButton',
                   command=lambda: win.destroy()).pack(side=tk.RIGHT)
        tree.bind('<<TreeviewSelect>>', lambda e: _show_preview())
        tree.bind('<Double-Button-1>', lambda e: _rename())
        _reload()

    # ---------- 现场抽签排序 ----------
    def _open_lottery(self):
        """现场抽签排序：原生窗口，同进程秒开。

        由 `lottery.LotteryAdmin`（管理窗口）+ `lottery.LotteryScreen`
        （扩展屏大屏动画）组成，数据沿用 抽签排序/data/state.json，
        支持 seed 复现校验、Excel/CSV/粘贴导入、历史记录与导出。
        """
        if self._activate_or_close_others('lottery'):
            return
        win = tk.Toplevel(self.root)
        # 先隐藏，等定位和全部控件构建完再显示，避免打开瞬间小窗
        # 在左上角闪一下、内容堆着再展开
        win.withdraw()
        win.title('现场抽签排序')
        _center_window(win, 1060, 720)
        try:
            app = lottery.LotteryAdmin(
                win, on_close=lambda: self._close_lottery_win())
            self._lottery_win = win
            self._lottery_app = app
        except Exception as e:
            messagebox.showerror(
                '错误', '打开现场抽签排序失败：\n%s' % e, parent=self.root)
            try:
                win.destroy()
            except Exception:
                pass
            return
        win.deiconify()
        win.lift()
        win.focus_force()

    def _close_lottery_win(self):
        """关闭抽签：关掉管理窗口 + 大屏窗口，停止背景音乐。"""
        app = getattr(self, '_lottery_app', None)
        if app is not None:
            try:
                app.close()
            except Exception:
                pass
        self._lottery_app = None
        win = self._lottery_win
        self._lottery_win = None
        if win is not None:
            try:
                if win.winfo_exists():
                    win.destroy()
            except Exception:
                pass

    # ---------- 长条屏合成投屏 ----------
    def _open_strip(self):
        """长条屏合成：两侧固定图 + 中间等比 PPT，合成整屏投到长条显示屏。"""
        if self._activate_or_close_others('strip'):
            return
        try:
            import ppt_converter
        except Exception as e:
            messagebox.showerror('错误', '加载失败：\n%s' % e, parent=self.root)
            return

        p = ui_theme.PALETTE
        win = tk.Toplevel(self.root)
        win.title('长条屏合成投屏')
        _center_window(win, 1180, 900)
        win.minsize(1020, 720)
        win.resizable(True, True)
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW',
                     lambda w=win: self._request_close(
                         '长条屏合成投屏', self._close_strip_win, w))
        ui_theme.style_window(win)
        self._strip_win = win

        body = tk.Frame(win, bg=p['bg'], padx=16, pady=14)
        body.pack(fill=tk.BOTH, expand=True)

        head = tk.Frame(body, bg=p['bg'])
        head.pack(fill=tk.X)
        tk.Label(head, text='长条屏合成投屏', font=ui_theme.font(13, 'bold'),
                 bg=p['bg'], fg=p['text']).pack(side=tk.LEFT)
        dot = tk.Label(head, text='●', font=ui_theme.font(11), bg=p['bg'],
                       fg=p['faint'])
        dot.pack(side=tk.LEFT, padx=(12, 4))
        state_label = tk.Label(head, text='未投屏', font=ui_theme.font(9),
                               bg=p['bg'], fg=p['muted'])
        state_label.pack(side=tk.LEFT)
        tk.Label(head, text='合成整屏后 1:1 投出；可用背景图 + PPT 任意摆放',
                 font=ui_theme.font(9), bg=p['bg'],
                 fg=p['faint']).pack(side=tk.RIGHT)

        mid = tk.Frame(body, bg=p['bg'])
        mid.pack(fill=tk.BOTH, expand=True, pady=(12, 8))

        # 左：设置（可滚动、可折叠）+ 中间 PPT 列表（固定在下方，永远看得见）
        left_wrap = tk.Frame(mid, bg=p['bg'])
        left_wrap.pack(side=tk.LEFT, fill=tk.Y)
        ppt_holder = tk.Frame(left_wrap, bg=p['bg'])
        ppt_holder.pack(side=tk.BOTTOM, fill=tk.X)
        fold_row = tk.Frame(left_wrap, bg=p['bg'])
        fold_row.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))
        collapse_var = tk.BooleanVar(value=False)
        collapse_btn = ttk.Button(fold_row, text='折叠设置',
                                  style='Ghost.TButton',
                                  command=lambda: _toggle_settings())
        collapse_btn.pack(side=tk.LEFT)
        tk.Label(fold_row, text='设置项多时可以折叠，直接看 PPT 列表',
                 font=ui_theme.font(8), bg=p['bg'],
                 fg=p['faint']).pack(side=tk.LEFT, padx=(8, 0))
        body_wrap = tk.Frame(left_wrap, bg=p['bg'])
        body_wrap.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        left_scroll = ttk.Scrollbar(body_wrap, orient=tk.VERTICAL)
        left_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        left_canvas = tk.Canvas(body_wrap, bg=p['bg'], highlightthickness=0,
                                bd=0, width=300,
                                yscrollcommand=left_scroll.set)
        left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left_scroll.configure(command=left_canvas.yview)
        left = tk.Frame(left_canvas, bg=p['bg'])
        left_item = left_canvas.create_window((0, 0), window=left, anchor='nw')

        def _sync_left(_event=None):
            try:
                left_canvas.configure(scrollregion=left_canvas.bbox('all'))
                left_canvas.itemconfigure(left_item,
                                          width=max(200,
                                                    left_canvas.winfo_width()))
            except tk.TclError:
                pass

        def _left_wheel(event):
            try:
                left_canvas.yview_scroll(-1 if event.delta > 0 else 1, 'units')
            except tk.TclError:
                pass

        left.bind('<Configure>', _sync_left)
        left_canvas.bind('<Configure>', _sync_left)
        left_canvas.bind('<MouseWheel>', _left_wheel)
        left.bind('<MouseWheel>', _left_wheel)

        def _group(title, parent=None):
            box = ttk.LabelFrame(parent if parent is not None else left,
                                 text=' %s ' % title, padding=(10, 8))
            box.pack(fill=tk.X, pady=(0, 10))
            return box

        g_screen = _group('显示与分区')
        monitors = video_player.list_monitors()
        default_index = 0
        for item in monitors:
            if not item['primary']:
                default_index = item['index']
                break

        def _screen_label(item):
            r = item['rect']
            return '显示器 %d（%d×%d）%s' % (
                item['index'] + 1, r[2] - r[0], r[3] - r[1],
                ' 主屏' if item['primary'] else ' 扩展屏')

        screen_var = tk.StringVar(
            value=_screen_label(monitors[default_index]) if monitors
            else '显示器 1')
        screen_combo = ttk.Combobox(g_screen, state='readonly', width=26,
                                    textvariable=screen_var,
                                    values=[_screen_label(m) for m in monitors]
                                    or ['显示器 1'])
        screen_combo.pack(anchor='w')
        row = tk.Frame(g_screen, bg=p['bg'])
        row.pack(fill=tk.X, pady=(8, 0))
        tk.Label(row, text='左区 %', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        left_var = tk.IntVar(value=15)
        left_spin = ttk.Spinbox(row, from_=0, to=45, width=4,
                                textvariable=left_var)
        left_spin.pack(side=tk.LEFT, padx=(4, 10))
        tk.Label(row, text='右区 %', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        right_var = tk.IntVar(value=15)
        right_spin = ttk.Spinbox(row, from_=0, to=45, width=4,
                                 textvariable=right_var)
        right_spin.pack(side=tk.LEFT, padx=(4, 0))
        row = tk.Frame(g_screen, bg=p['bg'])
        row.pack(fill=tk.X, pady=(6, 0))
        tk.Label(row, text='中间上下留白 %', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        margin_var = tk.IntVar(value=0)
        margin_spin = ttk.Spinbox(row, from_=0, to=25, width=4,
                                  textvariable=margin_var)
        margin_spin.pack(side=tk.LEFT, padx=(4, 0))
        # 直接在输入框里改百分比也立刻生效（拖动时由代码回写，加锁避免打架）
        def _on_ratio_var(*_args):
            if ratio_sync['busy'] or layout_var.get() == 'custom':
                return
            state['left_ratio'] = max(0.0, min(left_var.get() / 100.0, 0.45))
            state['right_ratio'] = max(0.0, min(right_var.get() / 100.0, 0.45))
            state['frames'] = {}
            _refresh_preview()

        left_var.trace_add('write', _on_ratio_var)
        right_var.trace_add('write', _on_ratio_var)
        tk.Label(g_screen, text='中间区域会按“等比缩放”放 PPT，绝不拉伸变形',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 justify='left', wraplength=250).pack(anchor='w', pady=(6, 0))
        auto_ratio_btn = ttk.Button(g_screen, text='按 PPT 16:9 自动分配',
                                    style='Ghost.TButton',
                                    command=lambda: _auto_ratio())
        auto_ratio_btn.pack(anchor='w', pady=(6, 0))

        g_side = _group('两侧固定图')
        bg_state = {'items': [x for x in ppt_converter.list_uploaded_backgrounds()]}
        side_names = ['（不用图）'] + [x['name'] for x in bg_state['items']]
        row = tk.Frame(g_side, bg=p['bg'])
        row.pack(fill=tk.X)
        tk.Label(row, text='左', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        left_img_var = tk.StringVar(value=side_names[1] if len(side_names) > 1
                                    else side_names[0])
        left_combo = ttk.Combobox(row, state='readonly', width=14,
                                  textvariable=left_img_var,
                                  values=side_names)
        left_combo.pack(side=tk.LEFT, padx=(6, 0))
        row = tk.Frame(g_side, bg=p['bg'])
        row.pack(fill=tk.X, pady=(4, 0))
        tk.Label(row, text='右', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        right_img_var = tk.StringVar(
            value=side_names[2] if len(side_names) > 2 else side_names[0])
        right_combo = ttk.Combobox(row, state='readonly', width=14,
                                   textvariable=right_img_var,
                                   values=side_names)
        right_combo.pack(side=tk.LEFT, padx=(6, 0))
        side_upload_btn = ttk.Button(g_side, text='上传两侧图片…',
                                     style='Ghost.TButton',
                                     command=lambda: _upload_side())
        side_upload_btn.pack(anchor='w', pady=(6, 0))

        # 版式：左右分区（旧） / 自定义背景 + PPT 任意位置（新）
        g_layout = _group('版式')
        layout_var = tk.StringVar(value='custom')
        ttk.Radiobutton(g_layout, text='自定义背景 + PPT 任意位置',
                        value='custom', variable=layout_var,
                        command=lambda: _sync_layout()).pack(anchor='w')
        ttk.Radiobutton(g_layout, text='左右分区（两侧固定图）', value='sides',
                        variable=layout_var,
                        command=lambda: _sync_layout()).pack(anchor='w',
                                                             pady=(2, 0))
        row = tk.Frame(g_layout, bg=p['bg'])
        row.pack(fill=tk.X, pady=(8, 0))
        tk.Label(row, text='背景图', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        bg_pick_var = tk.StringVar(value='（不用图）')
        bg_pick = ttk.Combobox(row, state='readonly', width=14,
                               textvariable=bg_pick_var, values=['（不用图）'])
        bg_pick.pack(side=tk.LEFT, padx=(6, 0))
        bg_pick.bind('<<ComboboxSelected>>', lambda e: _on_layout_change())
        bg_upload_btn = ttk.Button(g_layout, text='上传整屏背景图…',
                                   style='Ghost.TButton',
                                   command=lambda: _upload_bg())
        bg_upload_btn.pack(anchor='w', pady=(6, 0))
        ttk.Button(g_layout, text='管理背景图…', style='Ghost.TButton',
                   command=lambda: self._open_bg_manager(
                       win, on_change=lambda: _reload_bg_library())).pack(
            anchor='w', pady=(4, 0))
        box_row = tk.Frame(g_layout, bg=p['bg'])
        box_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(box_row, text='PPT 框 16:9 居中', style='Ghost.TButton',
                   command=lambda: _auto_box()).pack(side=tk.LEFT)
        ttk.Button(box_row, text='铺满', style='Ghost.TButton',
                   command=lambda: _box_full()).pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(g_layout, text='在右边预览上拖动黄框移动 PPT，\n'
                                '拖四个角改大小（等比缩放不改形）',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 justify='left', wraplength=250).pack(anchor='w', pady=(6, 0))

        def _toggle_settings():
            """折叠 / 展开上半部分的设置：折叠后 PPT 列表占满整栏。"""
            folded = not collapse_var.get()
            collapse_var.set(folded)
            try:
                if folded:
                    body_wrap.pack_forget()
                    collapse_btn.configure(text='展开设置')
                else:
                    body_wrap.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
                    collapse_btn.configure(text='折叠设置')
                    left_canvas.yview_moveto(0.0)
            except tk.TclError:
                pass
            _sync_left()

        g_ppt = _group('中间 PPT', ppt_holder)
        ppt_var = tk.StringVar(value='未选择')
        tk.Label(g_ppt, textvariable=ppt_var, font=ui_theme.font(8),
                 bg=p['bg'], fg=p['muted'], wraplength=250,
                 justify='left').pack(anchor='w')
        ttk.Button(g_ppt, text='添加 PPT（可多选）…', style='Accent.TButton',
                   command=lambda: _pick_ppt()).pack(anchor='w', pady=(6, 0))
        # PPT 列表：会场有很多份 PPT 时，按顺序一份接一份放；
        # 双击列表项直接切到那份 PPT。
        deck_box = tk.Frame(g_ppt, bg=p['bg'])
        deck_box.pack(fill=tk.X, pady=(6, 0))
        deck_tree = ttk.Treeview(deck_box, columns=('no', 'name', 'pages'),
                                 show='headings', height=5)
        deck_tree.heading('no', text='序')
        deck_tree.heading('name', text='PPT')
        deck_tree.heading('pages', text='页')
        deck_tree.column('no', width=28, anchor='center', stretch=False)
        deck_tree.column('name', width=150, anchor='w', stretch=True)
        deck_tree.column('pages', width=34, anchor='center', stretch=False)
        deck_tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        deck_scroll = ttk.Scrollbar(deck_box, orient=tk.VERTICAL,
                                    command=deck_tree.yview)
        deck_tree.configure(yscrollcommand=deck_scroll.set)
        deck_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        deck_row = tk.Frame(g_ppt, bg=p['bg'])
        deck_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(deck_row, text='移除', style='Ghost.TButton',
                   command=lambda: _remove_ppt()).pack(side=tk.LEFT)
        ttk.Button(deck_row, text='清空', style='Ghost.TButton',
                   command=lambda: _clear_ppt()).pack(side=tk.LEFT,
                                                      padx=(6, 0))
        row = tk.Frame(g_ppt, bg=p['bg'])
        row.pack(fill=tk.X, pady=(8, 0))
        auto_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, text='自动翻页', variable=auto_var,
                        command=lambda: _sync_auto()).pack(side=tk.LEFT)
        tk.Label(row, text='间隔', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT, padx=(8, 2))
        interval_var = tk.IntVar(value=15)
        ttk.Spinbox(row, from_=3, to=600, width=5,
                    textvariable=interval_var).pack(side=tk.LEFT)
        loop_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(g_ppt, text='放完循环', variable=loop_var).pack(
            anchor='w', pady=(4, 0))
        # 原样放映：用 WPS 演示 / PowerPoint 自己放，动画/切换/内嵌视频音频都在；
        # 放映窗口会被摆到中间 PPT 框里，两侧固定图照样保留。
        live_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(g_ppt, text='原样放映（带动画/视频，窗嵌中间框）',
                        variable=live_var).pack(anchor='w', pady=(4, 0))
        tk.Label(g_ppt, text='勾上后：用 WPS/PPT 直接放映这份 PPT，'
                             '翻页笔、自动翻页照旧；不勾则用导出的图片合成',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 wraplength=250, justify='left').pack(anchor='w', pady=(2, 0))

        # 右：预览 + 翻页
        right = tk.Frame(mid, bg=p['bg'])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(16, 0))
        ttk.Label(right, text='合成效果预览（拖黄色竖线可改左右分区）',
                  font=ui_theme.font(9), foreground=p['muted'],
                  background=p['bg']).pack(anchor='w')
        preview = tk.Canvas(right, bg='#0b1a33', highlightthickness=0, bd=0,
                            cursor='arrow')
        preview.pack(fill=tk.BOTH, expand=True, pady=(6, 8))
        # 幻灯片缩略图条（像 PPT 的幻灯片预览）：能看前后页、点一下跳过去，
        # 也可以滚轮/拖动翻看其它页
        film_row = tk.Frame(right, bg=p['bg'])
        film_row.pack(fill=tk.X, pady=(0, 6))
        film_holder = tk.Frame(film_row, bg='#0b1220', bd=0,
                               highlightthickness=1,
                               highlightbackground=p['border'])
        film_holder.pack(fill=tk.X, expand=True)
        film = tk.Canvas(film_holder, height=104, bg='#0b1220',
                         highlightthickness=0, bd=0, cursor='hand2')
        film.pack(fill=tk.BOTH, expand=True)
        tk.Label(right, text='缩略图：点一下跳页，滚轮或按住拖动看其它页',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 anchor='w').pack(anchor='w')
        page_row = tk.Frame(right, bg=p['bg'])
        page_row.pack(fill=tk.X)
        ttk.Button(page_row, text='上一页', style='Ghost.TButton',
                   command=lambda: _step(-1)).pack(side=tk.LEFT)
        ttk.Button(page_row, text='下一页', style='Ghost.TButton',
                   command=lambda: _step(1)).pack(side=tk.LEFT, padx=(6, 0))
        page_label = tk.Label(page_row, text='第 0 / 0 页',
                              font=ui_theme.font(9), bg=p['bg'],
                              fg=p['muted'])
        page_label.pack(side=tk.LEFT, padx=(12, 0))
        ttk.Button(page_row, text='刷新预览', style='Ghost.TButton',
                   command=lambda: _refresh_preview()).pack(side=tk.RIGHT)

        bar = tk.Frame(body, bg=p['bg'])
        bar.pack(fill=tk.X)
        start_btn = ttk.Button(bar, text='开始投屏', style='Accent.TButton',
                               command=lambda: _start_screen())
        start_btn.pack(side=tk.LEFT)
        ttk.Button(bar, text='停止投屏', style='Danger.TButton',
                   command=lambda: _stop_screen()).pack(side=tk.LEFT,
                                                        padx=(8, 0))
        tk.Label(body, text='提示：可一次加多份 PPT，按顺序依次播放（一份放完自动接下一份）；'
                            '双击列表切换，↑ ↓（或 PgUp/PgDn、← →）翻页，'
                            '空格切换自动翻页，ESC 退出投屏。',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 anchor='w').pack(fill=tk.X, pady=(8, 0))

        # ---------- 内部函数 ----------
        # decks：PPT 列表，每份 {path, name, pages}；cur=当前第几份，idx=当前第几页
        state = {'decks': [], 'cur': -1, 'idx': -1, 'cache': {},
                 'auto_job': None, 'thumb': None,
                 'preview_box': None, 'grip': None, 'frames': {},
                 'slide_box': [0.25, 0.10, 0.50, 0.80],   # PPT 框（比例）
                 'left_ratio': 0.15, 'right_ratio': 0.15,  # 左右区（浮点，跟手）
                 'drag': None, 'drag_origin': None}
        ratio_sync = {'busy': False}

        def _alive():
            try:
                return bool(win.winfo_exists())
            except tk.TclError:
                return False

        def _deck(index=None):
            """取第 index 份 PPT；不给就是当前这份。"""
            i = state['cur'] if index is None else index
            decks = state['decks']
            if 0 <= i < len(decks):
                return decks[i]
            return None

        def _pages():
            deck = _deck()
            return deck['pages'] if deck else []

        def _monitor_rect():
            """选中的显示器（插拔投影仪后按“主屏/扩展屏”重新对）。"""
            items = video_player.list_monitors()
            index = video_player.resolve_monitor_index(screen_var.get(), items)
            index = max(0, min(index, len(items) - 1))
            return items[index]

        def _refresh_screen_box(_event=None):
            items = video_player.list_monitors()
            names = [_screen_label(m) for m in items] or ['显示器 1']
            if list(screen_combo.cget('values')) != names:
                index = video_player.resolve_monitor_index(screen_var.get(),
                                                           items)
                screen_combo.configure(values=names)
                screen_var.set(_screen_label(items[index]))

        screen_combo.bind('<Button-1>', _refresh_screen_box, add='+')
        win.bind('<FocusIn>', _refresh_screen_box, add='+')

        def _side_image(var):
            name = var.get()
            for item in bg_state['items']:
                if item['name'] == name:
                    return strip_screen.load_image(item['image'])
            return None

        def _custom_bg_image():
            """自定义版式用的整屏背景图；没选图返回 None（用底色）。"""
            name = bg_pick_var.get()
            if not name or name == '（不用图）':
                return None
            for item in bg_state['items']:
                if item['name'] == name:
                    return strip_screen.load_image(item['image'])
            return None

        def _sync_bg_list(select=None):
            """上传后刷新“背景图”下拉。"""
            bg_state['items'] = ppt_converter.list_uploaded_backgrounds()
            names = ['（不用图）'] + [x['name'] for x in bg_state['items']]
            bg_pick.configure(values=names)
            if select and select in names:
                bg_pick_var.set(select)
            elif bg_pick_var.get() not in names:
                bg_pick_var.set(names[0])

        def _reload_bg_library():
            """背景图库改动后，把整屏背景图和两侧图的三个下拉都刷一遍。"""
            _sync_bg_list()
            names = ['（不用图）'] + [x['name'] for x in bg_state['items']]
            for combo, var in ((left_combo, left_img_var),
                               (right_combo, right_img_var)):
                combo.configure(values=names)
                if var.get() not in names:
                    var.set(names[0])
            state['frames'] = {}
            _refresh_preview()
            _on_layout_change()

        def _upload_bg():
            src = filedialog.askopenfilename(
                parent=win, title='选择整屏背景图',
                filetypes=[('图片文件', '*.png *.jpg *.jpeg *.bmp *.webp'),
                           ('所有文件', '*.*')])
            if not src:
                return
            stem = os.path.splitext(os.path.basename(src))[0]
            name = simpledialog.askstring('背景名称', '给这张背景起个名字：',
                                          initialvalue=stem, parent=win)
            if not name:
                return
            try:
                item = ppt_converter.add_uploaded_background(src, name)
            except Exception as e:
                messagebox.showerror('上传失败', str(e), parent=win)
                return
            _sync_bg_list(item['name'])
            state_label.config(text='背景图已保存：%s' % name)
            _on_layout_change()

        def _auto_box():
            """PPT 框：按 16:9 居中，占屏幕高度 90%。"""
            rect = _monitor_rect()['rect']
            width = max(2, rect[2] - rect[0])
            height = max(2, rect[3] - rect[1])
            box_h = 0.9
            box_w = min(1.0, (height * box_h) * 16.0 / 9.0 / float(width))
            state['slide_box'] = [(1.0 - box_w) / 2.0, (1.0 - box_h) / 2.0,
                                  box_w, box_h]
            _on_layout_change()

        def _box_full():
            """PPT 直接铺满整屏。"""
            state['slide_box'] = [0.0, 0.0, 1.0, 1.0]
            _on_layout_change()

        def _sync_layout():
            """切换版式：不用的控件灰掉，预览跟着重画。"""
            custom = layout_var.get() == 'custom'
            for widget in (left_spin, right_spin, margin_spin, auto_ratio_btn,
                           side_upload_btn):
                try:
                    widget.configure(state='disabled' if custom else 'normal')
                except tk.TclError:
                    pass
            for combo in (left_combo, right_combo):
                try:
                    combo.configure(state='disabled' if custom else 'readonly')
                except tk.TclError:
                    pass
            for widget in (bg_pick, bg_upload_btn):
                try:
                    widget.configure(
                        state=('disabled' if not custom
                               else ('readonly' if widget is bg_pick
                                     else 'normal')))
                except tk.TclError:
                    pass
            _on_layout_change()

        def _on_layout_change(_event=None):
            state['frames'] = {}
            _refresh_preview()
            if self._strip_screen is not None and state['idx'] >= 0:
                _push_to_screen()

        # 下拉框、版式一变就重画（不依赖事件是否送达）
        bg_pick_var.trace_add('write', lambda *a: _on_layout_change())
        layout_var.trace_add('write', lambda *a: _sync_layout())

        def _auto_ratio():
            """按 16:9 算出中间区宽度，把剩下的平均分给左右两侧。"""
            info = _monitor_rect()
            rect = info['rect']
            width = max(2, rect[2] - rect[0])
            height = max(2, rect[3] - rect[1])
            mid = min(width, int(round(height * 16.0 / 9.0)))
            side_ratio = max(0.0, min((width - mid) / 2.0 / width, 0.45))
            state['left_ratio'] = side_ratio
            state['right_ratio'] = side_ratio
            ratio_sync['busy'] = True
            try:
                left_var.set(int(round(side_ratio * 100)))
                right_var.set(int(round(side_ratio * 100)))
            finally:
                ratio_sync['busy'] = False
            side = int(round(side_ratio * 100))
            state_label.config(
                text='已按 16:9 分配：中间 %d 像素，左右各 %d%%' % (mid, side))
            state['frames'] = {}
            _refresh_preview()

        def _make_compositor():
            info = _monitor_rect()
            rect = info['rect']
            if layout_var.get() == 'custom':
                # 自定义版式：整屏背景图 + 任意位置的 PPT 框
                return strip_screen.Compositor(
                    rect[2] - rect[0], rect[3] - rect[1],
                    background_image=_custom_bg_image(),
                    slide_box=tuple(state['slide_box']))
            comp = strip_screen.Compositor(
                rect[2] - rect[0], rect[3] - rect[1],
                left_ratio=state['left_ratio'],
                right_ratio=state['right_ratio'],
                margin_ratio=margin_var.get() / 100.0,
                left_image=_side_image(left_img_var),
                right_image=_side_image(right_img_var))
            return comp

        def _page_image(index):
            """按需加载一页（缓存最近几页，几十页也不吃内存）。"""
            pages = _pages()
            if not (0 <= index < len(pages)):
                return None
            if index in state['cache']:
                return state['cache'][index]
            try:
                img = strip_screen.load_image(pages[index])
            except Exception:
                img = None
            state['cache'][index] = img
            if len(state['cache']) > 4:
                for key in list(state['cache'])[:-3]:
                    if key != index:
                        state['cache'].pop(key, None)
            return img

        def _compose():
            """合成整屏画面（带缓存：同一页/同一组比例只合成一次）。"""
            key = (state['cur'], state['idx'], layout_var.get(),
                   round(state['left_ratio'], 4),
                   round(state['right_ratio'], 4), margin_var.get(),
                   bg_pick_var.get(), tuple(state['slide_box']))
            frame = state['frames'].get(key)
            if frame is None:
                comp = _make_compositor()
                # “原样放映”时中间那块不画图（留黑），真正的画面由
                # WPS/PowerPoint 的放映窗口画上去
                page = None
                if state['idx'] >= 0 and not _live_active():
                    page = _page_image(state['idx'])
                frame = comp.render(page)
                state['frames'][key] = frame
                if len(state['frames']) > 8:
                    for old in list(state['frames'])[:-5]:
                        state['frames'].pop(old, None)
            return None, frame

        def _prefetch(index):
            """提前把下一页/上一页合成好，翻页时就不用现算，反应更快。"""
            if not _pages() or state['cur'] < 0:
                return
            total = len(_pages())
            if not (0 <= index < total):
                return
            key = (state['cur'], index, layout_var.get(),
                   round(state['left_ratio'], 4),
                   round(state['right_ratio'], 4), margin_var.get(),
                   bg_pick_var.get(),
                   tuple(state['slide_box']))
            if key in state['frames']:
                return
            try:
                comp = _make_compositor()
                state['frames'][key] = comp.render(_page_image(index))
                if len(state['frames']) > 8:
                    for old in list(state['frames'])[:-5]:
                        state['frames'].pop(old, None)
            except Exception:
                pass

        # ---------- 预览上直接拖分区 ----------
        # 左边界的黄线 = 左侧区宽度，右边界的黄线 = 右侧区宽度；
        # 按住拖动就能实时改，和上面两个百分比输入框是同一份数据。
        def _grip_x(which):
            box = state.get('preview_box')
            if not box:
                return None
            x0, _y0, w, _h = box
            if which == 'left':
                return x0 + w * max(0.0, state['left_ratio'])
            return x0 + w - w * max(0.0, state['right_ratio'])

        def _draw_grips():
            """预览上的可拖动标记：版式不同画的东西不同。"""
            preview.delete('grip')
            box = state.get('preview_box')
            if not box:
                return
            _x0, y0, _w, h = box
            if layout_var.get() == 'custom':
                rect = _box_pixels()
                if rect is None:
                    return
                bx0, by0, bx1, by1 = rect
                preview.create_rectangle(bx0, by0, bx1, by1,
                                         outline='#ffd54a', width=2,
                                         dash=(6, 4), tags=('grip',))
                for hx, hy in ((bx0, by0), (bx1, by0), (bx0, by1), (bx1, by1)):
                    preview.create_rectangle(hx - 6, hy - 6, hx + 6, hy + 6,
                                             fill='#ffd54a', outline='',
                                             tags=('grip',))
                return
            for which in ('left', 'right'):
                x = _grip_x(which)
                if x is None:
                    continue
                preview.create_line(x, y0, x, y0 + h, fill='#ffd54a',
                                    width=2, tags=('grip',))
                preview.create_rectangle(x - 5, y0 + h / 2.0 - 24,
                                         x + 5, y0 + h / 2.0 + 24,
                                         fill='#ffd54a', outline='',
                                         tags=('grip',))

        def _box_pixels():
            """PPT 框在预览画布里的像素矩形 (x0, y0, x1, y1)。"""
            box = state.get('preview_box')
            if not box:
                return None
            px, py, pw, ph = box
            sx, sy, sw, sh = state['slide_box']
            return (px + pw * sx, py + ph * sy,
                    px + pw * (sx + sw), py + ph * (sy + sh))

        def _grip_at(x):
            for which in ('left', 'right'):
                gx = _grip_x(which)
                if gx is not None and abs(x - gx) <= 10:
                    return which
            return None

        def _hit_test(x, y):
            """鼠标落在哪儿：分区黄线 / PPT 框的角 / PPT 框内部。"""
            if layout_var.get() == 'custom':
                rect = _box_pixels()
                if rect is None:
                    return None
                bx0, by0, bx1, by1 = rect
                for name, (hx, hy) in (('nw', (bx0, by0)), ('ne', (bx1, by0)),
                                       ('sw', (bx0, by1)), ('se', (bx1, by1))):
                    if abs(x - hx) <= 10 and abs(y - hy) <= 10:
                        return ('box', name)
                if bx0 <= x <= bx1 and by0 <= y <= by1:
                    return ('box', 'move')
                return None
            which = _grip_at(x)
            return ('grip', which) if which else None

        grip_job = {'job': None}

        def _schedule_drag_refresh():
            if grip_job['job'] is not None:
                return

            def run():
                grip_job['job'] = None
                _refresh_preview()
                _push_to_screen()
            grip_job['job'] = win.after(90, run)

        def _preview_motion(event):
            hit = _hit_test(event.x, event.y)
            cursor = 'arrow'
            if hit is not None:
                kind, part = hit
                if kind == 'grip':
                    cursor = 'sb_h_double_arrow'
                elif part == 'move':
                    cursor = 'fleur'
                elif part in ('nw', 'se'):
                    cursor = 'size_nw_se'
                else:
                    cursor = 'size_ne_sw'
            try:
                preview.configure(cursor=cursor)
            except tk.TclError:
                pass

        def _preview_press(event):
            state['drag'] = _hit_test(event.x, event.y)
            box = state.get('preview_box')
            if state['drag'] and box:
                px, py, pw, ph = box
                state['drag_origin'] = (
                    (event.x - px) / float(max(1, pw)),
                    (event.y - py) / float(max(1, ph)),
                    list(state['slide_box']))
            else:
                state['drag_origin'] = None
            if state['drag']:
                _preview_drag(event)

        def _preview_drag(event):
            drag = state.get('drag')
            box = state.get('preview_box')
            if not drag or not box:
                return
            kind, part = drag
            px, py, pw, ph = box
            if kind == 'grip':
                # 按像素跟手：内部存浮点比例，输入框只显示取整的百分比
                if part == 'left':
                    ratio = (event.x - px) / float(max(1, pw))
                else:
                    # 右边这条是“从右边缘往里”的宽度，方向别搞反
                    ratio = (px + pw - event.x) / float(max(1, pw))
                ratio = max(0.0, min(ratio, 0.45))
                other = (state['right_ratio'] if part == 'left'
                         else state['left_ratio'])
                if ratio + other > 0.90:        # 中间至少留 10%
                    ratio = max(0.0, 0.90 - other)
                if part == 'left':
                    state['left_ratio'] = ratio
                    target_var = left_var
                else:
                    state['right_ratio'] = ratio
                    target_var = right_var
                ratio_sync['busy'] = True
                try:
                    target_var.set(int(round(ratio * 100)))
                finally:
                    ratio_sync['busy'] = False
                _draw_grips()                   # 先给即时反馈
                _schedule_drag_refresh()        # 再按节流真正重新合成
                return
            if not state.get('drag_origin'):
                return
            rx = (event.x - px) / float(max(1, pw))
            ry = (event.y - py) / float(max(1, ph))
            rx0, ry0, box0 = state['drag_origin']
            left, top, width, height = box0
            if part == 'move':
                left = min(max(0.0, left + (rx - rx0)), max(0.0, 1.0 - width))
                top = min(max(0.0, top + (ry - ry0)), max(0.0, 1.0 - height))
            else:
                right = left + width
                bottom = top + height
                min_size = 0.04
                if 'w' in part:
                    left = min(max(0.0, rx), right - min_size)
                if 'e' in part:
                    right = max(min(1.0, rx), left + min_size)
                if 'n' in part:
                    top = min(max(0.0, ry), bottom - min_size)
                if 's' in part:
                    bottom = max(min(1.0, ry), top + min_size)
                width = right - left
                height = bottom - top
            state['slide_box'] = [left, top, width, height]
            _draw_grips()                       # 先给即时反馈
            _schedule_drag_refresh()            # 再按节流真正重新合成

        def _preview_release(_event=None):
            state['drag'] = None
            state['drag_origin'] = None

        preview.bind('<Motion>', _preview_motion)
        preview.bind('<Button-1>', _preview_press)
        preview.bind('<B1-Motion>', _preview_drag)
        preview.bind('<ButtonRelease-1>', _preview_release)

        def _refresh_preview():
            preview.delete('all')
            # 自定义版式下就算还没加载 PPT，也先把背景图显示出来，
            # 这样可以先调好 PPT 框的位置。
            if not _pages() and not (layout_var.get() == 'custom'
                                     and _custom_bg_image() is not None):
                state['thumb'] = None
                state['preview_box'] = None
                preview.create_text(max(180, preview.winfo_width() // 2),
                                    max(120, preview.winfo_height() // 2),
                                    text='选择 PPT 后自动预览',
                                    fill='#8fa4c8', font=ui_theme.font(10))
                _draw_film()
                return
            try:
                _comp, frame = _compose()
                box_w = max(240, preview.winfo_width())
                box_h = max(160, preview.winfo_height())
                scale = min(box_w / float(frame.width),
                            box_h / float(frame.height))
                thumb_w = max(1, int(frame.width * scale))
                thumb_h = max(1, int(frame.height * scale))
                thumb = frame.resize((thumb_w, thumb_h))
                from PIL import ImageTk
                photo = ImageTk.PhotoImage(thumb)
                state['thumb'] = photo
                x0 = (box_w - thumb_w) // 2
                y0 = (box_h - thumb_h) // 2
                state['preview_box'] = (x0, y0, thumb_w, thumb_h)
                preview.create_image(x0, y0, image=photo, anchor='nw')
                _draw_grips()
                _draw_film()
            except Exception as e:
                state['thumb'] = None
                state['preview_box'] = None
                preview.create_text(max(180, preview.winfo_width() // 2),
                                    max(120, preview.winfo_height() // 2),
                                    text='预览失败：%s' % e, fill='#ffb4b4',
                                    font=ui_theme.font(9))

        # ---------- 缩略图条：像 PPT 的幻灯片预览 ----------
        film_state = {'offset': 0.0, 'thumbs': {}, 'press': None,
                      'press_off': 0.0, 'moved': False}
        FILM_W, FILM_H, FILM_GAP = 120, 68, 10

        def _film_total():
            return len(_pages()) * (FILM_W + FILM_GAP) + FILM_GAP

        def _film_thumb(index):
            """第 index 页的小图（缓存，别重复解码大图）。"""
            key = (state['cur'], index)
            if key in film_state['thumbs']:
                return film_state['thumbs'][key]
            photo = None
            img = _page_image(index)
            if img is not None:
                try:
                    from PIL import ImageTk
                    small = img.copy()
                    small.thumbnail((FILM_W, FILM_H))
                    photo = ImageTk.PhotoImage(small)
                except Exception:
                    photo = None
            film_state['thumbs'][key] = photo
            return photo

        def _draw_film():
            """画缩略图条（只画可见范围，页多也不卡）。"""
            try:
                film.delete('all')
                width = max(60, film.winfo_width())
            except tk.TclError:
                return
            pages = _pages()
            total = _film_total()
            max_off = max(0.0, total - width)
            off = max(0.0, min(film_state['offset'], max_off))
            film_state['offset'] = off
            for i in range(len(pages)):
                x = FILM_GAP + i * (FILM_W + FILM_GAP) - off
                if x > width or x + FILM_W < 0:
                    continue
                current = (i == state['idx'])
                film.create_rectangle(
                    x - 2, 2, x + FILM_W + 2, FILM_H + 6,
                    outline='#ffd54a' if current else '#2b3f66',
                    width=3 if current else 1, fill='#101a2e')
                photo = _film_thumb(i)
                if photo is not None:
                    # 注意要 anchor='nw'：默认是中心点，图会跑到框外面去
                    film.create_image(
                        x + (FILM_W - photo.width()) // 2,
                        6 + (FILM_H - photo.height()) // 2,
                        image=photo, anchor='nw')
                else:
                    film.create_text(x + FILM_W / 2, 6 + FILM_H / 2,
                                     text='第 %d 页' % (i + 1),
                                     fill='#7c8797', font=ui_theme.font(8))
                film.create_text(x + FILM_W / 2, FILM_H + 18, text=str(i + 1),
                                 fill='#ffd54a' if current else '#8fa6cc',
                                 font=ui_theme.font(8))
            if not pages:
                film.create_text(max(80, width // 2), FILM_H / 2 + 6,
                                 text='加上 PPT 后这里显示每页缩略图',
                                 fill='#7c8797', font=ui_theme.font(9))

        def _film_follow():
            """翻页时把当前页挪进可视范围。"""
            try:
                width = max(60, film.winfo_width())
            except tk.TclError:
                return
            index = max(0, state['idx'])
            x0 = FILM_GAP + index * (FILM_W + FILM_GAP)
            if x0 - film_state['offset'] < 0:
                film_state['offset'] = max(0.0, x0 - FILM_GAP)
            elif x0 + FILM_W - film_state['offset'] > width:
                film_state['offset'] = x0 + FILM_W + FILM_GAP - width
            _draw_film()

        def _film_index_at(x):
            off = film_state['offset']
            index = int((x + off - FILM_GAP) // (FILM_W + FILM_GAP))
            return index if 0 <= index < len(_pages()) else -1

        def _film_click(event):
            if film_state['moved']:
                return
            index = _film_index_at(event.x)
            if index >= 0:
                _set_page(index)

        def _film_wheel(event):
            step = -1 if event.delta > 0 else 1
            film_state['offset'] += step * (FILM_W + FILM_GAP)
            _draw_film()

        def _film_press(event):
            film_state['press'] = event.x
            film_state['press_off'] = film_state['offset']
            film_state['moved'] = False

        def _film_drag(event):
            if film_state['press'] is None:
                return
            dx = event.x - film_state['press']
            if abs(dx) > 4:
                film_state['moved'] = True
            film_state['offset'] = film_state['press_off'] - dx
            _draw_film()

        def _film_release(_event=None):
            film_state['press'] = None

        film.bind('<Button-1>', _film_press)
        film.bind('<B1-Motion>', _film_drag)
        film.bind('<ButtonRelease-1>', lambda e: (_film_release(),
                                                  _film_click(e)))
        film.bind('<MouseWheel>', _film_wheel)
        film.bind('<Configure>', lambda e: _draw_film())

        def _push_to_screen():
            screen = self._strip_screen
            if screen is None or not screen.is_open():
                return
            try:
                _comp, frame = _compose()
                screen.show(frame)
            except Exception as e:
                state_label.config(text='投屏失败：%s' % e, fg=p['danger'])

        # ---------- 原样放映（WPS/PowerPoint 直接放，带中间框） ----------
        slide_player = {'obj': None, 'poll': None, 'last_pos': 0}

        def _live_active():
            return slide_player.get('obj') is not None

        def _live_box_rect():
            """中间 PPT 框在大屏上的绝对矩形（屏幕坐标）。"""
            rect = _monitor_rect()['rect']
            comp = _make_compositor()
            x, y, w, h = comp.slide_rect()
            return (rect[0] + x, rect[1] + y,
                    rect[0] + x + w, rect[1] + y + h)

        def _start_live():
            """起放映窗口并摆到中间框里；成不成都要给个明确结果。"""
            try:
                deck = _deck()
                if deck is None:
                    return False
                if _warm_screen() is None:
                    state_label.config(text='打不开投屏窗口', fg=p['danger'])
                    return False
                player = ppt_converter.SlideShowPlayer()
                box = _live_box_rect()
                if not player.start(deck['path'], rect=box, topmost=True):
                    state_label.config(
                        text='原样放映没起来：%s'
                             % (player.last_error or '未知原因'),
                        fg=p['danger'])
                    return False
                slide_player['obj'] = player
                slide_player['last_pos'] = 0
                _live_poll()
                return True
            except Exception as exc:
                state_label.config(text='原样放映失败：%s' % exc, fg=p['danger'])
                return False

        def _stop_live():
            job = slide_player.get('poll')
            if job is not None:
                try:
                    win.after_cancel(job)
                except tk.TclError:
                    pass
                slide_player['poll'] = None
            player = slide_player.get('obj')
            slide_player['obj'] = None
            if player is not None:
                try:
                    player.close()
                except Exception:
                    pass

        def _live_poll():
            """盯着放映页码：翻页笔直接翻，界面也跟着同步。"""
            player = slide_player.get('obj')
            if player is None:
                return
            pos = player.position()
            if pos and pos != slide_player.get('last_pos'):
                slide_player['last_pos'] = pos
                if 0 < pos <= len(_pages()):
                    state['idx'] = pos - 1
                    _update_page_label()
                    _film_follow()
                    _push_to_screen()
                    _refresh_preview()
            try:
                slide_player['poll'] = win.after(500, _live_poll)
            except tk.TclError:
                slide_player['poll'] = None

        def _update_page_label():
            total = len(_pages())
            current = state['idx'] + 1 if state['idx'] >= 0 else 0
            deck = _deck()
            if deck is not None:
                text = '第 %d / %d 页（第 %d/%d 份）' % (
                    current, total, state['cur'] + 1, len(state['decks']))
            else:
                text = '第 0 / 0 页'
            page_label.config(text=text)

        def _set_page(index):
            """在当前这份 PPT 里翻页。"""
            total = len(_pages())
            if total <= 0:
                return
            index = max(0, min(int(index), total - 1))
            if _live_active() and index != state['idx']:
                # 原样放映：翻页交给 WPS/PowerPoint（动画/视频不中断）
                try:
                    slide_player['obj'].goto(index + 1)
                except Exception:
                    pass
            state['idx'] = index
            _update_page_label()
            _film_follow()
            _push_to_screen()
            _refresh_preview()
            # 顺手把下一页/上一页先合成好，翻页时立刻出画面
            try:
                win.after(150, lambda i=index + 1: _prefetch(i))
                win.after(200, lambda i=index - 1: _prefetch(i))
            except tk.TclError:
                pass
            page = _page_image(index)
            if page is None:
                state_label.config(text='第 %d 页打不开' % (index + 1),
                                   fg=p['danger'])

        def _step(delta):
            """翻页：这份 PPT 翻完自动接上一份/下一份，按顺序依次继续。"""
            player = slide_player.get('obj')
            if player is not None:
                # 原样放映：翻页由 WPS/PowerPoint 执行，页码由轮询同步回来
                if delta > 0:
                    player.next()
                else:
                    player.prev()
                if delta > 0 and slide_player.get('last_pos', 0) >= len(_pages()):
                    slide_player['last_pos'] = 0
                else:
                    slide_player['last_pos'] = 0
                return
            decks = state['decks']
            if not decks or state['cur'] < 0:
                return
            total = len(_pages())
            index = (state['idx'] if state['idx'] >= 0 else 0) + delta
            if 0 <= index < total:
                _set_page(index)
                return
            nxt = state['cur'] + delta
            if nxt < 0 or nxt >= len(decks):
                if not loop_var.get():
                    return
                nxt = (nxt + len(decks)) % len(decks)
            _open_deck(nxt, 0 if delta > 0 else -1)

        def _sync_auto():
            """自动翻页定时器。"""
            if state['auto_job'] is not None:
                try:
                    win.after_cancel(state['auto_job'])
                except tk.TclError:
                    pass
                state['auto_job'] = None
            if auto_var.get() and _pages() and _alive():
                delay = max(3, int(interval_var.get())) * 1000
                state['auto_job'] = win.after(delay, _auto_tick)
            state_label.config(text='自动翻页：开' if auto_var.get()
                               else '自动翻页：关')

        def _auto_tick():
            state['auto_job'] = None
            if not _alive() or not auto_var.get():
                return
            _step(1)
            _sync_auto()

        def _render_decks(select=None):
            """把 PPT 列表画到左边的小列表里。"""
            deck_tree.delete(*deck_tree.get_children())
            for i, deck in enumerate(state['decks']):
                deck_tree.insert('', 'end', iid=str(i),
                                 values=(i + 1, deck['name'],
                                         len(deck['pages'])))
            if select is not None and 0 <= select < len(state['decks']):
                deck_tree.selection_set(str(select))
                deck_tree.see(str(select))

        def _selected_deck():
            sel = deck_tree.selection()
            return int(sel[0]) if sel else -1

        def _open_deck(index, page=0):
            """切到第 index 份 PPT（双击列表 / 翻页到头时用）。"""
            decks = state['decks']
            if not (0 <= index < len(decks)):
                return
            state['frames'] = {}
            deck = decks[index]
            total = len(deck['pages'])
            if total <= 0:
                return
            state['cur'] = index
            state['idx'] = -1
            state['cache'] = {}
            film_state['offset'] = 0.0      # 换了一份 PPT：缩略图从第一页看起
            if page < 0:
                page = total - 1
            _render_decks(index)
            ppt_var.set('%s（共 %d 页）' % (deck['name'], total))
            state_label.config(text='当前：%s（%d/%d 份）'
                                    % (deck['name'], index + 1, len(decks)))
            _set_page(page)

        def _on_deck_double(_event=None):
            i = _selected_deck()
            if i >= 0:
                _open_deck(i, 0)

        def _remove_ppt():
            i = _selected_deck()
            if i < 0:
                return
            del state['decks'][i]
            state['frames'] = {}
            if not state['decks']:
                state['cur'] = -1
                state['idx'] = -1
                state['cache'] = {}
                ppt_var.set('未选择')
                _render_decks()
                _update_page_label()
                _refresh_preview()
                return
            _open_deck(min(i, len(state['decks']) - 1), 0)

        def _clear_ppt():
            state['decks'] = []
            state['frames'] = {}
            state['cur'] = -1
            state['idx'] = -1
            state['cache'] = {}
            ppt_var.set('未选择')
            _render_decks()
            _update_page_label()
            _refresh_preview()

        def _pick_ppt():
            files = filedialog.askopenfilenames(
                parent=win, title='选择要在中间播放的 PPT（可多选）',
                filetypes=[('PPT 文件', '*.ppt *.pptx'),
                           ('所有文件', '*.*')])
            if not files:
                return
            info = _monitor_rect()
            rect = info['rect']
            mid_w = int((rect[2] - rect[0]) *
                        (1 - (state['left_ratio'] + state['right_ratio'])))
            width = max(1280, min(mid_w, 2560))
            added = []
            for pos, src in enumerate(files):
                stem = os.path.splitext(os.path.basename(src))[0]
                out_dir = os.path.join(os.path.dirname(src),
                                       stem + '_长条屏图片')
                state_label.config(
                    text='正在导出 PPT 页面…（%d/%d，这一步用 WPS/Office）'
                         % (pos + 1, len(files)))
                dot.config(fg=p['faint'])
                win.update_idletasks()
                try:
                    out_dir, count, paths = \
                        ppt_converter.convert_ppt_to_images(
                            src, out_dir, width=width,
                            on_progress=lambda cur, total, msg,
                            p=pos: (state_label.config(
                                text='导出第 %d/%d 份：%d/%d'
                                     % (p + 1, len(files), cur, total)),
                                win.update_idletasks()))
                except ppt_converter.PptConvertError as e:
                    messagebox.showerror('导出失败', str(e), parent=win)
                    return
                except Exception as e:
                    messagebox.showerror('导出失败', '发生未知错误：\n%s' % e,
                                         parent=win)
                    return
                state['decks'].append({
                    'path': src, 'name': os.path.basename(src),
                    'pages': list(paths)})
                added.append(len(state['decks']) - 1)
                if state['cur'] < 0:
                    state['cur'] = added[0]
            _render_decks(state['cur'])
            state['frames'] = {}
            _warm_screen()
            if self._strip_screen is not None:
                state_label.config(
                    text='大屏已就绪（正显示第 1 页）；已加入 %d 份 PPT，'
                         '双击切换，↑ ↓ 翻页' % len(added))
            else:
                state_label.config(
                    text='已加入 %d 份 PPT，双击列表可切换；'
                         '↑ ↓ / ← → / PgUp PgDn 翻页' % len(added))
            if state['idx'] < 0:
                _set_page(0)
            else:
                _update_page_label()

        def _upload_side():
            src = filedialog.askopenfilename(
                parent=win, title='选择两侧用的图片',
                filetypes=[('图片文件', '*.png *.jpg *.jpeg *.bmp *.webp'),
                           ('所有文件', '*.*')])
            if not src:
                return
            stem = os.path.splitext(os.path.basename(src))[0]
            name = simpledialog.askstring('图片名称', '给这张图起个名字：',
                                          initialvalue=stem, parent=win)
            if not name:
                return
            try:
                item = ppt_converter.add_uploaded_background(src, name)
            except Exception as e:
                messagebox.showerror('上传失败', str(e), parent=win)
                return
            bg_state['items'] = ppt_converter.list_uploaded_backgrounds()
            names = ['（不用图）'] + [x['name'] for x in bg_state['items']]
            for combo in (left_combo, right_combo):
                combo.configure(values=names)
            left_img_var.set(item['name'])
            state['frames'] = {}
            state_label.config(text='图片已保存：%s' % name)
            _refresh_preview()

        def _warm_screen():
            """提前把投屏窗口开出来（只放当前页），正式开始时不用等窗口创建。"""
            if not _pages():
                return None
            info = _monitor_rect()
            rect = info['rect']
            screen = self._strip_screen
            if screen is not None and screen.is_open() and \
                    tuple(screen.rect) == tuple(rect):
                return screen
            try:
                if screen is not None:
                    screen.close(notify=False)
                self._strip_screen = strip_screen.StripScreen(
                    self.root, rect,
                    on_close=_on_screen_closed,
                    on_next=lambda: _step(1),
                    on_prev=lambda: _step(-1),
                    on_toggle_auto=lambda: (
                        auto_var.set(not auto_var.get()), _sync_auto()))
            except Exception:
                self._strip_screen = None
            if self._strip_screen is not None:
                state_label.config(
                    text='大屏已就绪（显示当前页），点“开始投屏”立即切换',
                    fg=p['muted'])
                dot.config(fg=p['success'])
            return self._strip_screen

        def _start_screen():
            if not _pages():
                messagebox.showinfo('提示', '请先点“添加 PPT（可多选）…”。',
                                    parent=win)
                return
            info = _monitor_rect()
            rect = info['rect']
            if _warm_screen() is None:
                messagebox.showerror('投屏失败', '打不开投屏窗口。', parent=win)
                return
            if state['idx'] < 0:
                state['idx'] = 0
            if live_var.get():
                # 原样放映：动画/切换/内嵌视频音频都由 WPS/PowerPoint 负责
                if _start_live():
                    start_btn.configure(text='重新开始投屏')
                    dot.config(fg=p['success'])
                    state_label.config(
                        text='原样放映中（动画/视频都在）：翻页、翻页笔都能用',
                        fg=p['success'])
                    _sync_auto()
                    return
                if not messagebox.askyesno(
                        '原样放映没起来',
                        '这台电脑上的 WPS/PowerPoint 没能把放映窗口摆进中间框。'
                        '\n\n改成“图片合成”方式投屏吗？'
                        '（图片方式没有动画和视频）', parent=win):
                    return
            _set_page(state['idx'])
            start_btn.configure(text='重新开始投屏')
            dot.config(fg=p['success'])
            state_label.config(text='正在投屏：%s（%d×%d）'
                                    % (info.get('name', ''), rect[2] - rect[0],
                                       rect[3] - rect[1]))

        def _stop_screen():
            _stop_live()
            if self._strip_screen is not None:
                self._strip_screen.close(notify=False)
                self._strip_screen = None
            dot.config(fg=p['faint'])
            state_label.config(text='已停止投屏')

        def _on_screen_closed():
            self._strip_screen = None
            if not _alive():
                return
            dot.config(fg=p['faint'])
            state_label.config(text='投屏已退出（按 ESC）')

        # 列表交互：双击一份 PPT 就切过去；键盘上下/翻页键翻页
        deck_tree.bind('<Double-Button-1>', _on_deck_double)

        def _key_guard(_event=None):
            """输入框/下拉框里按上下键是调数值，别顺带翻页。"""
            try:
                focused = win.focus_get()
            except (tk.TclError, KeyError):
                focused = None
            cls = focused.winfo_class() if focused is not None else ''
            return any(cls.endswith(name) for name in
                       ('Entry', 'Spinbox', 'Combobox', 'Text', 'Listbox'))

        def _bind_key(seq, delta):
            def handler(event=None):
                if _key_guard(event):
                    return None
                _step(delta)
                return 'break'
            win.bind(seq, handler)
            return handler

        for _seq, _delta in (('<Down>', 1), ('<Up>', -1),
                             ('<Next>', 1), ('<Prior>', -1),
                             ('<Right>', 1), ('<Left>', -1)):
            _bind_key(_seq, _delta)

        _update_page_label()
        _render_decks()
        _sync_bg_list()
        _auto_box()
        _sync_layout()
        win._stop_live = _stop_live        # 关窗口时把原样放映一起关掉
        # 窗口拉伸时预览跟着重画（黄线位置也一起更新）
        preview.bind('<Configure>', lambda e: _refresh_preview(), add='+')

    def _close_strip_win(self):
        """关闭长条屏窗口：一并关掉投屏。"""
        stop_live = getattr(self._strip_win, '_stop_live', None)
        if callable(stop_live):
            try:
                stop_live()
            except Exception:
                pass
        if self._strip_screen is not None:
            try:
                self._strip_screen.close(notify=False)
            except Exception:
                pass
            self._strip_screen = None
        win = self._strip_win
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._strip_win = None

    # ------------------------------------------------------------------
    # 常用工具快捷启动器已下线（媒体处理改用内置 图片工具箱/音视频剪辑）
    # ------------------------------------------------------------------
    def _open_qrcode(self):
        """打开“二维码生成”子窗口：链接 / 文本转二维码，可加 logo。"""
        if self._activate_or_close_others('qrcode'):
            return
        win = self._qrcode_win
        if win is not None:
            try:
                alive = bool(win.winfo_exists())
            except tk.TclError:
                alive = False
            if alive:
                win.deiconify()
                win.lift()
                win.focus_force()
                return
            self._qrcode_win = None

        p = ui_theme.PALETTE
        win = tk.Toplevel(self.root)
        win.title('二维码生成')
        _center_window(win, 640, 600)
        win.resizable(False, False)
        ui_theme.style_window(win)
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW', self._close_qrcode_win)
        self._qrcode_win = win

        state = {'logo': None}          # 当前 logo 路径
        self._qr_photo = None           # 预览图引用（防被回收）
        self._qr_img = None             # 当前生成的 PIL 图（保存用）
        self._qr_text = None            # 当前生成内容（高清重生成用）

        tk.Label(win, text='内容（链接或普通文本）',
                 font=('微软雅黑', 10)).pack(anchor='w', padx=20,
                                             pady=(16, 4))
        text_box = tk.Text(win, font=('微软雅黑', 11), height=3,
                           wrap='char', relief='solid', bd=1,
                           highlightthickness=0)
        text_box.pack(fill=tk.X, padx=20)

        hint_label = tk.Label(win, text='链接会被手机直接识别为可打开的网页',
                              font=ui_theme.font(8), fg=p['faint'],
                              bg=p['bg'])
        hint_label.pack(anchor='w', padx=20, pady=(4, 0))

        def _pick_logo():
            path = filedialog.askopenfilename(
                parent=win, title='选择 logo 图片',
                filetypes=[('图片', '*.png *.jpg *.jpeg *.bmp *.gif '
                                    '*.webp'), ('所有文件', '*.*')])
            if not path:
                return
            state['logo'] = path
            name = os.path.basename(path)
            logo_label.config(
                text=name if len(name) <= 22 else name[:19] + '…',
                fg=p['text'])
            clear_btn.state(['!disabled'])

        def _clear_logo():
            state['logo'] = None
            logo_label.config(text='未选择（可不上传）', fg=p['muted'])
            clear_btn.state(['disabled'])

        logo_row = ttk.Frame(win)
        logo_row.pack(fill=tk.X, padx=20, pady=(10, 0))
        ttk.Button(logo_row, text='上传 logo…', style='Ghost.TButton',
                   command=_pick_logo).pack(side=tk.LEFT)
        clear_btn = ttk.Button(logo_row, text='清除', style='Ghost.TButton',
                               command=_clear_logo, state='disabled')
        clear_btn.pack(side=tk.LEFT, padx=(8, 0))
        logo_label = tk.Label(logo_row, text='未选择（可不上传）',
                             font=ui_theme.font(9), fg=p['muted'],
                             bg=p['bg'])
        logo_label.pack(side=tk.LEFT, padx=(10, 0))

        status_var = tk.StringVar(
            value='输入内容后点“生成二维码”')
        tk.Label(win, textvariable=status_var, font=ui_theme.font(9),
                 fg=p['muted'], bg=p['bg']).pack(pady=(10, 0))

        # 按钮先按底部 pack，保证二维码再大也不会把按钮挤出窗口
        btn_row = ttk.Frame(win)
        btn_row.pack(side=tk.BOTTOM, pady=(6, 16))

        qr_label = tk.Label(win, bg=p['bg'],
                            text='二维码预览', font=ui_theme.font(10),
                            fg=p['faint'], width=28, height=11)
        qr_label.pack(expand=True, pady=(8, 6))

        def _generate():
            from PIL import ImageTk
            text = text_box.get('1.0', 'end').strip()
            if not text:
                status_var.set('请先输入内容')
                return
            ecc = 'H' if state['logo'] else 'M'
            try:
                matrix = qrcode_lite.make_matrix(text, ecc=ecc)
            except ValueError as e:
                status_var.set(str(e))
                return
            except Exception as e:
                status_var.set('生成失败：%s' % e)
                return
            scale = max(3, int(300 // (len(matrix) + 12)))
            img = qrcode_lite.to_image(matrix, scale=scale)
            if state['logo']:
                try:
                    img = qrcode_lite.paste_logo(img, state['logo'])
                except Exception as e:
                    status_var.set('logo 处理失败：%s' % e)
                    return
            self._qr_photo = ImageTk.PhotoImage(img)
            qr_label.configure(image=self._qr_photo, width=0, height=0,
                               text='')
            self._qr_img = img
            self._qr_text = text
            ver = (len(matrix) - 17) // 4
            status_var.set('已生成（版本 %d，%s 级纠错%s）'
                           % (ver,
                              qrcode_lite._pick_version(text, ecc)[0],
                              '，含 logo' if state['logo'] else ''))

        def _save():
            if self._qr_img is None:
                status_var.set('请先生成二维码')
                return
            path = filedialog.asksaveasfilename(
                parent=win, title='保存二维码',
                defaultextension='.png',
                initialfile='二维码.png',
                filetypes=[('PNG 图片', '*.png')])
            if not path:
                return
            # 重新按高清尺寸生成（预览图是按屏幕缩放的）
            try:
                matrix = qrcode_lite.make_matrix(
                    self._qr_text, ecc='H' if state['logo'] else 'M')
                img = qrcode_lite.to_image(matrix, scale=24)
                if state['logo']:
                    img = qrcode_lite.paste_logo(img, state['logo'])
                img.convert('RGB').save(path)
            except Exception as e:
                messagebox.showerror('保存失败', str(e), parent=win)
                return
            status_var.set('已保存：%s' % path)
            safe_guard.write_log('二维码已保存：%s' % path)

        ttk.Button(btn_row, text='生成二维码', style='Accent.TButton',
                   command=_generate).pack(side=tk.LEFT)
        ttk.Button(btn_row, text='保存高清 PNG', style='Ghost.TButton',
                   command=_save).pack(side=tk.LEFT, padx=(10, 0))

    def _close_qrcode_win(self):
        """关闭二维码生成窗口。"""
        self._qr_photo = None
        self._qr_img = None
        self._qr_text = None
        win = self._qrcode_win
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._qrcode_win = None

    def _open_downloader(self):
        """打开“软件下载”子窗口：winget 源分类推荐 + 搜索下载。"""
        if self._activate_or_close_others('downloader'):
            return
        win = self._downloader_win
        if win is not None:
            try:
                alive = bool(win.winfo_exists())
            except tk.TclError:
                alive = False
            if alive:
                win.deiconify()
                win.lift()
                win.focus_force()
                return
            self._downloader_win = None

        win = tk.Toplevel(self.root)
        win.title('软件下载')
        _center_window(win, 800, 600)
        win.resizable(True, True)
        win.minsize(700, 520)
        ui_theme.style_window(win)
        win.transient(self.root)
        win.protocol('WM_DELETE_WINDOW', self._close_downloader_win)
        for base in (_BASE_DIR, os.path.dirname(os.path.abspath(__file__))):
            icon = os.path.join(base, '远程打印.ico')
            if os.path.exists(icon):
                try:
                    win.iconbitmap(icon)
                except Exception:
                    pass
                break
        self._downloader_win = win
        try:
            pane = winget_downloader.DownloaderPane(win)
        except Exception as e:
            messagebox.showerror(
                '软件下载', '软件下载打开失败：\n%s' % e, parent=win)
            self._close_downloader_win()
            return
        pane.pack(fill=tk.BOTH, expand=True)
        win.deiconify()
        win.lift()
        win.focus_force()

    def _close_downloader_win(self):
        """关闭软件下载子窗口。"""
        win = self._downloader_win
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._downloader_win = None

    def _open_screenshot(self):
        """打开截图子窗口（仿 ShareX 的四种模式 + 简易标注 + 全局热键）。"""
        if self._activate_or_close_others('screenshot'):
            return
        if _screenshot is None:
            messagebox.showerror(
                '截图', '截图模块加载失败，请检查 screenshot.py。',
                parent=self.root)
            return

        def _on_close():
            self._screenshot_win = None

        try:
            panel = _screenshot.open_panel(self.root, on_close=_on_close)
        except Exception as e:
            messagebox.showerror(
                '截图', '打开截图窗口失败：\n%s' % e, parent=self.root)
            return
        # 记下 panel.win，让 _activate_or_close_others('screenshot')
        # 下次点卡片时能直接复用这个窗口；关闭时由 on_close 清空
        self._screenshot_win = panel.win

    def _close_screenshot_win(self):
        """关闭截图子窗口（也用于切换到其它子窗口时主动关掉）。"""
        win = self._screenshot_win
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._screenshot_win = None

    def _open_music(self):


        """打开“会务音乐”子窗口：本地音频离线播放（Windows MCI，零依赖）。"""
        win = self._music_win
        if win is not None:
            try:
                if win.winfo_exists():
                    win.deiconify()
                    win.lift()
                    win.focus_force()
                    return
            except tk.TclError:
                pass

        win = tk.Toplevel(self.root)
        win.title('会务音乐')
        win.configure(bg=WINDOW_BG)
        win.resizable(False, False)
        ui_theme.style_window(win)
        for base in (_BASE_DIR, os.path.dirname(os.path.abspath(__file__))):
            icon = os.path.join(base, '远程打印.ico')
            if os.path.exists(icon):
                try:
                    win.iconbitmap(icon)
                except Exception:
                    pass
                break
        self._music_win = win

        state = {'list': [], 'names': [], 'idx': -1, 'playing': False,
                 'manual': False, 'loop': True, 'seek': False,
                 'play_t0': 0.0}   # 最近一次点播的时间，用于判断“没播起来”
        win._music_state = state

        # —— 播放引擎：Windows Media Player OCX（系统自带，mp3/wav 均可） ——
        try:
            import win32com.client
            wmp = win32com.client.Dispatch('WMPlayer.OCX')
        except Exception:
            messagebox.showerror('会务音乐', '无法初始化播放组件，请检查系统播放器。')
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._music_win = None
            return
        wmp.settings.volume = 100
        win._music_wmp = wmp

        def _play_state():
            """playState：3=播放中 2=已暂停 8=播放完 1=已停止。"""
            try:
                return int(wmp.playState)
            except Exception:
                return 0

        def _win_alive():
            """窗口是否还在：子窗口关闭后定时器要立刻停下来。"""
            try:
                return bool(win.winfo_exists())
            except tk.TclError:
                return False

        # 播放/暂停按钮图标（在界面区创建按钮后生效）
        ICON_PLAY = '\u25b6'    # ▶
        ICON_PAUSE = '\u23f8'   # ⏸

        def _set_play_icon(playing):
            """playing=True 显示暂停图标（再点一下会暂停），否则显示播放图标。"""
            try:
                btn_play.config(text=ICON_PAUSE if playing else ICON_PLAY)
            except Exception:
                pass

        # —— 停止时 2 秒淡出，不“啪”地切断声音 ——
        fade = {'job': None}  # 进行中的淡出定时器

        def _cancel_fade(restore=True):
            """取消淡出；restore=True 时把音量恢复到 100。"""
            if fade['job'] is not None:
                try:
                    win.after_cancel(fade['job'])
                except Exception:
                    pass
                fade['job'] = None
            if restore:
                try:
                    wmp.settings.volume = 100
                except Exception:
                    pass

        def _fade_out(done=None):
            """音量在 2 秒内线性降到 0 再真正停止；done 为停止后的回调。"""
            _cancel_fade()
            try:
                start = int(wmp.settings.volume) or 100
            except Exception:
                start = 100
            steps = {'left': 20, 'vol': start}  # 20 步 × 100ms = 2 秒

            def _tick():
                steps['left'] -= 1
                try:
                    wmp.settings.volume = max(
                        0, int(round(steps['vol'] * steps['left'] / 20.0)))
                except Exception:
                    pass
                if steps['left'] > 0 and _win_alive():
                    fade['job'] = win.after(100, _tick)
                else:
                    fade['job'] = None
                    try:
                        wmp.controls.stop()
                    except Exception:
                        pass
                    try:
                        wmp.settings.volume = 100
                    except Exception:
                        pass
                    if done is not None:
                        try:
                            done()
                        except Exception:
                            pass

            fade['job'] = win.after(100, _tick)

        def _current_tab():
            """当前页签：0=本地播放，1=在线音乐。控件只按它路由。"""
            try:
                return nb.index(nb.select())
            except Exception:
                return 0

        def _now_name():
            # 播放栏只呈现“当前页签”的曲目，本地/在线互不干扰
            if _current_tab() == 1:
                try:
                    i = online['sel'][online['pos']]
                    return res_list[i][0]
                except Exception:
                    return '未选择在线曲目'
            st = state
            if 0 <= st['idx'] < len(st['names']):
                return st['names'][st['idx']]
            return '未选择曲目'

        def _refresh():
            lbl_now.config(text='♪ ' + _now_name())
            _sync_status()

        def _sync_status():
            """按引擎真实状态刷新状态文字与播放/暂停图标。

            图标每 400ms 随 _poll 重新校准，不依赖点击瞬间的瞬时值，
            避免“开始播放后图标不变”的不同步问题。
            """
            st = state
            if fade['job'] is not None:
                lbl_status.config(text='淡出停止中…', fg='#888888')
                _set_play_icon(False)
                return
            if online.get('loading'):
                lbl_status.config(text='网络缓冲中…', fg='#d97706')
                _set_play_icon(False)
                return
            if _current_tab() == 1:
                # 在线页：只反映在线播放状态，本地播没播与本页无关
                if not online.get('playing'):
                    lbl_status.config(text='已停止', fg='#888888')
                    _set_play_icon(False)
                    return
                ps = _play_state()
                if ps == 2 or online.get('manual'):
                    lbl_status.config(text='已暂停（在线）', fg='#888888')
                    _set_play_icon(False)
                elif ps == 8:
                    lbl_status.config(text='已停止', fg='#888888')
                    _set_play_icon(False)
                else:
                    # 3=播放 6=缓冲 7=等待 9=切歌 11=重连，都算在响
                    lbl_status.config(text='播放中（在线）', fg='#2e8b57')
                    _set_play_icon(True)
                return
            st = state
            if not st['playing']:
                lbl_status.config(text='已停止', fg='#888888')
                _set_play_icon(False)
                return
            if st['manual']:
                lbl_status.config(text='已暂停', fg='#888888')
                _set_play_icon(False)
            elif _play_state() == 8:
                lbl_status.config(text='已停止', fg='#888888')
                _set_play_icon(False)
            elif _play_stalled(st):
                # 媒体载入后停在 Ready(10)/Stopped(1)：设备没就绪或格式不支持，
                # 不能继续显示“播放中”骗人
                lbl_status.config(
                    text=('无声（没有音频输出设备）' if not has_audio_output()
                          else '没能播放（格式或版权限制）'), fg='#c62828')
                _set_play_icon(False)
            else:
                lbl_status.config(text='播放中', fg='#2e8b57')
                _set_play_icon(True)

        def _play_stalled(st, grace=4.0):
            """播放请求发出后过了 grace 秒还没真正开始 → 判定没播起来。"""
            if time.monotonic() - float(st.get('play_t0') or 0) < grace:
                return False
            return _play_state() in (0, 1, 10)

        def _play_idx(i):
            st = state
            if not (0 <= i < len(st['list'])):
                return
            online['playing'] = False  # 本地播放时停止在线播放
            _clear_online_loading()
            st['idx'] = i
            st['manual'] = False
            _cancel_fade()  # 切歌/重播：取消淡出并恢复音量
            st['play_t0'] = time.monotonic()
            try:
                wmp.URL = st['list'][i]  # 设置播放源
                wmp.controls.play()
            except Exception:
                messagebox.showwarning(
                    '会务音乐',
                    '无法播放：\n%s\n\n可能不是支持的音频格式。'
                    % os.path.basename(st['list'][i]))
                return
            st['playing'] = True
            _refresh()
            _sync_selection()

        def _toggle():
            """播放/暂停：只作用于当前页签的播放器，本地/在线互不串扰。"""
            st = state
            if fade['job'] is not None:
                # 淡出途中按“播放/暂停”→ 取消淡出，恢复当前页签的播放
                _cancel_fade()
                if _current_tab() == 1:
                    sel = online['sel']
                    if sel:
                        pos = online['pos']
                        if not (0 <= pos < len(sel)):
                            pos = 0
                        _online_play(pos)
                    else:
                        _refresh()
                elif st['list']:
                    _play_idx(st['idx'] if st['idx'] >= 0 else 0)
                else:
                    _refresh()
                return
            if _current_tab() == 1:
                # —— 在线音乐页：只控制在线曲目 ——
                if online['playing']:
                    # 在响/缓冲 → 暂停；2=已暂停 → 继续
                    ps = _play_state()
                    if ps == 2 or online.get('manual'):
                        try:
                            wmp.controls.play()
                        except Exception:
                            pass
                        online['manual'] = False
                    else:
                        try:
                            wmp.controls.pause()
                        except Exception:
                            pass
                        online['manual'] = True
                    _refresh()
                    return
                sel = online['sel']
                if sel:
                    # 还没开播：直接播放选中集合里的当前位置
                    pos = online['pos']
                    if not (0 <= pos < len(sel)):
                        pos = 0
                    _online_play(pos)
                else:
                    messagebox.showinfo(
                        '在线音乐',
                        '请先在搜索结果里勾选曲目，或双击一首试听。')
                return
            # —— 本地播放页：只控制本地列表 ——
            if st['idx'] < 0:
                if st['list']:
                    _play_idx(0)
                else:
                    messagebox.showinfo('会务音乐',
                                        '请先点“选歌”添加音频。')
                return
            ps = _play_state()
            if ps == 3:  # 播放中 → 暂停
                wmp.controls.pause()
                st['manual'] = True
            elif ps == 2:  # 已暂停 → 继续
                wmp.controls.play()
                st['manual'] = False
            else:
                _play_idx(st['idx'])
            _refresh()
            _sync_selection()

        def _stop():
            st = state
            online['playing'] = False
            online['manual'] = False
            _clear_online_loading()
            st['manual'] = True
            st['playing'] = False
            if _play_state() == 3 and fade['job'] is None:
                # 正在响 → 2 秒淡出后再停，不突然切断声音
                _fade_out(_refresh)
            else:
                _cancel_fade()
                try:
                    wmp.controls.stop()
                except Exception:
                    pass
                _refresh()
            _sync_selection()

        def _step(delta):
            """上一首/下一首：只切当前页签的曲目，绝不串到另一页。

            手动切歌不受“是否正在播放 / 是否循环 / 是否只有一首”阻断：
            有可选曲目就一定响应（边界重播当前曲），没有才提示。
            """
            if _current_tab() == 1:
                # —— 在线音乐页：只在选中集合内切歌 ——
                sel = online['sel']
                if not sel:
                    messagebox.showinfo(
                        '在线音乐',
                        '请先在搜索结果里勾选曲目，或双击一首试听。')
                    return
                n = len(sel)
                if not (0 <= online['pos'] < n):
                    nxt = 0 if delta > 0 else n - 1
                elif online['loop']:
                    nxt = (online['pos'] + delta) % n
                else:
                    nxt = max(0, min(online['pos'] + delta, n - 1))
                _online_play(nxt)
                return
            # —— 本地播放页：只切本地列表 ——
            st = state
            if not st['list']:
                messagebox.showinfo('会务音乐',
                                    '本地列表还没有音频，请先点“选歌”。')
                return
            n = len(st['list'])
            if not (0 <= st['idx'] < n):
                nxt = 0 if delta > 0 else n - 1
            elif st['loop']:
                nxt = (st['idx'] + delta) % n
            else:
                nxt = max(0, min(st['idx'] + delta, n - 1))
            _play_idx(nxt)

        def _poll():
            st = state
            if not _win_alive():
                return
            _sync_progress()
            _refresh()  # 曲名 + 状态文字 + 播放/暂停图标整体校准
            if online['playing']:
                if online['loading']:
                    ps = _play_state()
                    if ps == 3:  # 3=播放中：缓冲结束，真正响起来了
                        title = online['load_title']
                        _clear_online_loading()
                        lbl_dl.config(text='在线播放：' + title,
                                      fg='#2e8b57')
                    elif (ps in (0, 1)
                          and time.monotonic() - online['load_t0'] > 15):
                        # 15 秒仍是“未定义/已停止”：源失效或网络不通
                        # （3=播放 6=缓冲 7=等待数据 9=切换 11=重连都算正常）
                        _clear_online_loading()
                        online['playing'] = False
                        lbl_dl.config(
                            text='网络加载超时，换一首或点“换一换”重试。',
                            fg='#c62828')
                    elif (ps == 10
                          and time.monotonic() - online['load_t0'] > 6):
                        # 10=Ready：媒体载入了却没真正开始播，多半是设备/版权
                        _clear_online_loading()
                        online['playing'] = False
                        lbl_dl.config(
                            text=('已载入但没有声音：没有检测到音频输出设备。'
                                  if not has_audio_output() else
                                  '这首没能播放（可能受版权限制），换一首。'),
                            fg='#c62828')
                        _refresh()
                if online['playing'] and _play_state() == 8:
                    # 在线曲目播完 → 切下一首（循环回到第一首；否则结束）
                    n = len(online['sel'])
                    if n > 0:
                        nxt = online['pos'] + 1
                        if nxt >= n and online['loop']:
                            nxt = 0
                        if nxt < n:
                            _online_play(nxt)
                        else:
                            online['playing'] = False
                            online['manual'] = False
                            _clear_online_loading()
                            lbl_dl.config(text='在线曲目已播放完毕。',
                                          fg='#2e8b57')
                            _refresh()
            if st['playing'] and not st['manual']:
                if _play_state() == 8:  # 一首播完
                    n = len(st['list'])
                    if n > 0:
                        nxt = st['idx'] + 1
                        if nxt >= n and st['loop']:
                            nxt = 0
                        if nxt < n:
                            _play_idx(nxt)
                        else:
                            st['playing'] = False
                            _refresh()
            try:
                win.after(400, _poll)
            except tk.TclError:
                pass

        def _pick_files():
            paths = filedialog.askopenfilenames(
                parent=win, title='选择音频文件',
                filetypes=[('音频文件', '*.mp3 *.wav *.ogg *.m4a *.wma *.flac'),
                           ('所有文件', '*.*')])
            if not paths:
                return
            for p in paths:
                if p not in state['list']:
                    state['list'].append(p)
                    state['names'].append(os.path.basename(p))
            _render()
            if state['idx'] < 0:
                _play_idx(0)

        def _rename():
            if not (0 <= state['idx'] < len(state['names'])):
                messagebox.showinfo('会务音乐',
                                    '请先在列表中选择要重命名的音乐。')
                return
            i = state['idx']
            new = simpledialog.askstring(
                '重命名音乐', '输入新名称：',
                initialvalue=state['names'][i], parent=win)
            if not new:
                return
            new = new.strip()
            if not new:
                return
            state['names'][i] = new
            _render()

        def _load_defaults():
            """把 files 目录下的默认会务音乐写入播放列表。"""
            for base in (_BASE_DIR, os.path.dirname(
                    os.path.abspath(__file__))):
                d = os.path.join(base, 'files')
                if not os.path.isdir(d):
                    continue
                for name in ['颁奖音乐.mp3', '领导上场音乐.mp3',
                             '暖场音乐(45分钟).mp3', '主持人上场音乐.mp3']:
                    p = os.path.join(d, name)
                    if os.path.exists(p) and p not in state['list']:
                        state['list'].append(p)
                        state['names'].append(name)
            _render()

        def _finish_close():
            _cancel_fade(restore=False)
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._music_win = None

        def _close_music():
            # 现场可能正在用音乐：关窗（X / 关闭按钮）前先确认，防误触
            if not ui_theme.confirm_close(win, '会务音乐'):
                return
            if _play_state() == 3 and fade['job'] is None:
                # 正在响 → 先隐藏窗口、声音 2 秒淡出后再销毁
                online['playing'] = False
                _clear_online_loading()
                state['playing'] = False
                state['manual'] = True
                try:
                    win.withdraw()
                except tk.TclError:
                    _finish_close()
                    return
                _fade_out(_finish_close)
                return
            _finish_close()

        # —— 进度相关 ——
        def _fmt(sec):
            sec = max(0, int(sec or 0))
            return '%02d:%02d' % (sec // 60, sec % 60)

        def _sync_progress():
            # 本地/在线共用一个引擎与同一条进度条
            try:
                md = wmp.currentMedia
            except Exception:
                md = None
            dur = 0
            pos = 0
            if md is not None:
                try:
                    dur = float(md.duration) or 0
                except Exception:
                    dur = 0
                try:
                    pos = float(wmp.controls.currentPosition) or 0
                except Exception:
                    pos = 0
            lbl_dur.config(text=_fmt(dur))
            lbl_pos.config(text=_fmt(pos))
            if dur > 0:
                sc.config(to=dur)
                if not state['seek']:
                    sc.set(pos)
            else:
                sc.config(to=1)
                sc.set(0)

        def _on_scale(v):
            if state['seek']:
                try:
                    if wmp.currentMedia is not None:
                        wmp.controls.currentPosition = float(v)
                except Exception:
                    pass

        def _seek_start(e=None):
            state['seek'] = True

        def _seek_end(e=None):
            state['seek'] = False
            _on_scale(sc.get())

        # —— 界面：播放器样式 ——
        # 顶栏：标题 + 关闭
        top = tk.Frame(win, bg=WINDOW_BG)
        top.pack(fill=tk.X, padx=16, pady=(12, 2))
        tk.Label(top, text='会务音乐', font=('微软雅黑', 15, 'bold'),
                 fg=PRIMARY, bg=WINDOW_BG).pack(side=tk.LEFT)
        tk.Button(top, text='关闭', font=('微软雅黑', 9),
                  bg=CARD_BG, relief=tk.FLAT, bd=1,
                  highlightbackground=BORDER_COLOR, cursor='arrow',
                  command=_close_music).pack(side=tk.RIGHT)

        # —— 统一播放栏：本地播放 / 在线音乐共用同一套显示与控制 ——
        # 当前曲目 + 状态
        frm_now = tk.Frame(win, bg=WINDOW_BG)
        frm_now.pack(fill=tk.X, padx=18, pady=(10, 0))
        lbl_now = tk.Label(frm_now, text='未选择曲目',
                           font=('微软雅黑', 12, 'bold'),
                           fg='#333333', bg=WINDOW_BG, anchor='w')
        lbl_now.pack(side=tk.LEFT, fill=tk.X, expand=True)
        lbl_status = tk.Label(frm_now, text='已停止', font=('微软雅黑', 9),
                              fg='#888888', bg=WINDOW_BG)
        lbl_status.pack(side=tk.RIGHT, padx=(8, 0))

        # 进度条 + 时间
        frm_prog = tk.Frame(win, bg=WINDOW_BG)
        frm_prog.pack(fill=tk.X, padx=20, pady=(8, 0))
        lbl_pos = tk.Label(frm_prog, text='00:00', font=('微软雅黑', 9),
                           fg='#666666', bg=WINDOW_BG, width=5, anchor='w')
        lbl_pos.pack(side=tk.LEFT)
        sc = ui_theme.SeekBar(frm_prog, from_=0, to=1, command=_on_scale,
                              on_seek_start=_seek_start,
                              on_seek_end=_seek_end, bg=WINDOW_BG)
        sc.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        lbl_dur = tk.Label(frm_prog, text='00:00', font=('微软雅黑', 9),
                           fg='#666666', bg=WINDOW_BG, width=5, anchor='e')
        lbl_dur.pack(side=tk.RIGHT)

        # 图标控制按钮（横排居中）：上一首 / 播放·暂停 / 下一首 / 停止
        # 播放后按钮自动变成暂停图标，暂停/停止后变回播放图标（_poll 每 400ms 校准）
        frm_ctrl = tk.Frame(win, bg=WINDOW_BG)
        frm_ctrl.pack(pady=(10, 6))
        icon_font = ('Segoe UI Symbol', 11)
        icon_font_big = ('Segoe UI Symbol', 13)

        def _icon_btn(parent, text, cmd, big=False):
            if big:
                btn = tk.Button(parent, text=text, font=icon_font_big,
                                bg=PRIMARY, fg='#ffffff', relief=tk.FLAT,
                                bd=0, width=4, height=1, cursor='arrow',
                                activebackground=ui_theme.PALETTE
                                .get('primary_dark', '#1d4ed8'),
                                activeforeground='#ffffff',
                                command=cmd)
            else:
                btn = tk.Button(parent, text=text, font=icon_font,
                                bg=CARD_BG, fg='#334155', relief=tk.FLAT,
                                bd=1, width=4, height=1, cursor='arrow',
                                highlightbackground=BORDER_COLOR,
                                activebackground='#e8eef5',
                                command=cmd)
            btn.pack(side=tk.LEFT, padx=8)
            return btn

        _icon_btn(frm_ctrl, '\u23ee', lambda: _step(-1))      # ⏮ 上一首
        btn_play = _icon_btn(frm_ctrl, ICON_PLAY, _toggle, big=True)
        _icon_btn(frm_ctrl, '\u23ed', lambda: _step(1))       # ⏭ 下一首
        _icon_btn(frm_ctrl, '\u23f9', _stop)                  # ⏹ 停止

        # 主区域：本地播放 / 在线音乐 两个页签（只放各自的内容，播放器两页一致）
        nb = ttk.Notebook(win)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=(2, 8))
        page1 = tk.Frame(nb, bg=WINDOW_BG)
        nb.add(page1, text='本地播放')
        page2 = tk.Frame(nb, bg=WINDOW_BG)
        nb.add(page2, text='在线音乐')

        # —— 音频输出设备提示条：没插音箱/声卡被禁用时明确告诉用户 ——
        # 这种情况下播放器状态仍是“播放中”，只是完全没声音，很容易被当成
        # 软件故障，所以单独给一条提示；插好设备后自动消失（点一下可立即重测）。
        dev_box = tk.Frame(win, bg='#fff4e5', highlightthickness=1,
                           highlightbackground='#f2c98a')
        tk.Label(dev_box,
                 text='⚠ 没有检测到可用的音频输出设备（音箱/耳机没插、'
                      '或在系统“声音设置”里被禁用），现在播放不会有声音。\n'
                      '插好设备后点这条提示重新检测。',
                 font=('微软雅黑', 9), fg='#b45309', bg='#fff4e5',
                 justify='left', anchor='w', wraplength=420,
                 cursor='hand2').pack(fill=tk.X, padx=8, pady=6)
        dev = {'ok': True}

        def _check_audio_dev(force=False):
            """检测输出设备，决定提示条显示/隐藏；返回当前是否有设备。"""
            ok = has_audio_output()
            if ok == dev['ok'] and not force:
                return ok
            dev['ok'] = ok
            try:
                if ok:
                    dev_box.pack_forget()
                else:
                    dev_box.pack(fill=tk.X, padx=18, pady=(0, 6),
                                 before=nb)
            except tk.TclError:
                pass
            return ok

        dev_box.bind('<Button-1>', lambda e: _check_audio_dev(force=True))
        for _c in dev_box.winfo_children():
            _c.bind('<Button-1>', lambda e: _check_audio_dev(force=True))

        # 切页即隔离：停掉另一页正在播放的内容。
        # 单一播放引擎同一时刻只能响一方，这样播放栏显示、听到的声音、
        # 上方控制按钮三者始终属于当前页签，不会“人在在线页却控制着本地”。
        music_tab = {'idx': 0}

        def _on_tab_change(_e=None):
            try:
                cur = nb.index(nb.select())
            except Exception:
                cur = 0
            if cur == music_tab['idx']:
                return
            music_tab['idx'] = cur
            _cancel_fade()
            try:
                wmp.controls.stop()
            except Exception:
                pass
            if cur == 1:
                # 切到在线页：本地播放停下
                state['playing'] = False
                state['manual'] = True
            else:
                # 切回本地页：在线播放停下
                online['playing'] = False
                online['manual'] = False
                _clear_online_loading()
            _refresh()
            _sync_selection()

        nb.bind('<<NotebookTabChanged>>', _on_tab_change)

        # —— 页1：本地播放（播放列表 + 工具行） ——
        # 播放列表（每行右侧带播放图标）
        frm_list = tk.Frame(page1, bg=WINDOW_BG)
        frm_list.pack(fill=tk.BOTH, expand=True, padx=14, pady=(8, 6))
        row_widgets = []   # [(行容器, 名称标签)]，用于只刷新高亮

        def _sync_selection():
            """只更新选中行的高亮，不重建列表。"""
            for idx, (row, lbl) in enumerate(row_widgets):
                color = '#dcebff' if idx == state['idx'] else WINDOW_BG
                for widget in (row, lbl):
                    try:
                        widget.configure(bg=color)
                    except tk.TclError:
                        pass

        canvas = tk.Canvas(frm_list, bg=WINDOW_BG, highlightthickness=0)
        sb = tk.Scrollbar(frm_list, orient=tk.VERTICAL, command=canvas.yview)
        inner = tk.Frame(canvas, bg=WINDOW_BG)
        inner_id = canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfigure(
            inner_id, width=e.width))
        for wgt in (canvas, inner):
            wgt.bind('<MouseWheel>', lambda e: canvas.yview_scroll(
                -int(e.delta / 120), 'units'))

        def _select(i):
            """单击选中：只改高亮色，不重建列表。

            点一下重建整个列表会把控件销毁，第二次点击落到新控件上，
            Tk 就不会判定为双击——双击播放会失效。
            """
            if not (0 <= i < len(state['names'])):
                return
            state['idx'] = i
            _sync_selection()

        def _remove_local(i):
            """从本地播放列表移除第 i 首（不删除文件）。"""
            st = state
            if not (0 <= i < len(st['list'])):
                return
            if st['idx'] == i:
                _stop()  # 移除正在播放的曲目 → 停止
            elif i < st['idx']:
                st['idx'] -= 1
            del st['list'][i]
            del st['names'][i]
            if st['idx'] >= len(st['list']):
                st['idx'] = len(st['list']) - 1
            _render()

        def _render():
            """按 state 重建播放列表，每行右侧带播放/移除图标。"""
            del row_widgets[:]
            for w in inner.winfo_children():
                w.destroy()
            for i, name in enumerate(state['names']):
                row = tk.Frame(inner, bg=WINDOW_BG)
                row.pack(fill=tk.X, padx=2, pady=1)
                tk.Button(row, text='▶', font=('微软雅黑', 9), width=2,
                          bg='#e8eef5', relief=tk.FLAT, bd=1,
                          cursor='arrow',
                          command=lambda i=i: _play_idx(i)).pack(
                    side=tk.LEFT, padx=(0, 6), pady=1)
                lbl = tk.Label(row, text=name, font=('微软雅黑', 10),
                               bg=WINDOW_BG, anchor='w', cursor='hand2')
                lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)
                # 双击播放：整行都能点（含名称右侧的空白处）
                for w in (row, lbl):
                    w.bind('<Button-1>', lambda e, i=i: _select(i))
                    w.bind('<Double-Button-1>',
                           lambda e, i=i: _play_idx(i))
                tk.Button(row, text='✕', font=('微软雅黑', 9), width=2,
                          bg='#fbe9e7', relief=tk.FLAT, bd=1,
                          fg='#c62828', cursor='arrow',
                          command=lambda i=i: _remove_local(i)).pack(
                    side=tk.RIGHT, padx=(6, 0), pady=1)
                row_widgets.append((row, lbl))
            _sync_selection()

        # 工具行：选歌 / 重命名 / 循环
        frm_tool = tk.Frame(page1, bg=WINDOW_BG)
        frm_tool.pack(fill=tk.X, padx=20, pady=(0, 12))
        loop_var = tk.BooleanVar(value=state['loop'])
        ttk.Checkbutton(frm_tool, text='列表循环播放', variable=loop_var,
                        command=lambda: state.update(
                            loop=loop_var.get())).pack(side=tk.LEFT)
        tk.Button(frm_tool, text='选歌', font=('微软雅黑', 9),
                  bg=CARD_BG, relief=tk.FLAT, bd=1,
                  highlightbackground=BORDER_COLOR, cursor='arrow',
                  command=_pick_files).pack(side=tk.LEFT, padx=(10, 0))
        tk.Button(frm_tool, text='重命名', font=('微软雅黑', 9),
                  bg=CARD_BG, relief=tk.FLAT, bd=1,
                  highlightbackground=BORDER_COLOR, cursor='arrow',
                  command=_rename).pack(side=tk.LEFT, padx=(10, 0))
        tk.Label(frm_tool, text='双击列表项即可播放', font=('微软雅黑', 8),
                 fg='#9aa3b2', bg=WINDOW_BG).pack(side=tk.RIGHT)

        _load_defaults()  # 默认会务音乐

        # —— 页2：在线音乐 ——
        # (标题, 歌曲id, fee)；直链是临时地址，播放/下载时用 id 现解析
        res_list = []
        online = {'sel': [], 'pos': -1, 'playing': False, 'manual': False,
                  'loop': False, 'page': 0,
                  'loading': False, 'load_title': '', 'load_t0': 0.0,
                  'load_dots': 0, 'load_job': None}  # 在线播放状态

        def _do_search():
            if not _win_alive():
                return
            kw = kw_var.get().strip()
            if not kw:
                messagebox.showinfo('在线音乐', '请输入搜索关键词。')
                return
            online['page'] = 0
            _stop_online_playback()
            rlb.delete(0, tk.END)
            res_list.clear()
            lbl_dl.config(text='搜索中…', fg='#888888')
            _start_worker(_search_worker, kw, 0)

        def _shuffle_search():
            """换一换：同一关键词随机翻到另一页，换一批曲目。"""
            if not _win_alive():
                return
            kw = kw_var.get().strip()
            if not kw:
                messagebox.showinfo('在线音乐', '请输入搜索关键词。')
                return
            import random
            page = online['page']
            nxt = page
            for _ in range(6):  # 尽量别和当前这批重复
                nxt = random.randrange(10)   # 前 10 页（每页 20 条）
                if nxt != page:
                    break
            online['page'] = nxt
            _stop_online_playback()
            rlb.delete(0, tk.END)
            res_list.clear()
            lbl_dl.config(text='换一批搜索中…', fg='#888888')
            _start_worker(_search_worker, kw, nxt * 20)

        def _call_in_ui(func):
            """把子线程的结果丢回界面线程；窗口已关闭时静默丢弃。"""
            try:
                win.after(0, func)
            except (tk.TclError, RuntimeError):
                pass

        def _start_worker(target, *args):
            """启动后台线程；线程内部异常只记日志，不影响主程序。"""
            import threading

            def _runner():
                try:
                    target(*args)
                except Exception as e:
                    safe_guard.write_log('后台任务出错：%r' % (e,))

            threading.Thread(target=_runner, daemon=True).start()

        # —— 在线播放“网络加载中”提示 ——
        def _set_online_loading(title):
            """开始缓冲：启动 250ms 一跳的省略号动画。"""
            online['loading'] = True
            online['load_title'] = title
            online['load_t0'] = time.monotonic()
            online['load_dots'] = 0
            job = online.get('load_job')
            if job:
                try:
                    win.after_cancel(job)
                except Exception:
                    pass
            _online_loading_tick()

        def _clear_online_loading():
            """缓冲结束（成功/失败/停止）：停掉省略号动画。"""
            online['loading'] = False
            job = online.get('load_job')
            if job:
                try:
                    win.after_cancel(job)
                except Exception:
                    pass
            online['load_job'] = None

        def _online_loading_tick():
            if not _win_alive() or not online['loading']:
                return
            n = online['load_dots'] % 8
            lbl_dl.config(text='网络加载中' + '.' * (n + 1), fg='#d97706')
            online['load_dots'] += 1
            online['load_job'] = win.after(250, _online_loading_tick)

        def _stop_online_playback():
            """停止在线播放并清理加载态（换一换/重新搜索时用）。"""
            if online['playing'] or online['loading']:
                try:
                    wmp.controls.stop()
                except Exception:
                    pass
            online['playing'] = False
            online['manual'] = False
            online['sel'] = []
            _clear_online_loading()

        def _search_worker(kw, offset=0):
            skipped = 0
            try:
                import urllib.request
                import urllib.parse
                import json
                items = []
                body = ('s=' + urllib.parse.quote(kw) +
                        '&type=1&limit=20&offset=%d' % offset)
                req = urllib.request.Request(
                    'https://music.163.com/api/search/get',
                    data=body.encode('utf-8'),
                    headers={
                        'Content-Type':
                            'application/x-www-form-urlencoded',
                        'Referer': 'https://music.163.com/',
                        'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=15) as r:
                    # 限制响应体积：异常/恶意响应不能把内存撑爆
                    raw = r.read(2 * 1024 * 1024)
                data = json.loads(raw.decode('utf-8', 'ignore'))
                songs = data.get('result', {}).get('songs', []) or []
                for s in songs:
                    sid = s.get('id')
                    if not sid:
                        continue
                    # fee: 0=免费 8=免费标准音质（能拿到直链）
                    #      1=VIP 4=付费专辑（免登录返回 -110/空，放不出来）
                    fee = int(s.get('fee') or 0)
                    if fee not in (0, 8):
                        skipped += 1
                        continue
                    name = s.get('name') or ''
                    artists = ','.join(
                        a.get('name') for a in (s.get('artists') or [])
                        if a.get('name')) or '未知歌手'
                    dur = int(s.get('duration') or 0) // 1000
                    title = '%s - %s  [%02d:%02d]' % (
                        name, artists, dur // 60, dur % 60)
                    # 只存歌曲 id：直链是临时地址（约 20 分钟失效），
                    # 而且网易云已关闭 outer/url 免登录直链，必须播放时现取。
                    items.append((title, int(sid), fee))
            except Exception:
                items = [('__ERR__', 0, 0)]
            _call_in_ui(lambda: _search_done(items, skipped))

        def _search_done(items, skipped=0):
            if not _win_alive():
                return
            if not items:
                if skipped:
                    lbl_dl.config(text='这一批只有付费曲目，已全部过滤；'
                                       '点“换一换”再看一批。', fg='#c62828')
                elif online.get('page'):
                    lbl_dl.config(text='这一批没有结果，再点一次“换一换”。',
                                  fg='#c62828')
                else:
                    lbl_dl.config(text='没有结果或搜索失败。', fg='#c62828')
                return
            first = items[0]
            if first[0] == '__ERR__':
                lbl_dl.config(text='搜索失败，请检查网络。', fg='#c62828')
                return
            for title, sid, fee in items:
                res_list.append((title, sid, fee))
                rlb.insert(tk.END, title)
            batch = ('（第 %d 批）' % (online['page'] + 1)
                     if online.get('page') else '')
            tip = '，已过滤 %d 首付费曲目' % skipped if skipped else ''
            lbl_dl.config(text='共 %d 首可播放%s%s，选中后点“下载到本地”。'
                          % (len(items), batch, tip), fg='#2e8b57')

        def _resolve_url(sid, timeout=12):
            """把歌曲 id 解析成当前可用的播放直链（失败返回 None）。

            网易云的 ``song/media/outer/url`` 免登录直链已失效（返回 404 页面），
            改用 ``song/enhance/player/url`` 现取。拿到的地址形如
            ``http://m701.music.126.net/2026.../<hash>/...mp3``，路径里带时间戳，
            约 20 分钟过期（接口 expi=1200），所以每次播放/下载都要现取，不能缓存。

            320k 拿不到时依次退到 192k / 128k：部分曲目只是高清码率受限，
            标准音质仍可放（接口 code 会变成 -110 / url 为空表示彻底无版权）。
            """
            import json
            import urllib.request
            for br in (320000, 192000, 128000):
                try:
                    url = ('https://music.163.com/api/song/enhance/player/url'
                           '?ids=[%s]&br=%d' % (sid, br))
                    req = urllib.request.Request(url, headers={
                        'User-Agent': 'Mozilla/5.0',
                        'Referer': 'https://music.163.com/'})
                    with urllib.request.urlopen(req, timeout=timeout) as r:
                        raw = r.read(256 * 1024)
                    data = json.loads(raw.decode('utf-8', 'ignore'))
                    item = (data.get('data') or [{}])[0]
                    if item.get('url'):
                        return item['url']
                except Exception:
                    continue
            return None

        def _online_play(pos):
            """从在线选中列表的 pos 位置开始播放（先解析直链，异步不卡界面）。"""
            sel = online['sel']
            if not (0 <= pos < len(sel)):
                online['playing'] = False
                return
            i = sel[pos]
            if not (0 <= i < len(res_list)):
                online['playing'] = False
                return
            online['pos'] = pos
            online['playing'] = True
            online['manual'] = False
            title, sid, fee = res_list[i]
            _cancel_fade()  # 在线播放顶掉淡出中的曲目：恢复音量
            if state['playing']:
                state['playing'] = False  # 在线播放顶掉本地播放
                state['manual'] = True
            _refresh()  # 统一播放栏立刻显示在线曲名
            _set_online_loading(title)   # 先提示“网络加载中…”
            _start_worker(_online_play_worker, pos, title, sid)

        def _online_play_worker(pos, title, sid):
            """后台取直链，取完回界面线程真正开播。"""
            url = _resolve_url(sid)
            _call_in_ui(lambda: _online_play_go(pos, title, url))

        def _online_play_go(pos, title, url):
            if not _win_alive():
                return
            if online['pos'] != pos or not online['playing']:
                return      # 已经切歌/停止，这次结果作废
            if not url:
                _clear_online_loading()
                online['playing'] = False
                online['manual'] = False
                _refresh()
                lbl_dl.config(
                    text='拿不到播放地址（版权受限或网络异常），换一首试试：'
                         + title, fg='#c62828')
                return
            try:
                wmp.URL = url
                wmp.controls.play()
            except Exception:
                _clear_online_loading()
                online['playing'] = False
                _refresh()
                lbl_dl.config(text='在线播放失败：' + title, fg='#c62828')

        def _start_online():
            """点“开始播放”：从第一首选中曲目开始循环播放选中集合。"""
            idxs = list(rlb.curselection())
            if not idxs:
                messagebox.showinfo('在线音乐',
                                    '请先在结果中选中要播放的曲目。')
                return
            online['sel'] = idxs
            _online_play(0)

        def _preview_result(e=None):
            # 双击：只试听该单曲（播完即停，不进入循环）
            idxs = list(rlb.curselection())
            if not idxs:
                return
            i = idxs[0]
            online['sel'] = [i]
            _online_play(0)

        def _remove_result():
            """从在线搜索结果中移除选中的曲目。"""
            idxs = sorted(rlb.curselection(), reverse=True)
            if not idxs:
                messagebox.showinfo('在线音乐', '请先选中要移除的结果。')
                return
            online['playing'] = False  # 移除时停止在线播放
            online['sel'] = []
            online['pos'] = -1
            _clear_online_loading()
            for k in idxs:
                del res_list[k]
                rlb.delete(k)
            lbl_dl.config(text='已移除 %d 条结果。' % len(idxs), fg='#2e8b57')

        def _download_result():
            sel = rlb.curselection()
            if not sel:
                messagebox.showinfo('在线音乐', '请先选中要下载的结果。')
                return
            title, sid, fee = res_list[sel[0]]
            lbl_dl.config(text='正在解析下载地址…', fg='#888888')
            _start_worker(_dl_worker, title, sid)

        def _dl_worker(title, sid):
            path = None
            started = False
            try:
                import urllib.request
                import re
                dl = _resolve_url(sid, timeout=15)
                if not dl:
                    _call_in_ui(lambda: _dl_fail(
                        '拿不到下载地址（版权受限或网络异常），换一首试试：'
                        + title))
                    return
                folder = os.path.join(_BASE_DIR, 'files', '在线音乐')
                if not os.path.isdir(folder):
                    os.makedirs(folder, exist_ok=True)
                # 去掉标题里的时长尾巴 [02:07]，再清理 Windows 非法字符
                name = re.sub(r'\s*\[\d{2}:\d{2}\]\s*$', '', title).strip()
                name = re.sub(r'[\\/:*?"<>|]', '_', name).strip() or '音乐'
                path = os.path.join(folder, name + '.mp3')
                base = name
                n = 1
                while os.path.exists(path):
                    n += 1
                    path = os.path.join(folder,
                                        '%s(%d).mp3' % (base, n))
                req = urllib.request.Request(
                    dl, headers={'User-Agent': 'Mozilla/5.0',
                                 'Referer': 'https://music.163.com/'})
                # 边下边写磁盘：不再把整首歌读进内存，也拦住超大文件
                with urllib.request.urlopen(req, timeout=120) as r:
                    total = int(r.headers.get('Content-Length') or 0)
                    if total > MAX_MUSIC_BYTES:
                        raise RuntimeError('文件过大')
                    with open(path, 'wb') as f:
                        started = True
                        copied = 0
                        while True:
                            chunk = r.read(256 * 1024)
                            if not chunk:
                                break
                            copied += len(chunk)
                            if copied > MAX_MUSIC_BYTES:
                                raise RuntimeError('文件过大')
                            f.write(chunk)
                _call_in_ui(lambda: _dl_done(path,
                                             os.path.basename(path)))
            except Exception:
                if started and path:
                    try:
                        os.remove(path)  # 删掉下了一半的残缺文件
                    except OSError:
                        pass
                _call_in_ui(lambda: _dl_fail('下载失败，请检查网络。'))

        def _dl_fail(msg):
            if not _win_alive():
                return
            lbl_dl.config(text=msg, fg='#c62828')

        def _dl_done(path, disp):
            if not _win_alive():
                return
            if path is None:
                lbl_dl.config(text='下载失败，请检查网络。', fg='#c62828')
                return
            state['list'].append(path)
            state['names'].append(disp)
            _render()
            short = disp if len(disp) <= 26 else disp[:24] + '…'
            lbl_dl.config(text='已下载：%s（已加入本地播放列表）' % short,
                          fg='#2e8b57')
            messagebox.showinfo('在线音乐', '下载完成：%s' % disp)

        # 搜索行（来源：网易云音乐）
        frm_search = tk.Frame(page2, bg=WINDOW_BG)
        frm_search.pack(fill=tk.X, padx=12, pady=(10, 4))
        kw_var = tk.StringVar(value='纯音乐')
        ttk.Entry(frm_search, textvariable=kw_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(frm_search, text='搜索', font=('微软雅黑', 9),
                  bg=CARD_BG, relief=tk.FLAT, bd=1,
                  highlightbackground=BORDER_COLOR, cursor='arrow',
                  command=_do_search).pack(side=tk.LEFT, padx=(6, 0))
        tk.Button(frm_search, text='换一换', font=('微软雅黑', 9),
                  bg=CARD_BG, relief=tk.FLAT, bd=1,
                  highlightbackground=BORDER_COLOR, cursor='arrow',
                  command=_shuffle_search).pack(side=tk.LEFT, padx=(6, 0))

        # 进度/播放控制统一使用页签上方的播放栏，两个页签界面格式一致
        # 结果列表
        frm_res_hdr = tk.Frame(page2, bg=WINDOW_BG)
        frm_res_hdr.pack(fill=tk.X, padx=12)
        tk.Label(frm_res_hdr, text='搜索结果（双击试听，多选后可移除）',
                 font=('微软雅黑', 9), fg='#7a8ca0',
                 bg=WINDOW_BG, anchor='w').pack(side=tk.LEFT)
        tk.Button(frm_res_hdr, text='移除选中', font=('微软雅黑', 9),
                  bg=CARD_BG, relief=tk.FLAT, bd=1,
                  highlightbackground=BORDER_COLOR, cursor='arrow',
                  command=_remove_result).pack(side=tk.RIGHT)
        frm_res = tk.Frame(page2, bg=WINDOW_BG)
        frm_res.pack(fill=tk.BOTH, expand=True, padx=12, pady=(2, 6))
        rlb = tk.Listbox(frm_res, font=('微软雅黑', 9),
                         selectmode=tk.EXTENDED, activestyle='none',
                         exportselection=False, height=10)
        rsb = tk.Scrollbar(frm_res, orient=tk.VERTICAL, command=rlb.yview)
        rlb.configure(yscrollcommand=rsb.set)
        rlb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rsb.pack(side=tk.RIGHT, fill=tk.Y)
        rlb.bind('<Double-Button-1>', _preview_result)

        # 右键菜单：直接下载到本地（也可试听 / 播放 / 移除）
        res_menu = tk.Menu(win, tearoff=0, bd=0,
                           bg=CARD_BG, fg=ui_theme.PALETTE['text'],
                           activebackground=ACCENT,
                           activeforeground=PRIMARY,
                           font=ui_theme.font(9))
        res_menu.add_command(label='下载到本地',
                             command=_download_result)
        res_menu.add_command(label='试听这首', command=_preview_result)
        res_menu.add_command(label='播放选中的曲目', command=_start_online)
        res_menu.add_separator()
        res_menu.add_command(label='从结果中移除', command=_remove_result)

        def _show_result_menu(event):
            """右键：先选中光标下的那条结果，再弹出菜单。"""
            index = rlb.nearest(event.y)
            if index >= 0 and index not in rlb.curselection():
                rlb.selection_clear(0, tk.END)
                rlb.selection_set(index)
                rlb.activate(index)
            try:
                res_menu.tk_popup(event.x_root, event.y_root)
            finally:
                try:
                    res_menu.grab_release()
                except tk.TclError:
                    pass

        rlb.bind('<Button-3>', _show_result_menu)

        # 下载行
        frm_dl = tk.Frame(page2, bg=WINDOW_BG)
        frm_dl.pack(fill=tk.X, padx=12, pady=(0, 4))
        onl_loop_var = tk.BooleanVar(value=online['loop'])
        ttk.Checkbutton(frm_dl, text='循环播放选中', variable=onl_loop_var,
                        command=lambda: online.update(
                            loop=onl_loop_var.get())).pack(side=tk.LEFT)
        tk.Button(frm_dl, text='播放选中', font=('微软雅黑', 9),
                  bg=PRIMARY, fg='#ffffff', relief=tk.FLAT, bd=0,
                  cursor='arrow',
                  activebackground=ui_theme.PALETTE.get('primary_dark',
                                                        '#1d4ed8'),
                  activeforeground='#ffffff',
                  command=_start_online).pack(side=tk.LEFT, padx=(8, 0))
        # 下载按钮固定在右侧：状态文字放到单独一行，否则长文件名会把它挤出去
        tk.Button(frm_dl, text='下载到本地', font=('微软雅黑', 9),
                  bg=CARD_BG, relief=tk.FLAT, bd=1,
                  highlightbackground=BORDER_COLOR, cursor='arrow',
                  command=_download_result).pack(side=tk.RIGHT)
        lbl_dl = tk.Label(page2, text='', font=('微软雅黑', 9),
                          fg='#2e8b57', bg=WINDOW_BG, anchor='w')
        lbl_dl.pack(fill=tk.X, padx=12, pady=(0, 8))

        win.protocol('WM_DELETE_WINDOW', _close_music)
        _center_window(win, 470, 620)
        win.after(300, _do_search)  # 打开时默认搜索"纯音乐"
        win.after(400, _poll)

        # 音频输出设备盯一眼：插上音箱/耳机后提示条自动消失
        def _dev_watch():
            if not _win_alive():
                return
            _check_audio_dev()
            try:
                win.after(2000, _dev_watch)
            except tk.TclError:
                pass

        win.after(200, _dev_watch)

    def _open_countdown(self):
        """打开 PPT 倒计时浮动窗口：无边框、可拖动、可透明、右键设置。

        主体只有一个倒计时数字；所有设置（时长、字号、颜色、透明、
        置顶、声音、PPT 联动等）都通过右键菜单完成。
        PPT 放映开始时自动开始倒计时（需 pywin32，检测 PowerPoint / WPS）。
        """
        if self._activate_or_close_others('countdown'):
            return

        # 色键透明用的“魔法色”：与文字色避开，被它涂满的区域会变透明。
        TRANSPARENT = '#FF00FF'

        win = tk.Toplevel(self.root)
        win.title('PPT倒计时')
        win.overrideredirect(True)  # 无边框
        # 默认投到扩展屏（投影仪/TV）由 native overlay 显示倒计时；
        # Tk 主浮窗（操作端）固定在主屏活动，与 native 各占一块屏，
        # 避免 Tk 跨屏时跟 native overlay 撞在同一块屏幕。
        monitors = video_player.list_monitors()
        default_index = 0
        primary_rect = None
        for item in monitors:
            if not item.get('primary'):
                default_index = item['index']
            elif primary_rect is None:
                primary_rect = item['rect']
        if primary_rect is None:
            # 没有 primary 标志时退到 Tk 合并虚拟桌面（单屏场景）
            primary_rect = (0, 0, win.winfo_screenwidth(),
                            win.winfo_screenheight())
        # Tk 主浮窗默认放主屏右上角——操作端常驻位置，便于坐在主屏右侧
        # 边控制边投影到扩展屏。
        win.geometry('260x100+%d+%d' % (
            max(primary_rect[0], primary_rect[2] - 280 - 40),
            primary_rect[1] + 40))
        win.attributes('-topmost', True)
        try:
            win.attributes('-transparentcolor', TRANSPARENT)
        except tk.TclError:
            pass
        win.configure(bg=TRANSPARENT)
        self._countdown_win = win

        state = {
            'running': False, 'job': None, 'end_time': 0.0,
            'total': 600, 'remaining': 600,
            'font_size': 64, 'color': '#C62828', 'bg_color': '#FFFFFF',
            'stroke_w': 0, 'stroke_color': None,  # 文字描边宽度/颜色(None=自动反差)
            'transparent': True, 'ontop': True, 'sound': True,
            'current_pos': 'tr',     # 位置按钮当前位置（tl/tr/bl/br/center），默认右上角
            'sound_file': None,  # 自定义到时提示音；None=默认提示音
            'warn_at': 60, 'warn_done': False,  # 剩余 N 秒声音提醒一次（0=关）
            'warn_red': True, 'link_ppt': True,
            'ppt_reset': True,  # PPT 退出放映时倒计时复位
            'alpha': 1.0, 'drag_x': 0, 'drag_y': 0,
            'ppt_job': None, 'ppt_was_on': False,
            'monitor_index': default_index,  # 当前所在的显示器（默认扩展屏）
            'sync_monitor_ui': None,  # 设置面板里的显示器下拉同步回调
            'settings_win': None,  # 右键设置面板（普通窗口，非 Menu）
            'run_btn': None,       # 设置面板里的“开始/暂停”按钮
            '_native_overlay': None,  # Win32 原生浮窗；扩展屏下启用，绕开 Tk 的 DWM bug
            '_native_failed': False,  # native 创建失败就别再试
            '_hidden': False,         # 当前是否处于"隐藏到托盘"状态
            '_tray': None,            # TrayIcon 实例；只在隐藏时存在
            '_hidden_was_running': False,  # 隐藏前是否在跑，恢复时由用户决定是否继续
            '_native_user_dragged': False,  # 用户手动拖 native 后标记为 True
            '_native_current_xy': None,    # native 当前位置 (x, y)；_native_pos_for 也读这个备用
            'native_margin': 0,            # native overlay 离屏幕物理边的边距（px）
                                           # 默认 0 → 真正贴边；想留点视觉空
                                           # 隙就调到 12/24/40，UI 里可以改。
            'native_zoom': 1.0,         # 扩展屏 native 字号缩放（不动窗口尺寸）。
                                           # 默认 1.0 → 与主屏字号一致；远处投
                                           # 影可调到 1.5 / 2.0 让字更大，但只
                                           # 影响字大小，窗口尺寸不动 → 还能贴边。
        }

        cv = tk.Canvas(win, bg=TRANSPARENT, highlightthickness=0,
                       cursor='hand2')
        cv.pack(fill=tk.BOTH, expand=True)

        # ---- 显示器：倒计时可以放到任意一块屏幕（默认扩展屏）----
        def _monitors():
            """当前所有显示器；取不到时退回主屏。"""
            items = video_player.list_monitors()
            if not items:
                items = [{'index': 0, 'primary': True,
                          'name': '显示器 1（%d×%d）  主屏' % (
                              win.winfo_screenwidth(),
                              win.winfo_screenheight()),
                          'rect': (0, 0, win.winfo_screenwidth(),
                                   win.winfo_screenheight())}]
            return items

        def _active_rect():
            """Tk 主浮窗的活动矩形 = 主屏矩形。

            历史上这里跟着 monitor_index 切，monitor_index 改了之后
            这块变成纯 Tk 主屏用的；位置按钮 / 离屏拉回 / 其它调用
            全部拿到主屏矩形，避免把 Tk 浮窗搬到扩展屏飞出可视区。
            """
            return _primary_rect()

        def _primary_rect():
            """主屏矩形 (left, top, right, bottom)；找不到时退到 Tk 合并桌面。"""
            for item in _monitors():
                if item.get('primary'):
                    return item['rect']
            return (0, 0, win.winfo_screenwidth(), win.winfo_screenheight())

        def _current_monitor_index():
            """按窗口中心点判断它现在在哪块屏上（拖动后可自动同步）。"""
            items = _monitors()
            try:
                cx = win.winfo_rootx() + win.winfo_width() // 2
                cy = win.winfo_rooty() + win.winfo_height() // 2
            except tk.TclError:
                return int(state.get('monitor_index', 0) or 0)
            for item in items:
                left, top, right, bottom = item['rect']
                if left <= cx < right and top <= cy < bottom:
                    return item['index']
            return int(state.get('monitor_index', 0) or 0)

        def _apply_monitor(index):
            """切换 native overlay 用的显示器。

            Tk 主浮窗（操作端）永远在主屏活动——这里**不搬 Tk 浮窗**，
            只切 native overlay 的位置。扩展屏投屏由 native overlay
            接管显示，Tk 浮窗留在主屏给用户操作。
            """
            items = _monitors()
            old = max(0, min(int(state.get('monitor_index', 0) or 0),
                             len(items) - 1))
            new = max(0, min(int(index or 0), len(items) - 1))
            state['monitor_index'] = new
            if new == old:
                return
            _sync_panel()
            _ensure_native_overlay()
            _present()
            _repaint(lift=True)
            _repaint_soon(kick=True)
            win.after(900, lambda: _diagnose('切换显示器'))

        # ---- 窗口存活判断（多个流程共用） ----
        def _alive(w):
            try:
                return bool(w is not None and w.winfo_exists())
            except tk.TclError:
                return False

        def _fmt(seconds):
            if seconds < 0:
                seconds = 0
            m, s = divmod(seconds, 60)
            if m >= 60:
                h, m = divmod(m, 60)
                return '%d:%02d:%02d' % (h, m, s)
            return '%02d:%02d' % (m, s)

        def _size_for(font_size):
            """按字号算出窗口该多大（Tk 主浮窗用）。"""
            f = tkfont.Font(font=('微软雅黑', int(font_size), 'bold'))
            return (max(int(f.measure(_fmt(state['remaining']))) + 40, 120),
                    max(int(f.metrics('linespace')) + 24, 60))

        def _bg():
            return TRANSPARENT if state['transparent'] else state['bg_color']

        def _fg():
            if (state['warn_red'] and state['running']
                    and 0 < state['remaining'] <= 60):
                return '#C62828'
            return state['color']

        def _luma(hexcolor):
            h = hexcolor.lstrip('#')
            if len(h) != 6:
                return 128
            r = int(h[0:2], 16)
            g = int(h[2:4], 16)
            b = int(h[4:6], 16)
            return 0.299 * r + 0.587 * g + 0.114 * b

        def _contrast(color):
            """与给定颜色形成高反差的颜色（白底黑字 / 黑底白字）。"""
            return '#000000' if _luma(color) >= 128 else '#FFFFFF'

        def _hex_rgb(color):
            """#RRGGBB → (r, g, b)；认不出来时给个中性灰。"""
            try:
                h = str(color).lstrip('#')
                if len(h) == 3:
                    h = ''.join(c * 2 for c in h)
                if len(h) != 6:
                    return (128, 128, 128)
                return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
            except (TypeError, ValueError):
                return (128, 128, 128)

        def _propagate_native_fill():
            """把"是否背景透明"这一决定同步给 native overlay。

            主屏 Tk 浮窗（属性面板里的"背景透明"勾选框）切到 切实心 时，
            扩展屏的 native overlay 必须跟着切——否则观众看到的主屏
            是黑底白字、扩展屏却是色键透明（看穿到桌面），视觉割裂。

            实现：
              transparent=True  → ``native.set_fill_color(None)``
                                  色键透明，``UpdateLayeredWindow``
                                  把色键像素合成到背景上；
              transparent=False → ``native.set_fill_color(rgb)``
                                  用 Tk 主浮窗同样的 ``bg_color``
                                  切实心涂底，跟主屏视觉一致。
            """
            native = state.get('_native_overlay')
            if native is None:
                return
            try:
                if state.get('transparent', True):
                    native.set_fill_color(None)
                else:
                    native.set_fill_color(_hex_rgb(state.get('bg_color')
                                                   or '#202020'))
            except Exception:
                pass

        def _near(color, target, tol=64):
            """两个颜色是否接近（用来判断屏幕采样点上有没有我们的字）。"""
            if color is None:
                return False
            return max(abs(color[i] - target[i]) for i in range(3)) <= tol

        def _effective_fg():
            """解决撞色：非透明背景且前景≈背景时，改用高反差文字色。"""
            fg = _fg()
            bg = _bg()
            if bg and bg != TRANSPARENT and fg.lower() == bg.lower():
                return _contrast(bg)
            return fg

        def _paint(canvas, font_size):
            """把一个画布画成倒计时（Tk 主浮窗的 canvas 走这里）。"""
            try:
                canvas.delete('all')
                text = _fmt(state['remaining'])
                fnt = ('微软雅黑', int(font_size), 'bold')
                fg = _effective_fg()
                cx = canvas.winfo_width() / 2
                cy = canvas.winfo_height() / 2
                # 文字描边（默认关闭 stroke_w=0），可手动设置宽度/颜色
                sw = int(state.get('stroke_w', 0) or 0)
                if sw > 0:
                    sc = state.get('stroke_color') or _contrast(fg)
                    off = max(1, int(round(sw)))
                    for (dx, dy) in [(off, 0), (-off, 0), (0, off),
                                     (0, -off), (off, off), (off, -off),
                                     (-off, off), (-off, -off)]:
                        canvas.create_text(cx + dx, cy + dy, text=text,
                                           font=fnt, fill=sc, anchor='center')
                item = canvas.create_text(cx, cy, text=text, font=fnt,
                                          fill=fg, anchor='center')
                # 记下文字的位置：自检时要用它算出“字应该出现在屏幕哪儿”
                canvas._text_item = item
            except tk.TclError:
                # 主循环外的调用（窗口正在销毁等）直接吞掉，别影响计时
                pass
            # 同时把文字镜像到 Win32 native overlay（扩展屏用）
            native = state.get('_native_overlay')
            if native is not None:
                try:
                    fg = _effective_fg()
                    r, g, b = _hex_rgb(fg)
                    # native 窗口尺寸跟 Tk 主浮窗一致，传入的字号也要
                    # 同步缩放才能跟窗口比例一致
                    boost = float(state.get('_native_boost') or 1.0)
                    # 位置独立：Tk 主浮窗已被锁在主屏活动，扩展屏的倒计时
                    # 由 native 独立显示给观众看；native 自己记录位置，
                    # 不再被 Tk 主浮窗牵着跑（避免撞在同一块屏幕）。
                    # native 字号用主屏浮窗基础字号 × boost。
                    fs = int(round(int(state.get('font_size') or 64) * boost))
                    # 描边同步：主屏 Tk 浮窗用 create_text 8 偏移画描边，
                    # native 也用同样的偏移画 → 主屏、扩展屏两边粗细一致。
                    # 描边宽度=0 时不画描边（与主屏一致）。
                    sw = int(state.get('stroke_w', 0) or 0)
                    if sw > 0:
                        sc_hex = state.get('stroke_color') or _contrast(fg)
                        sr, sg, sb = _hex_rgb(sc_hex)
                        native.set_text(_fmt(state['remaining']),
                                        color=(r, g, b), font_size=fs,
                                        stroke_w=sw,
                                        stroke_color=(sr, sg, sb))
                    else:
                        native.set_text(_fmt(state['remaining']),
                                        color=(r, g, b), font_size=fs)
                except Exception:
                    pass

        def _native_pos_for(rect, nw, nh, pos):
            """native overlay 在某块屏上的左上角坐标（按当前 pos 摆位）。

            ``rect = (left, top, right, bottom)`` 是目标屏幕矩形。
            ``nw / nh`` 是 native 窗口的实际宽高（已 ×boost）。
            边距取 ``state['native_margin']``，默认 0 → 真正贴 4 角。
            会议投影场景用户视觉预期是"零边距往角落里塞"，这里就
            真的给 0 了——投影边缘本来就有黑边 / 投影仪边框，多留
            的 8px 反而看着像没贴合。
            """
            left, top, right, bottom = rect
            try:
                margin = int(state.get('native_margin', 0) or 0)
            except Exception:
                margin = 0
            margin = max(0, margin)
            if pos == 'tl':
                return left + margin, top + margin
            if pos == 'tr':
                return right - nw - margin, top + margin
            if pos == 'bl':
                return left + margin, bottom - nh - margin
            if pos == 'br':
                return right - nw - margin, bottom - nh - margin
            # center
            cx = left + (right - left - nw) // 2
            cy = top + (bottom - top - nh) // 2
            # center 模式留 4 角 / 边沿视觉缓冲；不是要贴边的语义，
            # 这里用 min(margin, 16) 防止窗口"莫名其妙"漂到屏外
            cmargin = max(0, min(margin, 16))
            return max(left + cmargin, cx), max(top + cmargin, cy)

        def _move_native_overlay(pos):
            """把 native overlay 移到 ``pos`` 方位（在它所在的屏幕上）。

            如果用户已经手动拖动过 native（在扩展屏上自由安放），
            ``_native_user_dragged`` 是 True，**5 个位置按钮变成"归位
            快捷"**——一旦点了就把 native 拉回 pos 处重新归位，并清掉
            用户拖动标记，之后再按 pos 才能跟着 Tk 主浮窗走。
            """
            native = state.get('_native_overlay')
            if native is None:
                return
            try:
                monitors = _monitors()
                idx = int(state.get('monitor_index', 0) or 0)
                target = None
                for m in monitors:
                    if m['index'] == idx:
                        target = m.get('rect')
                        break
                if not target or len(target) < 4:
                    target = (0, 0, 1920, 1080)
                try:
                    nw, nh = native.size
                except Exception:
                    nw, nh = 260, 100
                nx, ny = _native_pos_for(target, nw, nh, pos)
                native.move(nx, ny)
                # 同步当前位置供下次诊断 / 复用
                state['_native_current_xy'] = (int(nx), int(ny))
                # 用户主动按了 pos 按钮 = 主动归位 → 清掉用户拖动标记
                state['_native_user_dragged'] = False
            except Exception:
                pass

        def _ensure_native_overlay():
            """根据 monitor_index 创建 / 销毁原生分层窗口。

            在扩展屏 / 非主屏上启用（绕开 Tk 在 DWM 合成上的 bug）；
            主屏不创建（主屏 Tk 渲染正常，副本能照常工作）。
            """
            if _native_countdown is None:
                return
            try:
                idx = int(state.get('monitor_index', 0) or 0)
            except Exception:
                idx = 0
            monitors = _monitors()
            is_extended = False
            try:
                for m in monitors:
                    if m['index'] == idx and not m.get('primary', False):
                        is_extended = True
                        break
            except Exception:
                is_extended = False

            native = state.get('_native_overlay')
            if is_extended:
                if native is None and not state.get('_native_failed'):
                    try:
                        # native overlay 位置独立于 Tk 主浮窗：
                        # Tk 主浮窗锁主屏（操作端），native 锁扩展屏
                        # （观众端）。不再用 win.winfo_x/y()，否则 Tk
                        # 移回主屏时 native 也会被拉到主屏。
                        ext_rect = None
                        for m in monitors:
                            if m['index'] == idx:
                                ext_rect = m.get('rect')
                                break
                        if not ext_rect or len(ext_rect) < 4:
                            ext_rect = (0, 0, 1920, 1080)
                        # native overlay 窗口尺寸与主屏 Tk 主浮窗一致——
                        # 不再 ×1.8 放大。早先 1.8× 是为了让投影端字更
                        # 大，但代价是 native 整体比主屏大 80%，看上去
                        # "边距很大"。现在 native 尺寸 = Tk 主浮窗尺寸，
                        # 字号由 native_countdown.py 内部按所在显示器 DPI
                        # 缩放（扩展屏 100% vs 主屏 125% 时字相对会大一
                        # 点），用户也能在属性面板独立加"native 字号缩放"
                        # 微调（state['native_zoom']，只动字号，不动窗口）。
                        w, h = _size_for(int(state.get('font_size', 48)))
                        # 万一 _size_for 给的太窄（极小字号），封个下限
                        nw = max(int(round(w)), 120)
                        nh = max(int(round(h)), 60)
                        # 如果 Tk 主浮窗已经被 _fit_window resize 过且
                        # 大于 _size_for 的值（用户手动改字号触发），以
                        # Tk 实际尺寸为准。Tk 还没初始化（winfo_width=1）
                        # 时仍用 _size_for 的值。
                        try:
                            tk_w_now = int(win.winfo_width())
                            if tk_w_now >= 120 and abs(tk_w_now - nw) > 4:
                                nw = tk_w_now
                        except tk.TclError:
                            pass
                        try:
                            tk_h_now = int(win.winfo_height())
                            if tk_h_now >= 60 and abs(tk_h_now - nh) > 4:
                                nh = tk_h_now
                        except tk.TclError:
                            pass
                        # 位置按 state['current_pos'] 走（tr/tl/br/bl/center），
                        # 这样点位置按钮的时候 native 会跟着 Tk 主浮窗一起切
                        pos = state.get('current_pos') or 'tr'
                        nx, ny = _native_pos_for(ext_rect, nw, nh, pos)
                        ov = _native_countdown.NativeCountdownOverlay(nx, ny, nw, nh)
                        # native overlay 也支持手动拖动——会务场景下操作者
                        # 在主屏但想微调大屏位置（比如避开投影画面里的某
                        # 个人），按住 native 边缘就能拖。代价：本窗口区
                        # 域内不能点桌面图标（投影端角落通常没图标）。
                        try:
                            ov.set_draggable(True)
                        except Exception:
                            pass
                        fg = _effective_fg()
                        r, g, b = _hex_rgb(fg)
                        # 字号 = 主屏字号 × DPI 缩放 × 用户独立缩放
                        # （native_countdown.py 内部已经会 ×dpi/96，外面
                        # 这里的 native_zoom 是用户额外叠加的缩放系数）
                        zoom = float(state.get('native_zoom') or 1.0)
                        fs = int(round(int(state['font_size']) * zoom))
                        # 描边：跟主屏 Tk 浮窗保持一致，扩展屏第一帧就
                        # 带描边，否则刚拉起来的时候看着跟主屏不一样。
                        sw = int(state.get('stroke_w', 0) or 0)
                        if sw > 0:
                            sc_hex = state.get('stroke_color') or _contrast(fg)
                            sr, sg, sb = _hex_rgb(sc_hex)
                            ov.set_text(_fmt(state['remaining']),
                                        color=(r, g, b), font_size=fs,
                                        stroke_w=sw,
                                        stroke_color=(sr, sg, sb))
                        else:
                            ov.set_text(_fmt(state['remaining']),
                                        color=(r, g, b), font_size=fs)
                        ov.show()
                        state['_native_overlay'] = ov
                        # _native_boost 现在不再乘到窗口上，但保留字段
                        # 以兼容老测试 / 旧逻辑；语义改成"字号缩放比"。
                        state['_native_boost'] = zoom
                        state['_native_origin_xy'] = (nx, ny)
                        state['_native_origin_wh'] = (nw, nh)
                        # 记一份 native 当前位置——pos 按钮 / 拖动后
                        # 都通过这个唯一的坐标源同步，别到处计算。
                        state['_native_current_xy'] = (nx, ny)
                        # 拖动期间实时回写坐标到 state（不触发 _move_pos）
                        # 这样万一 native 被用户搬走了，下次 pos 按钮还能
                        # 在它所在屏幕上按 pos 摆位，不会跨屏跳回。
                        def _on_native_move(_hwnd, x, y, _ov=ov):
                            try:
                                state['_native_current_xy'] = (int(x),
                                                                int(y))
                                # 关键：用户拖动后 state 标记成"手动调过"，
                                # 之后 pos 按钮只在用户没手动调过时自动归位；
                                # 调过的话，5 个 pos 按钮变成"复位快捷"。
                                state['_native_user_dragged'] = True
                            except Exception:
                                pass
                        try:
                            ov.set_drag_callback(_on_native_move)
                        except Exception:
                            pass
                        # 刚建好的 native 也要立刻同步"是否透明"。
                        # state['transparent'] 默认 True → 走色键路径；
                        # 用户在属性面板取消勾选 → 切切实心，跟主屏 Tk
                        # 浮窗的 bg_color 一致，避免主屏黑底白字而扩展
                        # 屏还是色键透明的视觉割裂。
                        try:
                            _propagate_native_fill()
                        except Exception:
                            pass
                        _diag_log('已建 native overlay（Win32 原生浮窗，'
                                  '尺寸=%dx%d  字号缩放=×%.2f  '
                                  '位置=扩展屏右上角(%d,%d)  可拖动）'
                                  % (nw, nh, zoom, nx, ny))
                    except Exception as exc:
                        state['_native_failed'] = True
                        _diag_log('创建 native overlay 失败（放弃）：%s' % exc)
            else:
                if native is not None:
                    try:
                        native.close()
                    except Exception:
                        pass
                    state['_native_overlay'] = None


        def _draw():
            _paint(cv, state['font_size'])
            _fit_window()

        _fit_font = None         # 主窗口上次已适配的字号

        def _resize_keep_center(target, font_size):
            """窗口随字号缩放，保持中心点不变（向四周扩展）。"""
            try:
                w, h = _size_for(font_size)
                x, y = target.winfo_x(), target.winfo_y()
                cw, ch = target.winfo_width(), target.winfo_height()
                if cw <= 1 or ch <= 1:
                    return
                target.geometry('%dx%d+%d+%d' % (
                    w, h, x + (cw - w) // 2, y + (ch - h) // 2))
                ui_theme.force_repaint(target)
            except tk.TclError:
                pass

        def _fit_window():
            nonlocal _fit_font
            fs = int(state['font_size'])
            if _fit_font != fs:
                _fit_font = fs
                _resize_keep_center(win, fs)

        def _refresh():
            try:
                bg = _bg()
                cv.configure(bg=bg)
                win.configure(bg=bg)
                _draw()
                _sync_run_btn()
                # 延迟设置窗口属性：在菜单回调期间直接调用
                # attributes(-topmost/-alpha) 会干扰菜单关闭导致面板置底，
                # 用 after(0) 推到当前事件处理完毕后再执行
                win.after(0, _apply_win_attrs)
            except tk.TclError:
                pass

        # 关键闭包引用挂到 state：_hide_countdown_to_tray /
        # _restore_countdown_from_tray 是类方法，拿不到局部闭包，
        # 通过 state 转发一下。这样从托盘或主程序入口恢复时调到的就是
        # 同一个函数实例（保持闭包内部状态一致）。
        state['_ensure_native_overlay_fn'] = _ensure_native_overlay
        state['_refresh_fn'] = _refresh
        state['_present_fn'] = lambda: _present()
        state['_repaint_fn'] = lambda: _repaint(lift=True)
        state['_repaint_soon_fn'] = lambda kick=False: _repaint_soon(kick=kick)

        def _sync_run_btn():
            """设置面板里的“开始/暂停”按钮跟着计时状态换字。"""
            btn = state.get('run_btn')
            if btn is None:
                return
            try:
                if not btn.winfo_exists():
                    state['run_btn'] = None
                    return
                btn.configure(text='暂停计时' if state['running']
                              else '开始计时')
            except tk.TclError:
                state['run_btn'] = None

        applied_attrs = {'key': None, 'topmost': None, 'alpha': None}

        def _apply_win_attrs():
            # 背景透明开关：
            #   勾选 → 色键透明浮层（浮在 PPT 上，但个别显示器不认）；
            #   取消 → 彻底不用分层窗口，变回普通实心窗口（系统一定会画出来）
            color_key = TRANSPARENT if state['transparent'] else ''
            eff_alpha = state['alpha'] if state['transparent'] else 1.0
            for target, canvas in ((win, cv),):
                if target is None or canvas is None:
                    continue
                try:
                    target.attributes('-topmost', state['ontop'])
                    target.attributes('-alpha', eff_alpha)
                    target.attributes('-transparentcolor', color_key)
                    # 同步画布/窗口背景为对应（透明或实心）颜色
                    canvas.configure(bg=_bg())
                    target.configure(bg=_bg())
                except tk.TclError:
                    pass
                if not state['transparent']:
                    # 去掉 WS_EX_LAYERED：这一步是“实心模式一定看得见”的关键
                    ui_theme.make_opaque(target)
                else:
                    # 从实心切回透明时，Tk 不会把分层样式补回来，得我们补
                    ui_theme.set_layered(target)
            # 属性改了才做一次“硬”重贴（含 SetWindowPos/RedrawWindow）。
            # 计时过程中每 0.5 秒刷新一次，不能每次都重贴，否则大屏会闪。
            changed = (applied_attrs['key'] != color_key
                       or applied_attrs['topmost'] != bool(state['ontop'])
                       or applied_attrs['alpha'] != eff_alpha)
            if changed:
                applied_attrs.update({'key': color_key,
                                      'topmost': bool(state['ontop']),
                                      'alpha': eff_alpha})
                _present()
            _repaint()
            # 透明/切实心 这个开关也驱动 native overlay（观众那边看到的
            # 倒计时要跟主屏同步：主屏实心 → 观众也实心，主屏透明 → 观众
            # 也透明）。每次改 _apply_win_attrs 都 propagate 一遍。
            _propagate_native_fill()

        def _repaint(lift=False):
            """主动让系统把倒计时窗口重画一遍。

            倒计时背景是色键透明：一旦系统没给它发重画消息，字没画出来，
            整个窗口就变成“隐形”，看着像倒计时不见了。
            移动之后顺手抬一次 z 序，确保它真的显示在大屏最前面。
            """
            try:
                if lift and state.get('ontop'):
                    win.lift()
                    win.attributes('-topmost', True)
            except tk.TclError:
                pass
            ui_theme.force_repaint(win)

        repaint_jobs = []

        def _present():
            """把浮窗重新合成到屏幕上（不藏不显，基本看不出闪烁）。

            色键透明 + 无边框的浮层在扩展屏上移动、改属性之后，Windows
            偶尔会漏掉一次分层窗口合成：窗口还在最前面，整块却是透明的，
            背景和字全看不见。Tk 自己重画救不回来，这里重新贴一遍分层
            属性，再强制重画窗口和它的包装窗口。
            """
            key = TRANSPARENT if state['transparent'] else ''
            alpha = state['alpha'] if state['transparent'] else 1.0
            for target in (win,):
                if not _alive(target):
                    continue
                try:
                    if state.get('ontop'):
                        target.attributes('-topmost', True)
                    target.attributes('-alpha', alpha)
                    target.attributes('-transparentcolor', key)
                except tk.TclError:
                    pass
                if state['transparent']:
                    ui_theme.set_layered(target)
                    ui_theme.present_overlay(target, key=key, alpha=alpha)
                else:
                    ui_theme.make_opaque(target)
                    ui_theme.force_repaint(target)

        def _kick_window():
            """重新贴一次画面（"修复显示"按钮用）。

            以前这里是 withdraw + deiconify 重建窗口，实测这一步本身就
            可能让分层窗口在扩展屏上彻底不再合成，所以只用安全的做法：
            重贴分层属性 + 强制重画。
            """
            _present()
            _repaint(lift=True)

        def _kick_extended_overlay():
            """点「开始计时」后强制让 DWM 重建扩展屏上分层窗口的 cache。

            之前版本在这里调了 ``make_opaque + set_layered`` —— 让窗口
            临时退回普通实心再切回分层。**这恰恰是 bug 的元凶**：切回
            分层后没有人触发 paint，DWM 里的 layered cache 是空的，
            屏幕 DC 上 cv 区域就显示出 cv 后面的 PowerPoint/桌面颜色。
            表现是扩展屏上的倒计时看上去整块"不见"了。

            现在的做法：
            1. 用 alpha 抖动让 ``SetLayeredWindowAttributes`` 真的把
               属性变更传给 DWM 触发 cache 重建（同值是 no-op）；
            2. ``win.update_idletasks + win.update`` 强制 Tk 处理
               pending paint 消息，把新字送进新 cache；
            3. 30ms 后再 present_overlay + force_repaint，再保险地
               走一遍整条合成。

            只对真正位于扩展屏的倒计时生效（DWM 在主屏不会漏合成）。
            """
            try:
                items = _monitors()
                cur = max(0, min(int(state.get('monitor_index', 0) or 0),
                                 len(items) - 1))
                if cur >= len(items) or items[cur]['primary']:
                    return
                if not state['transparent']:
                    return
                alpha = float(state.get('alpha', 1.0) or 1.0)
                # 1. 短暂把 alpha 调成 nudged（差 0.05 让 SetLayeredWindowAttributes
                #    不再是 no-op，DWM 真的会标记 cache 为 dirty 并重建）。
                nudged = max(0.05, min(0.95, alpha - 0.05)) \
                    if alpha >= 0.1 else min(0.95, alpha + 0.05)
                try:
                    ui_theme.present_overlay(
                        win, key=TRANSPARENT, alpha=nudged)
                except Exception:
                    pass
                # 2. 强制 Tk 处理 pending paint 消息，让 cache 有新内容
                try:
                    cv.update_idletasks()
                    win.update_idletasks()
                    win.update()
                except Exception:
                    pass
                # 3. 30ms 后恢复 alpha + 再 present + force_repaint
                def _restore_alpha():
                    try:
                        ui_theme.present_overlay(
                            win, key=TRANSPARENT, alpha=alpha)
                    except Exception:
                        pass
                    try:
                        ui_theme.force_repaint(win)
                    except Exception:
                        pass
                try:
                    win.after(30, _restore_alpha)
                except tk.TclError:
                    _restore_alpha()
            except Exception:
                pass

        # ---- 自检：倒计时到底有没有真的画在大屏上 ----
        # 色键透明浮层在个别显示器上会"整块隐形"（窗口在、字没画出来），
        # 从程序里看不出来。这里直接读屏幕像素来判定，并把每次判定写进
        # 倒计时诊断.log，找不到原因时这份日志就是证据。
        _diag = {'fail': 0, 'auto': False, 'solid_fail': False}

        # 启动期探针：立刻写一行确认模块跑到这里了（与按钮/计时无关）
        try:
            with open(
                os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             '倒计时诊断.log'),
                'a', encoding='utf-8') as _f:
                _f.write('[%s] %s\n' % (
                    time.strftime('%Y-%m-%d %H:%M:%S'),
                    '启动期探针：模块已加载到 _diag 段，_diag_log 已就绪'))
        except Exception:
            pass

        def _diag_path():
            try:
                base = os.path.dirname(safe_guard.log_path())
            except Exception:
                base = os.path.dirname(os.path.abspath(__file__))
            return os.path.join(base, '倒计时诊断.log')

        def _diag_log(msg):
            try:
                line = '[%s] %s\n' % (time.strftime('%Y-%m-%d %H:%M:%S'), msg)
                # 1. 原行为：写到 %APPDATA%\会务助手\ 下的日志（safe_guard 所在目录）
                try:
                    with open(_diag_path(), 'a', encoding='utf-8') as f:
                        f.write(line)
                except Exception:
                    pass
                # 2. 同时镜像写到源文件旁边（用户在桌面看的就是这一份），
                #    排查期方便用户立刻看到，不用去 %APPDATA% 翻。
                try:
                    mirror = os.path.join(
                        os.path.dirname(os.path.abspath(__file__)),
                        '倒计时诊断.log')
                    with open(mirror, 'a', encoding='utf-8') as f:
                        f.write(line)
                except Exception:
                    pass
            except Exception:
                pass

        def _sample_text():
            """在"文字应该在的屏幕位置"采样。

            同时跑两路：
            1. ``screen_pixels`` —— ``GetDC(None) + GetPixel``，读 GDI 缓存。
               DWM 在扩展屏上偶尔不会把 layered 窗口合成到屏幕，但 Tk 已经
               把字画进了 GDI 缓存，这条路仍会读到“画上了”的颜色，屏幕
               上其实看不见——它会误判成“画面正常”。
            2. ``screenshot_pixels`` —— ``BitBlt`` 把窗口区域截到内存 DC
               再读像素，读的是屏幕上“真实显示的颜色”。

            两路结果同时记到 detail；命中判定以 BitBlt 为主（更准）。
            返回 ``(命中数, 采样数, 细节字典)`` 或 None（两路都读不到时）。
            """
            try:
                item = getattr(cv, '_text_item', None)
                if item is None:
                    return None
                # 强制刷新，否则 cv.bbox 在 _refresh 刚跑完的一帧里可能拿到
                # 旧坐标（_start 之后立刻 _diagnose 时很容易踩这个坑）。
                try:
                    cv.update_idletasks()
                except Exception:
                    pass
                bbox = cv.bbox(item)
                if not bbox:
                    return None
                x1, y1, x2, y2 = bbox
                if x2 - x1 < 6 or y2 - y1 < 6:
                    return None
                base_x, base_y = cv.winfo_rootx(), cv.winfo_rooty()
                n = 7
                pts = []
                for i in range(n):
                    for j in range(n):
                        pts.append(
                            (int(base_x + x1 + (x2 - x1) * (i + 0.5) / n),
                             int(base_y + y1 + (y2 - y1) * (j + 0.5) / n)))
                # 关键诊断：两路同时跑，谁准一眼看出来
                cols_gdi = ui_theme.screen_pixels(pts)
                cols_screen = ui_theme.screenshot_pixels(pts)
                if cols_gdi is None and cols_screen is None:
                    return None
                fg = _hex_rgb(_effective_fg())
                sc = None
                if int(state.get('stroke_w', 0) or 0) > 0:
                    sc = _hex_rgb(state.get('stroke_color')
                                  or _contrast(_effective_fg()))
                # 主判定：BitBlt 截屏的色（更接近用户视觉）
                cols = cols_screen if cols_screen is not None else cols_gdi
                hits = 0
                for c in cols:
                    if c is None:
                        continue
                    if _near(c, fg) or (sc is not None and _near(c, sc)):
                        hits += 1
                # 49 个采样点里出现次数最多的颜色，做一个"屏幕在我这个
                # 区域里到底是什么颜色"的快速印象。失败时 dump 出来
                # 立刻就能看出来"是不是大屏上根本没画文字"。
                from collections import Counter
                palette = Counter(
                    tuple(c) for c in cols if c is not None
                ).most_common(3)
                # 对比两路命中情况：这是判断"GDI 缓存 vs 真屏幕"差异的关键
                hits_gdi = 0
                if cols_gdi is not None:
                    for c in cols_gdi:
                        if c is None:
                            continue
                        if _near(c, fg) or (sc is not None and _near(c, sc)):
                            hits_gdi += 1
                # 49 点中两路颜色不一样的点数（None 算 0）
                diff_pts = 0
                if cols_gdi is not None and cols_screen is not None:
                    for a, b in zip(cols_gdi, cols_screen):
                        if a is None or b is None:
                            continue
                        # 比较 RGB，不完全相同就算"两路颜色不一致"
                        if (abs(a[0] - b[0]) + abs(a[1] - b[1])
                                + abs(a[2] - b[2])) > 24:
                            diff_pts += 1
                method = 'BitBlt' if cols_screen is not None else 'GetDC'
                detail = {
                    'cv_screen': (base_x, base_y),
                    'bbox_rel': (x1, y1, x2, y2),
                    'bbox_screen': (base_x + x1, base_y + y1,
                                    base_x + x2, base_y + y2),
                    'fg': fg, 'stroke': sc,
                    'palette': palette,
                    'method': method,
                    'hits_screen': hits,
                    'hits_gdi': hits_gdi,
                    'diff_pts': diff_pts,
                }
                return hits, len(cols), detail
            except Exception:
                return None

        def _diagnose(tag, quiet=True, fast=False):
            """自检 + 记录；返回一句人话结论（给"诊断记录"按钮显示）。"""
            try:
                if not _alive(win):
                    return '倒计时窗口已经关闭'
                win.update_idletasks()
                items = _monitors()
                idx = max(0, min(int(state.get('monitor_index', 0) or 0),
                                 len(items) - 1))
                want = items[idx]
                x, y = win.winfo_x(), win.winfo_y()
                w, h = max(1, win.winfo_width()), max(1, win.winfo_height())
                cx, cy = x + w // 2, y + h // 2
                real = None
                for m in items:
                    l, t, r, b = m['rect']
                    if l <= cx < r and t <= cy < b:
                        real = m
                        break
                probe = _sample_text()
                if probe is None:
                    verdict = '无法自检（这台机器不让读屏幕像素）'
                elif probe[0] > 0:
                    verdict = '画面正常（采样命中 %d/%d，方法=%s）' % (
                        probe[0], probe[1],
                        probe[2].get('method', '?'))
                else:
                    verdict = '看不到画面（采样 %d 点都不是文字色，方法=%s）' % (
                        probe[1], probe[2].get('method', '?'))
                # 关键诊断：GetDC（GDI 缓存）和 BitBlt（真屏幕）是否一致？
                # 不一致 → 说明 DWM 没把这个 layered 窗口合成到屏幕上，
                # GDI 缓存说画上了，但用户视觉看不到。这是扩展屏的经典 bug。
                if probe is not None:
                    hg = probe[2].get('hits_gdi', 0)
                    hs = probe[2].get('hits_screen', 0)
                    dp = probe[2].get('diff_pts', 0)
                    if hg > 0 and hs == 0 and dp > 5:
                        verdict += ('  ⚠ GDI缓存命中 %d 但真屏幕 0：'
                                    'DWM 没合成 layered 窗口到屏幕' % hg)
                    elif hg == 0 and hs > 0 and dp > 5:
                        verdict += ('  ⚠ 真屏幕命中 %d 但 GDI 缓存 0：'
                                    '这台机器 GetDC 读到的是别的东西' % hs)
                    elif hg > 0 and hs > 0 and dp > 5:
                        verdict += ('  ⚠ 两路都命中但有 %d 点颜色不同' % dp)
                line = ('%s｜目标=%s 实际=%s｜窗口=%dx%d@%d,%d｜透明=%s｜%s｜显示器=%s'
                        % (tag,
                           '#%d%s' % (want['index'],
                                      '主屏' if want['primary'] else '扩展屏'),
                           ('#%d%s' % (real['index'],
                                       '主屏' if real['primary'] else '扩展屏'))
                           if real else '不落在任何显示器内',
                           w, h, x, y, state['transparent'], verdict,
                           ' / '.join('%d%s%s' % (
                               m['index'], '主' if m['primary'] else '扩',
                               m['rect']) for m in items)))
                _diag_log(line)
                if probe is None or probe[0] > 0:
                    return verdict
                # 失败时把 _sample_text 的详细信息 dump 到日志——
                # 看 49 点采到的颜色（应该是接近纯黑或纯白）就能判断
                # "是大屏真的没字，还是 sample_text 没采对位置"。
                detail = probe[2]
                pal = ', '.join('rgb%s×%d' % (rgb, cnt)
                                for rgb, cnt in detail.get('palette', []))
                _diag_log(('sample_detail｜cv屏幕=%s｜bbox相对canvas=%s'
                           '｜bbox绝对屏幕=%s｜期望fg=%s｜期望stroke=%s'
                           '｜采样49点里最常见的3个颜色：%s')
                          % (detail.get('cv_screen'),
                             detail.get('bbox_rel'),
                             detail.get('bbox_screen'),
                             detail.get('fg'),
                             detail.get('stroke') or '无',
                             pal or '无'))
                # 不再自动改透明模式 —— 之前切实心也没救回来，反而把
                # 状态搞乱让调试变难。这里只在日志里数连续失败次数。
                if not state['transparent'] and not _diag['solid_fail']:
                    _diag['solid_fail'] = True
                    _diag_log('注意：实心窗口也采不到画面，多半是这台机器'
                              '不允许读屏幕像素，或采样点位置错位')
                if _diag['auto']:
                    return verdict
                _diag['fail'] += 1
                _diag_log('连续 %d 次采不到文字色，按用户要求保持色键透明模式'
                          '（不自动切回实心；扩展屏已由 native overlay 接管显示）。'
                          % _diag['fail'])
                # ---- fallback 路径已禁用 ----
                # 用户要求"默认背景透明"。扩展屏已由 native overlay
                # 接管显示，主屏 Tk 浮窗透明模式稳定，不再需要 fallback。
                _diag['auto'] = True
                return verdict
            except Exception as exc:
                return '自检出错：%s' % exc

        def _dump_text_bbox(tag):
            """把当前画面的文字 bbox 写进诊断日志。

            主要是排查"开始计时之后程序以为画面异常但其实没异常"——
            如果采样点采不到文字色，但 bbox 在屏幕上的位置完全合理，
            那就是 sample 坐标错位，不是真的看不见。
            """
            try:
                if not _alive(win):
                    return
                try:
                    cv.update_idletasks()
                    win.update_idletasks()
                except Exception:
                    pass
                item = getattr(cv, '_text_item', None)
                if item is None:
                    _diag_log('dump_bbox[%s]：cv._text_item = None（可能 _refresh'
                              ' 里没建文字）' % tag)
                    return
                bbox = cv.bbox(item)
                base_x, base_y = cv.winfo_rootx(), cv.winfo_rooty()
                if not bbox:
                    _diag_log('dump_bbox[%s]：cv.bbox 返回空（文字可能还没渲染）'
                              % tag)
                    return
                x1, y1, x2, y2 = bbox
                _diag_log(('dump_bbox[%s]｜cv屏幕=(%d,%d)｜bbox相对=(%d,%d,%d,%d)'
                           '｜bbox绝对屏幕=(%d,%d)-(%d,%d)｜宽=%dx高=%d')
                          % (tag, base_x, base_y, x1, y1, x2, y2,
                             base_x + x1, base_y + y1,
                             base_x + x2, base_y + y2,
                             x2 - x1, y2 - y1))
            except Exception as exc:
                _diag_log('dump_bbox[%s] 异常：%s' % (tag, exc))

        def _repaint_soon(delays=(60, 180, 400, 800, 1500), kick=False):
            """移动位置后连补几次重画。

            扩展屏上的色键透明窗口，移完只有当帧重画容易“画不出来”，
            表现就是：点了位置按钮，大屏上的倒计时就看不见了。
            这里隔一段时间再补几次，系统总会画上。
            """
            for job in repaint_jobs:
                try:
                    win.after_cancel(job)
                except tk.TclError:
                    pass
            repaint_jobs[:] = []
            for delay in delays:
                try:
                    repaint_jobs.append(
                        win.after(delay, lambda: _repaint(lift=True)))
                except tk.TclError:
                    break
            if kick:
                # 补画之后再重贴一次分层属性（安全做法，不再隐藏/显示窗口）
                try:
                    repaint_jobs.append(
                        win.after(max(delays) + 250, _present))
                except tk.TclError:
                    pass

        def _set_time(total):
            state['total'] = max(1, min(int(total), 359999))
            state['remaining'] = state['total']
            _refresh()

        def _apply_time(m_var, s_var):
            try:
                total = int(m_var.get()) * 60 + int(s_var.get())
            except (tk.TclError, ValueError):
                return
            _set_time(total)

        def _start():
            _diag_log('进入 _start')  # 临时探针：确认点按钮真的进了这个函数
            if state['running']:
                _diag_log('_start 早 return: state[\'running\']=True')
                return
            if state['remaining'] <= 0:
                state['remaining'] = state['total']
            if state['remaining'] <= 0:
                _diag_log('_start 早 return: remaining<=0')
                return
            state['end_time'] = time.time() + state['remaining']
            state['running'] = True
            state['warn_done'] = False  # 本轮计时开始时重置中途提醒
            _diag_log('_start: 准备 _tick()')
            _tick()
            _diag_log('_start: _tick() 返回，准备 _present()')
            # 开始计时会连着重画：色键透明浮窗偶尔会在这时候被系统漏合成，
            # 表现就是"一点开始，大屏上的倒计时整块没了"，这里补一次硬重贴。
            _present()
            _diag_log('_start: _present() 返回，准备 _repaint()')
            _repaint(lift=True)
            _diag_log('_start: _repaint() 返回，准备 _kick_extended_overlay()')
            # 扩展屏上的色键透明窗口改 size 后，DWM 偶尔漏掉分层位图合成，
            # 仅靠上一步的 SetLayeredWindowAttributes + RedrawWindow 救不回来。
            # 临时去掉 WS_EX_LAYERED 再加回来，强制 DWM 重建分层位图——
            # 这一步骤只会作用在投影端，主屏上的副本不受影响。
            _kick_extended_overlay()
            _diag_log('_start: _kick_extended_overlay() 返回，准备清零 fail/auto')
            # 新一轮：失败计数清零（之前切实心没救回就别再切了）
            # 注意：一旦 auto=True 被设上了，说明这台机器的 fallback
            # 也不奏效——后续不应该反复触发切模式。但 _start 进入早期
            # 会清零，所以这里只清 fail 计数，保留 auto / solid_fail
            # 的历史标记，让 _diagnose 自己判断。
            _diag['fail'] = 0
            # 强制 tk 把这一帧画完再交给 DWM，否则 _sample_text 拿到的
            # bbox 可能是更新前的旧值。
            try:
                cv.update_idletasks()
                win.update_idletasks()
            except Exception:
                pass
            _diag_log('_start: 注册 1000ms 后 _diagnose，600ms 后 dump text bbox')
            # 1000ms 后立刻自检一次；同时 600ms 后 dump 文字 bbox，
            # 看"开始计时那一刻文字在屏幕的什么位置"。
            win.after(600, _dump_text_bbox, '开始计时+600ms')
            win.after(1000, _diagnose, '开始计时', True, False)

        def _pause():
            state['running'] = False
            if state['job'] is not None:
                try:
                    win.after_cancel(state['job'])
                except tk.TclError:
                    pass
                state['job'] = None
            _refresh()

        def _toggle_run():
            if state['running']:
                _pause()
            else:
                _start()

        def _reset():
            _pause()
            state['remaining'] = state['total']
            _refresh()

        def _play_sound():
            """到时提示音：有自定义声音文件则播放，否则用默认提示音。"""
            import winsound
            sf = state.get('sound_file')
            if sf:
                try:
                    ext = os.path.splitext(sf)[1].lower()
                    if ext == '.wav':  # wav 直接播放
                        winsound.PlaySound(
                            sf, winsound.SND_FILENAME | winsound.SND_ASYNC)
                        return
                    # mp3/mpeg 等：用系统 Windows Media Player 后台播放
                    import win32com.client
                    try:
                        wmp = win32com.client.Dispatch('WMPlayer.OCX')
                        wmp.settings.volume = 100
                        wmp.URL = sf
                        wmp.controls.play()
                    except Exception:
                        wmp = None
                    # 延迟关闭播放器，避免声音中断
                    def _quit(pl=wmp):
                        try:
                            pl.controls.stop()
                        except Exception:
                            pass
                    if wmp is not None:
                        try:
                            import threading
                            threading.Timer(30.0, _quit).start()
                        except Exception:
                            pass
                    return
                except Exception:
                    pass
            # 默认：系统提示音（比 Beep 更可靠），失败再退回 Beep
            try:
                winsound.PlaySound(
                    'SystemExclamation',
                    winsound.SND_ALIAS | winsound.SND_ASYNC)
            except Exception:
                winsound.Beep(880, 600)
                winsound.Beep(660, 1000)

        def _tick():
            try:
                if not state['running']:
                    return
                try:
                    if not win.winfo_exists():
                        state['running'] = False
                        return
                except tk.TclError:
                    state['running'] = False
                    return
                remain = int(state['end_time'] - time.time())
                if remain <= 0:
                    state['remaining'] = 0
                    state['running'] = False
                    _refresh()
                    if state['sound']:
                        try:
                            _play_sound()
                        except Exception:
                            pass
                    return
                state['remaining'] = remain
                # 剩余时间到设定阈值时声音提醒一次
                wa = int(state.get('warn_at', 0) or 0)
                if wa > 0 and 0 < remain <= wa and not state.get('warn_done'):
                    state['warn_done'] = True
                    try:
                        _play_sound()
                    except Exception:
                        pass
                _refresh()
                state['job'] = win.after(500, _tick)
            except tk.TclError:
                state['running'] = False

        def _ppt_on_show():
            """检测 PowerPoint / WPS 演示是否正在放映。"""
            try:
                import win32com.client as wc
            except Exception:
                return False
            for progid in ('PowerPoint.Application', 'KWPP.Application'):
                try:
                    app = wc.GetActiveObject(progid)
                except Exception:
                    continue
                try:
                    if app.SlideShowWindows.Count > 0:
                        return True
                except Exception:
                    pass
            return False

        def _check_ppt():
            try:
                if not win.winfo_exists():
                    state['ppt_job'] = None
                    return
                if not state.get('link_ppt'):
                    # 关掉“PPT放映自动开始”后不再检测放映状态：
                    # 既让开关真正生效，也省掉每 0.5 秒一次的系统调用
                    state['ppt_was_on'] = False
                    state['ppt_job'] = win.after(1000, _check_ppt)
                    return
                on = _ppt_on_show()
                if on and not state['ppt_was_on'] and not state['running']:
                    # PPT 放映开始：重置并启动倒计时
                    _reset()
                    _start()
                elif not on and state['ppt_was_on'] and state['running']:
                    # PPT 退出放映：按设置选择复位或仅停止
                    if state.get('ppt_reset'):
                        _reset()
                    else:
                        _pause()
                state['ppt_was_on'] = on
            except Exception:
                pass
            try:
                state['ppt_job'] = win.after(500, _check_ppt)
            except tk.TclError:
                state['ppt_job'] = None

        # ---- 拖动窗口 ----
        def _on_press(e):
            state['drag_x'] = e.x
            state['drag_y'] = e.y

        def _on_drag(e):
            x = win.winfo_x() + e.x - state['drag_x']
            y = win.winfo_y() + e.y - state['drag_y']
            # Tk 主浮窗（操作端）只在主屏活动：扩展屏的倒计时由
            # native overlay 接管显示，拖到主屏边界就被吸附回来，
            # 不允许跨屏，避免跟 native overlay 撞在同一块屏幕。
            pr = _primary_rect()
            left, top, right, bottom = pr
            try:
                w = max(int(win.winfo_width()), 1)
                h = max(int(win.winfo_height()), 1)
            except tk.TclError:
                w, h = 260, 100
            # 窗口整体必须落在主屏矩形内（不是只看左上角）。
            # 窗口比主屏还大时退化为左上角对齐。
            x_max = max(left, right - w)
            y_max = max(top, bottom - h)
            x = max(left, min(x, x_max))
            y = max(top, min(y, y_max))
            win.geometry('+%d+%d' % (x, y))
            _repaint()
            # 拖动期间也实时把设置面板贴着 Tk 主浮窗走，避免松开手再"跳一下"。
            # native overlay 不跟：它是观众端，位置由 pos 按钮控制，不被
            # 操作者拖动 Tk 主浮窗影响。
            try:
                _sync_panel()
            except Exception:
                pass

        def _on_release(_event=None):
            """拖完同步设置面板位置。

            注意：Tk 主浮窗永远在主屏活动（_on_drag 已限制边界），
            所以这里**不再更新 monitor_index**——拖动不会再跨屏，
            切"主屏"/"扩展屏"必须通过设置面板的下拉框显式切换。
            """
            _sync_panel()
            _repaint()
            _repaint_soon()

        def _on_wheel(e):
            delta = 4 if e.delta > 0 else -4
            state['font_size'] = max(8, min(state['font_size'] + delta, 400))
            _refresh()

        # 把 _on_drag 挂到 state 让测试 / 跨入口能调到同一个闭包实例
        state['_on_drag_fn'] = _on_drag

        for w in (win, cv):
            w.bind('<Button-1>', _on_press)
            w.bind('<B1-Motion>', _on_drag)
            w.bind('<ButtonRelease-1>', _on_release)
            w.bind('<Button-3>', lambda e: _open_settings())
            w.bind('<Double-Button-1>', lambda e: _reset())
            w.bind('<MouseWheel>', _on_wheel)
            # 快捷键：空格 开始/暂停
            w.bind('<space>', lambda e: _toggle_run())

        # 让窗口本身能接收键盘事件
        win.bind('<space>', lambda e: _toggle_run())
        # 尺寸变化时重画，保证文字始终居中
        win.bind('<Configure>', lambda e: _draw())

        def _set_size(x):
            state['font_size'] = max(8, min(int(x), 400))
            _refresh()

        def _ask_size():
            v = simpledialog.askinteger(
                '字号', '字号 (8-400)：',
                initialvalue=state['font_size'],
                parent=state['settings_win'], minvalue=8, maxvalue=400)
            if v:
                _set_size(v)

        def _ask_time():
            v = simpledialog.askinteger(
                '倒计时', '总秒数 (1-359999)：',
                initialvalue=state['total'],
                parent=state['settings_win'], minvalue=1, maxvalue=359999)
            if v:
                _set_time(v)

        def _pick(key):
            """弹出 Windows 系统颜色选择对话框，把结果写到 state[key]。

            踩过的坑：
              * 用 ``parent=`` 给 askcolor 传父窗口时，如果父窗口不是顶层
                可见的 / 被 grab / 被禁用，Windows 会把"选择颜色"对话框
                放到 Z 序底下 → 用户看着就像没弹出来。
              * 修法：先尝试带 parent（modal 在父窗口上），若父窗口失效
                或者抛异常，自动退回成"无父窗口"再弹一次。
              * 同时把所有失败原因写到运行日志，避免下次再盲调。
            """
            from tkinter import colorchooser
            sw = state.get('settings_win')

            def _safe_parent():
                """父窗口是否健在可见？不在就回 None。"""
                try:
                    if sw is not None and sw.winfo_exists():
                        if sw.state() != 'iconic':
                            return sw
                except tk.TclError:
                    return None
                return None

            parent = _safe_parent()
            try:
                # 先 lift + focus_force，让"父窗口"在 Z 序最上面，
                # 否则 Windows 弹出的对话框会落在它后面看着像没弹
                if parent is not None:
                    try:
                        parent.lift()
                        parent.focus_force()
                    except tk.TclError:
                        parent = None
                # 看门狗：万一对话框落进双屏死区 / 被压底，拉回面板所在屏
                guard_color_dialog_visible(
                    parent if parent is not None else self.root)
                _, hex_color = colorchooser.askcolor(
                    color=state.get(key) or '#FFFFFF',
                    title='选择颜色', parent=parent)
            except Exception as exc:
                _log('colorchooser 异常（带 parent=%s）：%r'
                     % (bool(parent), exc))
                # 父窗口有问题导致 askcolor 抛异常 → 不带 parent 再试
                try:
                    guard_color_dialog_visible(self.root)
                    _, hex_color = colorchooser.askcolor(
                        color=state.get(key) or '#FFFFFF',
                        title='选择颜色')
                except Exception as exc2:
                    _log('colorchooser 重试仍失败：%r' % exc2)
                    return False
            if hex_color:
                state[key] = hex_color
                _refresh()
                return True
            return False

        def _toggle(key):
            state[key] = not state[key]
            _refresh()

        def _flip_check(v, k):
            # 勾选框为开关，点击必然切换状态：直接翻转 state 并强制
            # variable 与 state 同步，确保勾选显示与实际状态一致。
            # （设置面板已改用 BooleanVar，这里保留给旧代码调用）
            state[k] = not state[k]
            try:
                v.set('1' if state[k] else '0')
            except tk.TclError:
                pass
            _refresh()

        def _toggle_link():
            state['link_ppt'] = not state['link_ppt']
            if state['link_ppt'] and state['ppt_job'] is None:
                _check_ppt()
            elif not state['link_ppt'] and state['ppt_job'] is not None:
                try:
                    win.after_cancel(state['ppt_job'])
                except tk.TclError:
                    pass
                state['ppt_job'] = None

        def _set_alpha(a):
            state['alpha'] = a
            _refresh()

        def _set_stroke(val):
            try:
                state['stroke_w'] = max(0, min(int(val), 60))
            except (tk.TclError, ValueError):
                return
            _refresh()

        def _set_stroke_auto():
            state['stroke_color'] = None
            _refresh()

        def _place_panel(sw):
            """把设置面板贴到 Tk 主浮窗旁边（永远跟浮窗走）。

            过去这里按 ``monitor_index`` 切：主屏 / 扩展屏锚点不一样，面板
            经常跑去另一个屏幕或者主程序窗口底下——跟 Tk 主浮窗的实际位置
            完全脱钩。现在 Tk 主浮窗锁主屏、native overlay 锁扩展屏，所以
            **面板永远以 Tk 主浮窗为锚**：

            1. 默认跟浮窗垂直对齐（中心线对齐），上下方向二选一——
               优先放浮窗**正下方**，放不下翻到**正上方**，避免跟浮窗撞。
            2. 浮窗已**非常贴近**屏幕边（pos 在角 + margin=8）时，面板"上 /
               下"都可能溢出，自动翻到**水平方向**（左/右）。
            3. 再做屏幕内边界保护，确保不出主屏 32px 余量。
            """
            try:
                sw.update_idletasks()
                win.update_idletasks()
                sw_w = sw.winfo_reqwidth()
                sw_h = sw.winfo_reqheight()
                wx = win.winfo_rootx()
                wy = win.winfo_rooty()
                ww = max(int(win.winfo_width()), 1)
                wh = max(int(win.winfo_height()), 1)
            except tk.TclError:
                return
            # 主屏矩形 = 边界盒
            items = _monitors()
            left = top = right = bottom = None
            for item in items:
                if item.get('primary'):
                    left, top, right, bottom = item['rect']
                    break
            if left is None:
                try:
                    sw_w_s = win.winfo_screenwidth()
                    sw_h_s = win.winfo_screenheight()
                except tk.TclError:
                    sw_w_s, sw_h_s = 1920, 1080
                left, top, right, bottom = 0, 0, sw_w_s, sw_h_s
            # 水平居中对齐浮窗中心
            cx = wx + ww // 2
            x_center = cx - sw_w // 2
            x_left = wx
            # ---- 选择垂直方向（下 vs 上）----
            gap = 6
            y_below = wy + wh + gap
            y_above = wy - sw_h - gap
            can_below = (y_below + sw_h) <= (bottom - 10)
            can_above = (y_above >= top + 10)
            if can_below:
                y = y_below
            elif can_above:
                y = y_above
            else:
                # 上下都放不下 → 翻成水平方向：左/右
                y = max(top + 8, min(wy, bottom - sw_h - 8))
                if (wx - sw_w - gap) >= (left + 10):
                    x_left = wx - sw_w - gap
                elif (wx + ww + gap + sw_w) <= (right - 10):
                    x_left = wx + ww + gap
                else:
                    x_left = cx - sw_w // 2
                x = x_left
                edge = 32
                x = min(max(x, left + edge), max(right - sw_w - edge,
                                                 left + edge))
                y = min(max(y, top + edge), max(bottom - sw_h - edge,
                                                 top + edge))
                sw.geometry('+%d+%d' % (x, y))
                return
            x = x_center
            # ---- 边界保护 ----
            edge = 32
            x = min(max(x, left + edge), max(right - sw_w - edge,
                                             left + edge))
            y = min(max(y, top + edge), max(bottom - sw_h - edge,
                                             top + edge))
            sw.geometry('+%d+%d' % (x, y))

        def _sync_panel():
            """计时器窗口移动后，设置面板重新贴到其周围（不乱跳）。"""
            sw = state.get('settings_win')
            if sw is None:
                return
            try:
                if not sw.winfo_exists():
                    return
                _place_panel(sw)
            except tk.TclError:
                pass

        def _move_pos(pos):
            try:
                # 记录当前位置——native overlay 跟着这个走
                state['current_pos'] = pos
                win.update_idletasks()
                # 按当前选中的显示器摆位，不跨屏跑回主屏
                left, top, right, bottom = _active_rect()
                sw = right - left
                sh = bottom - top
                w = win.winfo_width()
                h = win.winfo_height()
                if pos == 'center':
                    x, y = (sw - w) // 2, (sh - h) // 2
                elif pos == 'tl':
                    x, y = 0, 0
                elif pos == 'tr':
                    x, y = sw - w, 0
                elif pos == 'bl':
                    x, y = 0, sh - h
                elif pos == 'br':
                    x, y = sw - w, sh - h
                win.geometry('+%d+%d' % (left + x, top + y))
                # 关键：先把位移真正生效，再重画。
                # 否则 _fit_window/_place_panel 读到的还是旧坐标，会把新位置盖回去，
                # 表现就是“点了位置按钮，倒计时不见了”。
                win.update_idletasks()
                _refresh()
                _present()
                _repaint(lift=True)
                _repaint_soon(kick=True)
                _sync_panel()
                # native overlay（扩展屏那个）跟着位置按钮走，
                # 但 Tk 主浮窗已经在主屏 _move_pos 摆完了
                _move_native_overlay(pos)
                # 移完等画面稳定下来，自检一次“大屏上到底有没有画出来”
                win.after(900, lambda: _diagnose('点位置'))
            except tk.TclError:
                pass

        # ---- 右键设置面板 ----
        # 用普通有边框 Toplevel 替代 Menu 组件：menu 挂在无边框透明窗口上，
        # Windows 的 grab/焦点管理会失效导致不自动关闭/置底。
        # 普通窗口关闭可靠，彻底消除该问题。设置实时生效，最后点“完成”关闭。
        def _settings_done():
            sw = state.get('settings_win')
            if sw is not None:
                try:
                    sw.destroy()
                except tk.TclError:
                    pass
                state['settings_win'] = None
            state['sync_monitor_ui'] = None
            state['run_btn'] = None
            state['check_vars'] = {}

        def _open_settings():
            # 已在打开则置前（先恢复最小化）
            sw = state.get('settings_win')
            if sw is not None:
                try:
                    if sw.winfo_exists():
                        if sw.state() == 'iconic':
                            sw.deiconify()
                        sw.lift()
                        sw.focus_force()
                        return
                except tk.TclError:
                    pass
            sw = tk.Toplevel(self.root)
            sw.title('倒计时设置')
            sw.resizable(False, False)
            # 先藏起来再建内容：Toplevel 一创建就会按默认位置（WC 级联，
            # 常常落在浮窗左边）先露出来一次，等 _place_panel 摆好再跳位，
            # 看着就是“左边闪一下的窗口”。withdraw 后等摆好位置再显示。
            sw.withdraw()
            ui_theme.style_window(sw)
            try:
                # 设置面板始终置顶，方便操作，不随“窗口置顶”勾选变化
                sw.attributes('-topmost', True)
            except tk.TclError:
                pass
            state['settings_win'] = sw

            p = ui_theme.PALETTE
            body = tk.Frame(sw, bg=p['bg'], padx=14, pady=12)
            body.pack(fill=tk.BOTH, expand=True)
            tk.Label(body, text='倒计时设置', font=ui_theme.font(12, 'bold'),
                     bg=p['bg'], fg=p['text']).pack(anchor='w')
            tk.Label(body, text='调整后立即生效，看完点“关闭面板”',
                     font=ui_theme.font(8), bg=p['bg'],
                     fg=p['faint']).pack(anchor='w', pady=(1, 10))

            def _section(title, collapsible=False):
                """一栏设置：带标题的分组框，内部两列对齐。

                ``collapsible=True`` 时点标题栏可折叠/展开整栏内容
                （标题末尾 ▾ 展开、▸ 收起），面板不至于太长。
                """
                lf = ttk.Labelframe(body, text=' %s ' % title,
                                    padding=(10, 8))
                lf.pack(fill=tk.X, pady=(0, 8))
                lf.grid_columnconfigure(1, weight=1)
                if not collapsible:
                    return lf

                opened = {'on': True}
                managed = {}    # 子控件 → 最初的布局方式（收起后就查不到了）

                def _apply():
                    for w in lf.winfo_children():
                        mgr = managed.get(w) or w.winfo_manager()
                        if mgr:
                            managed[w] = mgr
                        if opened['on']:
                            if mgr == 'grid':
                                w.grid()
                            elif mgr == 'pack':
                                w.pack()
                        else:
                            if mgr == 'grid':
                                w.grid_remove()
                            elif mgr == 'pack':
                                w.pack_forget()
                    if opened['on']:
                        # height=0 → 按内容自动撑开
                        lf.configure(text=' %s ▾' % title, height=0)
                    else:
                        # 收起时留一条标题栏，否则整栏会缩成 1px 看不见
                        lf.configure(text=' %s ▸' % title, height=30)

                def _toggle_fold(_event=None):
                    # 只有点在顶部标题条上才折叠/展开，点内容区空白不响应
                    if (_event is not None and opened['on']
                            and getattr(_event, 'y', 0) > 26):
                        return
                    opened['on'] = not opened['on']
                    _apply()
                    # 高度变了重新贴一次，别让展开后的面板跑出屏幕
                    try:
                        _place_panel(lf.winfo_toplevel())
                    except tk.TclError:
                        pass

                lf.configure(cursor='hand2')
                lf.bind('<Button-1>', _toggle_fold)
                lf._toggle_fold = _toggle_fold
                return lf

            LBL = {'sticky': 'e', 'padx': (0, 8), 'pady': 3}

            # —— 页1：时长 ——
            sec_time = _section('时长')
            m_var = tk.IntVar(value=state['total'] // 60)
            s_var = tk.IntVar(value=state['total'] % 60)

            def _sync_time(*_):
                _apply_time(m_var, s_var)

            # 直接输入数字或点箭头都实时生效（Spinbox command 只响应箭头）
            m_var.trace_add('write', _sync_time)
            s_var.trace_add('write', _sync_time)
            ttk.Label(sec_time, text='倒计时').grid(row=0, column=0, **LBL)
            frm_time = tk.Frame(sec_time, bg=p['bg'])
            frm_time.grid(row=0, column=1, sticky='w')
            ttk.Spinbox(frm_time, from_=0, to=5999, width=4,
                        textvariable=m_var).pack(side='left')
            ttk.Label(frm_time, text='分').pack(side='left', padx=(3, 10))
            ttk.Spinbox(frm_time, from_=0, to=59, width=3,
                        textvariable=s_var).pack(side='left')
            ttk.Label(frm_time, text='秒').pack(side='left', padx=(3, 0))

            # 预设下拉：3 4 5 8 10 15 30 60 分钟
            presets = (180, 240, 300, 480, 600, 900, 1800, 3600)
            pre_box = ttk.Combobox(
                sec_time, state='readonly', width=9, cursor='arrow',
                values=['%d 分钟' % (s // 60) for s in presets])

            def _preset_label():
                for sec in presets:
                    if state['total'] == sec:
                        return '%d 分钟' % (sec // 60)
                return ''

            def _on_preset(_=None):
                txt = pre_box.get()
                if txt.endswith('分钟'):
                    try:
                        _set_time(int(txt.replace('分钟', '')) * 60)
                    except ValueError:
                        pass

            pre_box.set(_preset_label())
            pre_box.bind('<<ComboboxSelected>>', _on_preset)
            ttk.Label(sec_time, text='常用时长').grid(row=1, column=0, **LBL)
            pre_box.grid(row=1, column=1, sticky='w', pady=3)

            # 开始/暂停：双击浮窗、按空格也能开关，但按钮最直观
            frm_run = tk.Frame(sec_time, bg=p['bg'])
            frm_run.grid(row=2, column=0, columnspan=2, sticky='w',
                         pady=(8, 0))
            run_btn = ttk.Button(frm_run, text='开始计时', width=10,
                                 style='Accent.TButton', command=_toggle_run)
            run_btn.pack(side='left')
            ttk.Button(frm_run, text='复位', width=6, style='Ghost.TButton',
                       command=_reset).pack(side='left', padx=(8, 0))
            state['run_btn'] = run_btn
            _sync_run_btn()

            # —— 显示组（默认折叠，点标题栏展开）——
            sec_show = _section('显示', collapsible=True)

            # 投屏显示器：默认扩展屏（投影仪/TV），可切到任意一块屏
            mon_label_of = lambda item: item['name']
            mon_var = tk.StringVar()
            mon_box = ttk.Combobox(
                sec_show, state='readonly', width=24, cursor='arrow',
                textvariable=mon_var,
                values=[mon_label_of(m) for m in _monitors()])

            def _sync_monitor_ui():
                items = _monitors()
                idx = max(0, min(int(state.get('monitor_index', 0) or 0),
                                 len(items) - 1))
                mon_box.configure(values=[mon_label_of(m) for m in items])
                mon_var.set(mon_label_of(items[idx]))

            def _on_monitor_sel(_event=None):
                items = _monitors()
                for item in items:
                    if mon_label_of(item) == mon_var.get():
                        _apply_monitor(item['index'])
                        return

            mon_box.bind('<<ComboboxSelected>>', _on_monitor_sel)
            _sync_monitor_ui()
            state['sync_monitor_ui'] = _sync_monitor_ui
            ttk.Label(sec_show, text='投屏显示器').grid(row=0, column=0, **LBL)
            mon_box.grid(row=0, column=1, sticky='w', pady=3)

            tk.Label(sec_show,
                     text='默认投到“扩展屏”（投影仪/TV）；主屏浮窗用来看和'
                          '操作，可拖动、滚轮缩放，双击复位倒计时。',
                     font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                     justify='left', wraplength=250).grid(
                row=7, column=0, columnspan=2, sticky='w', pady=(4, 0))

            # 窗口位置（五个方位）
            ttk.Label(sec_show, text='位置').grid(row=1, column=0, **LBL)
            frm_pos = tk.Frame(sec_show, bg=p['bg'])
            frm_pos.grid(row=1, column=1, sticky='w', pady=3)
            for lbl, pos in [('中央', 'center'), ('左上', 'tl'),
                             ('右上', 'tr'), ('左下', 'bl'), ('右下', 'br')]:
                ttk.Button(frm_pos, text=lbl, width=4,
                           command=lambda q=pos: _move_pos(q)).pack(
                    side='left', padx=(0, 4))

            # 字号 / 描边粗细
            sz_var = tk.IntVar(value=state['font_size'])

            def _on_size(*_):
                _set_size(sz_var.get())

            sz_var.trace_add('write', _on_size)  # 输入或点箭头都实时生效
            sw_var = tk.IntVar(value=state['stroke_w'])

            def _on_stroke(*_):
                _set_stroke(sw_var.get())

            sw_var.trace_add('write', _on_stroke)  # 输入或点箭头都实时生效
            ttk.Label(sec_show, text='字号').grid(row=2, column=0, **LBL)
            frm_size = tk.Frame(sec_show, bg=p['bg'])
            frm_size.grid(row=2, column=1, sticky='w', pady=3)
            ttk.Spinbox(frm_size, from_=8, to=400, width=4,
                        textvariable=sz_var).pack(side='left')
            ttk.Label(frm_size, text='描边粗细').pack(
                side='left', padx=(14, 6))
            ttk.Spinbox(frm_size, from_=0, to=60, width=4,
                        textvariable=sw_var).pack(side='left')

            # 文字颜色：按钮 + 色块，一眼看到当前颜色
            ttk.Label(sec_show, text='文字颜色').grid(row=3, column=0, **LBL)
            frm_color = tk.Frame(sec_show, bg=p['bg'])
            frm_color.grid(row=3, column=1, sticky='w', pady=3)
            fg_swatch = tk.Label(frm_color, width=3, bd=0, bg=state['color'],
                                 cursor='hand2', highlightthickness=1,
                                 highlightbackground=p['border'])
            bg_swatch = tk.Label(frm_color, width=3, bd=0,
                                 bg=state['bg_color'], cursor='hand2',
                                 highlightthickness=1,
                                 highlightbackground=p['border'])

            def _pick_swatch(key, swatch):
                if _pick(key):
                    try:
                        swatch.configure(bg=state[key])
                    except tk.TclError:
                        pass

            ttk.Button(frm_color, text='前景', width=5,
                       command=lambda: _pick_swatch('color', fg_swatch)).pack(
                side='left')
            fg_swatch.pack(side='left', padx=(4, 14))
            ttk.Button(frm_color, text='背景', width=5,
                       command=lambda: _pick_swatch('bg_color', bg_swatch)
                       ).pack(side='left')
            bg_swatch.pack(side='left', padx=(4, 0))
            # 色块带手型光标，用户自然会直接点色块选色——绑上同样的动作
            fg_swatch.bind('<Button-1>',
                           lambda e: _pick_swatch('color', fg_swatch))
            bg_swatch.bind('<Button-1>',
                           lambda e: _pick_swatch('bg_color', bg_swatch))

            # 描边颜色
            ttk.Label(sec_show, text='描边颜色').grid(row=4, column=0, **LBL)
            frm_stroke = tk.Frame(sec_show, bg=p['bg'])
            frm_stroke.grid(row=4, column=1, sticky='w', pady=3)
            stroke_swatch = tk.Label(
                frm_stroke, width=3, bd=0, cursor='hand2',
                bg=state.get('stroke_color') or _contrast(_fg()),
                highlightthickness=1, highlightbackground=p['border'])

            def _pick_stroke():
                # 描边宽度为 0 时描边不绘制，选色看不到效果：
                # 自动设为 2 并同步输入框，让所选颜色立即可见
                if not _pick('stroke_color'):
                    return
                try:
                    stroke_swatch.configure(bg=state['stroke_color'])
                except tk.TclError:
                    pass
                if int(state.get('stroke_w', 0) or 0) == 0:
                    state['stroke_w'] = 2
                    try:
                        sw_var.set(2)
                    except tk.TclError:
                        pass

            def _stroke_auto():
                _set_stroke_auto()
                try:
                    stroke_swatch.configure(bg=_contrast(_fg()))
                except tk.TclError:
                    pass

            ttk.Button(frm_stroke, text='选色', width=5,
                       command=_pick_stroke).pack(side='left')
            stroke_swatch.pack(side='left', padx=(4, 10))
            stroke_swatch.bind('<Button-1>', lambda e: _pick_stroke())
            ttk.Button(frm_stroke, text='自动反差', width=9,
                       command=_stroke_auto).pack(side='left')

            # 不透明度
            alpha_var = tk.IntVar(value=int(state['alpha'] * 100))

            def _on_alpha(v):
                alpha_var.set(int(v))
                _set_alpha(int(v) / 100.0)

            ttk.Label(sec_show, text='不透明度').grid(row=5, column=0, **LBL)
            frm_alpha = tk.Frame(sec_show, bg=p['bg'])
            frm_alpha.grid(row=5, column=1, sticky='w', pady=3)
            a_scale = ui_theme.SeekBar(frm_alpha, from_=10, to=100,
                                       width=118, command=_on_alpha,
                                       bg=p['bg'])
            a_scale.set(int(state['alpha'] * 100))
            a_scale.pack(side='left')
            ttk.Label(frm_alpha, textvariable=alpha_var, width=4,
                      anchor='e').pack(side='left', padx=(6, 0))

            # 扩展屏字号缩放（只动字号，不动窗口尺寸 → 仍能贴 4 角）
            # 默认 1.0 = 跟主屏字号一致；调到 1.5/2.0 让远处投影更清楚。
            zoom_pct_var = tk.IntVar(value=int(float(state['native_zoom']) * 100))

            def _on_native_zoom(v):
                try:
                    pct = int(float(v))
                except Exception:
                    pct = 100
                zoom_pct_var.set(pct)
                state['native_zoom'] = pct / 100.0
                # 不销毁重建：_paint 每次都用 _native_boost 算字号传给
                # native.set_text(font_size=...)，set_text 内部发现字号
                # 变了会自己重建字体。同步 boost 后直接重画即可——
                # 之前在这里 close 旧 overlay 再 _refresh()，但
                # _refresh 不会重建 native → 扩展屏倒计时直接消失。
                state['_native_boost'] = pct / 100.0
                if state.get('_native_overlay') is None:
                    # 还没建（比如刚切到扩展屏）就补建一次
                    state['_native_failed'] = False
                    try:
                        _ensure_native_overlay()
                    except Exception:
                        pass
                _refresh()

            ttk.Label(sec_show, text='扩展屏字号缩放').grid(
                row=6, column=0, **LBL)
            frm_zoom = tk.Frame(sec_show, bg=p['bg'])
            frm_zoom.grid(row=6, column=1, sticky='w', pady=3)
            z_scale = ui_theme.SeekBar(frm_zoom, from_=100, to=300,
                                       width=118, command=_on_native_zoom,
                                       bg=p['bg'])
            z_scale.set(int(float(state['native_zoom']) * 100))
            z_scale.pack(side='left')
            ttk.Label(frm_zoom, textvariable=zoom_pct_var, width=5,
                      anchor='e').pack(side='left', padx=(6, 0))

            # —— 行为组（默认折叠，点标题栏展开）——
            sec_behave = _section('行为', collapsible=True)
            checks = tk.Frame(sec_behave, bg=p['bg'])
            checks.grid(row=0, column=0, columnspan=2, sticky='w')
            checks.grid_columnconfigure(1, weight=1)
            for idx, (lbl, key) in enumerate([
                    ('背景透明', 'transparent'),
                    ('到时声音提示', 'sound'),
                    ('不足1分钟变红', 'warn_red'),
                    ('PPT放映自动开始', 'link_ppt'),
                    ('PPT退出倒计时复位', 'ppt_reset')]):
                var = tk.BooleanVar(value=bool(state[key]))
                # 自动切换实心模式时要能同步勾选框，所以记一份引用
                vars_map = state.get('check_vars')
                if vars_map is None:
                    vars_map = state['check_vars'] = {}
                vars_map[key] = var

                def _on_check(k=key, v=var):
                    state[k] = bool(v.get())
                    _refresh()

                ttk.Checkbutton(checks, text=lbl, variable=var,
                                cursor='arrow', command=_on_check).grid(
                    row=idx // 2, column=idx % 2, sticky='w',
                    padx=(0, 12), pady=2)

            tk.Label(sec_behave,
                     text='大屏上看不到倒计时？①点下面的“诊断记录”看结论；'
                          '②取消勾选“背景透明”改成实心窗口，一定看得见',
                     font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                     wraplength=330, justify='left').grid(
                row=3, column=0, columnspan=2, sticky='w', pady=(6, 0))

            frm_fix = tk.Frame(sec_behave, bg=p['bg'])
            frm_fix.grid(row=4, column=0, columnspan=2, sticky='w',
                         pady=(6, 0))

            def _diag_report():
                """手动诊断：结论直接弹出来，同时写进诊断日志。"""
                text = _diagnose('手动诊断')
                _diag_log('手动诊断结论：%s' % text)
                try:
                    mons = _monitors()
                    mi = max(0, min(int(state.get('monitor_index', 0) or 0),
                                    len(mons) - 1))
                    where = mons[mi]['name']
                except Exception:
                    where = '未知'
                try:
                    messagebox.showinfo(
                        '倒计时自检',
                        '%s\n\n目标屏：%s\n诊断日志：\n%s\n\n'
                        '若结论是“看不到画面”，请取消勾选“背景透明”，'
                        '改用实心窗口。' % (text, where, _diag_path()),
                        parent=state.get('settings_win'))
                except tk.TclError:
                    pass

            ttk.Button(frm_fix, text='修复显示', style='Ghost.TButton',
                       command=_kick_window).pack(side='left')
            ttk.Button(frm_fix, text='诊断记录', style='Ghost.TButton',
                       command=_diag_report).pack(side='left', padx=(8, 0))

            # 到时提示音：内置 1-5 号可直接选择，也可自定义声音文件
            ttk.Label(sec_behave, text='到时提示音').grid(row=1, column=0,
                                                          **LBL)

            snd_dir = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), 'files')
            snd_box = ttk.Combobox(
                sec_behave, state='readonly', width=14, cursor='arrow',
                values=('默认提示音', '提示音1', '提示音2', '提示音3',
                        '提示音4', '提示音5', '自定义…'))

            def _snd_label():
                sf = state.get('sound_file')
                if not sf:
                    return '默认提示音'
                base = os.path.basename(sf)
                if os.path.dirname(sf) == snd_dir and base.endswith(
                        '.mp3') and base[:-4].isdigit():
                    return '提示音%d' % int(base[:-4])
                return base if len(base) <= 16 else base[:13] + '…'

            def _pick_sound():
                p = filedialog.askopenfilename(
                    parent=state['settings_win'], title='选择到时提示音',
                    filetypes=[('音频文件', '*.wav;*.mp3;*.mpeg'),
                               ('所有文件', '*.*')])
                if p:
                    state['sound_file'] = p
                    snd_box.set(_snd_label())

            def _on_snd_sel(event=None):
                val = snd_box.get()
                if val == '默认提示音':
                    state['sound_file'] = None
                elif val == '自定义…':
                    _pick_sound()
                else:
                    try:
                        n = int(val.replace('提示音', ''))
                    except ValueError:
                        return
                    state['sound_file'] = os.path.join(
                        snd_dir, '%d.mp3' % n)
                snd_box.set(_snd_label())

            snd_box.set(_snd_label())
            snd_box.bind('<<ComboboxSelected>>', _on_snd_sel)
            snd_box.grid(row=1, column=1, sticky='w', pady=3)

            # 剩余时间提醒：剩余 N 秒时响一次（0=关闭）
            warn_var = tk.IntVar(value=int(state.get('warn_at', 0) or 0))

            def _on_warn(*_):
                try:
                    state['warn_at'] = max(0, min(int(warn_var.get()), 3599))
                except (tk.TclError, ValueError):
                    return

            warn_var.trace_add('write', _on_warn)
            ttk.Label(sec_behave, text='剩余提醒').grid(row=2, column=0, **LBL)
            frm_warn = tk.Frame(sec_behave, bg=p['bg'])
            frm_warn.grid(row=2, column=1, sticky='w', pady=3)
            ttk.Spinbox(frm_warn, from_=0, to=3599, width=5,
                        textvariable=warn_var).pack(side='left')
            ttk.Label(frm_warn, text='秒时提醒一次（0＝关闭）').pack(
                side='left', padx=(6, 0))

            # —— 底部：复位 / 隐藏到托盘 / 退出计时器 / 关闭面板 ——
            frm_save = tk.Frame(body, bg=p['bg'])
            frm_save.pack(fill=tk.X, pady=(2, 0))
            ttk.Button(frm_save, text='关闭面板', style='Accent.TButton',
                       command=_settings_done).pack(side=tk.RIGHT)
            ttk.Button(frm_save, text='复位计时', style='Ghost.TButton',
                       command=_reset).pack(side=tk.RIGHT, padx=(0, 8))
            ttk.Button(frm_save, text='退出计时器', style='Danger.TButton',
                       command=self._close_countdown_requested).pack(side=tk.LEFT)
            ttk.Button(frm_save, text='隐藏到托盘', style='Ghost.TButton',
                       command=_hide_to_tray).pack(side=tk.LEFT, padx=(8, 0))

            # 折叠/展开时面板宽度别跟着跳：先记下三栏当前宽度固定住，
            # 再把「显示」「行为」默认收起（点标题栏展开），面板更短。
            try:
                body.update_idletasks()
                for _lf in (sec_time, sec_show, sec_behave):
                    _lf.configure(width=_lf.winfo_reqwidth())
            except tk.TclError:
                pass
            for _lf in (sec_show, sec_behave):
                _lf._toggle_fold()

            # —— 点击面板空白处可拖动 ——
            _SKIP = ('Button', 'Checkbutton', 'Scale', 'Spinbox',
                     'Entry', 'Listbox', 'Combobox', 'Menubutton',
                     'Notebook', 'Tab', 'Canvas', 'Text', 'Labelframe')

            def _is_interactive(w):
                cls = w.winfo_class() if hasattr(w, 'winfo_class') else ''
                if cls.endswith('Label'):
                    return True
                return any(cls.endswith(s) for s in _SKIP)

            # 拖动统一在顶层 sw 处理：空白处按下可拖；交互控件由回调跳过。
            # 光标统一使用普通箭头，不设移动图标。
            sw._dragging = False
            sw._grab_dx = 0
            sw._grab_dy = 0

            def _sw_press(e):
                if _is_interactive(e.widget):
                    sw._dragging = False
                    return
                sw._grab_dx = e.x_root - sw.winfo_x()
                sw._grab_dy = e.y_root - sw.winfo_y()
                sw._dragging = True

            def _sw_drag(e):
                if not sw._dragging:
                    return
                sw.geometry('+%d+%d' % (e.x_root - sw._grab_dx,
                                        e.y_root - sw._grab_dy))

            sw.configure(cursor='arrow')
            sw.bind('<Button-1>', _sw_press, add='+')
            sw.bind('<B1-Motion>', _sw_drag, add='+')

            # 弹出到主窗口正下方（放不下则翻到上方）
            # 位置算好了再 deiconify：这样窗口第一次出现就在最终位置，
            # 不会先闪一下再跳。
            try:
                _place_panel(sw)
                sw.update_idletasks()
                sw.deiconify()
            except tk.TclError:
                pass
            sw.lift()
            sw.focus_force()

        win._countdown_state = state
        _refresh()
        # 在扩展屏上同时建一个 Win32 原生分层窗口（绕开 Tk 的 DWM 合成 bug）。
        # 主屏上 Tk 渲染正常，不创建 native overlay（避免双层叠加）。
        _ensure_native_overlay()
        # 无边框透明窗口刚摆到扩展屏上时，系统有时不发重画消息（字没画出来
        # 就等于整块隐形），头几秒多催几次
        for _delay in (200, 600, 1200, 2200):
            try:
                win.after(_delay, _repaint)
            except tk.TclError:
                break
        # 打开后先自检一次：这时候能看到画面，后面出事才好对比
        win.after(1500, lambda: _diagnose('打开窗口'))

        def _keep_visible():
            """透明浮窗偶尔会被系统漏画成“隐形”，每隔几秒补一次。"""
            try:
                if not win.winfo_exists():
                    return
                _repaint()
            except tk.TclError:
                pass
            try:
                win.after(3000, _keep_visible)
            except tk.TclError:
                pass

        ticks = {}

        def _ensure_on_screen():
            """倒计时要是跑到了看不见的地方（拔显示器/换分辨率），拉回可见区。"""
            try:
                if not win.winfo_exists():
                    return
                win.update_idletasks()
                w = max(1, win.winfo_width())
                h = max(1, win.winfo_height())
                cx = win.winfo_x() + w // 2
                cy = win.winfo_y() + h // 2
                items = _monitors()
                visible = False
                for item in items:
                    left, top, right, bottom = item['rect']
                    if left <= cx < right and top <= cy < bottom:
                        visible = True
                        break
                if not visible:
                    left, top, right, bottom = _active_rect()
                    x = max(left, min(right - w - 20,
                                      left + (right - left - w) // 2))
                    y = max(top, min(bottom - h - 20,
                                     top + (bottom - top - h) // 2))
                    win.geometry('+%d+%d' % (x, y))
                    win.update_idletasks()
                    # monitor_index 不再根据 Tk 浮窗位置推断——它只
                    # 表示 native overlay 用哪块屏，由用户在设置面板
                    # 显式指定。Tk 主浮窗永远在主屏，monitor_index
                    # 与 Tk 浮窗位置解耦。
                    _refresh()
                    _repaint(lift=True)
                    sync = state.get('sync_monitor_ui')
                    if sync is not None:
                        sync()
                    _sync_panel()
                else:
                    # 定期补一次重画：色键透明窗口在扩展屏上偶尔会自己
                    # 变成“隐形”（窗口在、字没画出来），补画就恢复了
                    _repaint(lift=True)
                    # 每 5 次（约 10 秒）再硬重贴一次：连字带背景一起
                    # 没了的那种“整块隐形”，只有重贴分层属性才救得回来
                    ticks['n'] = ticks.get('n', 0) + 1
                    if ticks['n'] % 5 == 0:
                        _present()
            except tk.TclError:
                pass
            finally:
                # 无论如何都要把下一次排上：以前这句写在 try 里，
                # 一旦哪次抛异常，这个“保命”循环就再也不跑了。
                try:
                    win.after(2000, _ensure_on_screen)
                except tk.TclError:
                    pass

        win.after(3200, _keep_visible)
        win.after(1500, _ensure_on_screen)
        if state['link_ppt']:
            _check_ppt()

        # ---- 隐藏到托盘 ----
        # 闭包定义在所有依赖都准备好之后：_pause /
        # _ensure_native_overlay / _settings_done 都在它之前已经定义。
        def _hide_to_tray():
            """把倒计时窗口藏到任务栏右下角图标里；计时暂停。

            触发：右键设置面板 → 点"隐藏到托盘"。
            行为：暂停计时（用户要求"隐藏后暂停"）→ 隐藏 Tk 主浮窗 →
                  隐藏 native overlay → 关闭设置面板 → 建/刷新托盘图标。
            """
            self._hide_countdown_to_tray(win, state)

        def _restore_from_tray():
            """从托盘恢复倒计时窗口。

            触发：双击托盘图标 / 托盘菜单 → 显示倒计时。
            行为：保持隐藏时的暂停状态（用户自己决定要不要继续）→
                  重新显示 Tk 主浮窗 → 重新建/显示 native overlay
                  （如果 monitor_index 仍在扩展屏）→ 移除托盘图标。
            """
            self._restore_countdown_from_tray(win, state)

    # ---- 隐藏到托盘 / 恢复：单独提出来，因为 _activate_or_close_others
    # ---- 也要调用恢复逻辑（用户在主程序点"PPT倒计时"按钮也能恢复）

    def _hide_countdown_to_tray(self, win, state):
        """把指定倒计时窗口藏到任务栏右下角图标里；计时暂停。"""
        if state.get('_hidden'):
            return
        # 暂停计时——和 _pause() 闭包同样的逻辑（这里取不到闭包，
        # 直接复制一遍；只是把 job 取消 + running=False）。
        state['running'] = False
        if state.get('job') is not None:
            try:
                win.after_cancel(state['job'])
            except tk.TclError:
                pass
            state['job'] = None
        state['_hidden_was_running'] = bool(state.get('running', False))
        state['_hidden'] = True
        try:
            win.withdraw()
        except tk.TclError:
            pass
        native = state.get('_native_overlay')
        if native is not None:
            try:
                native.hide()
            except Exception:
                pass
        # 关掉设置面板
        sw = state.get('settings_win')
        if sw is not None:
            try:
                sw.destroy()
            except tk.TclError:
                pass
            state['settings_win'] = None
            state['sync_monitor_ui'] = None
            state['run_btn'] = None
        # 建托盘图标。首次隐藏时才建；之后重复隐藏只刷 tooltip
        if state.get('_tray') is None and _tray_icon is not None:
            try:
                state['_tray'] = _tray_icon.TrayIcon(
                    tip='会务助手 - 倒计时（已隐藏）',
                    icon_id=_tray_icon.IDI_INFORMATION,
                    on_double_click=lambda: self._restore_countdown_from_tray(
                        win, state),
                    menu=[
                        _tray_icon.MenuItem(
                            '显示倒计时',
                            lambda: self._restore_countdown_from_tray(
                                win, state)),
                        _tray_icon.MenuItem(None),
                        _tray_icon.MenuItem(
                            '关闭计时器',
                            lambda: self._close_countdown_requested()),
                    ],
                )
                state['_tray'].show()
            except Exception:
                state['_tray'] = None
        elif state.get('_tray') is not None:
            try:
                state['_tray'].set_tip('会务助手 - 倒计时（已隐藏）')
            except Exception:
                pass

    def _restore_countdown_from_tray(self, win, state):
        """从托盘恢复指定倒计时窗口。

        保持隐藏时的暂停状态（用户自己决定要不要继续）；
        如果隐藏前压根没在跑，恢复后就还是没在跑——这符合"暂停后想用
        再调出来"的语义。
        """
        if not state.get('_hidden'):
            return
        state['_hidden'] = False
        try:
            win.deiconify()
            win.lift()
            win.focus_force()
        except tk.TclError:
            pass
        # native overlay：调 _ensure_native_overlay 会按 monitor_index
        # 决定建 / 不建；之前 hide 时没销毁的对象会保留原位置不动。
        # 之后再 show() 一下让窗口可见（hide 只是 SetWindowPos 不显示，
        # DIB 和分层属性都还在）。
        was_native = state.get('_native_overlay')
        ensure_native = state.get('_ensure_native_overlay_fn')
        if ensure_native is not None:
            try:
                ensure_native()
            except Exception:
                pass
        native = state.get('_native_overlay')
        if native is not None and not getattr(native, '_visible', True):
            try:
                native.show()
            except Exception:
                pass
        # 移掉托盘图标
        tray = state.get('_tray')
        if tray is not None:
            try:
                tray.remove()
            except Exception:
                pass
            state['_tray'] = None
        # 重画一下：色键透明浮窗刚 deiconify 出来有时是空的
        try:
            win.update_idletasks()
        except tk.TclError:
            pass
        refresh = state.get('_refresh_fn')
        if refresh is not None:
            try:
                refresh()
            except Exception:
                pass
        # 关键：Tk 色键透明浮窗在 withdraw→deiconify 之后，DWM 里
        # 的 layered cache 是空的，光靠 _refresh() 不够——必须像
        # _move_pos 末尾那样补一次 _present（重贴分层属性）+ _repaint
        # （强制重画）+ _repaint_soon（隔几毫秒多次重画），不然就
        # 会看到那种"主浮窗出来是一块空白，要点位置按钮才显示"的
        # 诡异现象。
        for fn_name in ('_present_fn', '_repaint_fn', '_repaint_soon_fn'):
            fn = state.get(fn_name)
            if fn is not None:
                try:
                    fn() if fn_name != '_repaint_soon_fn' else fn(kick=True)
                except Exception:
                    pass

    def _close_countdown_requested(self):
        """设置面板/托盘里主动退出倒计时：现场环节关键，先确认。"""
        win = self._countdown_win
        if self._confirm_close('PPT倒计时', win or self.root):
            self._close_countdown_win()

    def _close_countdown_win(self):
        """用户单独关闭倒计时浮动窗口。"""
        win = self._countdown_win
        if win is not None:
            state = getattr(win, '_countdown_state', None)
            if state is not None:
                state['running'] = False
                for key in ('job', 'ppt_job'):
                    if state.get(key) is not None:
                        try:
                            win.after_cancel(state[key])
                        except tk.TclError:
                            pass
                        state[key] = None
                # 移掉托盘图标——否则下次启动会有"幽灵图标"留几秒
                tray = state.get('_tray')
                if tray is not None:
                    try:
                        tray.remove()
                    except Exception:
                        pass
                    state['_tray'] = None
                state['_hidden'] = False
                # 销毁设置面板
                sw = state.get('settings_win')
                if sw is not None:
                    try:
                        sw.destroy()
                    except tk.TclError:
                        pass
                    state['settings_win'] = None
                # 销毁 Win32 native overlay（先 hide 再 close）
                native = state.get('_native_overlay')
                if native is not None:
                    try:
                        native.close()
                    except Exception:
                        pass
                    state['_native_overlay'] = None
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._countdown_win = None

    def _confirm_close(self, name, parent=None, main=False):
        """弹“关闭确认”；parent 尽量给具体功能窗，避免弹窗压在后面。"""
        return ui_theme.confirm_close(parent or self.root, name=name, main=main)

    def _close_main_requested(self):
        """首页点 X：退出整个程序前二次确认，防止现场误触。"""
        if self._confirm_close('', main=True):
            self._close_all()

    def _request_close(self, name, closer, parent=None):
        """功能窗点 X：统一二次确认后再走真正的关闭逻辑。"""
        if self._confirm_close(name, parent):
            closer()

    def _start_focus_guard(self):
        """启动现场守护：屏幕常亮 + 弹窗广告自动最小化。"""
        try:
            guard = focus_guard.FocusGuard(
                on_event=lambda ev: self.root.after(
                    0, self._on_guard_event))
            guard.start()
        except Exception:
            guard = None
        self._guard = guard
        self._update_guard_label()

    def _on_guard_event(self):
        """守护线程拦截到一个弹窗（已切回主线程）：刷新状态/面板。"""
        try:
            self._update_guard_label()
        except tk.TclError:
            pass
        win = self._guard_win
        if win is not None and win.winfo_exists():
            try:
                self._refresh_guard_list()
            except tk.TclError:
                pass

    def _update_guard_label(self):
        guard = self._guard
        if guard is None:
            text, fg = '● 防护不可用', '#fca5a5'
        elif guard.awake_on and guard.block_on:
            extra = (' · 已拦截 %d 个弹窗' % guard.count
                     if guard.count else '')
            text, fg = '● 现场防护中' + extra, '#dbeafe'
        elif guard.awake_on:
            text, fg = '● 屏幕常亮中', '#dbeafe'
        elif guard.block_on:
            extra = (' · 已拦截 %d 个弹窗' % guard.count
                     if guard.count else '')
            text, fg = '● 弹窗拦截中' + extra, '#dbeafe'
        else:
            text, fg = '○ 防护已暂停', '#cbd5e1'
        self._guard_label.config(text=text, fg=fg)

    def _open_guard_panel(self):
        win = self._guard_win
        if win is not None:
            try:
                win.lift()
                win.focus_force()
                return
            except tk.TclError:
                self._guard_win = None
        p = ui_theme.PALETTE
        win = tk.Toplevel(self.root)
        win.title('现场防护')
        win.configure(bg=WINDOW_BG)
        win.resizable(False, False)
        win.transient(self.root)
        _center_window(win, 460, 520)
        try:
            win.grab_set()
        except Exception:
            pass

        top = tk.Frame(win, bg=CARD_BG, bd=0, highlightthickness=1,
                       highlightbackground=BORDER_COLOR)
        top.pack(fill=tk.X, padx=14, pady=(14, 8))
        tk.Label(top, text='现场防护', font=ui_theme.font(13, 'bold'),
                 fg=p['text'], bg=CARD_BG).pack(anchor='w', padx=14,
                                                 pady=(12, 2))
        tk.Label(top,
                 text='软件运行期间阻止屏保和睡眠；检测到弹窗广告自动\n'
                      '最小化（不关闭进程），可在此查看并恢复。退出软件\n'
                      '后自动恢复系统原设置。',
                 font=ui_theme.font(9), fg=p['faint'], bg=CARD_BG,
                 justify=tk.LEFT).pack(anchor='w', padx=14, pady=(0, 10))

        opt = tk.Frame(win, bg=WINDOW_BG)
        opt.pack(fill=tk.X, padx=18)
        guard = self._guard
        var_awake = tk.BooleanVar(
            value=bool(guard and guard.awake_on))
        var_block = tk.BooleanVar(
            value=bool(guard and guard.block_on))

        def _toggle_awake():
            if guard is not None:
                guard.set_awake(var_awake.get())
                self._update_guard_label()

        def _toggle_block():
            if guard is not None:
                guard.set_block(var_block.get())
                self._update_guard_label()

        ttk.Checkbutton(opt, text='屏幕常亮（阻止屏保 / 黑屏 / 睡眠）',
                        variable=var_awake,
                        command=_toggle_awake).pack(anchor='w', pady=2)
        ttk.Checkbutton(opt, text='弹窗广告自动最小化',
                        variable=var_block,
                        command=_toggle_block).pack(anchor='w', pady=2)

        tk.Label(win, text='拦截记录', font=ui_theme.font(9, 'bold'),
                 fg=p['text'], bg=WINDOW_BG).pack(
            anchor='w', padx=18, pady=(12, 4))

        list_frame = tk.Frame(win, bg=BORDER_COLOR, bd=0)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=14)
        cols = ('time', 'title', 'process')
        tree = ttk.Treeview(list_frame, columns=cols, show='headings',
                            height=6)
        tree.heading('time', text='时间')
        tree.heading('title', text='弹窗标题')
        tree.heading('process', text='来源进程')
        tree.column('time', width=70, anchor='w', stretch=False)
        tree.column('title', width=240, anchor='w')
        tree.column('process', width=120, anchor='w')
        vsb = ttk.Scrollbar(list_frame, orient='vertical',
                            command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._guard_tree = tree

        def _restore():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo('现场防护', '请先选择一条记录。',
                                    parent=win)
                return
            if guard is None:
                return
            iid = sel[0]
            rec = guard.history[int(iid)]
            if guard.restore(rec['hwnd']):
                tree.item(iid, tags=('restored',))
                messagebox.showinfo('现场防护',
                                    '已恢复该窗口。', parent=win)
            else:
                messagebox.showinfo(
                    '现场防护',
                    '该窗口已经关闭，无法恢复。', parent=win)

        tree.tag_configure('restored', foreground=p['faint'])

        def _clear():
            if guard is not None:
                guard.clear_history()
                self._refresh_guard_list()
                self._update_guard_label()

        btns = tk.Frame(win, bg=WINDOW_BG)
        btns.pack(fill=tk.X, padx=14, pady=12)
        ui_theme.FlatButton(btns, text='恢复选中窗口',
                            command=_restore).pack(side=tk.LEFT)
        ui_theme.FlatButton(btns, text='清空记录', kind='secondary',
                            command=_clear).pack(side=tk.LEFT,
                                                 padx=(8, 0))
        def _close_panel():
            self._guard_win = None
            self._guard_tree = None
            win.destroy()

        win.protocol('WM_DELETE_WINDOW', _close_panel)
        ui_theme.FlatButton(btns, text='关闭',
                            command=_close_panel).pack(side=tk.RIGHT)
        self._guard_win = win
        self._refresh_guard_list()

    def _refresh_guard_list(self):
        win = self._guard_win
        tree = getattr(self, '_guard_tree', None)
        if win is None or tree is None:
            return
        guard = self._guard
        try:
            existing = set(tree.get_children())
        except tk.TclError:
            return
        needed = set()
        if guard is not None:
            for i, rec in enumerate(guard.history):
                iid = str(i)
                needed.add(iid)
                if iid not in existing:
                    tree.insert('', 'end', iid=iid,
                                values=(rec['time'], rec['title'],
                                        rec['process']))
        for iid in existing - needed:
            tree.delete(iid)
        if needed:
            tree.see(str(max(int(x) for x in needed)))

    def _close_all(self):
        """关闭主界面：各子窗口随之关闭，不会残留。"""
        if self._guard is not None:
            try:
                self._guard.stop()
            except Exception:
                pass
        self._close_welcome_win()
        self._close_ppt_win()
        self._close_print_win()
        self._close_countdown_win()
        self._close_music_win()
        self._close_inbox_win()
        self._close_video_win()
        self._close_lottery_win()
        self._close_downloader_win()
        self.root.destroy()

    def _close_music_win(self):
        """关闭会务音乐子窗口（停止播放）。"""
        win = self._music_win
        if win is not None:
            wmp = getattr(win, '_music_wmp', None)
            if wmp is not None:
                try:
                    wmp.controls.stop()
                except Exception:
                    pass
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._music_win = None


def _activate_existing_instance(window_title):
    """已有实例在跑时，把它从最小化/遮挡状态拉到前台。

    返回 True = 激活成功；False = 没找到窗口（极少见，比如第一实例还在
    初始化就被双击了）。本函数绝不抛异常 —— 任何步骤失败都返回 False。

    注意：minimized → restored 必须走 SendMessageTimeout(WM_SYSCOMMAND,
    SC_RESTORE) 而不是 ShowWindow(SW_RESTORE)。前者走 Tk 的 wm_state
    'normal' 路径，能正确恢复 980x680+ 的 geometry；后者会触发 Tk 的
    "几何重计算"路径，但此时 Tk 内部的 stored size 已被破坏，导致
    窗口被还原到 16x39（仅 title bar）—— 用户看到的就是"很小的框，
    什么也没有"。
    """
    try:
        import ctypes
        from ctypes import wintypes
        # 用本地 DLL handle 设 argtypes，不污染 ctypes.windll 共享对象
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        FindWindowW = user32.FindWindowW
        FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
        FindWindowW.restype = wintypes.HWND
        IsIconic = user32.IsIconic
        IsIconic.argtypes = [wintypes.HWND]
        IsIconic.restype = wintypes.BOOL
        SendMessageTimeoutW = user32.SendMessageTimeoutW
        SendMessageTimeoutW.argtypes = [
            wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM,
            ctypes.c_uint, ctypes.c_uint, ctypes.c_void_p]
        SendMessageTimeoutW.restype = ctypes.c_long
        ShowWindow = user32.ShowWindow
        ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
        ShowWindow.restype = wintypes.BOOL
        SetForegroundWindow = user32.SetForegroundWindow
        SetForegroundWindow.argtypes = [wintypes.HWND]
        SetForegroundWindow.restype = wintypes.BOOL
        BringWindowToTop = user32.BringWindowToTop
        BringWindowToTop.argtypes = [wintypes.HWND]
        BringWindowToTop.restype = wintypes.BOOL

        WM_SYSCOMMAND = 0x0112
        SC_RESTORE = 0xF120
        SMTO_ABORTIFHUNG = 0x0002

        hwnd = FindWindowW(None, window_title)
        if not hwnd:
            return False
        # 关键路径：发 WM_SYSCOMMAND/SC_RESTORE 给 Tk 主消息循环
        # Tk 会跑 wm_state 'normal' + raise，正确恢复 geometry
        # 500ms 超时；Tk 卡死时 fallback 到 ShowWindow(SW_RESTORE)
        if IsIconic(hwnd):
            SendMessageTimeoutW(hwnd, WM_SYSCOMMAND, SC_RESTORE, 0,
                                SMTO_ABORTIFHUNG, 500, None)
            # 再保险一次：如果上面没生效（极端情况），用 ShowWindow
            if IsIconic(hwnd):
                ShowWindow(hwnd, 9)        # SW_RESTORE
        # 已经可见的窗口就只是抢焦点
        SetForegroundWindow(hwnd)
        BringWindowToTop(hwnd)
        return True
    except Exception:
        return False


def main():
    # 单实例：Windows 命名互斥体保证同一账号同一会话下只开一个
    import ctypes
    _MUTEX_NAME = 'Local\\会务助手单实例'
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, False, _MUTEX_NAME)
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        # 已有实例在跑 → 把它拉到前台，本进程静默退出
        ok = _activate_existing_instance(WINDOW_TITLE)
        try:
            safe_guard.write_log(
                '检测到已有实例%s拉起，重复打开的窗口被本进程静默退出'
                % ('已' if ok else '未（未找到 HWND），'))
        except Exception:
            pass
        return

    root = tk.Tk()
    # 让任何其它进程都能把这个窗口拉到前台（解决双击 exe
    # 时"第二次启动被 SetForegroundWindow 拒绝"的 Windows 限制）
    try:
        ctypes.windll.user32.AllowSetForegroundWindow(-1)   # ASFW_ANY
    except Exception:
        pass
    # 全局兜底：任何界面回调/子线程出错只记录日志并提示，不让程序直接退出
    guard = safe_guard.install(app_name='会务助手', gui=True, root=root)
    safe_guard.write_log('会务助手启动')
    App(root)

    def check_components():
        try:
            import ppt_converter
            ok = ppt_converter.presentation_app_available()
        except Exception:
            ok = False
        if not ok:
            messagebox.showwarning(
                '缺少组件',
                '未检测到 WPS Office 或 Microsoft PowerPoint。\n\n'
                '“会务PPT转换”和“生成欢迎语”依赖 WPS 演示或'
                ' PowerPoint 组件，可能无法正常使用。\n\n'
                '打印座位牌等功能不受影响。')

    root.after(400, check_components)
    try:
        root.mainloop()
    finally:
        safe_guard.write_log('会务助手退出')
        del guard


if __name__ == '__main__':
    sys.exit(safe_guard.main_guard(app_name='会务助手', gui=True, func=main))
