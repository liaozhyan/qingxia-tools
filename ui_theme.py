# -*- coding: utf-8 -*-
"""界面主题：统一的现代扁平风格（配色、字体、控件样式、渐变标题栏、卡片）。

只在 Windows 自带的 Tk 上做视觉升级，不引入任何第三方界面库，
打包体积不变、也不增加崩溃风险。
"""

import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk


PALETTE = {
    'bg': '#f4f6fb',
    'surface': '#ffffff',
    'surface_alt': '#eef2fa',
    'hover': '#f7faff',
    'border': '#e5eaf3',
    'border_hover': '#bcd0f5',
    'text': '#1f2937',
    'muted': '#6b7280',
    'faint': '#9aa3b2',
    'primary': '#2563eb',
    'primary_dark': '#1d4ed8',
    'primary_soft': '#e8f0fe',
    'primary_grad': '#0ea5e9',
    'success': '#16a34a',
    'warning': '#d97706',
    'danger': '#dc2626',
}

ACCENTS = {
    'ppt': '#2563eb',
    'welcome': '#0d9488',
    'print': '#ea580c',
    'countdown': '#e11d48',
    'music': '#7c3aed',
    'inbox': '#059669',
    'wechat': '#07c160',
    'video': '#b45309',
    'lottery': '#65a30d',
    'strip': '#0891b2',
    'tools': '#475569',
    'qrcode': '#db2777',
    'downloader': '#ca8a04',
    'screenshot': '#0d9488',
    'prompter': '#0ea5e9',
    'lan': '#0891b2',
    'photo': '#16a34a',
    'mail': '#dc2626',
}

_FONT_FAMILY = None
_FONT_OVERRIDE = None


def set_font_override(fam):
    """给后续创建的字体临时指定族（抽签大屏用）；传空则恢复自动探测。

    注意：只影响“之后新建”的字体，已在界面上建好的控件字体不变。
    """
    global _FONT_OVERRIDE
    _FONT_OVERRIDE = (fam or '').strip() or None


def font_family():
    """挑一个系统里有的现代中文字体。"""
    if _FONT_OVERRIDE:
        return _FONT_OVERRIDE
    global _FONT_FAMILY
    if _FONT_FAMILY:
        return _FONT_FAMILY
    candidates = ('Microsoft YaHei UI', '微软雅黑', 'Microsoft YaHei',
                  'Segoe UI', 'Tahoma')
    try:
        families = set(tkfont.families())
    except Exception:
        return 'Tahoma'  # 还没有窗口，先不缓存，等 Tk 就绪后再探测
    for name in candidates:
        if name in families:
            _FONT_FAMILY = name
            break
    else:
        _FONT_FAMILY = 'Tahoma'
    return _FONT_FAMILY


def font(size=10, weight='normal'):
    if weight == 'normal':
        return (font_family(), size)
    return (font_family(), size, weight)


def confirm_close(parent, name='', main=False):
    """会务功能窗口关闭前的统一二次确认，防止现场误触 X 中断会议环节。

    name 为功能名（如“会务音乐”）；main=True 用于主界面退出。
    返回 True 表示用户确认关闭。对话框本身出异常时放行关闭，
    避免任何极端情况下把窗口关不掉。
    """
    from tkinter import messagebox
    try:
        if main:
            msg = ('关闭主界面将退出整个会务助手，所有已打开的会务功能'
                   '（欢迎屏、打印、会务音乐、视频/长条屏投屏、倒计时、'
                   '局域网互传等）都会立即停止。\n\n'
                   '会务现场操作事关环节衔接，请再次确认：现在要退出吗？')
        else:
            msg = ('【%s】正在会务现场使用中。\n\n'
                   '关闭后该功能会立即停止，可能导致正在进行的环节中断'
                   '（如播放的音乐、投屏画面、倒计时、正在收发的文件等）。\n\n'
                   '为避免误操作，请确认：现在要关闭%s吗？'
                   % (name, name))
        return messagebox.askokcancel(
            '关闭确认', msg, parent=parent, icon='warning',
            default='cancel')
    except Exception:
        return True


def mix(color_a, color_b, ratio):
    """按比例混合两个 #RRGGBB 颜色（ratio=0 → a，1 → b）。"""
    try:
        a = color_a.lstrip('#')
        b = color_b.lstrip('#')
        out = []
        for i in (0, 2, 4):
            va = int(a[i:i + 2], 16)
            vb = int(b[i:i + 2], 16)
            out.append(int(round(va + (vb - va) * ratio)))
        return '#%02x%02x%02x' % tuple(out)
    except Exception:
        return color_a


def apply_theme(root):
    """把现代样式套到 ttk 控件上；可重复调用。"""
    p = PALETTE
    try:
        # 默认底色：让子窗口里没写颜色的 Frame、Label 也统一
        # （注意：不要在这里设 *Font。字体名带空格时 Tk 解析会失败，
        #   回退成“宋体 18”，界面字体会整体变大，这里保持系统默认字体）
        root.option_add('*Frame.background', p['bg'])
        root.option_add('*Label.background', p['bg'])
        root.option_add('*Label.foreground', p['text'])
        root.option_add('*Button.background', p['surface'])
        root.option_add('*Button.foreground', p['text'])
        root.option_add('*Button.activeBackground', p['primary_soft'])
        root.option_add('*Button.activeForeground', p['primary_dark'])
        root.option_add('*Button.highlightBackground', p['border'])
        root.option_add('*Scale.background', p['bg'])
        root.option_add('*Checkbutton.background', p['bg'])
        root.option_add('*Radiobutton.background', p['bg'])
    except Exception:
        pass

    try:
        style = ttk.Style(root)
    except Exception:
        return
    try:
        style.theme_use('clam')
    except Exception:
        pass

    base = font_family()
    try:
        # 默认字号保持 9pt（和系统默认一致），不要整体放大
        style.configure('.', background=p['bg'],
                        foreground=p['text'], focuscolor=p['primary_soft'])

        style.configure('TButton', padding=(9, 4),
                        relief='flat', borderwidth=0,
                        background=p['surface_alt'], foreground=p['text'])
        style.map('TButton',
                  background=[('pressed', p['border']),
                              ('active', p['primary_soft']),
                              ('disabled', p['surface_alt'])],
                  foreground=[('disabled', p['faint'])])

        style.configure('Accent.TButton', background=p['primary'],
                        foreground='#ffffff', font=(base, 9, 'bold'),
                        padding=(11, 4))
        style.map('Accent.TButton',
                  background=[('pressed', p['primary_dark']),
                              ('active', p['primary_dark']),
                              ('disabled', mix(p['primary'], '#ffffff', 0.62))],
                  foreground=[('disabled', '#ffffff')])

        style.configure('Ghost.TButton', background=p['surface'],
                        foreground=p['primary'], padding=(9, 4))
        style.map('Ghost.TButton',
                  background=[('pressed', p['primary_soft']),
                              ('active', p['primary_soft']),
                              ('disabled', p['surface'])],
                  foreground=[('disabled', p['faint'])])

        style.configure('Danger.TButton', background=p['surface'],
                        foreground=p['danger'], padding=(9, 4))
        style.map('Danger.TButton',
                  background=[('active', '#fdeaea'), ('pressed', '#fbdada')])

        for name in ('TEntry', 'TSpinbox'):
            style.configure(name, fieldbackground=p['surface'],
                            background=p['surface'], foreground=p['text'],
                            bordercolor=p['border'], lightcolor=p['border'],
                            darkcolor=p['border'], insertcolor=p['text'],
                            padding=3, relief='flat')
        style.map('TEntry',
                  bordercolor=[('focus', p['primary'])],
                  lightcolor=[('focus', p['primary'])],
                  darkcolor=[('focus', p['primary'])])

        style.configure('TCombobox', fieldbackground=p['surface'],
                        background=p['surface'], foreground=p['text'],
                        bordercolor=p['border'], lightcolor=p['border'],
                        darkcolor=p['border'], arrowcolor=p['muted'],
                        fieldpadding=(4, 1))
        # 下拉框聚焦时不要变成蓝框（点完一直留着蓝边很扎眼），
        # 只保证只读状态下底色是白的、各态前景/箭色齐全，
        # 避免展开时尺寸抖动、选中后文字"空白"
        style.map('TCombobox',
                  fieldbackground=[('readonly', p['surface'])],
                  background=[('readonly', p['surface'])],
                  foreground=[('readonly', p['text']),
                              ('disabled', p['faint'])],
                  arrowcolor=[('active', p['primary_dark'])])
        # 下拉列表（popdown Listbox）若不设颜色/字体，项和选中态会"空白"
        for _opt, _val in (
                ('*TCombobox*Listbox.background', p['surface']),
                ('*TCombobox*Listbox.foreground', p['text']),
                ('*TCombobox*Listbox.selectBackground', p['primary']),
                ('*TCombobox*Listbox.selectForeground', '#ffffff'),
                ('*TCombobox*Listbox.font', (base, 9))):
            try:
                root.option_add(_opt, _val)
            except Exception:
                pass

        style.configure('TRadiobutton', background=p['bg'],
                        foreground=p['text'], focuscolor=p['bg'],
                        indicatorcolor=p['surface'])
        style.map('TRadiobutton',
                  indicatorcolor=[('selected', p['primary']),
                                  ('pressed', p['primary_soft'])],
                  foreground=[('disabled', p['faint'])])

        # 复选框用自绘的对勾图（clam 默认画的是「X」）
        _install_checkbutton(style, root, p)

        style.configure('TLabelframe', background=p['bg'],
                        bordercolor=p['border'], relief='solid',
                        borderwidth=1, padding=6)
        style.configure('TLabelframe.Label', background=p['bg'],
                        foreground=p['muted'])

        style.configure('TNotebook', background=p['bg'], borderwidth=0,
                        tabmargins=(0, 4, 0, 0))
        style.configure('TNotebook.Tab', font=(base, 9),
                        background=p['surface_alt'], foreground=p['muted'],
                        padding=(14, 6), borderwidth=0)
        style.map('TNotebook.Tab',
                  background=[('selected', p['surface'])],
                  foreground=[('selected', p['primary'])],
                  # clam 主题给选中态单独设了 padding {6 4 6 2}，
                  # 不覆盖的话点一下标签就会变高 2px、变窄 8px，来回跳
                  padding=[('selected', (14, 6)), ('!selected', (14, 6))])

        style.configure('TProgressbar', troughcolor=p['surface_alt'],
                        background=p['primary'], bordercolor=p['surface_alt'],
                        lightcolor=p['primary'], darkcolor=p['primary'],
                        thickness=8)

        style.configure('TScale', background=p['bg'],
                        troughcolor=p['surface_alt'], bordercolor=p['border'])

        style.configure('Vertical.TScrollbar', background=p['surface_alt'],
                        troughcolor=p['bg'], bordercolor=p['bg'],
                        arrowcolor=p['muted'], borderwidth=0,
                        lightcolor=p['surface_alt'],
                        darkcolor=p['surface_alt'])
        style.map('Vertical.TScrollbar',
                  background=[('active', p['border_hover'])])

        style.configure('TSeparator', background=p['border'])

        # 列表（收件记录等）
        style.configure('Treeview', background=p['surface'],
                        fieldbackground=p['surface'], foreground=p['text'],
                        rowheight=26, borderwidth=0,
                        font=(base, 9))
        style.map('Treeview',
                  background=[('selected', p['primary_soft'])],
                  foreground=[('selected', p['primary_dark'])])
        style.configure('Treeview.Heading', background=p['surface_alt'],
                        foreground=p['muted'], relief='flat',
                        borderwidth=0, padding=(8, 6), font=(base, 9))
        style.map('Treeview.Heading',
                  background=[('active', p['border'])])
    except Exception:
        pass


def style_window(win):
    """子窗口统一底色，避免和主界面风格不一致。"""
    try:
        win.configure(bg=PALETTE['bg'])
    except Exception:
        pass
    apply_theme(win)


def force_repaint(win):
    """让 Windows 立刻把窗口（含里面的画布）重画一遍。

    无边框窗口直接摆到扩展屏/投影仪上时，系统有时不给它发重画消息，
    结果大屏上一直是一块黑（本机预览却正常）。这里主动补一次重绘。
    """
    try:
        import ctypes
        user32 = ctypes.windll.user32
        rdw_invalidate = 0x0001
        rdw_updatenow = 0x0100
        rdw_allchildren = 0x0080
        user32.RedrawWindow(
            ctypes.c_void_p(int(win.winfo_id())), None, None,
            rdw_invalidate | rdw_updatenow | rdw_allchildren)
        return True
    except Exception:
        return False


def present_window(win):
    """把窗口真正“亮”到大屏上：抬 z 序 + 置顶 + 重画。

    无边框窗口直接摆到扩展屏/投影仪时，系统有时既不重画也不合成，
    屏幕上是黑的。实测抬一次 z 序（lift/topmost）才会显示出来，
    所以每次换画面都顺手做一遍。
    """
    try:
        win.lift()
    except Exception:
        pass
    try:
        win.attributes('-topmost', True)
    except Exception:
        pass
    return force_repaint(win)


_LAYER_HELPERS = {}


def _layer_api():
    """取（并缓存）分层窗口相关的 user32 入口，句柄按 64 位处理。"""
    if _LAYER_HELPERS:
        return _LAYER_HELPERS
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        get_long = getattr(user32, 'GetWindowLongPtrW', None) \
            or user32.GetWindowLongW
        get_long.argtypes = [wintypes.HWND, ctypes.c_int]
        get_long.restype = ctypes.c_ssize_t
        user32.GetParent.argtypes = [wintypes.HWND]
        user32.GetParent.restype = wintypes.HWND
        user32.SetLayeredWindowAttributes.argtypes = [
            wintypes.HWND, wintypes.COLORREF, ctypes.c_ubyte, wintypes.DWORD]
        user32.SetLayeredWindowAttributes.restype = wintypes.BOOL
        user32.SetWindowPos.argtypes = [
            wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
            ctypes.c_int, ctypes.c_int, ctypes.c_uint]
        user32.SetWindowPos.restype = wintypes.BOOL
        user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.ShowWindow.restype = wintypes.BOOL
        user32.RedrawWindow.argtypes = [
            wintypes.HWND, ctypes.c_void_p, ctypes.c_void_p,
            ctypes.c_uint]
        user32.RedrawWindow.restype = wintypes.BOOL
        _LAYER_HELPERS.update({
            'user32': user32, 'get_long': get_long, 'wintypes': wintypes,
            'ptr': ctypes.c_void_p})
    except Exception:
        _LAYER_HELPERS['user32'] = None
    return _LAYER_HELPERS


def _window_handles(win):
    """Tk 窗口底层句柄（含包装窗口）：分层属性在包装窗口上才生效。"""
    api = _layer_api()
    user32 = api.get('user32')
    if user32 is None:
        return []
    handles = []
    try:
        h = int(win.winfo_id())
        guard = 0
        while h and guard < 8:
            handles.append(h)
            h = int(user32.GetParent(h) or 0)
            guard += 1
    except Exception:
        pass
    return handles


def present_overlay(win, key=None, alpha=None):
    """色键透明浮窗掉画面时，从 Win32 层面把它重新合成回屏幕。

    无边框 + ``-transparentcolor`` 的窗口（倒计时那种浮层）在扩展屏上
    移动、或者开始计时重画之后，Windows 偶尔会漏掉一次分层窗口合成：
    窗口还在、也在最前面，可整块都是透明的——背景和字全看不见。
    Tk 层面重画救不回来，必须重新贴一遍分层属性并强制重画。
    """
    ok = force_repaint(win)
    if not sys.platform.startswith('win'):
        return ok
    api = _layer_api()
    user32 = api.get('user32')
    if user32 is None:
        return ok
    try:
        if key is None:
            try:
                key = win.attributes('-transparentcolor')
            except Exception:
                key = None
        if alpha is None:
            try:
                alpha = float(win.attributes('-alpha'))
            except Exception:
                alpha = 1.0
        keyref, use_key = 0, False
        if key:
            try:
                r, g, b = win.winfo_rgb(str(key))
                keyref = (((r >> 8) & 0xFF)
                          | (((g >> 8) & 0xFF) << 8)
                          | (((b >> 8) & 0xFF) << 16))
                use_key = True
            except Exception:
                use_key = False
        try:
            aval = max(0, min(int(round(float(alpha) * 255)), 255))
        except Exception:
            aval = 255
        # LWA_COLORKEY=1 / LWA_ALPHA=2
        flags = 0x00000002 | (0x00000001 if use_key else 0)
        # SWP_NOSIZE|SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE
        # 注意：不要用 SWP_FRAMECHANGED——它会要求系统重算窗口边框，
        # 分层窗口（色键透明）在扩展屏上被这么一折腾常常整块不再合成。
        swp = 0x0001 | 0x0002 | 0x0004 | 0x0010
        # RDW_INVALIDATE|RDW_ERASE|RDW_FRAME|RDW_UPDATENOW|RDW_ALLCHILDREN
        rdw = 0x0001 | 0x0004 | 0x0010 | 0x0100 | 0x0080
        for h in _window_handles(win):
            try:
                ex = int(api['get_long'](h, -20))  # GWL_EXSTYLE
            except Exception:
                ex = 0
            if ex & 0x00080000:  # WS_EX_LAYERED
                try:
                    user32.SetLayeredWindowAttributes(
                        h, keyref, aval, flags)
                    user32.SetWindowPos(h, None, 0, 0, 0, 0, swp)
                except Exception:
                    pass
            try:
                user32.RedrawWindow(h, None, None, rdw)
            except Exception:
                pass
        return True
    except Exception:
        return ok


def make_opaque(win):
    """把窗口彻底变回普通窗口（去掉分层样式），用于“实心”模式。

    色键透明窗口在个别显示器/远程显示上会整块不显示；不透明时干脆不用
    分层窗口，用最普通的窗口，系统一定会画出来。
    """
    if not sys.platform.startswith('win'):
        return False
    api = _layer_api()
    user32 = api.get('user32')
    if user32 is None:
        return False
    try:
        import ctypes
        set_long = getattr(user32, 'SetWindowLongPtrW', None) \
            or user32.SetWindowLongW
        set_long.argtypes = [api['wintypes'].HWND, ctypes.c_int,
                             ctypes.c_ssize_t]
        set_long.restype = ctypes.c_ssize_t
        # SWP_NOSIZE|SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE|SWP_FRAMECHANGED
        swp = 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020
        for h in _window_handles(win):
            try:
                ex = int(api['get_long'](h, -20))  # GWL_EXSTYLE
            except Exception:
                continue
            if ex & 0x00080000:  # WS_EX_LAYERED
                try:
                    set_long(h, -20, ex & ~0x00080000)
                    user32.SetWindowPos(h, None, 0, 0, 0, 0, swp)
                    user32.RedrawWindow(h, None, None, 0x0001 | 0x0004 | 0x0100)
                except Exception:
                    pass
        force_repaint(win)
        return True
    except Exception:
        return False


def set_layered(win):
    """确保窗口带着 WS_EX_LAYERED（从“实心”切回“透明”时必须做）。

    实测：我们自己把分层样式去掉之后，Tk 再设 -transparentcolor 并不会
    把样式加回来（它以为还设着），于是窗口会变成一块实心色块。
    这里主动补上样式，色键/透明度由随后的 present_overlay 贴。
    """
    if not sys.platform.startswith('win'):
        return False
    api = _layer_api()
    user32 = api.get('user32')
    if user32 is None:
        return False
    try:
        import ctypes
        handles = _window_handles(win)
        if not handles:
            return False
        h = handles[-1]        # 最外层：Tk 把色键设在它上面
        set_long = getattr(user32, 'SetWindowLongPtrW', None) \
            or user32.SetWindowLongW
        set_long.argtypes = [api['wintypes'].HWND, ctypes.c_int,
                             ctypes.c_ssize_t]
        set_long.restype = ctypes.c_ssize_t
        ex = int(api['get_long'](h, -20))
        if not (ex & 0x00080000):
            set_long(h, -20, ex | 0x00080000)
            # 改扩展样式必须跟一次 SWP_FRAMECHANGED 才生效（只做这一次）
            user32.SetWindowPos(h, None, 0, 0, 0, 0,
                                0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020)
        return True
    except Exception:
        return False


def screen_pixels(points):
    """读取屏幕上若干点的颜色，返回 [(r,g,b) 或 None]；读不到返回 None。

    部分远程/虚拟显示器不给普通程序读屏幕（GetPixel 直接失败），
    这时返回 None，调用方按“无法自检”处理，不影响正常显示。
    """
    if not sys.platform.startswith('win'):
        return None
    try:
        import ctypes
        user32 = ctypes.windll.user32
        gdi32 = ctypes.windll.gdi32
        user32.GetDC.restype = ctypes.c_void_p
        user32.GetDC.argtypes = [ctypes.c_void_p]
        user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        gdi32.GetPixel.restype = ctypes.c_uint32
        gdi32.GetPixel.argtypes = [ctypes.c_void_p, ctypes.c_int,
                                   ctypes.c_int]
        hdc = user32.GetDC(None)
        if not hdc:
            return None
        out = []
        got = False
        try:
            for x, y in points:
                c = int(gdi32.GetPixel(ctypes.c_void_p(hdc), int(x), int(y)))
                if c == 0xFFFFFFFF:      # CLR_INVALID：读不到
                    out.append(None)
                    continue
                got = True
                out.append(((c >> 0) & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF))
        finally:
            user32.ReleaseDC(None, ctypes.c_void_p(hdc))
        return out if got else None
    except Exception:
        return None


def screenshot_pixels(points):
    """截屏后读指定点位的像素颜色，返回 [(r,g,b) 或 None]；截不到返回 None。

    和 ``screen_pixels`` 的区别：
    - ``screen_pixels`` 直接 ``GetDC(None) + GetPixel``，读到的是 GDI 缓存。
      DWM 在扩展屏上偶尔不会把 layered 窗口合成到屏幕，但 Tk 已经把字画
      进了 GDI 缓存——``GetPixel`` 还能读到“画上了”的颜色，可屏幕其实
      看不见。这种情况下 ``screen_pixels`` 会误判成“画面正常”。
    - ``screenshot_pixels`` 先 ``BitBlt`` 把屏幕区域拷到内存 DC，再读像素，
      才是“屏幕上真的有什么颜色”。扩展屏上分层窗口不可见时，它会读到
      桌面/PPT 透出来的颜色而不是前景文字色。

    调用方同时跑两路对比，就能区分“缓存说画上了”和“屏幕真画上了”。
    """
    if not sys.platform.startswith('win'):
        return None
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        gdi32 = ctypes.windll.gdi32

        if not points:
            return None
        xs = [int(p[0]) for p in points]
        ys = [int(p[1]) for p in points]
        rx, ry = min(xs), min(ys)
        rx2, ry2 = max(xs), max(ys)
        rw, rh = rx2 - rx + 1, ry2 - ry + 1
        if rw <= 0 or rh <= 0:
            return None

        user32.GetDC.restype = ctypes.c_void_p
        user32.GetDC.argtypes = [ctypes.c_void_p]
        user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        gdi32.CreateCompatibleDC.restype = ctypes.c_void_p
        gdi32.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
        gdi32.CreateCompatibleBitmap.restype = ctypes.c_void_p
        gdi32.CreateCompatibleBitmap.argtypes = [
            ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
        gdi32.SelectObject.restype = ctypes.c_void_p
        gdi32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        gdi32.GetPixel.restype = ctypes.c_uint32
        gdi32.GetPixel.argtypes = [ctypes.c_void_p, ctypes.c_int,
                                   ctypes.c_int]
        gdi32.BitBlt.restype = ctypes.c_int
        gdi32.BitBlt.argtypes = [
            ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
            ctypes.c_int, ctypes.c_int,
            ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
            ctypes.c_uint]
        gdi32.DeleteDC.argtypes = [ctypes.c_void_p]
        gdi32.DeleteObject.argtypes = [ctypes.c_void_p]

        hdc_screen = user32.GetDC(None)
        if not hdc_screen:
            return None
        hdc_mem = gdi32.CreateCompatibleDC(hdc_screen)
        if not hdc_mem:
            user32.ReleaseDC(None, ctypes.c_void_p(hdc_screen))
            return None
        hbmp = gdi32.CreateCompatibleBitmap(hdc_screen, rw, rh)
        if not hbmp:
            gdi32.DeleteDC(ctypes.c_void_p(hdc_mem))
            user32.ReleaseDC(None, ctypes.c_void_p(hdc_screen))
            return None
        old_obj = gdi32.SelectObject(hdc_mem, hbmp)
        try:
            # SRCCOPY = 0x00CC0020；把屏幕 (rx, ry, rw, rh) 拷到内存 DC 的 (0, 0)
            ok = int(gdi32.BitBlt(
                hdc_mem, 0, 0, rw, rh, hdc_screen, rx, ry, 0x00CC0020))
            if not ok:
                return None
            out = []
            got = False
            for px, py in points:
                lx = int(px) - rx
                ly = int(py) - ry
                c = int(gdi32.GetPixel(hdc_mem, lx, ly))
                if c == 0xFFFFFFFF:
                    out.append(None)
                    continue
                got = True
                out.append(((c >> 0) & 0xFF, (c >> 8) & 0xFF,
                            (c >> 16) & 0xFF))
            return out if got else None
        finally:
            try:
                if old_obj:
                    gdi32.SelectObject(hdc_mem, old_obj)
            except Exception:
                pass
            try:
                gdi32.DeleteObject(ctypes.c_void_p(hbmp))
            except Exception:
                pass
            try:
                gdi32.DeleteDC(ctypes.c_void_p(hdc_mem))
            except Exception:
                pass
            user32.ReleaseDC(None, ctypes.c_void_p(hdc_screen))
    except Exception:
        return None


# 自绘的勾选框图（Tk 的 clam 主题选中的是「X」形记号，不是勾）
_CHECK_IMAGES = []


def _line_pixels(out, x0, y0, x1, y1, width=1):
    """把一条直线上的像素（可加粗）收集到集合里，用于手绘勾。"""
    dx = abs(x1 - x0)
    dy = -abs(y1 - y0)
    sx = 1 if x0 < x1 else -1
    sy = 1 if y0 < y1 else -1
    err = dx + dy
    while True:
        for ox in range(width):
            for oy in range(width):
                out.add((x0 + ox, y0 + oy))
        if x0 == x1 and y0 == y1:
            break
        e2 = 2 * err
        if e2 >= dy:
            err += dy
            x0 += sx
        if e2 <= dx:
            err += dx
            y0 += sy


def _make_check_images(master, p):
    """生成勾选框的四种状态图：白底方框，选中时画蓝色对勾。"""
    size = 13
    gap = 7            # 方框与文字之间留出的透明间距

    def make(selected, disabled):
        border = '#d7dde8' if disabled else '#c3ccdd'
        fill = '#f2f4f8' if disabled else '#ffffff'
        check = '#b9c4d6' if disabled else p['primary']
        # 图片比方框宽一点，右侧留透明区域当作间距（ttk 的布局不支持 padding）
        img = tk.PhotoImage(master=master, width=size + gap, height=size)
        img.put(border, to=(0, 0, size, size))
        img.put(fill, to=(1, 1, size - 1, size - 1))
        if selected:
            pixels = set()
            _line_pixels(pixels, 3, 6, 5, 9, 2)
            _line_pixels(pixels, 5, 9, 10, 3, 2)
            for (x, y) in pixels:
                if 1 <= x < size - 1 and 1 <= y < size - 1:
                    img.put(check, to=(x, y, x + 1, y + 1))
        return img

    imgs = [make(False, False), make(True, False),
            make(False, True), make(True, True)]
    _CHECK_IMAGES.extend(imgs)
    return imgs


def _install_checkbutton(style, master, p):
    """把 ttk 复选框的指示器换成自绘的方框 + 对勾。"""
    try:
        if 'Flat.Checkbutton.indicator' in style.element_names():
            return  # 同一个解释器里已经装过，不重复造图
    except Exception:
        pass
    try:
        unchecked, checked, dis_unchecked, dis_checked = \
            _make_check_images(master, p)
    except Exception:
        return
    try:
        style.element_create('Flat.Checkbutton.indicator', 'image',
                             unchecked,
                             ('selected', checked),
                             ('disabled', dis_unchecked),
                             ('disabled selected', dis_checked))
    except tk.TclError:
        pass  # 同一个解释器里可能已经建过
    try:
        style.layout('TCheckbutton', [
            ('Checkbutton.padding', {'sticky': 'nswe', 'children': [
                ('Flat.Checkbutton.indicator', {'side': 'left', 'sticky': ''}),
                ('Checkbutton.focus', {'sticky': 'nswe', 'children': [
                    ('Checkbutton.label', {'sticky': 'nswe'})]}),
            ]}),
        ])
        style.configure('TCheckbutton', background=p['bg'],
                        foreground=p['text'], focuscolor=p['bg'],
                        padding=(0, 2))
        style.map('TCheckbutton',
                  foreground=[('disabled', p['faint'])])
    except tk.TclError:
        pass


def round_rect(canvas, x1, y1, x2, y2, radius, **kwargs):
    """在画布上画圆角矩形（Tk 没有原生圆角，用平滑多边形模拟）。"""
    r = max(0, min(radius, (x2 - x1) / 2.0, (y2 - y1) / 2.0))
    points = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)


class GradientBand(tk.Canvas):
    """顶部渐变标题栏。"""

    def __init__(self, master, title='', subtitle='', height=60,
                 color_from=PALETTE['primary'],
                 color_to=PALETTE['primary_grad'], **kw):
        tk.Canvas.__init__(self, master, height=height, highlightthickness=0,
                           bd=0, bg=color_from, **kw)
        self._title = title
        self._subtitle = subtitle
        self._c1 = color_from
        self._c2 = color_to
        self._h = height
        self.bind('<Configure>', lambda e: self._draw())
        self._draw()

    def _draw(self):
        try:
            w = max(self.winfo_width(), 1)
            h = self._h
            self.delete('all')
            steps = 48
            for i in range(steps):
                x0 = int(w * i / float(steps))
                x1 = int(w * (i + 1) / float(steps)) + 1
                self.create_rectangle(
                    x0, 0, x1, h, width=0,
                    fill=mix(self._c1, self._c2, i / float(steps - 1)))
            # 右上角装饰圆：极轻微提亮，只做层次，不破坏渐变
            for r, ratio in ((96, 0.05), (72, 0.09), (48, 0.13)):
                self.create_oval(w - r - 8, -r - 14, w + r - 8, r - 14,
                                 fill=mix(self._c2, '#ffffff', ratio),
                                 outline='')
            # 紧凑单行：标题居中，副标题跟在标题右侧同一行
            cy = h / 2.0
            title_font = font(15, 'bold')
            self.create_text(26, cy, text=self._title, anchor='w',
                             font=title_font, fill='#ffffff')
            if self._subtitle:
                try:
                    tx = 26 + tkfont.Font(font=title_font).measure(
                        self._title) + 14
                except tk.TclError:
                    tx = 180
                self.create_text(tx, cy + 1, text=self._subtitle, anchor='w',
                                 font=font(10),
                                 fill=mix(self._c1, '#ffffff', 0.80))
        except tk.TclError:
            pass


class ModernCard(tk.Canvas):
    """现代卡片：圆角色块图标 + 标题 + 说明 + 箭头，鼠标悬停有反馈。"""

    # 字体真实行高缓存：不写死像素，随系统 DPI 缩放自适应
    _LINE_CACHE = {}

    PAD_X = 18
    BADGE = 38
    TEXT_X = 74
    # 横条卡片纵向内边距/间距：配合字体真实行高，保证文字不压框线
    ROW_PAD_TOP = 10
    ROW_PAD_BOTTOM = 10
    ROW_TITLE_GAP = 6
    TILE_PAD = 16
    TILE_BADGE = 42
    TILE_LINE_H = 16

    @classmethod
    def _line_height(cls, size, bold=False):
        """查询某档字体的真实行高（随 DPI 变化），带缓存。"""
        key = (size, bold)
        if key not in cls._LINE_CACHE:
            try:
                cls._LINE_CACHE[key] = tkfont.Font(
                    font=font(size, 'bold' if bold else 'normal')
                ).metrics('linespace')
            except tk.TclError:
                cls._LINE_CACHE[key] = size + 5
        return cls._LINE_CACHE[key]

    @classmethod
    def row_height(cls, desc_lines):
        """按字体真实行高计算横条卡片高度：上内边距 + 标题 + 间距
        + 说明各行 + 下内边距，保证文字与框线之间留白对称、不压线。"""
        return (cls.ROW_PAD_TOP
                + cls._line_height(12, bold=True)
                + cls.ROW_TITLE_GAP
                + max(1, desc_lines) * cls._line_height(9)
                + cls.ROW_PAD_BOTTOM)

    def __init__(self, master, title, description, command=None,
                 icon='', accent=PALETTE['primary'], width=440,
                 layout='row'):
        self._desc = (description or '').split('\n')
        self._layout = layout
        if layout == 'tile':
            height = (self.TILE_PAD + self.TILE_BADGE + 12 + 22
                      + max(1, len(self._desc)) * self.TILE_LINE_H + 14)
        else:
            height = self.row_height(len(self._desc))
        tk.Canvas.__init__(self, master, height=height, width=width,
                           highlightthickness=0, bd=0,
                           bg=master['bg'], cursor='hand2')
        self._title = title
        self._icon = icon or (title[:1] if title else '')
        self._accent = accent
        self._command = command
        self._hover = False
        self._card_w = width
        self.bind('<Configure>', self._on_configure)
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<Button-1>', self._on_click)
        self._draw()

    def _on_configure(self, event):
        if abs(event.width - self._card_w) > 1:
            self._card_w = event.width
            self._draw()

    def _on_enter(self, _event=None):
        self._hover = True
        self._draw()

    def _on_leave(self, _event=None):
        self._hover = False
        self._draw()

    def _on_click(self, _event=None):
        if callable(self._command):
            self._command()

    def set_description(self, text):
        """更新说明文字（行数不变，高度不变）。"""
        lines = (text or '').split('\n')
        if lines == self._desc:
            return
        if len(lines) != len(self._desc):
            return          # 行数变了会撑高卡片，忽略
        self._desc = lines
        self._draw()

    def _draw(self):
        try:
            p = PALETTE
            w = max(self._card_w, 60)
            h = self.winfo_height()
            if h <= 1:
                h = int(self['height'])
            self.delete('all')
            fill = p['hover'] if self._hover else p['surface']
            outline = p['border_hover'] if self._hover else p['border']
            round_rect(self, 1, 1, w - 1, h - 2, 14, fill=fill,
                       outline=outline, width=1)
            if self._layout == 'tile':
                self._draw_tile(w, h, p)
            else:
                self._draw_row(w, h, p)
        except tk.TclError:
            pass

    def _draw_row(self, w, h, p):
        """横条式：左侧色块图标，右侧标题 + 说明，最右侧箭头。"""
        bx = self.PAD_X
        by = (h - self.BADGE) // 2      # 图标在卡片中垂直居中
        round_rect(self, bx, by, bx + self.BADGE, by + self.BADGE, 12,
                   fill=self._accent, outline='')
        self.create_text(bx + self.BADGE / 2.0, by + self.BADGE / 2.0 + 1,
                         text=self._icon, font=font(15, 'bold'),
                         fill='#ffffff')
        title_ls = self._line_height(12, bold=True)
        desc_ls = self._line_height(9)
        top = self.ROW_PAD_TOP
        self.create_text(self.TEXT_X, top, text=self._title,
                         anchor='nw', font=font(12, 'bold'), fill=p['text'])
        y = top + title_ls + self.ROW_TITLE_GAP
        for line in self._desc:
            self.create_text(self.TEXT_X, y, text=line, anchor='nw',
                             font=font(9), fill=p['muted'])
            y += desc_ls
        self.create_text(w - 20, h / 2.0 - 2, text='›', anchor='center',
                         font=font(17, 'bold'),
                         fill=self._accent if self._hover else '#c6cfe0')

    def _draw_tile(self, w, h, p):
        """磁贴式：上方色块图标，下面标题 + 说明，右上角箭头。"""
        pad = self.TILE_PAD
        size = self.TILE_BADGE
        round_rect(self, pad, pad, pad + size, pad + size, 12,
                   fill=self._accent, outline='')
        self.create_text(pad + size / 2.0, pad + size / 2.0 + 1,
                         text=self._icon, font=font(16, 'bold'),
                         fill='#ffffff')
        self.create_text(w - pad, pad - 2, text='›', anchor='ne',
                         font=font(16, 'bold'),
                         fill=self._accent if self._hover else '#c6cfe0')
        title_y = pad + size + 12
        self.create_text(pad, title_y, text=self._title, anchor='nw',
                         font=font(12, 'bold'), fill=p['text'])
        y = title_y + 22
        for line in self._desc:
            self.create_text(pad, y, text=line, anchor='nw', font=font(9),
                             fill=p['muted'])
            y += self.TILE_LINE_H


class FlatButton(tk.Label):
    """扁平按钮：比系统按钮更现代，鼠标悬停有反馈。"""

    def __init__(self, master, text, command=None, kind='primary', **kw):
        p = PALETTE
        self._kind = kind
        self._command = command
        self._enabled = True
        fnt = kw.pop('font', None)
        padx = kw.pop('padx', 18)
        pady = kw.pop('pady', 7)
        cur = kw.pop('cursor', 'hand2')
        if kind == 'primary':
            self._bg, self._fg = p['primary'], '#ffffff'
            self._hover_bg = p['primary_dark']
            weight = 'bold'
        elif kind == 'danger':
            self._bg, self._fg = '#fdeaea', p['danger']
            self._hover_bg = '#fbdada'
            weight = 'normal'
        else:
            self._bg, self._fg = p['surface'], p['primary']
            self._hover_bg = p['primary_soft']
            weight = 'normal'
        tk.Label.__init__(self, master, text=text, bg=self._bg, fg=self._fg,
                          font=fnt if fnt is not None else font(10, weight),
                          padx=padx, pady=pady,
                          cursor=cur, bd=0,
                          highlightthickness=1 if kind != 'primary' else 0,
                          highlightbackground=p['border'], **kw)
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<Button-1>', self._on_enter)
        self.bind('<ButtonRelease-1>', self._on_release)

    def _on_enter(self, _event=None):
        if self._enabled:
            self.configure(bg=self._hover_bg)

    def _on_leave(self, _event=None):
        if self._enabled:
            self.configure(bg=self._bg)

    def _on_release(self, _event=None):
        if not self._enabled:
            return
        self.configure(bg=self._hover_bg)
        if callable(self._command):
            self._command()

    def set_enabled(self, enabled):
        self._enabled = bool(enabled)
        try:
            if enabled:
                self.configure(bg=self._bg, fg=self._fg, cursor='hand2')
            else:
                self.configure(bg=PALETTE['surface_alt'],
                               fg=PALETTE['faint'], cursor='arrow')
        except tk.TclError:
            pass


class SeekBar(tk.Canvas):
    """可拖动的进度条：圆角轨道 + 圆角进度 + 圆形滑块。

    接口和 tk.Scale 基本一致（get / set / config(to=...)），
    另外把「开始拖动 / 结束拖动」做成回调，避免依赖 Tk 的绑定顺序。
    """

    INSET = 6.0      # 轨道两端留白，保证滑块不会被裁掉

    def __init__(self, master, from_=0.0, to=1.0, value=None, command=None,
                 on_seek_start=None, on_seek_end=None, height=22,
                 bar_height=6, trough='#dfe6f2', fill=None, bg=None,
                 width=200, **kw):
        p = PALETTE
        tk.Canvas.__init__(self, master, width=width, height=height,
                           highlightthickness=0, bd=0,
                           bg=bg or p['bg'], cursor='hand2', **kw)
        self._from = float(from_)
        self._to = float(to)
        self._value = self._clamp(
            float(from_ if value is None else value))
        self._command = command
        self._seek_start = on_seek_start
        self._seek_end = on_seek_end
        self._bar_h = bar_height
        self._trough = trough
        self._fill = fill or p['primary']
        self._hover = False
        self._dragging = False
        self._enabled = True
        self.bind('<Configure>', lambda e: self._redraw())
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<Button-1>', self._on_press)
        self.bind('<B1-Motion>', self._on_drag)
        self.bind('<ButtonRelease-1>', self._on_release)
        self._redraw()

    # ---------- 对外接口（尽量和 tk.Scale 一致） ----------
    def get(self):
        return self._value

    def set(self, value):
        try:
            self._value = self._clamp(float(value))
        except (TypeError, ValueError):
            return
        self._redraw()

    def configure(self, cnf=None, **kw):
        for key in ('to', 'from_', 'command'):
            if key in kw:
                raw = kw.pop(key)
                if key == 'to':
                    self._to = float(raw)
                elif key == 'from_':
                    self._from = float(raw)
                else:
                    self._command = raw
        tk.Canvas.configure(self, cnf, **kw)
        self._redraw()

    config = configure

    def set_enabled(self, enabled):
        try:
            self.configure(cursor='hand2' if enabled else 'arrow')
        except tk.TclError:
            pass
        self._enabled = bool(enabled)

    # ---------- 内部计算 ----------
    def _clamp(self, value):
        lo, hi = min(self._from, self._to), max(self._from, self._to)
        return max(lo, min(value, hi))

    def _fraction(self):
        span = self._to - self._from
        if abs(span) < 1e-9:
            return 0.0
        return (self._value - self._from) / span

    def _track(self):
        w = max(self.winfo_width(), 1)
        return self.INSET, max(self.INSET + 8, w - self.INSET)

    def _value_from_x(self, x):
        left, right = self._track()
        if right <= left:
            return self._from
        f = (float(x) - left) / float(right - left)
        f = max(0.0, min(1.0, f))
        return self._from + f * (self._to - self._from)

    def _notify(self):
        if callable(self._command):
            try:
                self._command(self._value)
            except Exception:
                pass

    # ---------- 事件 ----------
    def _on_enter(self, _event=None):
        self._hover = True
        self._redraw()

    def _on_leave(self, _event=None):
        self._hover = False
        self._redraw()

    def _on_press(self, event):
        if not getattr(self, '_enabled', True):
            return
        self._dragging = True
        # 先通知业务方“开始拖动”，再改值，这样拖动过程能实时跳转
        if callable(self._seek_start):
            try:
                self._seek_start()
            except Exception:
                pass
        self._value = self._value_from_x(event.x)
        self._redraw()
        self._notify()

    def _on_drag(self, event):
        if not getattr(self, '_enabled', True):
            return
        if not self._dragging:
            self._dragging = True
        self._value = self._value_from_x(event.x)
        self._redraw()
        self._notify()

    def _on_release(self, event):
        if self._dragging and event is not None:
            self._value = self._value_from_x(event.x)
        self._dragging = False
        self._redraw()
        self._notify()
        if callable(self._seek_end):
            try:
                self._seek_end()
            except Exception:
                pass

    # ---------- 绘制 ----------
    def _redraw(self):
        try:
            w = max(self.winfo_width(), 1)
            h = max(self.winfo_height(), 1)
        except tk.TclError:
            return
        if w <= 1:
            return
        self.delete('all')
        cy = h / 2.0
        r = self._bar_h / 2.0
        left, right = self._track()
        round_rect(self, left, cy - r, right, cy + r, r,
                   fill=self._trough, outline='')
        f = max(0.0, min(1.0, self._fraction()))
        cx = left + (right - left) * f
        if cx - left > 0.6:
            round_rect(self, left, cy - r, cx, cy + r, r,
                       fill=self._fill, outline='')
        # 圆形滑块：悬停/拖动时放大并出现柔和光晕
        tr = 5.0 + (1.5 if (self._hover or self._dragging) else 0.0)
        if self._hover or self._dragging:
            round_rect(self, cx - tr - 4, cy - tr - 4,
                       cx + tr + 4, cy + tr + 4, tr + 4,
                       fill=mix(self._fill, '#ffffff', 0.80), outline='')
        self.create_oval(cx - tr, cy - tr, cx + tr, cy + tr,
                         fill='#ffffff', outline=self._fill, width=2)
