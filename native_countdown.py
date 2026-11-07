# -*- coding: utf-8 -*-
"""native_countdown.py - Win32 原生倒计时浮窗，绕开 Tk 在扩展屏上的 DWM bug。

问题
----
Tk 的 ``-transparentcolor`` 在有些扩展屏 / USB 显示器 / 远程桌面里，DWM 不把
canvas 内容合成到屏幕帧缓冲；窗口位置正确、GDI 缓存里有字，但用户看着就是
一块透明。任何补丁（重新贴色键 / 切 WS_EX_LAYERED / 切实心）都救不回来——
因为合成路径本身在扩展屏上就废了。

本模块解决思路
--------------
完全绕过 Tk 的浮层路径，自己用 Win32 ``CreateWindowExW`` 起一个
``WS_EX_LAYERED | WS_EX_TOPMOST`` 的 ``WS_POPUP`` 窗口；里面 32-bit BGRA
的位图直接通过 ``UpdateLayeredWindow(ULW_COLORKEY)`` 推到 DWM。
背景涂 ``#FF00FF``（色键），文字用 GDI ``DrawTextW`` 直绘，都不依赖 Tk
的 canvas → 屏幕这条链路。

最低依赖：Python 3.x + ctypes（系统自带），不需要 pywin32 / PyQt。

API
---
``ov = NativeCountdownOverlay(x, y, w, h)``
    创建一个屏幕位置 (x, y) 上 w×h 的浮窗。
``ov.set_text('05:00', color=(0xC6, 0x28, 0x28), font_size=64)``
    更新倒计时文字并立即把更新后的位图刷到屏幕。
``ov.move(nx, ny)``
    整窗口拖到新位置。
``ov.show()`` / ``ov.hide()``
    显示 / 隐藏（hide 之后窗口还在，只是不刷位图）。
``ov.close()``
    销毁窗口、释放 GDI 资源。

线程模型：所有方法必须在调 ``CreateWindowExW`` 的同一个线程上调用。
"""
import sys
import ctypes
from ctypes import wintypes


# ============================================================================
# Win32 常量
# ============================================================================

WS_EX_LAYERED       = 0x00080000
WS_EX_TOPMOST       = 0x00000008
WS_EX_TOOLWINDOW    = 0x00000080
WS_EX_TRANSPARENT   = 0x00000020   # 鼠标穿透（点到窗口直接透到下一层）
WS_POPUP            = 0x80000000
WS_VISIBLE          = 0x10000000

ULW_COLORKEY        = 0x00000001
ULW_ALPHA           = 0x00000002
ULW_OPAQUE          = 0x00000004

DT_CENTER           = 0x00000001
DT_VCENTER          = 0x00000004
DT_SINGLELINE       = 0x00000020
DT_NOPREFIX         = 0x00000800

TRANSPARENT_MODE    = 1   # GDI SetBkMode

DIB_RGB_COLORS      = 0
BI_RGB              = 0

# SRCCOPY = 0x00CC0020

SW_SHOWNOACTIVATE   = 4
SW_HIDE             = 0
SWP_NOMOVE          = 0x0002
SWP_NOSIZE          = 0x0001
SWP_NOACTIVATE      = 0x0010
SWP_NOZORDER        = 0x0004

HWND_TOPMOST        = -1

# 色键（背景涂这个颜色 → 通过 ULW_COLORKEY 变成透明）。
# 必须是不大可能出现在文字 / 桌面里的"魔法色"。
COLOR_KEY_BGR       = 0xFF00FF  # BGRA 中的 0xFF0000FF 在 COLORREF 里是 0xFF00FF

# GDI 字体的 quality 常量
ANTIALIASED_QUALITY   = 4
CLEARTYPE_QUALITY     = 5
DEFAULT_QUALITY       = 0

# 字符集：GB2312_CHARSET 等价于中文字符 + 默认西文；用 DEFAULT_CHARSET 让 GDI 选
DEFAULT_CHARSET     = 1


# ============================================================================
# ctypes 结构
# ============================================================================

class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ('biSize', ctypes.c_uint32),
        ('biWidth', ctypes.c_int32),
        ('biHeight', ctypes.c_int32),
        ('biPlanes', ctypes.c_uint16),
        ('biBitCount', ctypes.c_uint16),
        ('biCompression', ctypes.c_uint32),
        ('biSizeImage', ctypes.c_uint32),
        ('biXPelsPerMeter', ctypes.c_int32),
        ('biYPelsPerMeter', ctypes.c_int32),
        ('biClrUsed', ctypes.c_uint32),
        ('biClrImportant', ctypes.c_uint32),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [
        ('bmiHeader', BITMAPINFOHEADER),
        ('bmiColors', ctypes.c_uint32 * 3),
    ]


class BLENDFUNCTION(ctypes.Structure):
    _fields_ = [
        ('BlendOp', ctypes.c_byte),
        ('BlendFlags', ctypes.c_byte),
        ('SourceConstantAlpha', ctypes.c_byte),
        ('AlphaFormat', ctypes.c_byte),
    ]


class POINT(ctypes.Structure):
    _fields_ = [
        ('x', ctypes.c_long),
        ('y', ctypes.c_long),
    ]


class SIZE(ctypes.Structure):
    _fields_ = [
        ('cx', ctypes.c_long),
        ('cy', ctypes.c_long),
    ]


class RECT(ctypes.Structure):
    _fields_ = [
        ('left', ctypes.c_long),
        ('top', ctypes.c_long),
        ('right', ctypes.c_long),
        ('bottom', ctypes.c_long),
    ]


class WNDCLASSEXW(ctypes.Structure):
    _fields_ = [
        ('cbSize', ctypes.c_uint),
        ('style', ctypes.c_uint),
        ('lpfnWndProc', ctypes.c_void_p),
        ('cbClsExtra', ctypes.c_int),
        ('cbWndExtra', ctypes.c_int),
        ('hInstance', ctypes.c_void_p),
        ('hIcon', ctypes.c_void_p),
        ('hCursor', ctypes.c_void_p),
        ('hbrBackground', ctypes.c_void_p),
        ('lpszMenuName', ctypes.c_wchar_p),
        ('lpszClassName', ctypes.c_wchar_p),
        ('hIconSm', ctypes.c_void_p),
    ]


# ============================================================================
# Win32 函数 setup
# ============================================================================

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
kernel32 = ctypes.windll.kernel32

# GetModuleHandleW 是 kernel32 的，不在 user32
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
kernel32.GetModuleHandleW.restype = wintypes.HMODULE

user32.RegisterClassExW.argtypes = [ctypes.POINTER(WNDCLASSEXW)]
user32.RegisterClassExW.restype = ctypes.c_uint16

user32.CreateWindowExW.argtypes = [
    ctypes.c_uint, wintypes.LPCWSTR, wintypes.LPCWSTR,
    ctypes.c_uint, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, ctypes.c_void_p,
]
user32.CreateWindowExW.restype = wintypes.HWND

user32.DestroyWindow.argtypes = [wintypes.HWND]
user32.DestroyWindow.restype = wintypes.BOOL

user32.SetWindowPos.argtypes = [
    wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
    ctypes.c_int, ctypes.c_int, ctypes.c_uint,
]
user32.SetWindowPos.restype = wintypes.BOOL

user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
user32.ShowWindow.restype = wintypes.BOOL

user32.UpdateLayeredWindow.argtypes = [
    wintypes.HWND,
    wintypes.HDC,
    ctypes.POINTER(POINT),
    ctypes.POINTER(SIZE),
    wintypes.HDC,
    ctypes.POINTER(POINT),
    ctypes.c_uint,
    ctypes.POINTER(BLENDFUNCTION),
    ctypes.c_uint,
]
user32.UpdateLayeredWindow.restype = wintypes.BOOL

user32.DefWindowProcW.argtypes = [wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM]
user32.DefWindowProcW.restype = ctypes.c_long

# SetWindowLongPtrW / GetWindowLongPtrW —— 动态切 WS_EX_TRANSPARENT 走它。
user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int,
                                    ctypes.c_ssize_t]
user32.SetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
GWL_EXSTYLE = -20
SWP_FRAMECHANGED = 0x0020
SWP_NOACTIVATE = 0x0010
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOZORDER = 0x0004

# Drag 相关的 Win32 常量
HTCLIENT = 1
HTCAPTION = 2
WM_NCHITTEST = 0x0084
WM_MOVE = 0x0003

user32.GetDC.argtypes = [wintypes.HWND]
user32.GetDC.restype = wintypes.HDC
user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
user32.ReleaseDC.restype = ctypes.c_int

user32.LoadCursorW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR]
user32.LoadCursorW.restype = wintypes.HANDLE

# ---- DPI 查询（扩展屏字号缩放要靠它）----
# shcore.dll 是 Win8+ 系统库；Windows 7 没有，正常返回 96 fallback。
MDT_EFFECTIVE_DPI = 0
try:
    _shcore = ctypes.windll.shcore
    _shcore.GetDpiForMonitor.argtypes = [
        wintypes.HMONITOR, ctypes.c_uint,
        ctypes.POINTER(ctypes.c_uint), ctypes.POINTER(ctypes.c_uint),
    ]
    _shcore.GetDpiForMonitor.restype = ctypes.c_long
    _HAS_SHCORE = True
except Exception:
    _shcore = None
    _HAS_SHCORE = False

user32.MonitorFromPoint.argtypes = [POINT, ctypes.c_uint]
user32.MonitorFromPoint.restype = wintypes.HMONITOR

# MONITOR_DEFAULTTONEAREST = 0x00000002
_MONITOR_DEFAULTTONEAREST = 0x00000002


def dpi_for_point(x, y):
    """屏幕坐标 (x, y) 所在显示器的 effective DPI（一般就是 96/120/144/192）。

    拿不到（Win7 或外部异常）就返回 96（=100% 缩放）。
    """
    if not _HAS_SHCORE:
        return 96
    try:
        hmon = user32.MonitorFromPoint(POINT(int(x), int(y)),
                                       _MONITOR_DEFAULTTONEAREST)
        if not hmon:
            return 96
        dpi_x = ctypes.c_uint(96)
        dpi_y = ctypes.c_uint(96)
        hr = _shcore.GetDpiForMonitor(hmon, MDT_EFFECTIVE_DPI,
                                      ctypes.byref(dpi_x), ctypes.byref(dpi_y))
        if hr != 0:  # S_OK
            return 96
        return int(dpi_x.value) or 96
    except Exception:
        return 96

gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
gdi32.CreateCompatibleDC.restype = wintypes.HDC

gdi32.CreateDIBSection.argtypes = [
    wintypes.HDC, ctypes.POINTER(BITMAPINFO), ctypes.c_uint,
    ctypes.POINTER(ctypes.c_void_p), wintypes.HANDLE, ctypes.c_uint,
]
gdi32.CreateDIBSection.restype = wintypes.HBITMAP

gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
gdi32.SelectObject.restype = wintypes.HGDIOBJ

gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
gdi32.DeleteDC.argtypes = [wintypes.HDC]

gdi32.CreateFontW.argtypes = [
    ctypes.c_int,    # nHeight
    ctypes.c_int,    # nWidth
    ctypes.c_int,    # nEscapement
    ctypes.c_int,    # nOrientation
    ctypes.c_int,    # fnWeight
    ctypes.c_uint,   # fdwItalic
    ctypes.c_uint,   # fdwUnderline
    ctypes.c_uint,   # fdwStrikeOut
    ctypes.c_uint,   # fdwCharSet
    ctypes.c_uint,   # fdwOutputPrecision
    ctypes.c_uint,   # fdwClipPrecision
    ctypes.c_uint,   # fdwQuality
    ctypes.c_uint,   # fdwPitchAndFamily
    wintypes.LPCWSTR,# lpszFace
]
gdi32.CreateFontW.restype = wintypes.HFONT

gdi32.SetTextColor.argtypes = [wintypes.HDC, ctypes.c_uint32]
gdi32.SetBkColor.argtypes = [wintypes.HDC, ctypes.c_uint32]
gdi32.SetBkMode.argtypes = [wintypes.HDC, ctypes.c_int]

# DrawTextW 在 user32 里，不在 gdi32
user32.DrawTextW.argtypes = [
    wintypes.HDC, wintypes.LPCWSTR, ctypes.c_int, ctypes.POINTER(RECT), ctypes.c_uint,
]
user32.DrawTextW.restype = ctypes.c_int

gdi32.CreateSolidBrush.argtypes = [ctypes.c_uint32]
gdi32.CreateSolidBrush.restype = wintypes.HBRUSH

# FillRect 在 user32 里
user32.FillRect.argtypes = [wintypes.HDC, ctypes.POINTER(RECT), wintypes.HBRUSH]
user32.FillRect.restype = ctypes.c_int


# ============================================================================
# Window class 注册（每个进程只一次）
# ============================================================================

_CLASS_NAME = 'WorkBuddy_NativeCountdown_v1'
_CLASS_REGISTERED = False
# 必须常驻，否则 wndproc 函数对象被 GC，窗口收到消息就 crash
_WNDPROC_REF = None


def _register_class(hinstance):
    """注册窗口类。所有 NativeCountdownOverlay 共用一个 window class。

    wnd_proc 现在还支持简单 drag：传 ``HWND_DRAG_ENABLED`` 走 HTCAPTION 路径，
    Windows 系统自己处理 WM_NCLBUTTONDOWN 拖动窗口，拖完回调
    ``_WINPROC_DRAG_REFS[hwnd]``（如果有）。
    """
    global _CLASS_REGISTERED, _WNDPROC_REF, _WINPROC_DRAG_REFS, _WINPROC_MOVE_REFS
    if _CLASS_REGISTERED:
        return True

    def wnd_proc(hwnd, msg, wp, lp):
        try:
            if msg == WM_NCHITTEST:
                # 仅当窗口被显式开启 drag 模式时，把命中测试当作标题栏，
                # 这样按住 native 窗口就能拖动。Windows 会自己发起
                # WM_NCLBUTTONDOWN + WM_SYSCOMMAND + 实时 WM_MOVE，
                # 调用方拿到最终的 (x, y) 即可同步 state。
                if _WINPROC_DRAG_REFS.get(int(hwnd)):
                    return HTCAPTION
                return HTCLIENT
            if msg == WM_MOVE:
                pos = wintypes.PARAM(wp).value & 0xFFFF
                # 注：WM_MOVE 的 x/y 在 lp 低 16 位和高 16 位
                lpx = wintypes.PARAM(lp).value & 0xFFFF
                lpy = (wintypes.PARAM(lp).value >> 16) & 0xFFFF
                if lpx > 0x7FFF:
                    lpx -= 0x10000
                if lpy > 0x7FFF:
                    lpy -= 0x10000
                cb = _WINPROC_MOVE_REFS.get(int(hwnd))
                if cb:
                    try:
                        cb(int(hwnd), int(lpx), int(lpy))
                    except Exception:
                        pass
        except Exception:
            pass
        return int(user32.DefWindowProcW(hwnd, msg, wp, lp))

    WNDPROC = ctypes.WINFUNCTYPE(
        ctypes.c_long, wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM)
    # 拿到回调指针，并把回调本身挂到全局，防止被 GC
    _WNDPROC_REF = WNDPROC(wnd_proc)

    wc = WNDCLASSEXW()
    wc.cbSize = ctypes.sizeof(WNDCLASSEXW)
    wc.style = 0
    wc.lpfnWndProc = ctypes.cast(_WNDPROC_REF, ctypes.c_void_p).value
    wc.cbClsExtra = 0
    wc.cbWndExtra = 0
    wc.hInstance = hinstance
    wc.hIcon = None
    # 不显式设光标（NULL），让窗口跟随系统当前光标
    wc.hCursor = None
    wc.hbrBackground = None
    wc.lpszMenuName = None
    wc.lpszClassName = _CLASS_NAME
    wc.hIconSm = None

    atom = user32.RegisterClassExW(ctypes.byref(wc))
    if not atom:
        return False
    _CLASS_REGISTERED = True
    return True


# drag 模式标志：{hwnd(int): bool}
_WINPROC_DRAG_REFS = {}
# 拖动结束时回调：{hwnd(int): callable(hwnd, x, y)}
_WINPROC_MOVE_REFS = {}


# ============================================================================
# 主类
# ============================================================================

class NativeCountdownOverlay:
    """Win32 原生倒计时浮窗。

    用法::

        ov = NativeCountdownOverlay(5480, 20, 260, 100)
        ov.set_text('05:00', color=(0xC6, 0x28, 0x28), font_size=64)
        ov.show()
        # ... update often:
        ov.set_text('04:59')
        ov.move(new_x, new_y)
        # ... when done:
        ov.close()
    """

    # 字体候选：Microsoft YaHei UI 优先（中英文都好看）
    _FONT_CANDIDATES = (
        'Microsoft YaHei UI',
        '微软雅黑',
        'Microsoft YaHei',
        'PingFang SC',
        'SimHei',
        'Segoe UI',
        'Arial',
    )

    # 字号缩放：DPI 之外不再额外放大。调用方应该按物理尺寸传窗口，
    # 我们只处理 DPI 缩放（per-monitor DPI awareness 差异）。
    # 早先版本里这个值是 1.6，给扩展屏多放大字号；但那样会让字
    # 在物理尺寸固定的情况下溢出，更好的做法是让调用方传更大的窗口。
    FONT_BOOST = 1.0

    def __init__(self, x, y, width, height, hinstance=None):
        if not sys.platform.startswith('win'):
            raise RuntimeError('NativeCountdownOverlay 仅在 Windows 下可用')

        self._x = int(x)
        self._y = int(y)
        self._w = max(int(width), 16)
        self._h = max(int(height), 16)

        self._hwnd = None
        self._hdc_screen = None
        self._hdc_mem = None
        self._hbmp = None
        self._hbmp_old = None
        self._hfont = None
        self._hfont_old = None
        self._hbrush_key = None
        # 背景填充 brush：None → 用色键（透明），有效 → 切实心
        # 注意：色键透明由 UpdateLayeredWindow(ULW_COLORKEY) 处理，不需要
        # 任何 brush；这里存的是"切实心模式下的填充 brush"。
        self._hbrush_fill = None
        self._fill_mode = 'key'  # 'key' = 色键透明；'solid' = 用 _hbrush_fill
        self._fill_color = None  # 最近一次 set_fill_color(rgb) 入参，调试用
        self._visible = False
        # 当前展示状态
        self._text = ''
        self._color = (0xC6, 0x28, 0x28)  # 默认红
        # ---- 描边状态：默认关闭 ----
        # 跟主屏 Tk 浮窗的 create_text 8 偏移画法保持一致，让"主屏
        # 描边、扩展屏描边"看上去是同一笔字。描边色未指定时 _render
        # 会自动用文字色反差色（与主屏逻辑对齐）。
        self._stroke_w = 0
        self._stroke_color = None
        # 字号会按所在显示器的 DPI 缩放（扩展屏 125% → 字号 ×1.25），
        # 否则字到了扩展屏会明显变小。这里缓存一份逻辑字号。
        self._font_size = max(int(min(self._h, 200) * 0.55), 12)
        # 该显示器 effective DPI；先按创建位置算一次
        self._dpi = dpi_for_point(self._x, self._y)

        # 注册 window class
        if hinstance is None:
            hinstance = kernel32.GetModuleHandleW(None)
        if not _register_class(hinstance):
            raise RuntimeError('注册 window class 失败')

        # 创建窗口：先 WS_POPUP（默认隐藏），配完位图再 ShowWindow
        # 注意：WS_EX_TRANSPARENT 让窗口"鼠标穿透"——所有点击事件透到下层 Tk
        # 窗口，所以用户在浮窗上按住还能拖动浮窗本身（drag handler 装在 Tk 上）
        ex_style = (WS_EX_LAYERED | WS_EX_TOPMOST | WS_EX_TOOLWINDOW
                    | WS_EX_TRANSPARENT)
        hwnd = user32.CreateWindowExW(
            ex_style,
            _CLASS_NAME,
            'WorkBuddyNativeCountdown',
            WS_POPUP,
            self._x, self._y, self._w, self._h,
            None, None, hinstance, None,
        )
        if not hwnd:
            err = ctypes.GetLastError() if hasattr(ctypes, 'GetLastError') else 0
            raise RuntimeError('CreateWindowExW 失败 (GetLastError=%d)' % err)
        self._hwnd = hwnd

        self._init_gdi()

    # --------------------------------------------------------------------- GDI

    def _init_gdi(self):
        """建立内存 DC + 32-bit DIB + 字体 + 色键画刷。"""
        # 屏幕 DC（用于 UpdateLayeredWindow 的目标）
        self._hdc_screen = user32.GetDC(None)
        if not self._hdc_screen:
            raise RuntimeError('GetDC(None) 失败')

        self._hdc_mem = gdi32.CreateCompatibleDC(self._hdc_screen)
        if not self._hdc_mem:
            raise RuntimeError('CreateCompatibleDC 失败')

        # 32-bit BGRA DIB，顶向下
        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = self._w
        bmi.bmiHeader.biHeight = -self._h
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = BI_RGB
        bmi.bmiHeader.biSizeImage = self._w * self._h * 4

        bits = ctypes.c_void_p()
        self._hbmp = gdi32.CreateDIBSection(
            self._hdc_mem, ctypes.byref(bmi), DIB_RGB_COLORS,
            ctypes.byref(bits), None, 0,
        )
        if not self._hbmp:
            raise RuntimeError('CreateDIBSection 失败')

        self._hbmp_old = gdi32.SelectObject(self._hdc_mem, self._hbmp)

        # 字体：按该显示器 DPI 缩放。CreateFontW 的 height 是逻辑像素，
        # 但我们的 DIB 是物理像素；要让物理像素下的字看起来跟主屏一样大，
        # 就把字号乘 (dpi / 96)。FONT_BOOST 现在固定 1.0（窗口尺寸
        # 由调用方按需放大，字号与窗口比例不变）。
        try:
            scaled_size = max(8, int(round(self._font_size * self._dpi
                                           / 96.0 * self.FONT_BOOST)))
        except Exception:
            scaled_size = self._font_size
        self._hfont = self._make_font(scaled_size)
        if self._hfont:
            self._hfont_old = gdi32.SelectObject(self._hdc_mem, self._hfont)

        # 色键画刷（纯 #FF00FF，让那块像素透明）
        self._hbrush_key = gdi32.CreateSolidBrush(COLOR_KEY_BGR)

        # 先涂一次色键，否则第一次 UpdateLayeredWindow 时未初始化像素是黑色
        self._fill_key()

    def _fill_key(self):
        """按 self._fill_mode 涂背景：key → 色键；solid → 切实心填充色。

        名义叫 _fill_key 是历史原因（最初只有 key），逻辑上现在是"涂底"。
        """
        rc = RECT(0, 0, self._w, self._h)
        if self._fill_mode == 'key' or self._hbrush_fill is None:
            brush = self._hbrush_key
        else:
            brush = self._hbrush_fill
        user32.FillRect(self._hdc_mem, ctypes.byref(rc), brush)

    def _make_font(self, size):
        """按候选顺序试一张中英兼顾的字体；失败就拿系统默认。"""
        size = max(8, int(size))
        for name in self._FONT_CANDIDATES:
            try:
                h = gdi32.CreateFontW(
                    int(size), 0, 0, 0,           # height, width, esc, orient
                    700,                          # weight (FW_BOLD = 700)
                    0, 0, 0,                      # italic, underline, strikeout
                    DEFAULT_CHARSET,
                    0, 1,                          # OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS
                    CLEARTYPE_QUALITY,
                    0,                             # DEFAULT_PITCH
                    name,
                )
                if h:
                    return h
            except Exception:
                continue
        return None

    def _render(self):
        """把当前 _text / _color 画到内存 DIB 上（背景色键）。

        描边实现跟主屏 Tk create_text 8 偏移保持一致：
          1. 先用描边色在 8 个方向各画一遍（不透明背景仍显示），
             相当于"在文字周围铺一圈描边色"；
          2. 再用前景色画一次中心文字，盖住描边的中心重叠部分，
             只留 4 角 / 4 边的外延 → 视觉上就是 1px~Npx 的描边。
        这样能保证主屏、扩展屏两边描边粗细看上去一样。
        """
        self._fill_key()
        if not self._text:
            return
        # 设置 GDI 文字属性
        if self._hfont:
            gdi32.SelectObject(self._hdc_mem, self._hfont)

        # 整窗口矩形
        rc = RECT(0, 0, self._w, self._h)
        flags = DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX

        # ---- 描边：8 方向偏移画描边色（每张只画一次，靠 Win32 自己
        # 处理重叠）----
        sw = int(getattr(self, '_stroke_w', 0) or 0)
        sc = getattr(self, '_stroke_color', None)
        if sw > 0 and sc is not None:
            try:
                sr, sg, sb = int(sc[0]), int(sc[1]), int(sc[2])
                stroke_c = (int(sb) << 16) | (int(sg) << 8) | int(sr)
                gdi32.SetTextColor(self._hdc_mem,
                                   ctypes.c_uint32(stroke_c))
                gdi32.SetBkMode(self._hdc_mem, TRANSPARENT_MODE)
                off = max(1, int(round(sw)))
                for dx, dy in ((off, 0), (-off, 0), (0, off), (0, -off),
                               (off, off), (off, -off),
                               (-off, off), (-off, -off)):
                    rc_s = RECT(dx, dy, self._w + dx, self._h + dy)
                    user32.DrawTextW(
                        self._hdc_mem, ctypes.c_wchar_p(self._text), -1,
                        ctypes.byref(rc_s), flags,
                    )
            except Exception:
                # 描边失败不能影响前景文字绘制
                pass

        # ---- 前景文字（覆盖描边的中心重叠）----
        r, g, b = self._color[:3]
        # COLORREF 是 0x00BBGGRR
        text_color = (int(b) << 16) | (int(g) << 8) | int(r)
        gdi32.SetTextColor(self._hdc_mem, ctypes.c_uint32(text_color))
        gdi32.SetBkMode(self._hdc_mem, TRANSPARENT_MODE)
        user32.DrawTextW(
            self._hdc_mem, ctypes.c_wchar_p(self._text), -1,
            ctypes.byref(rc), flags,
        )

    def _push(self):
        """把当前 DIB 通过 UpdateLayeredWindow 推到屏幕上。"""
        if not self._hwnd or not self._hdc_mem or not self._hdc_screen:
            return False
        try:
            p_dst = POINT(self._x, self._y)
            sz = SIZE(self._w, self._h)
            p_src = POINT(0, 0)
            blend = BLENDFUNCTION(0, 0, 255, 0)   # AC_SRC_OVER, alpha=255
            ok = user32.UpdateLayeredWindow(
                self._hwnd,
                self._hdc_screen,
                ctypes.byref(p_dst),
                ctypes.byref(sz),
                self._hdc_mem,
                ctypes.byref(p_src),
                COLOR_KEY_BGR,
                ctypes.byref(blend),
                ULW_COLORKEY,
            )
            return bool(ok)
        except Exception as e:
            return False

    # ----------------------------------------------------------------- public

    def set_text(self, text, color=None, font_size=None,
                 stroke_w=0, stroke_color=None, refresh=True):
        """设置倒计时文字（立刻贴一次屏幕）。

        参数
        ----
        text : str
            要显示的文本。
        color : (r, g, b) | None
            文字颜色，``None`` 保持上次颜色不变。
        font_size : int | None
            逻辑字号；GDI 实际字号 = font_size × (该屏 DPI / 96)。
        stroke_w : int
            文字描边宽度（像素）。``0`` = 不描边（默认）。
        stroke_color : (r, g, b) | None
            描边颜色。``None`` 时按调用方外部决定；本方法只在
            ``stroke_w > 0`` 且 ``stroke_color`` 是合法 ``(r,g,b)`` 元组
            时才画描边。
        """
        self._text = '' if text is None else str(text)
        if color is not None:
            try:
                self._color = (int(color[0]), int(color[1]), int(color[2]))
            except Exception:
                pass
        # ---- 描边参数：缓存到实例字段，_render 时读 ----
        try:
            self._stroke_w = max(0, int(stroke_w or 0))
        except Exception:
            self._stroke_w = 0
        if stroke_color is not None:
            try:
                self._stroke_color = (int(stroke_color[0]),
                                      int(stroke_color[1]),
                                      int(stroke_color[2]))
            except Exception:
                # 非法 → 不动
                pass
        if font_size is not None and int(font_size) != self._font_size:
            self._font_size = max(8, int(font_size))
            # 重建字体（DPI 缩放后的物理字号）
            if self._hfont_old:
                try:
                    gdi32.SelectObject(self._hdc_mem, self._hfont_old)
                except Exception:
                    pass
            if self._hfont:
                try:
                    gdi32.DeleteObject(self._hfont)
                except Exception:
                    pass
            try:
                scaled = max(8, int(round(self._font_size * self._dpi
                                          / 96.0 * self.FONT_BOOST)))
            except Exception:
                scaled = self._font_size
            self._hfont = self._make_font(scaled)
            if self._hfont:
                self._hfont_old = gdi32.SelectObject(self._hdc_mem, self._hfont)
        self._render()
        if refresh and self._visible:
            self._push()

    def move(self, x, y, refresh=True):
        """移动窗口到屏幕坐标 (x, y)。跨屏时自动按新显示器 DPI 重算字号。"""
        self._x = int(x)
        self._y = int(y)
        # 跨屏 DPI 变了 → 重算字体大小，保证在每块屏上都跟逻辑字号一致
        try:
            new_dpi = dpi_for_point(self._x, self._y)
        except Exception:
            new_dpi = self._dpi
        if new_dpi != self._dpi:
            self._dpi = new_dpi
            if self._hfont_old:
                try:
                    gdi32.SelectObject(self._hdc_mem, self._hfont_old)
                except Exception:
                    pass
            if self._hfont:
                try:
                    gdi32.DeleteObject(self._hfont)
                except Exception:
                    pass
            try:
                scaled = max(8, int(round(self._font_size * self._dpi
                                          / 96.0 * self.FONT_BOOST)))
            except Exception:
                scaled = self._font_size
            self._hfont = self._make_font(scaled)
            if self._hfont:
                self._hfont_old = gdi32.SelectObject(self._hdc_mem, self._hfont)
            # 重画当前文字（字号变了，layout 也得变）
            self._render()
        if self._hwnd:
            try:
                user32.SetWindowPos(
                    self._hwnd, HWND_TOPMOST,
                    self._x, self._y, self._w, self._h,
                    SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOZORDER,
                )
            except Exception:
                pass
        if refresh and self._visible:
            self._push()

    def show(self):
        """显示窗口（第一次显示前必须 set_text 至少一次）。"""
        if not self._hwnd:
            return False
        self._render()
        ok = self._push()
        try:
            user32.ShowWindow(self._hwnd, SW_SHOWNOACTIVATE)
        except Exception:
            pass
        self._visible = True
        return bool(ok)

    def hide(self):
        if self._hwnd:
            try:
                user32.ShowWindow(self._hwnd, SW_HIDE)
            except Exception:
                pass
        self._visible = False

    def refresh(self):
        """强制用当前 DIB 重贴一次（与 set_text 不同：不重画，只重发）。"""
        if self._visible:
            self._push()

    def set_fill_color(self, rgb):
        """切换 native overlay 是否"切实心"。

        透明模式（默认）：``rgb=None`` → 背景继续涂色键，``UpdateLayeredWindow
        (ULW_COLORKEY)`` 自然把它变透明（看穿到桌面 / PPT）。

        切实心模式：``rgb=(r, g, b)`` → 创建一只纯色 brush 涂满背景，文字
        居中画在中间，整张图完全不透明（用户在属性面板里把"背景透明"关掉
        时，跟随主屏浮窗一起切到切实心，保证扩展屏和主屏视觉一致）。

        切换后会立刻重画并刷新到屏幕（如果当前可见）。
        """
        # 清理旧 fill brush（如果有）
        if self._hbrush_fill:
            try:
                gdi32.DeleteObject(self._hbrush_fill)
            except Exception:
                pass
            self._hbrush_fill = None
        if rgb is None:
            # 切回色键透明模式
            self._fill_mode = 'key'
            self._fill_color = None
        else:
            try:
                r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
            except Exception:
                self._fill_mode = 'key'
                self._fill_color = None
            else:
                # COLORREF = 0x00BBGGRR
                cref = (int(b) << 16) | (int(g) << 8) | int(r)
                self._hbrush_fill = gdi32.CreateSolidBrush(cref)
                self._fill_mode = 'solid'
                self._fill_color = (r, g, b)
        # 重新画一次并推送到屏幕（如果可见）
        try:
            if self._hdc_mem:
                self._render()
            if self._visible:
                self._push()
        except Exception:
            pass

    def set_draggable(self, draggable):
        """开 / 关 native 窗口的拖动能力。

        打开后：本窗口去掉 ``WS_EX_TRANSPARENT``，鼠标事件进来，
        整个窗口都按 HTCAPTION 处理（按住就能拖）。
        关闭后：恢复 ``WS_EX_TRANSPARENT``，鼠标事件穿透回桌面。
        """
        try:
            hwnd = int(self._hwnd) if self._hwnd else 0
        except Exception:
            hwnd = 0
        if not hwnd:
            return
        cur = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        TRANSPARENT = WS_EX_TRANSPARENT
        new = (int(cur) & ~TRANSPARENT) if draggable else (
            int(cur) | TRANSPARENT)
        if new != cur:
            user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new)
            user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER
                                | SWP_NOACTIVATE | SWP_FRAMECHANGED)
        # 同步 drag flag 表
        _WINPROC_DRAG_REFS[hwnd] = bool(draggable)

    def set_drag_callback(self, callback):
        """窗口 WM_MOVE（拖动期间）回调 ``callback(hwnd, x, y)``。

        设为 None 取消。一般接 ``cb(self._hwnd, x, y)`` 把当前 (x, y)
        同步到 state，然后下一次 _move_pos / _native_pos_for 会自然用
        上。
        """
        try:
            hwnd = int(self._hwnd) if self._hwnd else 0
        except Exception:
            hwnd = 0
        if not hwnd:
            return
        if callback is None:
            _WINPROC_MOVE_REFS.pop(hwnd, None)
        else:
            _WINPROC_MOVE_REFS[hwnd] = callback

    def close(self):
        """销毁窗口、释放所有 GDI 资源。多次调用安全。"""
        # 先把 hwnd 从全局 drag 表里摘掉，避免回调打到已销毁的窗口
        try:
            h = int(self._hwnd) if self._hwnd else 0
        except Exception:
            h = 0
        if h:
            _WINPROC_DRAG_REFS.pop(h, None)
            _WINPROC_MOVE_REFS.pop(h, None)
        try:
            if self._hwnd:
                user32.DestroyWindow(self._hwnd)
        except Exception:
            pass
        self._hwnd = None

        if self._hdc_mem:
            try:
                if self._hbmp_old:
                    gdi32.SelectObject(self._hdc_mem, self._hbmp_old)
                if self._hfont_old:
                    gdi32.SelectObject(self._hdc_mem, self._hfont_old)
            except Exception:
                pass

        for h, fn in (
            (self._hbmp, gdi32.DeleteObject),
            (self._hfont, gdi32.DeleteObject),
            (self._hbrush_key, gdi32.DeleteObject),
            (self._hbrush_fill, gdi32.DeleteObject),
        ):
            if h:
                try:
                    fn(h)
                except Exception:
                    pass
        self._hbrush_fill = None
        self._fill_mode = 'key'
        self._fill_color = None
        if self._hdc_mem:
            try:
                gdi32.DeleteDC(self._hdc_mem)
            except Exception:
                pass
        if self._hdc_screen:
            try:
                user32.ReleaseDC(None, self._hdc_screen)
            except Exception:
                pass

        self._hdc_mem = None
        self._hdc_screen = None
        self._hbmp = None
        self._hbmp_old = None
        self._hfont = None
        self._hfont_old = None
        self._hbrush_key = None

    # ---------- properties ----------
    @property
    def hwnd(self):
        return self._hwnd

    @property
    def position(self):
        return (self._x, self._y)

    @property
    def size(self):
        return (self._w, self._h)
