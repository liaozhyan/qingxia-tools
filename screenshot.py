# -*- coding: utf-8 -*-
"""会务截图模块：仿 ShareX 设计，Python + Tkinter 实现。

四种截图模式：
    全屏 / 窗口 / 区域 / 滚动拼接

简易标注：矩形 / 箭头 / 文字 / 马赛克，撤销/重做、保存 / 复制。
自动归档：默认存到 ``会务收件/截图/YYYY-MM-DD/``，按时间戳命名。

设计取舍：
    * 不依赖 mss / pynput / keyboard 等第三方库——
      截屏用 Pillow 自带的 ``ImageGrab``（项目本就用 Pillow），
      全局热键走 Win32 ``RegisterHotKey`` + 一个 message-only 窗口。
      这样跟项目其它模块保持一致（``ui_theme.py`` 也都是 ctypes），
      不会因为缺一个 pip 包就整块崩。
    * 模块对外只暴露一个入口函数 ``open_panel``，被 ``会务助手.py``
      调一行就能挂出子窗口，关闭由主程序负责。
"""

import os
import sys
import time
import ctypes
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ctypes import wintypes
from ctypes import c_void_p

try:
    from PIL import Image, ImageTk, ImageGrab, ImageDraw, ImageFont
except ImportError:                              # Pillow 一定在，但兜个底
    Image = ImageTk = ImageGrab = ImageDraw = ImageFont = None

import ui_theme
import video_player                              # 拿 list_monitors / 显示器矩形


# ============================================================
# 常量 / 路径
# ============================================================

# 项目根目录（dist/ 会务助手.exe 运行时，__file__ 在临时目录，要走 sys._MEIPASS）
if getattr(sys, 'frozen', False):
    _BASE_DIR = sys._MEIPASS
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 截图归档目录：<项目根>/会务收件/截图/YYYY-MM-DD/
INBOX_DIRNAME = '会务收件'
SCREENSHOT_DIRNAME = '截图'

# 全局热键的虚拟键码（Win32 VK_*，跟键盘布局无关）
VK = {
    'shift':  0x10, 'control': 0x11, 'alt': 0x12, 'menu': 0x12,
    'a': 0x41, 's': 0x53, 'q': 0x51, 'x': 0x58, 'z': 0x5A,
    'f1': 0x70, 'f2': 0x71, 'f3': 0x72,
    'esc': 0x1B, 'enter': 0x0D, 'space': 0x20,
}
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_NOREPEAT = 0x4000
MOD_WIN = 0x0008


def _today_dir(base=None):
    """今天的截图归档目录（不存在会自动建）。"""
    base = base or _BASE_DIR
    root = os.path.join(base, INBOX_DIRNAME, SCREENSHOT_DIRNAME,
                        time.strftime('%Y-%m-%d'))
    try:
        os.makedirs(root, exist_ok=True)
    except Exception:
        pass
    return root


def _gen_filename(prefix='截图'):
    """按时间戳生成文件名（不含目录）。"""
    return '%s_%s.png' % (prefix, time.strftime('%H%M%S'))


# ============================================================
# 0. 子窗口摆位 / 唤出（与会务助手主程序保持一致的行为约定）
# ============================================================

def monitor_rect_for(master):
    """找主界面所在那块显示器的矩形；找不到就退回主屏 / 整屏。"""
    rect = None
    try:
        rx = master.winfo_rootx() + max(0, master.winfo_width()) // 2
        ry = master.winfo_rooty() + max(0, master.winfo_height()) // 2
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
                if item.get('primary'):
                    rect = item['rect']
                    break
        except Exception:
            rect = None
    return rect


def center_window(win, width, height, master=None):
    """把子窗口摆到主界面所在显示器的中间。

    与会务助手主程序里的 ``_center_window`` 同一个套路，**刻意不用
    ``max(0, x)`` 钳位**：多显示器时副屏坐标可能是负数，钳成 0 会让窗口
    跑到主屏左上角。尺寸超出显示器时自动缩到能放下为止。
    """
    master = master or getattr(win, 'master', None)
    rect = monitor_rect_for(master) if master is not None else None
    if rect is None:
        try:
            rect = (0, 0, win.winfo_screenwidth(), win.winfo_screenheight())
        except Exception:
            rect = (0, 0, 1280, 800)
    left, top, right, bottom = rect
    mon_w = max(200, right - left)
    mon_h = max(200, bottom - top)
    width = int(min(width, mon_w - 40))
    height = int(min(height, mon_h - 60))
    width = max(320, width)
    height = max(240, height)
    x = left + max(0, (mon_w - width) // 2)
    y = top + max(0, (mon_h - height) // 2)
    try:
        win.geometry('%dx%d+%d+%d' % (width, height, x, y))
    except Exception:
        pass
    return width, height


def present_window(win, topmost_ms=0):
    """按项目约定把新窗口"亮"出来。

    新建的 Toplevel 在部分机器上系统不给发重画消息，界面上就是一块空白
    （白屏 / 黑屏）。主程序里统一靠 `deiconify + lift + focus + 强制重画`
    兜底，这里照做；再补一次延迟重画，防止第一次重画时内容还没布局完。
    """
    try:
        win.deiconify()
    except Exception:
        pass
    try:
        win.update_idletasks()
    except Exception:
        pass
    try:
        win.lift()
    except Exception:
        pass
    try:
        win.focus_force()
    except Exception:
        pass
    try:
        ui_theme.force_repaint(win)
    except Exception:
        pass

    if topmost_ms:
        try:
            win.attributes('-topmost', True)

            def _undo():
                try:
                    win.attributes('-topmost', False)
                except Exception:
                    pass
            win.after(int(topmost_ms), _undo)
        except Exception:
            pass

    def _repaint_again():
        try:
            ui_theme.force_repaint(win)
        except Exception:
            pass
    try:
        win.after(260, _repaint_again)
    except Exception:
        pass


def _force_foreground(win):
    """把窗口真正抢到前台（绕过 Windows 前台锁定）。

    截完图弹标注器时，现场常开着长条屏投屏 / 倒计时这类
    ``overrideredirect + topmost`` 的置顶窗；光靠 ``lift/focus_force``
    可能既抢不到前台、又被它们压在下面，表现为"工具栏点了没反应"。
    这里用经典的 AttachThreadInput 技巧把前台输入状态借过来再
    SetForegroundWindow；非 Windows 或任何一步失败都静默忽略。
    """
    try:
        import ctypes
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        user32.GetAncestor.restype = ctypes.c_void_p
        user32.GetAncestor.argtypes = [ctypes.c_void_p, ctypes.c_uint]
        user32.GetForegroundWindow.restype = ctypes.c_void_p
        user32.GetWindowThreadProcessId.argtypes = [
            ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
        user32.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
        user32.BringWindowToTop.argtypes = [ctypes.c_void_p]
        user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
        user32.AttachThreadInput.argtypes = [ctypes.c_ulong,
                                             ctypes.c_ulong, ctypes.c_int]
        kernel32.GetCurrentThreadId.restype = ctypes.c_ulong

        hwnd = user32.GetAncestor(int(win.winfo_id()), 2) or int(
            win.winfo_id())
        fg = user32.GetForegroundWindow()
        pid = ctypes.c_ulong()
        fg_tid = user32.GetWindowThreadProcessId(fg, ctypes.byref(pid))
        cur_tid = kernel32.GetCurrentThreadId()
        attached = 0
        if fg_tid and fg_tid != cur_tid:
            attached = user32.AttachThreadInput(cur_tid, fg_tid, True)
        try:
            user32.ShowWindow(hwnd, 9)          # SW_RESTORE
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
        finally:
            if attached:
                user32.AttachThreadInput(cur_tid, fg_tid, False)
    except Exception:
        pass


# ============================================================
# 1. 屏幕捕获引擎
# ============================================================
#
# 抓屏底层用 GDI 的 BitBlt + GetDIBits（直接向屏幕 DC 取像素），
# 这段实现**改造自 python-mss 的 Windows 后端**：
#     https://github.com/BoboTiG/python-mss  （MIT License）
#
# 为什么是"搬代码"而不是 pip 装 mss：
#   mss 在 __init__ 里会调用 shcore.SetProcessDpiAwareness(2)，把**整个进程**
#   切成"逐显示器 DPI 感知"。本程序（Tk 界面）是 DPI 不感知、靠系统虚拟化
#   显示的；一旦被切成 DPI 感知，主界面会整体缩水、坐标全乱。
#   所以只取它的抓屏核心，不引入这个副作用。
#
# 抓屏失败时自动回退 Pillow 的 ImageGrab，两条腿走路。

CAPTUREBLT = 0x40000000
SRCCOPY = 0x00CC0020
DIB_RGB_COLORS = 0


def _private_winfunc(dll, name, restype, *argtypes):
    """按函数地址取一个**私有**的带类型副本。

    ctypes 里给 ``user32.XXX.argtypes`` 赋值改的是**进程级共享对象**——
    同进程的 ui_theme / video_player 等模块全都会跟着变。一旦别人把
    ``RegisterClassExW.argtypes`` 设成他们自己的 ``WNDCLASSEXW`` 指针，
    我们再传自己的 ``WNDCLASSEX`` 就会被直接拒掉：

        ArgumentError: expected LP_WNDCLASSEXW instance
                       instead of pointer to WNDCLASSEX

    所以这几个 Win32 函数统统走"取地址 + 自己声明类型"的私有副本，
    既不受别人影响，也不污染别人。
    """
    try:
        addr = ctypes.cast(getattr(dll, name), c_void_p).value
    except Exception:
        return None
    try:
        proto = ctypes.WINFUNCTYPE(restype, *argtypes)
    except Exception:
        return None
    return proto(addr) if addr else None


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ('biSize', wintypes.DWORD), ('biWidth', wintypes.LONG),
        ('biHeight', wintypes.LONG), ('biPlanes', wintypes.WORD),
        ('biBitCount', wintypes.WORD), ('biCompression', wintypes.DWORD),
        ('biSizeImage', wintypes.DWORD), ('biXPelsPerMeter', wintypes.LONG),
        ('biYPelsPerMeter', wintypes.LONG), ('biClrUsed', wintypes.DWORD),
        ('biClrImportant', wintypes.DWORD),
    ]


class _BITMAPINFO(ctypes.Structure):
    _fields_ = [('bmiHeader', _BITMAPINFOHEADER),
                ('bmiColors', wintypes.DWORD * 3)]


class _GdiGrabber(object):
    """GDI 抓屏（BitBlt + GetDIBits），改造自 python-mss/windows.py。"""

    # (dll, 函数名, 返回类型, 参数类型)
    _FUNCS = (
        ('user32', 'GetWindowDC', c_void_p, (c_void_p,)),
        ('user32', 'ReleaseDC', ctypes.c_int, (c_void_p, c_void_p)),
        ('gdi32', 'CreateCompatibleDC', c_void_p, (c_void_p,)),
        ('gdi32', 'DeleteDC', ctypes.c_int, (c_void_p,)),
        ('gdi32', 'CreateCompatibleBitmap', c_void_p,
         (c_void_p, ctypes.c_int, ctypes.c_int)),
        ('gdi32', 'SelectObject', c_void_p, (c_void_p, c_void_p)),
        ('gdi32', 'BitBlt', ctypes.c_int,
         (c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
          c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_uint)),
        ('gdi32', 'GetDIBits', ctypes.c_int,
         (c_void_p, c_void_p, ctypes.c_uint, ctypes.c_uint, c_void_p,
          c_void_p, ctypes.c_uint)),
        ('gdi32', 'DeleteObject', ctypes.c_int, (c_void_p,)),
    )

    def __init__(self):
        self._ok = False
        try:
            dlls = {
                'user32': ctypes.WinDLL('user32', use_last_error=True),
                'gdi32': ctypes.WinDLL('gdi32', use_last_error=True),
            }
            for dll, name, restype, argtypes in self._FUNCS:
                fn = _private_winfunc(dlls[dll], name, restype, *argtypes)
                if fn is None:
                    return
                setattr(self, '_' + name, fn)
            self._ok = True
        except Exception:
            traceback.print_exc()
            self._ok = False

    @property
    def available(self):
        return self._ok and Image is not None

    def grab(self, rect):
        """抓 rect=(left, top, right, bottom)，返回 PIL.Image（RGB）或 None。"""
        if not self.available:
            return None
        left, top, right, bottom = [int(round(v)) for v in rect]
        w = right - left
        h = bottom - top
        if w <= 0 or h <= 0:
            return None
        hdc_src = self._GetWindowDC(0)
        if not hdc_src:
            return None
        hdc_mem = hbmp = None
        try:
            hdc_mem = self._CreateCompatibleDC(hdc_src)
            if not hdc_mem:
                return None
            hbmp = self._CreateCompatibleBitmap(hdc_src, w, h)
            if not hbmp:
                return None
            self._SelectObject(hdc_mem, hbmp)
            if not self._BitBlt(hdc_mem, 0, 0, w, h, hdc_src, left, top,
                                SRCCOPY | CAPTUREBLT):
                return None
            bmi = _BITMAPINFO()
            bmi.bmiHeader.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
            bmi.bmiHeader.biWidth = w
            bmi.bmiHeader.biHeight = -h            # 负值 = top-down
            bmi.bmiHeader.biPlanes = 1
            bmi.bmiHeader.biBitCount = 32
            bmi.bmiHeader.biCompression = 0        # BI_RGB
            buf = ctypes.create_string_buffer(w * h * 4)
            got = self._GetDIBits(hdc_mem, hbmp, 0, h, buf,
                                  ctypes.byref(bmi), DIB_RGB_COLORS)
            if not got:
                return None
            return Image.frombuffer('RGB', (w, h), buf.raw,
                                    'raw', 'BGRX', 0, 1).copy()
        except Exception:
            traceback.print_exc()
            return None
        finally:
            try:
                if hbmp:
                    self._DeleteObject(hbmp)
                if hdc_mem:
                    self._DeleteDC(hdc_mem)
                self._ReleaseDC(0, hdc_src)
            except Exception:
                pass


_GDI_GRABBER = None


def _gdi_grabber():
    """进程内单例（初始化一次，后续复用）。"""
    global _GDI_GRABBER
    if _GDI_GRABBER is None:
        _GDI_GRABBER = _GdiGrabber()
    return _GDI_GRABBER


def _dim_image(img, factor=0.42):
    """把图压暗（factor 越小越暗），用于区域截图里"选区外"的遮罩效果。"""
    try:
        from PIL import ImageEnhance
        return ImageEnhance.Brightness(img).enhance(factor)
    except Exception:
        try:
            return img.point(lambda v: int(v * factor))
        except Exception:
            return img


# ============================================================
# 2.5 DPI 缩放：本程序是 DPI 不感知的，但抓屏必须按物理像素算
# ============================================================
# 这是"区域截图画面被放大 / 滚动截图错位"的根因：
#
#   本程序（Tk）是 **DPI 不感知** 的，Windows 把它的坐标系**虚拟化**成
#   逻辑像素（150% 缩放下 3840x2160 物理屏只有 2560x1440 逻辑）。于是
#       GetSystemMetrics(SM_CXVIRTUALSCREEN) -> 2560   （逻辑）
#       GetDeviceCaps(HORZRES)                -> 2560   （也是逻辑，DC 同样被虚拟化）
#   而 GDI 的 BitBlt 走的是**设备坐标 = 物理像素**，根本不管 DPI：
#       BitBlt(hdc, 0, 0, 2560, 1440, hdc_screen, 0, 0)
#   抓到的其实是**屏幕左上 2560x1440 物理像素**——也就是整屏的左上 2/3。
#   这张图再铺到 2560x1440 逻辑（=3840x2160 物理）的覆盖层上，内容就是
#   物理 1:1 显示，比屏幕上看到的大 1.5 倍 → 用户感觉"画面被放大了"。
#   滚动截图用同一个 _grab，于是也跟着错位。
#
# 拿真实缩放比只能靠 GetDpiForWindow（unaware 进程里它也返回真值）：
#       GetDpiForSystem()            -> 96    （被虚拟化，不可用）
#       GetDpiForWindow(桌面窗口)     -> 144   （真实值，可用）
# 所以：**所有传给 BitBlt 的 rect 都要乘上这个比例**换算成物理像素。

_DPI_SCALE_CACHE = None


def _dpi_scale():
    """物理像素 / 逻辑像素 的比例（100% -> 1.0，150% -> 1.5）。"""
    global _DPI_SCALE_CACHE
    if _DPI_SCALE_CACHE is not None:
        return _DPI_SCALE_CACHE
    scale = 1.0
    try:
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        fn = _private_winfunc(user32, 'GetDpiForWindow', ctypes.c_uint,
                              c_void_p)
        if fn is not None:
            dpi = fn(user32.GetDesktopWindow())
            if dpi:
                scale = dpi / 96.0
    except Exception:
        scale = 1.0
    # 兜底：拿不到就用"物理虚拟屏宽 / 逻辑虚拟屏宽"估
    if scale == 1.0:
        try:
            user32 = ctypes.WinDLL('user32', use_last_error=True)
            gdi32 = ctypes.WinDLL('gdi32', use_last_error=True)
            logical = user32.GetSystemMetrics(78)
            hdc = user32.GetDC(0)
            try:
                physical = gdi32.GetDeviceCaps(hdc, 8)
            finally:
                user32.ReleaseDC(0, hdc)
            if logical > 0 and physical > logical:
                scale = physical / float(logical)
        except Exception:
            pass
    _DPI_SCALE_CACHE = scale if scale > 0 else 1.0
    return _DPI_SCALE_CACHE


def _dpi_reset_cache():
    """测试/显示器变化时清缓存。"""
    global _DPI_SCALE_CACHE
    _DPI_SCALE_CACHE = None


# 缩图质量（Pillow 新旧版本枚举位置不同，做兼容）
try:
    _LANCZOS = Image.Resampling.LANCZOS
except Exception:
    try:
        _LANCZOS = Image.LANCZOS
    except Exception:
        _LANCZOS = 1


class CaptureEngine(object):
    """所有截图模式的统一入口。

    PIL.ImageGrab 在 Windows 上抓的是**虚拟桌面**——多块显示器拼起来的大画布，
    坐标系就是 GDI 的虚拟屏幕坐标。``list_monitors`` 给的 ``rect`` 也是这套坐标。
    """

    @staticmethod
    def _grab(bbox):
        """抓一块区域，返回 PIL.Image（RGB）。失败返回 None。

        ★ **bbox 是逻辑坐标**（Tk / DPI 不感知视角，也是
        ``monitor_rect_for`` / ``virtual_desktop_rect`` 给的坐标系）。
        GDI 的 BitBlt 走**物理**像素，所以这里先按 ``_dpi_scale()`` 换算，
        否则 150% 缩放下只会抓到屏幕的左上 2/3（见文件里 DPI 那段注释）。

        优先走 GDI（BitBlt，快、能抓分层/硬件合成窗口），失败回退
        Pillow 的 ImageGrab（``all_screens=True`` 拿虚拟桌面）。
        """
        scale = _dpi_scale()
        if scale != 1.0:
            bbox = tuple(int(round(v * scale)) for v in bbox)
        left, top, right, bottom = bbox
        if right <= left or bottom <= top:
            return None
        g = _gdi_grabber()
        if g.available:
            img = g.grab((left, top, right, bottom))
            if img is not None:
                return img
        if ImageGrab is None:
            return None
        try:
            img = ImageGrab.grab(bbox=(left, top, right, bottom),
                                 all_screens=True)
            return img.convert('RGB')
        except Exception:
            try:
                img = ImageGrab.grab(bbox=(left, top, right, bottom))
                return img.convert('RGB')
            except Exception:
                return None

    @staticmethod
    def fullscreen(monitor_index=None):
        """截整块屏幕（指定显示器矩形）。"""
        monitors = video_player.list_monitors()
        if not monitors:
            return None
        if monitor_index is None:
            monitor_index = 0
        monitor_index = max(0, min(monitor_index, len(monitors) - 1))
        rect = monitors[monitor_index]['rect']
        left, top, right, bottom = rect
        if right <= left or bottom <= top:
            return None
        return CaptureEngine._grab((left, top, right, bottom))

    @staticmethod
    def region(rect):
        """截指定区域：rect=(left, top, right, bottom)，虚拟屏幕坐标。"""
        left, top, right, bottom = rect
        if right <= left or bottom <= top:
            return None
        return CaptureEngine._grab((left, top, right, bottom))

    @staticmethod
    def virtual_desktop_image():
        """把整块虚拟桌面（多屏合并）抓成一张图。

        区域截图用：先冻结整屏，再在冻结图上框选，坐标零误差。
        """
        return CaptureEngine.region(CaptureEngine.virtual_desktop_rect())

    @staticmethod
    def window(hwnd):
        """截指定窗口（hwnd 是窗口句柄）。

        优先用 ``PrintWindow``（能抓到被遮挡/最小化的窗口，HWND=0 表示整个屏幕），
        失败时回退到 GDI BitBlt（窗口必须在屏幕可见范围内）。

        ★ 这里所有 Win32 调用都走**私有带类型副本**。本进程里别的模块
        （ui_theme / video_player / strip_screen）会给 user32 的**共享函数对象**
        设 argtypes；一旦别人把某个句柄参数声明成 int32，我们再传 64 位句柄就炸：

            ArgumentError: argument 2: <class 'OverflowError'>: int too long to convert

        这正是"窗口截图 / 滚动截图点了没反应"的元凶之一。
        """
        if not hwnd:
            return None
        if Image is None:
            return None
        try:
            user32 = ctypes.WinDLL('user32', use_last_error=True)
            gdi32 = ctypes.WinDLL('gdi32', use_last_error=True)
            _u = lambda n, r, *a: _private_winfunc(user32, n, r, *a)
            _g = lambda n, r, *a: _private_winfunc(gdi32, n, r, *a)
            GetWindowRect = _u('GetWindowRect', ctypes.c_int, c_void_p, c_void_p)
            GetWindowDC = _u('GetWindowDC', c_void_p, c_void_p)
            ReleaseDC = _u('ReleaseDC', ctypes.c_int, c_void_p, c_void_p)
            PrintWindow = _u('PrintWindow', ctypes.c_int,
                            c_void_p, c_void_p, ctypes.c_uint)
            CreateCompatibleDC = _g('CreateCompatibleDC', c_void_p, c_void_p)
            CreateCompatibleBitmap = _g('CreateCompatibleBitmap', c_void_p,
                                        c_void_p, ctypes.c_int, ctypes.c_int)
            SelectObject = _g('SelectObject', c_void_p, c_void_p, c_void_p)
            GetDIBits = _g('GetDIBits', ctypes.c_int,
                           c_void_p, c_void_p, ctypes.c_uint, ctypes.c_uint,
                           c_void_p, c_void_p, ctypes.c_uint)
            DeleteObject = _g('DeleteObject', ctypes.c_int, c_void_p)
            DeleteDC = _g('DeleteDC', ctypes.c_int, c_void_p)
            if None in (GetWindowRect, GetWindowDC, ReleaseDC, PrintWindow,
                        CreateCompatibleDC, CreateCompatibleBitmap, SelectObject,
                        GetDIBits, DeleteObject, DeleteDC):
                return None

            # 拿窗口矩形（逻辑坐标）；乘 DPI 比例换算成物理像素去渲染，
            # 这样 150% 缩放下也是高清的（与区域截图一致）。
            rect = wintypes.RECT()
            if not GetWindowRect(c_void_p(hwnd), ctypes.byref(rect)):
                return None
            scale = _dpi_scale()
            width = max(1, int(round((rect.right - rect.left) * scale)))
            height = max(1, int(round((rect.bottom - rect.top) * scale)))
            hdc_window = GetWindowDC(c_void_p(hwnd))
            if not hdc_window:
                return None
            try:
                hdc_mem = CreateCompatibleDC(hdc_window)
                if not hdc_mem:
                    return None
                try:
                    hbmp = CreateCompatibleBitmap(hdc_window, width, height)
                    if not hbmp:
                        return None
                    try:
                        SelectObject(hdc_mem, hbmp)
                        # PW_RENDERFULLCONTENT = 2（DWM 合成的内容也能抓到）
                        PrintWindow(c_void_p(hwnd), hdc_mem, 2)
                        # 构造 BITMAPINFOHEADER 让 GetDIBits 写 RGB
                        class BMIH(ctypes.Structure):
                            _fields_ = [
                                ('biSize', wintypes.DWORD),
                                ('biWidth', wintypes.LONG),
                                ('biHeight', wintypes.LONG),
                                ('biPlanes', wintypes.WORD),
                                ('biBitCount', wintypes.WORD),
                                ('biCompression', wintypes.DWORD),
                                ('biSizeImage', wintypes.DWORD),
                                ('biXPelsPerMeter', wintypes.LONG),
                                ('biYPelsPerMeter', wintypes.LONG),
                                ('biClrUsed', wintypes.DWORD),
                                ('biClrImportant', wintypes.DWORD),
                            ]
                        bmi = BMIH()
                        bmi.biSize = ctypes.sizeof(BMIH)
                        bmi.biWidth = width
                        bmi.biHeight = -height       # top-down
                        bmi.biPlanes = 1
                        bmi.biBitCount = 32
                        bmi.biCompression = 0        # BI_RGB
                        buf = ctypes.create_string_buffer(width * height * 4)
                        got = GetDIBits(hdc_mem, hbmp, 0, height,
                                        buf, ctypes.byref(bmi), 0)
                        if not got:
                            return None
                        # BGRA → RGB
                        img = Image.frombuffer('RGBA', (width, height),
                                               buf.raw, 'raw', 'BGRA', 0, 1)
                        return img.convert('RGB')
                    finally:
                        DeleteObject(hbmp)
                finally:
                    DeleteDC(hdc_mem)
            finally:
                ReleaseDC(c_void_p(hwnd), hdc_window)
        except Exception:
            traceback.print_exc()
            return None

    @staticmethod
    def list_visible_windows():
        """枚举所有可见顶层窗口，返回 [{hwnd, title, rect}]。

        用于"窗口截图"模式：用户在列表里挑一个。
        """
        out = []

        ENUMPROC = ctypes.WINFUNCTYPE(
            ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

        def _cb(hwnd, _lp):
            try:
                user32 = ctypes.windll.user32
                if not user32.IsWindowVisible(hwnd):
                    return True
                # 跳过没有标题的窗口
                length = user32.GetWindowTextLengthW(hwnd)
                if length <= 0:
                    return True
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value.strip()
                if not title:
                    return True
                # 跳过工具窗口（没有出现在任务栏的）
                ex_style = user32.GetWindowLongW(hwnd, -20)   # GWL_EXSTYLE
                if ex_style & 0x00000080:                      # WS_EX_TOOLWINDOW
                    return True
                # 跳过本进程自己的窗口（截图选区覆盖层等）
                pid = wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if pid.value == os.getpid():
                    return True
                rect = wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                w, h = rect.right - rect.left, rect.bottom - rect.top
                if w < 40 or h < 40:
                    return True
                out.append({
                    'hwnd': int(hwnd),
                    'title': title,
                    'rect': (rect.left, rect.top, rect.right, rect.bottom),
                })
            except Exception:
                pass
            return True

        try:
            ctypes.windll.user32.EnumWindows(ENUMPROC(_cb), 0)
        except Exception:
            pass
        # 当前活动窗口排第一
        try:
            fg = ctypes.windll.user32.GetForegroundWindow()
            for i, item in enumerate(out):
                if item['hwnd'] == fg:
                    out.insert(0, out.pop(i))
                    break
        except Exception:
            pass
        return out

    @staticmethod
    def virtual_desktop_rect():
        """虚拟桌面（多屏合并）的整体 rect，**逻辑坐标**（unaware 视角）。

        注意：这个值在 150% 缩放时是 2560x1440，而屏幕物理是 3840x2160。
        要跟 BitBlt 对话必须乘 ``dpi_scale()``，``_grab`` 里已经做了。
        """
        try:
            user32 = ctypes.windll.user32
            x = user32.GetSystemMetrics(76)    # SM_XVIRTUALSCREEN
            y = user32.GetSystemMetrics(77)    # SM_YVIRTUALSCREEN
            w = user32.GetSystemMetrics(78)    # SM_CXVIRTUALSCREEN
            h = user32.GetSystemMetrics(79)    # SM_CYVIRTUALSCREEN
            return (x, y, x + w, y + h)
        except Exception:
            return (0, 0, tk.Tk().winfo_screenwidth(), tk.Tk().winfo_screenheight())

    @staticmethod
    def dpi_scale():
        """物理像素/逻辑像素比例（100% -> 1.0，150% -> 1.5）。"""
        return _dpi_scale()


# ============================================================
# 2. 全局热键（Win32 RegisterHotKey + 独立消息泵线程）
# ============================================================

# ============================================================
# 2. 全局热键（RegisterHotKey + message-only window）
# ============================================================

class _HotkeyWindow(object):
    """隐藏的 message-only window，WM_HOTKEY 都会送到这里。

    创建一次，所有热键共享。线程里跑 GetMessageW 泵，收到热键就回调。
    """

    _WM_HOTKEY = 0x0312

    def __init__(self):
        self._hwnd = None
        self._thread = None
        self._callbacks = {}                    # id -> (modifiers, vk, func)
        self._next_id = 1
        self._lock = threading.Lock()
        self._ready = threading.Event()
        self._closed = False
        self._reg_fn = None                     # 私有 RegisterHotKey 副本
        self._unreg_fn = None                   # 私有 UnregisterHotKey 副本

    def start(self):
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._run, daemon=True,
                                        name='hotkey-pump')
        self._thread.start()
        # 等窗口建好再返回，避免 race
        self._ready.wait(timeout=2.0)

    def _run(self):
        """在线程里：建 message-only window，跑消息泵。"""
        try:
            user32 = ctypes.windll.user32
            # 全部走私有带类型副本，绕开进程级 argtypes 被别人改过的问题
            RegisterClassExW = _private_winfunc(
                user32, 'RegisterClassExW', wintypes.ATOM, c_void_p)
            CreateWindowExW = _private_winfunc(
                user32, 'CreateWindowExW', wintypes.HWND,
                wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR,
                wintypes.DWORD, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                ctypes.c_int, wintypes.HWND, c_void_p, c_void_p, c_void_p)
            DefWindowProcW = _private_winfunc(
                user32, 'DefWindowProcW', wintypes.LPARAM,
                wintypes.HWND, wintypes.UINT, wintypes.WPARAM,
                wintypes.LPARAM)
            GetMessageW = _private_winfunc(
                user32, 'GetMessageW', ctypes.c_int,
                c_void_p, wintypes.HWND, wintypes.UINT, wintypes.UINT)
            TranslateMessage = _private_winfunc(
                user32, 'TranslateMessage', wintypes.BOOL, c_void_p)
            DispatchMessageW = _private_winfunc(
                user32, 'DispatchMessageW', c_void_p, c_void_p)
            if None in (RegisterClassExW, CreateWindowExW, DefWindowProcW,
                        GetMessageW, TranslateMessage, DispatchMessageW):
                return

            # 注册一个临时窗口类（必须，否则 CreateWindowEx 失败）
            WNDCLASSEX = self._wndclassex()

            class_name = 'HuiWuHotkey_%d' % os.getpid()

            def_wnd_proc = ctypes.WINFUNCTYPE(
                ctypes.c_int64, wintypes.HWND, wintypes.UINT,
                wintypes.WPARAM, wintypes.LPARAM)

            def _wnd_proc(hwnd, msg, wparam, lparam):
                if msg == self._WM_HOTKEY:
                    hotkey_id = int(wparam) & 0xFFFF
                    with self._lock:
                        cb = self._callbacks.get(hotkey_id)
                    if cb is not None:
                        try:
                            cb[2]()             # func
                        except Exception:
                            traceback.print_exc()
                # 一定要返回整数（回调声明的是整型返回码）
                return int(DefWindowProcW(hwnd, msg, wparam, lparam) or 0)

            # 把 closure 钩到 C 回调上（避免被回收）
            self._wnd_proc_ref = def_wnd_proc(_wnd_proc)

            wc = WNDCLASSEX()
            wc.cbSize = ctypes.sizeof(WNDCLASSEX)
            # WINFUNCTYPE 实例要 cast 成 c_void_p 才能放进结构体字段
            wc.lpfnWndProc = ctypes.cast(self._wnd_proc_ref, c_void_p).value
            wc.hInstance = 0
            wc.lpszClassName = class_name
            # 注册失败也不直接放弃：类可能已存在（同一进程重复注册），
            # 这时后面的 CreateWindowExW 依然能用这个类名建窗口。
            try:
                RegisterClassExW(ctypes.byref(wc))
            except Exception:
                pass

            # HWND_MESSAGE = -3，message-only window，不显示也不抢焦点
            self._hwnd = CreateWindowExW(
                0, class_name, 'huihu_hotkey',
                0, 0, 0, 0, 0, -3, 0, 0, 0)
            if not self._hwnd:
                return

            self._ready.set()

            # 消息泵
            msg = wintypes.MSG()
            while not self._closed:
                ret = GetMessageW(ctypes.byref(msg), None, 0, 0)
                if ret <= 0:
                    break
                TranslateMessage(ctypes.byref(msg))
                DispatchMessageW(ctypes.byref(msg))
        except Exception:
            traceback.print_exc()

    @staticmethod
    def _wndclassex():
        class WNDCLASSEX(ctypes.Structure):
            _fields_ = [
                ('cbSize', wintypes.UINT),
                ('style', wintypes.UINT),
                ('lpfnWndProc', ctypes.c_void_p),
                ('cbClsExtra', ctypes.c_int),
                ('cbWndExtra', ctypes.c_int),
                ('hInstance', ctypes.c_void_p),
                ('hIcon', ctypes.c_void_p),
                ('hCursor', ctypes.c_void_p),
                ('hbrBackground', ctypes.c_void_p),
                ('lpszMenuName', wintypes.LPCWSTR),
                ('lpszClassName', wintypes.LPCWSTR),
                ('hIconSm', ctypes.c_void_p),
            ]
        return WNDCLASSEX

    def _hotkey_api(self):
        """惰性拿 RegisterHotKey / UnregisterHotKey 的私有副本。"""
        if self._reg_fn is None:
            user32 = ctypes.windll.user32
            self._reg_fn = _private_winfunc(
                user32, 'RegisterHotKey', wintypes.BOOL,
                wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT)
            self._unreg_fn = _private_winfunc(
                user32, 'UnregisterHotKey', wintypes.BOOL,
                wintypes.HWND, ctypes.c_int)
        return self._reg_fn, self._unreg_fn

    def register(self, modifiers, vk, func):
        """注册一个全局热键；返回 hotkey id，失败返回 None。"""
        if not self._ready.wait(timeout=2.0):
            return None
        if self._hwnd is None:
            return None
        with self._lock:
            hotkey_id = self._next_id
            self._next_id += 1
            self._callbacks[hotkey_id] = (modifiers, vk, func)
        try:
            reg_fn, _unreg = self._hotkey_api()
            if reg_fn is None:
                raise RuntimeError('RegisterHotKey 不可用')
            ok = reg_fn(self._hwnd, hotkey_id,
                        modifiers | MOD_NOREPEAT, vk)
            if not ok:
                with self._lock:
                    self._callbacks.pop(hotkey_id, None)
                return None
            return hotkey_id
        except Exception:
            with self._lock:
                self._callbacks.pop(hotkey_id, None)
            return None

    def unregister(self, hotkey_id):
        if hotkey_id is None or self._hwnd is None:
            return
        try:
            _reg, unreg_fn = self._hotkey_api()
            if unreg_fn is not None:
                unreg_fn(self._hwnd, hotkey_id)
        except Exception:
            pass
        with self._lock:
            self._callbacks.pop(hotkey_id, None)


# 单例：消息泵线程只跑一个
_HOTKEY_WINDOW = None
_HOTKEY_LOCK = threading.Lock()


def _hotkey_manager():
    global _HOTKEY_WINDOW
    with _HOTKEY_LOCK:
        if _HOTKEY_WINDOW is None:
            _HOTKEY_WINDOW = _HotkeyWindow()
            _HOTKEY_WINDOW.start()
        return _HOTKEY_WINDOW


def parse_hotkey(spec):
    """把 'Ctrl+Shift+S' 解析成 (modifiers, vk)；解析失败返回 None。

    支持：Ctrl / Shift / Alt / Win + 一个主键（字母 / F1~F12 / Esc / Enter）。
    """
    if not spec:
        return None
    parts = [p.strip().lower() for p in spec.split('+') if p.strip()]
    if not parts:
        return None
    modifiers = 0
    main_keys = []
    for p in parts:
        if p in ('ctrl', 'control', 'ctl'):
            modifiers |= MOD_CONTROL
        elif p == 'shift':
            modifiers |= MOD_SHIFT
        elif p in ('alt', 'menu'):
            modifiers |= MOD_ALT
        elif p in ('win', 'meta', 'cmd'):
            modifiers |= MOD_WIN
        else:
            main_keys.append(p)
    if len(main_keys) != 1:
        return None
    key = main_keys[0].upper()
    vk = VK.get(key) or VK.get(key.lower())
    if vk is None:
        # 允许 'A' / 'a' 这种
        if len(key) == 1 and 'A' <= key <= 'Z':
            vk = ord(key)
        else:
            return None
    return modifiers, vk


# ============================================================
# 3. 区域选择覆盖层（在"冻结的屏幕图"上框选）
# ============================================================

class RegionOverlay(object):
    """全屏覆盖层：在**冻结的屏幕截图**上拖框选区。

    这是"区域截图不准 / 不能用"的根治办法，也是 ShareX / Flameshot /
    Snipaste 的通用做法：

        * 先把整块虚拟桌面抓成一张**静态图**，铺满覆盖层；
        * 用户在这张图上框选，框到的是**图片像素坐标**，最后直接
          ``frozen.crop(rect)`` 裁出来——所见即所得，零误差；
        * 旧实现让用户对着"活桌面"框，再按框的坐标去截屏；中间隔着
          窗口重绘 / 遮罩淡出 / 面板恢复，帧和帧对不齐，于是"框的位置
          和截到的内容差一截"。

    同时修掉"释放坐标陈旧"的问题：旧代码在鼠标松开时用的是**最后一次
    Motion 事件**留下的旧坐标（``self._cur``），而不是松开那一刻的真实
    坐标，快速拖框时会偏几十上百像素。现在 ``_on_release`` 直接读
    ``event.x/event.y``。
    """

    def __init__(self, master, on_done, on_cancel=None, frozen=None):
        """on_done(rect) -> rect 是**物理像素坐标** (left, top, right, bottom)，
        取自冻结原图，可直接 ``frozen.crop(rect)``。
        on_cancel()    -> 用户按 ESC / 右键 / 只点了一下没拖"""
        self._master = master
        self._on_done = on_done
        self._on_cancel = on_cancel or (lambda: None)

        v = CaptureEngine.virtual_desktop_rect()
        self._vleft, self._vtop, self._vright, self._vbottom = v
        self._vw = max(1, self._vright - self._vleft)   # 逻辑宽
        self._vh = max(1, self._vbottom - self._vtop)   # 逻辑高

        self._frozen = frozen
        if self._frozen is None:
            self._frozen = CaptureEngine.region(v)

        # ★ 关键：冻结图是**物理像素**（150% 缩放时 3840x2160），而覆盖层窗口
        #   是**逻辑尺寸**（2560x1440）。两者不能用同一套坐标：
        #     * 窗口/canvas 用逻辑坐标 → 选框画得跟鼠标严丝合缝；
        #     * 显示用一张降到逻辑尺寸的副本 → 所见 = 屏幕，不再"放大 1.5 倍"；
        #     * 裁剪用原始物理图 + 逻辑选区×比例 → 输出仍是高清。
        #   比例直接从图片实际尺寸反推（不依赖全局 DPI 探测，永远自洽）。
        self._scale = 1.0
        if self._frozen is not None:
            try:
                iw = self._frozen.size[0]
                if iw > 0:
                    self._scale = iw / float(self._vw)
            except Exception:
                self._scale = 1.0
        else:
            try:
                self._scale = CaptureEngine.dpi_scale()
            except Exception:
                self._scale = 1.0

        # 显示副本（逻辑尺寸）：拖框时贴"亮框"也用它
        self._display = self._frozen
        if self._frozen is not None:
            try:
                if self._frozen.size != (self._vw, self._vh):
                    self._display = self._frozen.resize(
                        (self._vw, self._vh), _LANCZOS)
            except Exception:
                self._display = self._frozen

        self._start = None           # 起点（canvas 逻辑坐标）
        self._cur = None             # 当前点（canvas 逻辑坐标）
        self._rect_id = None
        self._size_id = None
        self._bright_id = None
        self._photo_dim = None
        self._photo_bright = None
        self._pending = False
        self._finished = False

        self.win = tk.Toplevel(master)
        self.win.overrideredirect(True)
        self.win.geometry('%dx%d+%d+%d' % (self._vw, self._vh,
                                           self._vleft, self._vtop))
        try:
            self.win.attributes('-topmost', True)
        except tk.TclError:
            pass
        self.win.configure(bg='#0b1220')

        self.canvas = tk.Canvas(self.win, width=self._vw, height=self._vh,
                                bg='#0b1220', highlightthickness=0, bd=0,
                                cursor='cross')
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # 背景：**压暗**的冻结图（逻辑尺寸副本；只建一次，拖动时不重建）
        if self._display is not None:
            try:
                self._photo_dim = ImageTk.PhotoImage(
                    _dim_image(self._display, 0.42))
                self.canvas.create_image(0, 0, image=self._photo_dim,
                                         anchor='nw', tags='bg')
            except Exception:
                traceback.print_exc()

        # 选框（画在图像坐标上，跟最终裁剪区域完全一致）
        self._rect_id = self.canvas.create_rectangle(
            0, 0, 0, 0, outline='#22d3ee', width=2, dash=(7, 4),
            state='hidden')
        self._size_id = self.canvas.create_text(
            0, 0, text='', fill='#22d3ee', font=ui_theme.font(11, 'bold'),
            anchor='nw', state='hidden')

        # 顶部提示条
        self.canvas.create_rectangle(0, 0, self._vw, 56, fill='#0b1220',
                                     outline='', tags='hintbg', state='normal')
        self.canvas.create_text(
            self._vw // 2, 28,
            text='按住左键拖框选区，松开即截    ·    ESC 取消',
            fill='#e6f7ff', font=ui_theme.font(12, 'bold'), tags='hint')
        self.canvas.tag_raise('hintbg')
        self.canvas.tag_raise('hint')

        # 事件：win 和 canvas 都绑一份，避免焦点落到别处丢事件
        for w in (self.win, self.canvas):
            w.bind('<ButtonPress-1>', self._on_press)
            w.bind('<B1-Motion>', self._on_drag)
            w.bind('<ButtonRelease-1>', self._on_release)
            w.bind('<Escape>', lambda e: self._finish(None))
            w.bind('<Return>', lambda e: self._finish(self._selection()))
            w.bind('<Button-3>', lambda e: self._finish(None))

        # 置顶 + 抓焦点 + 强制重画（防白屏）
        try:
            self.win.lift()
            self.win.focus_force()
            self.canvas.focus_set()
        except tk.TclError:
            pass
        try:
            ui_theme.force_repaint(self.win)
        except Exception:
            pass
        try:
            self.win.after(150, self._repaint)
        except Exception:
            pass

    # ---------- 坐标 ----------
    def _to_canvas(self, cx, cy):
        """鼠标事件坐标 → canvas 逻辑坐标（夹到窗口内）。

        注意这里**不做缩放**：画布就是逻辑尺寸，鼠标坐标天然一致。
        逻辑→物理的换算只在 ``_finish`` 里做一次（裁剪时才需要）。
        """
        cx = max(0, min(int(round(cx)), self._vw))
        cy = max(0, min(int(round(cy)), self._vh))
        return cx, cy

    def _to_physical(self, rect):
        """canvas 逻辑选区 → 冻结**物理图**的像素坐标。"""
        if rect is None:
            return None
        if self._scale and self._scale != 1.0:
            return tuple(int(round(v * self._scale)) for v in rect)
        return tuple(int(v) for v in rect)

    def _selection(self):
        """当前选区（canvas 逻辑坐标）；太小返回 None。"""
        if self._start is None or self._cur is None:
            return None
        x1 = min(self._start[0], self._cur[0])
        y1 = min(self._start[1], self._cur[1])
        x2 = max(self._start[0], self._cur[0])
        y2 = max(self._start[1], self._cur[1])
        if x2 - x1 < 2 or y2 - y1 < 2:
            return None
        return (int(x1), int(y1), int(x2), int(y2))

    # ---------- 事件 ----------
    def _on_press(self, event):
        self._start = self._to_canvas(event.x, event.y)
        self._cur = self._start
        self._schedule_refresh()

    def _on_drag(self, event):
        if self._start is None:
            return
        self._cur = self._to_canvas(event.x, event.y)
        self._schedule_refresh()

    def _on_release(self, event):
        """★ 关键修复：读"松开这一刻"的真实坐标，而不是最后一次 Motion 的旧值。"""
        if self._start is None:
            return
        self._cur = self._to_canvas(event.x, event.y)
        rect = self._selection()
        if rect is None:
            # 只是点了一下、没真正拖：清掉，让用户重来（不直接退出）
            self._start = None
            self._cur = None
            self._schedule_refresh()
            return
        self._finish(rect)

    # ---------- 刷新（合并抖动，避免狂建图） ----------
    def _schedule_refresh(self):
        if self._pending or self._finished:
            return
        self._pending = True
        try:
            self.win.after(0, self._do_refresh)
        except tk.TclError:
            self._pending = False

    def _do_refresh(self):
        self._pending = False
        if self._finished:
            return
        rect = self._selection()
        if rect is None:
            if self._bright_id is not None:
                self.canvas.itemconfigure(self._bright_id, state='hidden')
            self.canvas.itemconfigure(self._rect_id, state='hidden')
            self.canvas.itemconfigure(self._size_id, state='hidden')
            return
        x1, y1, x2, y2 = rect

        # 亮框：把冻结图对应区域原样贴回来 → "框内明亮、框外压暗"
        # 用 _display（逻辑尺寸），才能跟 canvas 坐标 1:1 对上；
        # 用物理原图会错位（尺寸差 1.5 倍）。
        if self._display is not None:
            try:
                crop = self._display.crop((x1, y1, x2, y2))
                self._photo_bright = ImageTk.PhotoImage(crop)
                if self._bright_id is None:
                    self._bright_id = self.canvas.create_image(
                        x1, y1, image=self._photo_bright, anchor='nw')
                    self.canvas.tag_lower(self._bright_id, self._rect_id)
                else:
                    self.canvas.itemconfigure(self._bright_id,
                                              image=self._photo_bright,
                                              state='normal')
                    self.canvas.coords(self._bright_id, x1, y1)
            except Exception:
                traceback.print_exc()

        self.canvas.coords(self._rect_id, x1, y1, x2, y2)
        self.canvas.itemconfigure(self._rect_id, state='normal')
        self.canvas.coords(self._size_id, x1 + 8, y1 + 8)
        # 尺寸标签显示**物理像素**（裁剪出来的实际大小），比逻辑值更有意义
        px1, py1, px2, py2 = self._to_physical(rect)
        self.canvas.itemconfigure(
            self._size_id, text='%d × %d' % (px2 - px1, py2 - py1),
            state='normal')

        try:
            self.canvas.tag_raise('hintbg')
            self.canvas.tag_raise('hint')
        except tk.TclError:
            pass

    def _repaint(self):
        try:
            ui_theme.force_repaint(self.win)
        except Exception:
            pass

    # ---------- 收尾 ----------
    def _finish(self, rect):
        if self._finished:
            return
        self._finished = True
        try:
            self.win.destroy()
        except tk.TclError:
            pass
        if rect is None:
            try:
                self._on_cancel()
            except Exception:
                traceback.print_exc()
        else:
            # ★ 逻辑选区 → 物理像素（冻结原图坐标），这样裁出来是高清的
            try:
                rect = self._to_physical(rect)
            except Exception:
                traceback.print_exc()
            try:
                self._on_done(rect)
            except Exception:
                traceback.print_exc()


# ============================================================
# 4. 简易标注编辑器
# ============================================================

class AnnotationEditor(object):
    """截完图后弹出编辑器：矩形 / 箭头 / 文字 / 马赛克 + 撤销 / 重做 + 保存 / 复制。

    用 ``tk.Canvas`` 直接画矩形、文字箭头，存到 ``self._shapes`` 列表，
    撤销/重做用栈实现；保存走 PIL。
    """

    _TOOLS = ('选择', '矩形', '箭头', '文字', '马赛克')

    def __init__(self, master, image, on_saved=None, on_cancel=None):
        self._master = master
        self._image = image.copy()                      # 原始底图（不再修改）
        self._work = image.copy()                       # 当前工作图层
        self._shapes = []                               # 已画的所有形状
        self._redo_stack = []
        self._on_saved = on_saved or (lambda path: None)
        self._on_cancel = on_cancel or (lambda: None)
        self._draft = None                              # 鼠标拖动过程中的临时形状
        self._tool = tk.StringVar(value='矩形')
        self._color = tk.StringVar(value='#dc2626')
        self._stroke = tk.IntVar(value=3)
        self._mosaic = tk.IntVar(value=12)
        self._font_size = tk.IntVar(value=20)

        self.win = tk.Toplevel(master)
        self.win.title('截图标注（Ctrl+S 保存，Esc 取消，Ctrl+Z 撤销）')
        try:
            ui_theme.style_window(self.win)
        except Exception:
            pass
        self.win.transient(master)
        self.win.configure(bg='#f4f6fb')

        # 画布尺寸自适应：图片长边按 900 缩一下，避免 4K 屏撑爆
        max_side = 900
        iw, ih = self._image.size
        scale = min(1.0, max_side / float(max(iw, ih)))
        self._scale = scale
        self._cw = max(320, int(round(iw * scale)))
        self._ch = max(200, int(round(ih * scale)))
        # 显示用的图（缩放后）
        self._display = self._image.resize((self._cw, self._ch))
        self._photo = ImageTk.PhotoImage(self._display)

        self._build_toolbar()

        # 画布
        body = tk.Frame(self.win, bg='#f4f6fb')
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self.canvas = tk.Canvas(body, width=self._cw, height=self._ch,
                                bg='#ffffff', highlightthickness=1,
                                highlightbackground='#cbd5e1', cursor='cross')
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._canvas_img = self.canvas.create_image(
            0, 0, image=self._photo, anchor='nw')

        # 事件
        self.canvas.bind('<ButtonPress-1>', self._on_press)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<ButtonRelease-1>', self._on_release)
        self.canvas.bind('<Double-Button-1>', self._on_text_commit)
        for seq in ('<Control-s>', '<Control-S>'):
            self.win.bind(seq, lambda e: self.save())
        for seq in ('<Control-z>', '<Control-Z>'):
            self.win.bind(seq, lambda e: self.undo())
        for seq in ('<Control-y>', '<Control-Y>',
                    '<Control-Shift-z>', '<Control-Shift-Z>'):
            self.win.bind(seq, lambda e: self.redo())
        self.win.bind('<Escape>', lambda e: self.cancel())
        self.win.bind('<Control-c>', lambda e: self.copy_to_clipboard())

        # 关闭协议
        self.win.protocol('WM_DELETE_WINDOW', self.cancel)
        # 摆到主界面所在显示器的中间（多屏 / 高 DPI 下也不会跑到左上角）
        try:
            self.win.update_idletasks()
            rw = max(360, self.win.winfo_reqwidth())
            rh = max(260, self.win.winfo_reqheight())
        except Exception:
            rw = max(360, self._cw + 32)
            rh = max(260, self._ch + 96)
        center_window(self.win, rw, rh, master=master)
        # ★ 标注期间保持置顶：现场常开的长条屏投屏 / 倒计时等无边框置顶窗
        #   会把普通窗口压在下面，导致工具栏"点了没反应"（点击被上层窗吞掉）。
        #   保存/取消时本窗口直接销毁，置顶状态随之消失，不留下副作用。
        try:
            self.win.attributes('-topmost', True)
        except tk.TclError:
            pass
        present_window(self.win)
        _force_foreground(self.win)
        # 窗口是在选区遮罩的鼠标事件调用栈里创建的，再延迟补一次抬前台，
        # 防止遮罩销毁瞬间前台状态交接失败
        self.win.after(150, self._bring_to_front)

        self._refresh_undo_state()

    def _bring_to_front(self):
        try:
            if not self.win.winfo_exists():
                return
            self.win.lift()
            self.win.attributes('-topmost', True)
            self.win.focus_force()
        except tk.TclError:
            return
        _force_foreground(self.win)

    # ----- UI -----
    def _build_toolbar(self):
        p = ui_theme.PALETTE
        bar = tk.Frame(self.win, bg=p['surface'])
        bar.pack(fill=tk.X, padx=8, pady=8)

        # 工具单选
        tool_frame = tk.Frame(bar, bg=p['surface'])
        tool_frame.pack(side=tk.LEFT)
        for name in self._TOOLS:
            ttk.Radiobutton(tool_frame, text=name, variable=self._tool,
                            value=name).pack(side=tk.LEFT, padx=2)

        # 颜色（6 个预设）
        color_frame = tk.Frame(bar, bg=p['surface'])
        color_frame.pack(side=tk.LEFT, padx=(20, 6))
        tk.Label(color_frame, text='颜色', font=ui_theme.font(9),
                 bg=p['surface'], fg=p['muted']).pack(side=tk.LEFT, padx=(0, 4))
        for col in ('#dc2626', '#2563eb', '#16a34a', '#f59e0b',
                    '#7c3aed', '#0f172a'):
            tk.Label(color_frame, text='', bg=col, width=2,
                     relief='ridge', bd=1, cursor='hand2').pack(
                side=tk.LEFT, padx=1)
        # 颜色按钮点击事件
        for child in color_frame.winfo_children()[1:]:
            col = child.cget('bg')
            child.bind('<Button-1>', lambda e, c=col: self._color.set(c))

        # 笔粗
        tk.Label(bar, text='粗细', font=ui_theme.font(9),
                 bg=p['surface'], fg=p['muted']).pack(
            side=tk.LEFT, padx=(10, 4))
        ttk.Spinbox(bar, from_=1, to=20, width=4, textvariable=self._stroke,
                    ).pack(side=tk.LEFT)

        # 马赛克块大小（仅在选马赛克时生效）
        tk.Label(bar, text='马赛克', font=ui_theme.font(9),
                 bg=p['surface'], fg=p['muted']).pack(
            side=tk.LEFT, padx=(10, 4))
        ttk.Spinbox(bar, from_=4, to=64, width=4, textvariable=self._mosaic,
                    ).pack(side=tk.LEFT)

        # 文字大小
        tk.Label(bar, text='字号', font=ui_theme.font(9),
                 bg=p['surface'], fg=p['muted']).pack(
            side=tk.LEFT, padx=(10, 4))
        ttk.Spinbox(bar, from_=10, to=80, width=4, textvariable=self._font_size,
                    ).pack(side=tk.LEFT)

        # 操作按钮
        ttk.Button(bar, text='撤销', style='Ghost.TButton',
                   command=self.undo, width=6).pack(side=tk.LEFT, padx=(20, 2))
        ttk.Button(bar, text='重做', style='Ghost.TButton',
                   command=self.redo, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text='复制', style='Ghost.TButton',
                   command=self.copy_to_clipboard, width=6).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(bar, text='保存…', style='Accent.TButton',
                   command=self.save_as, width=8).pack(side=tk.LEFT, padx=(8, 2))
        ttk.Button(bar, text='保存到收件', style='Accent.TButton',
                   command=self.save, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text='取消', style='Ghost.TButton',
                   command=self.cancel, width=6).pack(side=tk.LEFT, padx=2)

    # ----- 形状数据结构 -----
    # 每条形状：
    #   ('rect',   canvas_rect=(x1,y1,x2,y2), color, stroke, image_rect=(orig坐标))
    #   ('arrow',  canvas_pts=[(x1,y1),(x2,y2)], color, stroke, image_pts=...)
    #   ('text',   canvas_xy=(x,y), text=..., color, font_size, image_xy=...)
    #   ('mosaic', canvas_rect, image_rect)
    def _to_image(self, x, y):
        return (x / self._scale, y / self._scale)

    def _to_canvas(self, x, y):
        return (x * self._scale, y * self._scale)

    def _snap_rect(self, x1, y1, x2, y2):
        return (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))

    # ----- 鼠标交互 -----
    # 工具栏单选显示的是中文名，形状数据内部一律用英文 kind
    # （_redraw_draft / _apply_shape 全按 rect/arrow/text/mosaic 分支）
    _TOOL_KIND = {'矩形': 'rect', '箭头': 'arrow',
                  '文字': 'text', '马赛克': 'mosaic'}

    def _on_press(self, event):
        kind = self._TOOL_KIND.get(self._tool.get())
        if kind is None:
            # “选择”或未知工具：画布按下不产生任何形状
            return
        if kind == 'text':
            # 文字模式：单击放文字（双击在拖动后 commit）
            self._draft = {
                'kind': 'text',
                'canvas_xy': (event.x, event.y),
                'image_xy': self._to_image(event.x, event.y),
                'text': '',
                'color': self._color.get(),
                'font_size': self._font_size.get(),
            }
            return
        # 其他工具：开始拖框
        self._draft = {
            'kind': kind,
            'canvas_start': (event.x, event.y),
            'canvas_end': (event.x, event.y),
            'color': self._color.get(),
            'stroke': self._stroke.get(),
            'mosaic': self._mosaic.get(),
        }
        self._redraw_draft()

    def _on_drag(self, event):
        if self._draft is None:
            return
        if self._draft['kind'] == 'text':
            return
        self._draft['canvas_end'] = (event.x, event.y)
        self._redraw_draft()

    def _on_release(self, event):
        if self._draft is None:
            return
        kind = self._draft['kind']
        if kind == 'text':
            # 文字模式单击：弹输入框
            self._ask_text_input(event.x_root, event.y_root)
            return
        x1, y1 = self._draft['canvas_start']
        x2, y2 = self._draft['canvas_end']
        rx1, ry1, rx2, ry2 = self._snap_rect(x1, y1, x2, y2)
        # 箭头允许细长形状：只看起点终点距离，不强制宽高阈值
        if kind == 'arrow':
            if abs(x2 - x1) < 4 and abs(y2 - y1) < 4:
                self._draft = None
                self.canvas.delete('draft')
                return
        else:
            if rx2 - rx1 < 4 or ry2 - ry1 < 4:
                self._draft = None
                self.canvas.delete('draft')
                return
        # 把画布坐标转成原始图坐标存进 shape
        ix1, iy1 = self._to_image(rx1, ry1)
        ix2, iy2 = self._to_image(rx2, ry2)
        shape = {
            'kind': kind,
            'canvas_rect': (rx1, ry1, rx2, ry2),
            'image_rect': (ix1, iy1, ix2, iy2),
            'canvas_start': (rx1, ry1),
            'canvas_end': (rx2, ry2),
            'color': self._draft['color'],
            'stroke': self._draft['stroke'],
        }
        if kind == 'arrow':
            # ★ 箭头必须保留真实起笔/收笔方向：归一化后的 rect 会把
            # “从右往左画”翻成“从左往右”，导致箭头永远朝右；
            # 箭头头部画在收笔端 (x2,y2)，所以这里存**未归一化**的原始端点。
            shape['canvas_start'] = (x1, y1)
            shape['canvas_end'] = (x2, y2)
            shape['image_start'] = self._to_image(x1, y1)
            shape['image_end'] = self._to_image(x2, y2)
        if kind == 'mosaic':
            shape['mosaic'] = self._draft['mosaic']
        self._push_shape(shape)
        self._draft = None
        # ★ 草稿是画布上的临时线条，形状已烙进底图；不删掉的话，
        # 反向拖箭头时草稿箭头和烙进底图的箭头会同时显示成两个。
        self.canvas.delete('draft')

    def _on_text_commit(self, _event):
        pass

    def _ask_text_input(self, x_root, y_root):
        # 用 simpledialog.askstring 但要让它跟随覆盖层
        from tkinter import simpledialog
        text = simpledialog.askstring(
            '输入文字', '请输入要添加的文字：',
            parent=self.win)
        if not text:
            self._draft = None
            return
        d = self._draft
        d['text'] = text
        shape = {
            'kind': 'text',
            'canvas_xy': d['canvas_xy'],
            'image_xy': d['image_xy'],
            'text': text,
            'color': d['color'],
            'font_size': d['font_size'],
        }
        self._push_shape(shape)
        self._draft = None

    def _redraw_draft(self):
        # 抹掉之前的 draft，重新画
        self.canvas.delete('draft')
        if self._draft is None:
            return
        kind = self._draft['kind']
        if kind in ('rect', 'mosaic'):
            x1, y1 = self._draft['canvas_start']
            x2, y2 = self._draft['canvas_end']
            rx1, ry1, rx2, ry2 = self._snap_rect(x1, y1, x2, y2)
            color = '#dc2626' if kind == 'mosaic' else self._draft['color']
            dash = (4, 4) if kind == 'mosaic' else None
            self.canvas.create_rectangle(
                rx1, ry1, rx2, ry2, outline=color, width=2,
                tags='draft', dash=dash)
        elif kind == 'arrow':
            x1, y1 = self._draft['canvas_start']
            x2, y2 = self._draft['canvas_end']
            self.canvas.create_line(
                x1, y1, x2, y2, fill=self._draft['color'],
                width=self._draft['stroke'], arrow=tk.LAST,
                tags='draft')

    # ----- 形状栈 -----
    def _push_shape(self, shape):
        self._shapes.append(shape)
        self._redo_stack.clear()
        self._recompose()
        self._refresh_undo_state()

    def undo(self):
        if not self._shapes:
            return
        self._redo_stack.append(self._shapes.pop())
        self._recompose()
        self._refresh_undo_state()

    def redo(self):
        if not self._redo_stack:
            return
        self._shapes.append(self._redo_stack.pop())
        self._recompose()
        self._refresh_undo_state()

    def _refresh_undo_state(self):
        # 重新画所有形状（直接从 work 复制 display）
        # 清掉所有形状图层，再重新画
        pass

    def _recompose(self):
        """在 ``self._work``（原图）上重新画所有形状，缩放成 display 给画布。"""
        work = self._image.copy()
        draw = ImageDraw.Draw(work)
        for s in self._shapes:
            self._apply_shape(draw, work, s)
        self._work = work
        # 更新显示
        self._display = work.resize((self._cw, self._ch))
        self._photo = ImageTk.PhotoImage(self._display)
        self.canvas.itemconfigure(self._canvas_img, image=self._photo)

    @staticmethod
    def _hex_to_rgb(s):
        s = s.lstrip('#')
        return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))

    def _apply_shape(self, draw, image, shape):
        kind = shape['kind']
        if kind == 'rect':
            x1, y1, x2, y2 = shape['image_rect']
            for k in range(max(1, int(round(shape['stroke'])))):
                draw.rectangle(
                    (x1 - k, y1 - k, x2 + k, y2 + k),
                    outline=shape['color'])
        elif kind == 'arrow':
            # 箭头：优先用落形时存好的原图端点（保留方向），
            # 兼容旧 shape 时再从 canvas 端点现转。
            if shape.get('image_start') is not None:
                x1, y1 = shape['image_start']
                x2, y2 = shape['image_end']
            else:
                cx1, cy1 = shape.get('canvas_start')
                cx2, cy2 = shape.get('canvas_end')
                x1, y1 = self._to_image(cx1, cy1)
                x2, y2 = self._to_image(cx2, cy2)
            stroke = max(1, int(round(shape['stroke'])))
            color = shape['color']
            draw.line([(x1, y1), (x2, y2)], fill=color, width=stroke)
            # 箭头头部（两条短线拼成的三角形）
            import math
            angle = math.atan2(y2 - y1, x2 - x1)
            head_len = max(8.0, stroke * 4)
            head_angle = math.pi / 6
            ax1 = x2 - head_len * math.cos(angle - head_angle)
            ay1 = y2 - head_len * math.sin(angle - head_angle)
            ax2 = x2 - head_len * math.cos(angle + head_angle)
            ay2 = y2 - head_len * math.sin(angle + head_angle)
            draw.polygon([(x2, y2), (ax1, ay1), (ax2, ay2)], fill=color)
        elif kind == 'text':
            x, y = shape['image_xy']
            color = shape['color']
            size = max(10, int(shape['font_size']))
            try:
                font = ImageFont.truetype('msyh.ttc', size)
            except Exception:
                try:
                    font = ImageFont.truetype('arial.ttf', size)
                except Exception:
                    font = ImageFont.load_default()
            draw.text((x, y), shape['text'], fill=color, font=font)
        elif kind == 'mosaic':
            x1, y1, x2, y2 = shape['image_rect']
            block = max(4, int(shape.get('mosaic', 12)))
            crop = image.crop((int(x1), int(y1), int(x2), int(y2)))
            iw, ih = crop.size
            sw = max(1, iw // block)
            sh = max(1, ih // block)
            small = crop.resize((sw, sh), Image.BILINEAR)
            pixelated = small.resize((iw, ih), Image.NEAREST)
            image.paste(pixelated, (int(x1), int(y1)))

    # ----- 保存 -----
    def _save_to(self, path):
        try:
            self._work.save(path, 'PNG')
            return True
        except Exception as e:
            try:
                messagebox.showerror('保存失败', str(e), parent=self.win)
            except Exception:
                pass
            return False

    def save(self):
        """保存到默认收件目录，文件名按时间戳。"""
        if Image is None:
            return
        d = _today_dir()
        path = os.path.join(d, _gen_filename('截图_标注'))
        if self._save_to(path):
            try:
                self._on_saved(path)
            except Exception:
                pass
            try:
                self.win.destroy()
            except tk.TclError:
                pass

    def save_as(self):
        if Image is None:
            return
        path = filedialog.asksaveasfilename(
            parent=self.win, title='保存截图',
            defaultextension='.png',
            filetypes=[('PNG 图片', '*.png'),
                       ('JPEG 图片', '*.jpg;*.jpeg')],
            initialdir=_today_dir(),
            initialfile=_gen_filename('截图'),
        )
        if not path:
            return
        if self._save_to(path):
            try:
                self._on_saved(path)
            except Exception:
                pass
            try:
                self.win.destroy()
            except tk.TclError:
                pass

    def copy_to_clipboard(self):
        """复制到 Windows 剪贴板（用 PIL 自带 ImageWin 不可靠，
        走 ctypes 写 CF_DIB 更稳）。"""
        try:
            from PIL import ImageWin
            dib = ImageWin.Dib(self._work)
            dib.copy_to_clipboard()
            return True
        except Exception as e:
            traceback.print_exc()
            try:
                messagebox.showwarning('复制失败',
                                       '无法复制到剪贴板：%s' % e,
                                       parent=self.win)
            except Exception:
                pass
            return False

    def cancel(self):
        try:
            self.win.destroy()
        except tk.TclError:
            pass
        try:
            self._on_cancel()
        except Exception:
            pass


# ============================================================
# 5. 截图面板（子窗口）
# ============================================================

def _set_click_through(win, enable=True):
    """让窗口对鼠标透明——鼠标事件（含滚轮）穿透到下面的窗口。

    滚动截图时用：遮罩铺满全屏，但用户还要能正常滚动下面的页面。
    需要 WS_EX_LAYERED | WS_EX_TRANSPARENT 两个扩展样式。

    ★ 同样走**私有带类型副本**：这两个函数的句柄参数一旦被别人设过
    argtypes（声明成 int32），64 位句柄会被截断 → 样式设不上 →
    鼠标穿透失效 → "滚动截图滚不动"。
    """
    if not sys.platform.startswith('win'):
        return False
    try:
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        GWL_EXSTYLE = -20
        WS_EX_LAYERED = 0x00080000
        WS_EX_TRANSPARENT = 0x00000020
        get_long = _private_winfunc(user32, 'GetWindowLongPtrW', ctypes.c_void_p,
                                    c_void_p, ctypes.c_int)
        set_long = _private_winfunc(user32, 'SetWindowLongPtrW', c_void_p,
                                    c_void_p, ctypes.c_int, c_void_p)
        if get_long is None or set_long is None:
            # 回退到 32 位入口（32 位 Python 上只有这两个）
            get_long = _private_winfunc(user32, 'GetWindowLongW', ctypes.c_long,
                                        c_void_p, ctypes.c_int)
            set_long = _private_winfunc(user32, 'SetWindowLongW', ctypes.c_long,
                                        c_void_p, ctypes.c_int, ctypes.c_long)
        get_parent = _private_winfunc(user32, 'GetParent', c_void_p, c_void_p)
        if None in (get_long, set_long, get_parent):
            return False

        # Tk 的窗口外面可能还包了一层，扩展样式要设在包装窗口上
        handles = []
        try:
            h = int(win.winfo_id())
            guard = 0
            while h and guard < 8:
                handles.append(h)
                h = int(get_parent(c_void_p(h)) or 0)
                guard += 1
        except Exception:
            return False
        for h in handles:
            try:
                ex = int(get_long(c_void_p(h), GWL_EXSTYLE) or 0)
            except Exception:
                continue
            if enable:
                ex |= (WS_EX_LAYERED | WS_EX_TRANSPARENT)
            else:
                ex &= ~WS_EX_TRANSPARENT
            try:
                set_long(c_void_p(h), GWL_EXSTYLE, c_void_p(ex))
            except Exception:
                pass
        return True
    except Exception:
        return False


class ScreenshotPanel(object):
    """截图模块的主子窗口：四个按钮 + 截图记录 + 全局热键开关。

    提供 ``open_panel(master, on_close=None)`` 给会务助手调用。
    """

    HOTKEY_REGION = 'Ctrl+Shift+S'
    HOTKEY_FULL = 'Ctrl+Shift+F'
    HOTKEY_WINDOW = 'Ctrl+Shift+W'

    def __init__(self, master, on_close=None):
        self._master = master
        self._on_close = on_close
        self._hotkey_ids = []
        self._overlay = None
        self._record_paths = []
        self._active_mode = None

        self.win = tk.Toplevel(master)
        self.win.title('会务截图（仿 ShareX，全内置）')
        try:
            ui_theme.style_window(self.win)
        except Exception:
            pass
        self.win.transient(master)
        self.win.configure(bg=ui_theme.PALETTE['bg'])
        self.win.protocol('WM_DELETE_WINDOW', self.close)

        # 摆到主界面所在显示器的中间（多屏 / 高 DPI 下也不会跑到左上角）
        _w, _h = center_window(self.win, 820, 600, master=master)
        try:
            self.win.minsize(min(720, _w), min(520, _h))
        except Exception:
            pass

        self._build_ui()
        self._load_records()
        self._register_hotkeys()

        # 内容建完后再"亮"出来，避免系统漏画导致的空白（白屏）窗口
        present_window(self.win)

    # ---------- UI ----------
    def _build_ui(self):
        p = ui_theme.PALETTE
        body = tk.Frame(self.win, bg=p['bg'])
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=14)

        # 标题
        head = tk.Frame(body, bg=p['bg'])
        head.pack(fill=tk.X)
        tk.Label(head, text='会务截图', font=ui_theme.font(14, 'bold'),
                 bg=p['bg'], fg=p['text']).pack(side=tk.LEFT)
        tk.Label(head,
                 text='截屏、标注、保存一条龙；归档到 “会务收件/截图/日期/”',
                 font=ui_theme.font(9), bg=p['bg'],
                 fg=p['faint']).pack(side=tk.LEFT, padx=(10, 0))

        # 模式按钮
        modes = tk.Frame(body, bg=p['bg'])
        modes.pack(fill=tk.X, pady=(14, 6))
        for text, desc, cmd in (
                ('区域截图', '拖框选区，截完直接打开标注', self._capture_region),
                ('全屏截图', '截当前主屏（多屏可选）', self._capture_full),
                ('窗口截图', '列出所有可见窗口，挑一个截', self._capture_window),
                ('滚动截图', '一段一段往下拼，适合长页面', self._capture_scrolling),
        ):
            box = tk.Frame(modes, bg=p['surface'],
                           highlightthickness=1,
                           highlightbackground=p['border'])
            box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, ipady=10)
            ttk.Button(box, text=text, style='Accent.TButton',
                       command=cmd, width=14).pack(pady=(6, 2))
            tk.Label(box, text=desc, font=ui_theme.font(8),
                     bg=p['surface'], fg=p['muted'],
                     wraplength=140, justify=tk.CENTER).pack()

        # 热键提示
        hk_frame = tk.Frame(body, bg=p['bg'])
        hk_frame.pack(fill=tk.X, pady=(8, 4))
        tk.Label(hk_frame, text='全局热键（随时按，无需切到本窗口）',
                 font=ui_theme.font(9, 'bold'), bg=p['bg'],
                 fg=p['muted']).pack(side=tk.LEFT)
        for name, spec in (('区域', self.HOTKEY_REGION),
                           ('全屏', self.HOTKEY_FULL),
                           ('窗口', self.HOTKEY_WINDOW)):
            ttk.Label(hk_frame, text='%s %s' % (name, spec),
                      background=p['surface_alt'],
                      foreground=p['text'],
                      padding=(8, 2)).pack(side=tk.LEFT, padx=4)

        # 记录列表
        list_box = ttk.LabelFrame(body, text='本次会话的截图')
        list_box.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        cols = ('time', 'mode', 'size', 'path')
        self.tree = ttk.Treeview(list_box, columns=cols, show='headings',
                                 height=8)
        self.tree.heading('time', text='时间')
        self.tree.heading('mode', text='模式')
        self.tree.heading('size', text='尺寸')
        self.tree.heading('path', text='路径')
        self.tree.column('time', width=80, anchor='w')
        self.tree.column('mode', width=70, anchor='w')
        self.tree.column('size', width=90, anchor='w')
        self.tree.column('path', width=520, anchor='w')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        sb = ttk.Scrollbar(list_box, orient=tk.VERTICAL,
                           command=self.tree.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4), pady=4)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind('<Double-1>', self._on_open_record)

        # 底部按钮
        foot = tk.Frame(body, bg=p['bg'])
        foot.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(foot, text='打开归档目录', style='Ghost.TButton',
                   command=self._open_inbox).pack(side=tk.LEFT)
        ttk.Button(foot, text='刷新列表', style='Ghost.TButton',
                   command=self._load_records).pack(side=tk.LEFT, padx=6)

        # 状态
        self._status = tk.StringVar(value='就绪')
        tk.Label(foot, textvariable=self._status, font=ui_theme.font(8),
                 bg=p['bg'], fg=p['muted']).pack(side=tk.RIGHT)

    # ---------- 记录 ----------
    def _load_records(self):
        self.tree.delete(*self.tree.get_children())
        # 从 _today_dir 扫今天的截图
        d = _today_dir()
        items = []
        try:
            for name in os.listdir(d):
                if not name.lower().endswith(('.png', '.jpg', '.jpeg')):
                    continue
                full = os.path.join(d, name)
                try:
                    st = os.stat(full)
                    items.append((st.st_mtime, full))
                except Exception:
                    pass
        except Exception:
            pass
        items.sort(reverse=True)
        for mtime, full in items[:50]:
            try:
                with Image.open(full) as im:
                    size = '%d × %d' % im.size
            except Exception:
                size = '?'
            t = time.strftime('%H:%M:%S', time.localtime(mtime))
            self.tree.insert('', 'end', values=(t, '截图', size, full))

    def _add_record(self, path, mode='区域'):
        try:
            with Image.open(path) as im:
                size = '%d × %d' % im.size
        except Exception:
            size = '?'
        t = time.strftime('%H:%M:%S')
        self.tree.insert('', 0, values=(t, mode, size, path))

    def _on_open_record(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        item = self.tree.item(sel[0])
        path = item['values'][-1]
        if not path or not os.path.exists(path):
            return
        try:
            os.startfile(path)                    # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror('打开失败', str(e), parent=self.win)

    def _open_inbox(self):
        d = _today_dir()
        try:
            os.startfile(d)                        # type: ignore[attr-defined]
        except Exception:
            pass

    # ---------- 截图入口 ----------
    def _capture_full(self):
        """全屏截图：截主界面所在那块显示器。

        抓之前必须先把面板收起来——否则面板自己会被截进去（旧实现就有这个毛病）。
        """
        self._status.set('截取全屏…')
        rect = (monitor_rect_for(self.win)
                or CaptureEngine.virtual_desktop_rect())
        try:
            self.win.withdraw()
        except tk.TclError:
            pass

        def _go():
            try:
                img = CaptureEngine.region(rect)
            except Exception:
                traceback.print_exc()
                img = None
            self._restore_panel()
            if img is None:
                self._status.set('截屏失败')
                return
            self._after_capture(img, mode='全屏')

        self.win.after(220, _go)

    def _restore_panel(self):
        """把面板从"收起"状态恢复回来（区域截图前后用）。"""
        try:
            self.win.deiconify()
        except tk.TclError:
            pass
        try:
            present_window(self.win)
        except Exception:
            pass

    def _capture_region(self):
        """区域截图：先把整屏"冻"成一张图，再在图上框选。

        顺序很关键——必须是
            收面板 → 等它真的消失 → **抓图** → 在抓好的图上框选
        而不是
            收面板 → 框选 → 恢复面板 → 再抓图   ← 旧实现的顺序，
        后者会把刚恢复的面板/遮罩一起截进去，也容易框歪。
        """
        self._status.set('请拖框选区…')
        try:
            self.win.withdraw()                 # withdraw 比 iconify 更彻底
        except tk.TclError:
            pass
        try:
            self.win.update_idletasks()
        except tk.TclError:
            pass
        # 220ms 给系统把窗口真正移出屏幕的时间，然后抓图
        self.win.after(220, self._freeze_and_select)

    def _freeze_and_select(self):
        base = CaptureEngine.virtual_desktop_image()
        if base is None:
            self._restore_panel()
            self._status.set('截屏失败')
            return

        def _done(rect):
            # rect 已经是**冻结图的像素坐标**，直接裁即可，零误差
            self._restore_panel()
            if rect is None:
                self._status.set('已取消')
                return
            try:
                img = base.crop(rect)
            except Exception:
                traceback.print_exc()
                img = None
            if img is None or img.width < 2 or img.height < 2:
                self._status.set('选区太小')
                return
            self._after_capture(img, mode='区域')

        def _cancel():
            self._restore_panel()
            self._status.set('已取消')

        self._overlay = RegionOverlay(self.win, _done, _cancel, frozen=base)

    def _capture_window(self):
        try:
            wins = CaptureEngine.list_visible_windows()
        except Exception as e:
            self._status.set('枚举窗口失败：%s' % e)
            return
        if not wins:
            messagebox.showinfo('没有窗口',
                                '没找到可截的可见顶层窗口。',
                                parent=self.win)
            return
        # 弹选择窗
        pick = tk.Toplevel(self.win)
        pick.title('选择要截的窗口')
        try:
            ui_theme.style_window(pick)
        except Exception:
            pass
        pick.transient(self.win)
        center_window(pick, 520, 420, master=self.win)

        _p = ui_theme.PALETTE
        listbox = tk.Listbox(pick, font=ui_theme.font(10),
                             activestyle='dotbox',
                             bg=_p['surface'], fg=_p['text'],
                             selectbackground=_p['primary'],
                             selectforeground='#ffffff',
                             highlightthickness=1,
                             highlightbackground=_p['border'])
        listbox.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        for w in wins:
            r = w['rect']
            label = '%s  (%d × %d @ %d,%d)' % (
                w['title'][:60], r[2] - r[0], r[3] - r[1], r[0], r[1])
            listbox.insert('end', label)
        listbox.selection_set(0)
        listbox.focus_set()

        chosen = {'hwnd': None}

        def _ok(_e=None):
            sel = listbox.curselection()
            if not sel:
                return
            chosen['hwnd'] = wins[sel[0]]['hwnd']
            pick.destroy()

        def _cancel():
            pick.destroy()

        listbox.bind('<Double-1>', _ok)
        listbox.bind('<Return>', _ok)
        btn_bar = tk.Frame(pick, bg=_p['bg'])
        btn_bar.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Button(btn_bar, text='截这个窗口', style='Accent.TButton',
                   command=_ok).pack(side=tk.RIGHT)
        ttk.Button(btn_bar, text='取消', style='Ghost.TButton',
                   command=_cancel).pack(side=tk.RIGHT, padx=6)
        pick.bind('<Escape>', lambda e: _cancel())
        present_window(pick)
        pick.wait_window()
        if chosen['hwnd'] is None:
            self._status.set('已取消')
            return
        img = CaptureEngine.window(chosen['hwnd'])
        if img is None:
            self._status.set('窗口截图失败')
            return
        self._after_capture(img, mode='窗口')

    def _capture_scrolling(self):
        """滚动截图（长页面拼接）。

        交互设计（比"自动滚动"可靠得多）：
            * 遮罩铺满全屏，但设成**鼠标穿透**——用户可以正常滚轮滚页面；
            * 按 **空格 / ↓** 截一帧（截之前先把遮罩藏起来，不然遮罩会被截进去）；
            * 按 **回车** 结束，把所有帧纵向拼成一张长图；
            * 按 **ESC** 取消。
        """
        self._status.set('滚动截图：滚到位置按空格截一帧，回车结束')
        # 收面板（withdraw 比 iconify 更彻底，不会被截进去）
        scroll_rect = (monitor_rect_for(self.win)
                       or CaptureEngine.virtual_desktop_rect())
        try:
            self.win.withdraw()
        except tk.TclError:
            pass

        captured = []

        def _grab_current():
            try:
                return CaptureEngine.region(scroll_rect)
            except Exception:
                return None

        overlay = tk.Toplevel(self._master)
        overlay.overrideredirect(True)
        try:
            overlay.attributes('-topmost', True)
            # 遮罩要足够透，用户才能看清下面的页面内容
            overlay.attributes('-alpha', 0.22)
        except tk.TclError:
            pass
        virtual = CaptureEngine.virtual_desktop_rect()
        vw = virtual[2] - virtual[0]
        vh = virtual[3] - virtual[1]
        overlay.geometry('%dx%d+%d+%d' % (vw, vh, virtual[0], virtual[1]))
        hint = tk.StringVar(value=(
            '滚动截图\n\n'
            '用鼠标滚轮滚到下一个位置，按 空格 截一帧\n'
            '回车 结束并保存    ESC 取消\n\n'
            '已截 0 帧'))
        label = tk.Label(overlay, textvariable=hint, bg='#101820',
                         fg='#ffffff', font=ui_theme.font(13, 'bold'),
                         justify=tk.CENTER)
        label.pack(expand=True)
        # 鼠标穿透：滚轮事件直接给到下面的应用
        overlay.after(60, lambda: _set_click_through(overlay, True))

        busy = {'v': False}

        def _flash(color):
            try:
                label.configure(bg=color)
                overlay.after(320, lambda: label.configure(bg='#101820'))
            except tk.TclError:
                pass

        def _do_grab():
            busy['v'] = False
            img = _grab_current()
            # 先把遮罩放回来，用户才知道截图完成了
            try:
                overlay.deiconify()
                overlay.lift()
                overlay.attributes('-topmost', True)
                overlay.after(60, lambda: _set_click_through(overlay, True))
            except tk.TclError:
                return
            if img is None:
                _flash('#7c1d1d')
                hint.set('截屏失败：请确认没有开启独占全屏的应用\n\n'
                         '空格 重试   回车 结束    ESC 取消')
                return
            captured.append(img)
            hint.set('滚动截图\n\n'
                     '用鼠标滚轮滚到下一个位置，按 空格 截一帧\n'
                     '回车 结束并保存    ESC 取消\n\n'
                     '已截 %d 帧' % len(captured))
            _flash('#0e3a1f')

        def _grab_frame(_e=None):
            """截一帧：先藏遮罩 → 等它真的消失 → 再截屏。"""
            if busy['v']:
                return
            busy['v'] = True
            _flash('#1e3a8a')
            try:
                overlay.withdraw()
            except tk.TclError:
                busy['v'] = False
                return
            # 120ms 给系统把窗口真正移出屏幕的时间
            overlay.after(140, _do_grab)

        def _finish(_e=None):
            try:
                _set_click_through(overlay, False)
                overlay.destroy()
            except tk.TclError:
                pass
            self._restore_panel()
            if not captured:
                self._status.set('没有截到任何画面')
                return
            self._status.set('拼接 %d 帧…' % len(captured))
            stitched = self._stitch_vertical(captured)
            self._after_capture(stitched, mode='滚动')

        def _cancel(_e=None):
            try:
                _set_click_through(overlay, False)
                overlay.destroy()
            except tk.TclError:
                pass
            self._restore_panel()
            self._status.set('已取消')

        for seq in ('<space>', '<Down>', '<Prior>'):
            overlay.bind(seq, _grab_frame)
        overlay.bind('<Return>', _finish)
        overlay.bind('<Escape>', _cancel)
        # 让遮罩拿到键盘焦点（鼠标穿透不影响键盘）
        try:
            overlay.focus_force()
        except tk.TclError:
            pass
        # 打开后自动先截一帧当前画面
        overlay.after(260, _grab_frame)

    @staticmethod
    def _match_overlap(prev, cur):
        """估算 cur 的顶部与 prev 的底部重叠多少行（像素）。

        做法：把两帧下采样成 160px 宽的灰度图，在候选重叠高度里
        找"prev 尾部 == cur 头部"最接近的那个高度。匹配质量太差就返回 0
        （宁可不裁，也不要因为误匹配把内容切掉）。

        只在 160px 宽的小图上做，逐行差求和，开销可接受。
        """
        try:
            W = 160
            ph = max(2, int(prev.height * W / max(1, prev.width)))
            ch = max(2, int(cur.height * W / max(1, cur.width)))
            p = prev.convert('L').resize((W, ph))
            c = cur.convert('L').resize((W, ch))
            # 重叠高度按"较小帧高度的 15%~60%"搜索
            limit = min(ph, ch)
            hi = max(2, int(limit * 0.6))
            lo = max(2, int(limit * 0.15))
            if hi <= lo:
                return 0
            # 先把 prev 尾部与 cur 头部都转成 bytes，避免循环里反复 crop
            p_bytes = p.tobytes()
            c_bytes = c.tobytes()
            best_ov, best_score = 0, None
            for ov in range(lo, hi + 1):
                a = p_bytes[(ph - ov) * W:]
                b = c_bytes[:ov * W]
                n = min(len(a), len(b))
                if n <= 0:
                    continue
                # 每 4 个字节采一个样，进一步加速
                score = 0
                for i in range(0, n, 4):
                    score += abs(a[i] - b[i])
                if best_score is None or score < best_score:
                    best_score, best_ov = score, ov
            if best_ov <= 0 or best_score is None:
                return 0
            samples = (best_ov * W) // 4
            if samples <= 0:
                return 0
            avg = best_score / float(samples)
            # 平均灰度差 > 14 认为不是同一块内容，放弃裁剪
            if avg > 14:
                return 0
            # 换算回原图行数
            return max(0, int(round(best_ov * prev.height / float(ph))))
        except Exception:
            return 0

    @classmethod
    def _stitch_vertical(cls, images):
        """把多帧画面纵向拼接成长图。

        相邻帧若滚动幅度小于一屏，会有重叠区——这里用 ``_match_overlap``
        估算重叠行数并裁掉，避免长图里出现重复内容。
        匹配不可靠时退化为"整帧累加"（宁可多一段，也不丢内容）。
        """
        if not images:
            return None
        if len(images) == 1:
            return images[0].copy()

        # 宽度统一（多屏混合时可能有差异，居中留白）
        max_w = max(im.width for im in images)
        frames = []
        for im in images:
            if im.width != max_w:
                canvas = Image.new('RGB', (max_w, im.height), '#ffffff')
                canvas.paste(im, ((max_w - im.width) // 2, 0))
                frames.append(canvas)
            else:
                frames.append(im)

        # 逐帧累加，第二帧起尝试去掉与前一帧的重叠
        placed = [frames[0]]
        for cur in frames[1:]:
            overlap = cls._match_overlap(placed[-1], cur)
            if overlap > 0 and overlap < cur.height:
                placed.append(cur.crop((0, overlap, cur.width, cur.height)))
            else:
                placed.append(cur)

        total_h = sum(im.height for im in placed)
        out = Image.new('RGB', (max_w, total_h), '#ffffff')
        y = 0
        for im in placed:
            out.paste(im, (0, y))
            y += im.height
        return out

    # ---------- 标注 & 保存 ----------
    def _after_capture(self, image, mode='区域'):
        if image is None:
            self._status.set('截图失败')
            return
        self._status.set('已截到 %d × %d，可标注后保存' % image.size)
        try:
            self.win.deiconify()
            self.win.lift()
        except tk.TclError:
            pass

        def _on_saved(path):
            self._status.set('已保存：%s' % path)
            self._add_record(path, mode=mode)

        def _on_cancel():
            self._status.set('已取消')

        try:
            AnnotationEditor(self.win, image, on_saved=_on_saved,
                             on_cancel=_on_cancel)
        except Exception as e:
            traceback.print_exc()
            # 没编辑器也至少存一份
            d = _today_dir()
            path = os.path.join(d, _gen_filename('截图_%s' % mode))
            try:
                image.save(path, 'PNG')
                self._add_record(path, mode=mode)
                self._status.set('已保存：%s' % path)
            except Exception as e2:
                self._status.set('保存失败：%s' % e2)

    # ---------- 全局热键 ----------
    def _register_hotkeys(self):
        mgr = _hotkey_manager()
        for spec, cmd in ((self.HOTKEY_REGION, self._capture_region),
                          (self.HOTKEY_FULL, self._capture_full),
                          (self.HOTKEY_WINDOW, self._capture_window)):
            parsed = parse_hotkey(spec)
            if not parsed:
                continue
            modifiers, vk = parsed
            hid = mgr.register(modifiers, vk, cmd)
            if hid is not None:
                self._hotkey_ids.append(hid)
            else:
                # 占用或失败：退回 Tk bind_all（仅本进程前台时生效）
                pass

    def _unregister_hotkeys(self):
        if not self._hotkey_ids:
            return
        try:
            mgr = _hotkey_manager()
            for hid in self._hotkey_ids:
                mgr.unregister(hid)
        except Exception:
            pass
        self._hotkey_ids = []

    # ---------- 生命周期 ----------
    def close(self, force=False):
        self._unregister_hotkeys()
        try:
            self.win.destroy()
        except tk.TclError:
            pass
        if callable(self._on_close):
            try:
                self._on_close()
            except Exception:
                pass


# ============================================================
# 公共入口
# ============================================================

def open_panel(master, on_close=None):
    """给会务助手.py 调：打开截图子窗口。返回 ScreenshotPanel 实例。"""
    return ScreenshotPanel(master, on_close=on_close)
