# -*- coding: utf-8 -*-
"""Windows 系统托盘图标（任务栏右下角通知区）。

不依赖 pywin32 / pystray 等第三方库，纯 ctypes 调 Shell_NotifyIcon +
CreateWindowExW + CreatePopupMenu。

设计要点：

- 用 ``HWND_MESSAGE``（message-only window）创建隐藏窗口，跟调用线程
  的现有消息循环（Tk 的 ``mainloop`` / 普通 ``PumpWaitingMessages``）
  共享 pump——我们的 ``WndProc`` 会被调用线程的消息循环驱动，无需
  自己起线程。这是和 Tk 共存的关键。

- 右键菜单用 ``CreatePopupMenu + AppendMenuW + TrackPopupMenu``；
  菜单项点击通过 ``WM_COMMAND`` 回调。

- 双击左键 = 自定义回调（一般是"恢复窗口"）。
  右键 = 弹出菜单。

- 图标默认用系统内置 ``IDI_INFORMATION``；调用方可以传自定义
  ``LoadIconW`` 句柄（例如 ``LoadImageW`` 加载的 .ico）。

- 注意 ``NOTIFYICONDATAW`` 在不同 Windows 版本 size 不同（Win 7+
  是 ``NOTIFYICONDATAW_V3_SIZE`` = 956，新版带 ``guidItem`` 后是
  968+）。本模块按 ``ctypes.sizeof`` 自动算，不会硬编码。

本模块自身不写日志、不弹窗；任何调用失败都返回 False 并把异常
抛给调用方。
"""

import ctypes
from ctypes import wintypes
import sys


# ===================== Win32 绑定 =====================

user32 = ctypes.WinDLL('user32', use_last_error=True)
shell32 = ctypes.WinDLL('shell32', use_last_error=True)
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)


# 类型别名
UINT = wintypes.UINT
HWND = wintypes.HWND
HICON = wintypes.HICON
HMENU = wintypes.HMENU
HINSTANCE = wintypes.HINSTANCE
LRESULT = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, HWND, UINT, wintypes.WPARAM, wintypes.LPARAM)
HBRUSH = wintypes.HANDLE


# 常量
IDI_APPLICATION = 32512
IDI_INFORMATION = 32516
IDI_WARNING = 32515
IDI_ERROR = 32513

NIM_ADD = 0x00000000
NIM_MODIFY = 0x00000001
NIM_DELETE = 0x00000002
NIM_SETVERSION = 0x00000004

NIF_MESSAGE = 0x00000001
NIF_ICON = 0x00000002
NIF_TIP = 0x00000004
NIF_SHOWTIP = 0x00000080

NOTIFYICON_VERSION_4 = 4

HWND_MESSAGE = -3

WM_USER = 0x0400
WM_COMMAND = 0x0111
WM_DESTROY = 0x0002
WM_RBUTTONUP = 0x0205
WM_LBUTTONDBLCLK = 0x0203

TPM_RIGHTBUTTON = 0x0002
TPM_RETURNCMD = 0x0100
TPM_NONOTIFY = 0x0080

MF_STRING = 0x00000000
MF_SEPARATOR = 0x00008000
MF_GRAYED = 0x00000001

CS_HREDRAW = 0x0002
CS_VREDRAW = 0x0001

CW_USEDEFAULT = 0x80000000

# 用一个唯一的 message 字符串避免和其它 tray 冲突
_TRAY_MSG = 'TrayIconMsg_{:x}'.format(id(object()) & 0xFFFF)
WM_TRAY_NOTIFY = user32.RegisterWindowMessageW(_TRAY_MSG)
if not WM_TRAY_NOTIFY:
    WM_TRAY_NOTIFY = WM_USER + 1


# ===================== 结构体 =====================

class GUID(ctypes.Structure):
    _fields_ = [
        ('Data1', ctypes.c_ulong),
        ('Data2', ctypes.c_ushort),
        ('Data3', ctypes.c_ushort),
        ('Data4', ctypes.c_ubyte * 8),
    ]


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.DWORD),
        ('hWnd', wintypes.HWND),
        ('uID', wintypes.UINT),
        ('uFlags', wintypes.UINT),
        ('uCallbackMessage', wintypes.UINT),
        ('hIcon', wintypes.HICON),
        ('szTip', ctypes.c_wchar * 128),
        ('dwState', wintypes.DWORD),
        ('dwStateMask', wintypes.DWORD),
        ('szInfo', ctypes.c_wchar * 256),
        ('uVersion', wintypes.UINT),
        ('szInfoTitle', ctypes.c_wchar * 64),
        ('dwInfoFlags', wintypes.DWORD),
        ('guidItem', GUID),
        ('hBalloonIcon', wintypes.HICON),
    ]


class WNDCLASSEXW(ctypes.Structure):
    _fields_ = [
        ('cbSize', wintypes.UINT),
        ('style', wintypes.UINT),
        ('lpfnWndProc', WNDPROC),
        ('cbClsExtra', ctypes.c_int),
        ('cbWndExtra', ctypes.c_int),
        ('hInstance', wintypes.HINSTANCE),
        ('hIcon', wintypes.HICON),
        ('hCursor', wintypes.HANDLE),
        ('hbrBackground', HBRUSH),
        ('lpszMenuName', wintypes.LPCWSTR),
        ('lpszClassName', wintypes.LPCWSTR),
        ('hIconSm', wintypes.HICON),
    ]


# ===================== 函数原型 =====================

user32.CreateWindowExW.restype = HWND
user32.CreateWindowExW.argtypes = [
    wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    HWND, HMENU, HINSTANCE, wintypes.LPVOID,
]

user32.DefWindowProcW.restype = LRESULT
user32.DefWindowProcW.argtypes = [HWND, UINT, wintypes.WPARAM, wintypes.LPARAM]

user32.RegisterClassExW.restype = wintypes.ATOM
user32.RegisterClassExW.argtypes = [ctypes.POINTER(WNDCLASSEXW)]

user32.DestroyWindow.argtypes = [HWND]

user32.LoadIconW.restype = HICON
# LoadIconW 第二参数可以是字符串（注册类名）也可以是 MAKEINTRESOURCE(int)
# ——用 c_void_p 让 ctypes 不强校验，调用方传 int 或 str 都可以。
user32.LoadIconW.argtypes = [wintypes.HINSTANCE, ctypes.c_void_p]

shell32.Shell_NotifyIconW.restype = wintypes.BOOL
shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD, ctypes.POINTER(NOTIFYICONDATAW)]

user32.CreatePopupMenu.restype = HMENU
user32.CreatePopupMenu.argtypes = []

user32.AppendMenuW.restype = wintypes.BOOL
user32.AppendMenuW.argtypes = [HMENU, wintypes.UINT, wintypes.UINT, wintypes.LPCWSTR]

user32.TrackPopupMenu.restype = wintypes.BOOL
user32.TrackPopupMenu.argtypes = [HMENU, wintypes.UINT, ctypes.c_int, ctypes.c_int,
                                 ctypes.c_int, HWND, wintypes.LPVOID]

user32.SetForegroundWindow.argtypes = [HWND]
user32.DestroyMenu.argtypes = [HMENU]
user32.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
user32.PostQuitMessage.argtypes = [ctypes.c_int]

user32.RegisterWindowMessageW.restype = wintypes.UINT
user32.RegisterWindowMessageW.argtypes = [wintypes.LPCWSTR]

kernel32.GetModuleHandleW.restype = HINSTANCE
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]


# 全局 tray 实例表：菜单项点击时通过 ID 反查 TrayIcon
_REGISTRY = {}

# 防止 WndProc / 回调被 GC 回收——WINFUNCTYPE 只是 thin wrapper，必须保留引用
_WINPROC_REFS = []


# ===================== 菜单项定义 =====================

class MenuItem(object):
    """单个菜单项。``label=None`` 表示分隔线。"""
    __slots__ = ('id', 'label', 'callback', 'enabled')

    _next_id = 1000

    def __init__(self, label, callback=None, enabled=True):
        self.id = MenuItem._next_id
        MenuItem._next_id += 1
        self.label = label
        self.callback = callback
        self.enabled = enabled


# ===================== 主类 =====================

class TrayIcon(object):
    """单个系统托盘图标。

    用法::

        t = TrayIcon(tip='倒计时', icon_id=IDI_INFORMATION,
                     on_double_click=lambda: print('dbl'),
                     menu=[MenuItem('显示', show_cb),
                           MenuItem('退出', exit_cb)])
        t.show()           # 显示图标
        # ... 主程序运行 ...
        t.remove()         # 关闭程序前调用，否则图标留在托盘

    同一个进程可以创建多个 ``TrayIcon``，每个用独立 uID 区分。
    """

    _next_uid = 1

    def __init__(self, tip='', icon_id=IDI_INFORMATION, hicon=None,
                 on_double_click=None, menu=None):
        """
        :param tip:           tooltip（鼠标悬停显示，最多 127 字）
        :param icon_id:       系统图标 ID（如 ``IDI_INFORMATION``）；与 hicon 互斥
        :param hicon:         直接传 ``HICON`` 句柄（自定义 .ico）；优先于 icon_id
        :param on_double_click:  双击托盘图标回调
        :param menu:          ``[MenuItem, ...]`` 列表；None 表示无菜单
        """
        self._tip = (tip or '')[:127]
        self._uid = TrayIcon._next_uid
        TrayIcon._next_uid += 1
        self._hwnd = None
        self._wndclass_registered = False
        self._added = False
        self._on_double_click = on_double_click
        self._menu_items = list(menu or [])
        self._menu_id_to_item = {}
        self._destroyed = False

        # 解析图标
        if hicon is not None:
            self._hicon = hicon
        else:
            self._hicon = user32.LoadIconW(None, icon_id)
            if not self._hicon:
                # LoadIconW 在 IDI_* 上偶尔返回 None（取决于系统主题），改用 LoadCursor
                self._hicon = user32.LoadIconW(None, IDI_APPLICATION)

        self._hinstance = kernel32.GetModuleHandleW(None)

    # ---------- 生命周期 ----------

    def show(self):
        """注册窗口类 → 建 message-only window → ``NIM_ADD``。失败抛异常。"""
        if self._destroyed:
            raise RuntimeError('TrayIcon 已销毁，不能再次 show()')
        if self._added:
            return

        # 1. 注册窗口类（每个 TrayIcon 独立类名，避免 WndProc 串扰）
        class_name = 'TrayIconCls_{}_{:x}'.format(self._uid, id(self) & 0xFFFF)

        def _wnd_proc(hwnd, msg, wparam, lparam):
            try:
                if msg == WM_TRAY_NOTIFY:
                    # lParam: 低位是鼠标消息（WM_LBUTTONDBLCLK / WM_RBUTTONUP）
                    ev = lparam & 0xFFFF
                    if ev == WM_LBUTTONDBLCLK:
                        cb = self._on_double_click
                        if cb:
                            try:
                                cb()
                            except Exception:
                                pass
                    elif ev == WM_RBUTTONUP:
                        self._show_menu()
                    return 0
                if msg == WM_COMMAND:
                    # 菜单点击：wParam 低位是 item id
                    item_id = wparam & 0xFFFF
                    self._dispatch_menu_item(item_id)
                    return 0
                if msg == WM_DESTROY:
                    return 0
            except Exception:
                pass
            return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        # 保留 WINFUNCTYPE 引用，否则 ctypes 会 GC 掉导致崩溃
        wnd_proc = WNDPROC(_wnd_proc)
        _WINPROC_REFS.append(wnd_proc)

        wcx = WNDCLASSEXW()
        wcx.cbSize = ctypes.sizeof(WNDCLASSEXW)
        wcx.style = 0
        wcx.lpfnWndProc = wnd_proc
        wcx.cbClsExtra = 0
        wcx.cbWndExtra = 0
        wcx.hInstance = self._hinstance
        wcx.hIcon = self._hicon
        wcx.hCursor = None
        wcx.hbrBackground = None
        wcx.lpszMenuName = None
        wcx.lpszClassName = class_name
        wcx.hIconSm = None

        atom = user32.RegisterClassExW(ctypes.byref(wcx))
        if not atom:
            err = ctypes.get_last_error()
            # 0x5842 (1410) = ERROR_CLASS_ALREADY_EXISTS: 同名类已注册就复用
            if err != 1410:
                raise ctypes.WinError(err)

        # 2. 建 message-only window
        hwnd = user32.CreateWindowExW(
            0, class_name, 'TrayIconHidden',
            0, 0, 0, 0, 0,
            HWND_MESSAGE,  # parent: message-only, 不出现在桌面/任务栏
            None, self._hinstance, None)
        if not hwnd:
            err = ctypes.get_last_error()
            raise ctypes.WinError(err)

        self._hwnd = hwnd
        _REGISTRY[hwnd] = self

        # 3. NIM_ADD
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = hwnd
        nid.uID = self._uid
        nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP | NIF_SHOWTIP
        nid.uCallbackMessage = WM_TRAY_NOTIFY
        nid.hIcon = self._hicon
        nid.szTip = self._tip

        if not shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid)):
            err = ctypes.get_last_error()
            user32.DestroyWindow(hwnd)
            _REGISTRY.pop(hwnd, None)
            self._hwnd = None
            raise ctypes.WinError(err)

        # 4. 声明使用 NOTIFYICON_VERSION_4（支持现代行为：双击等）
        nid2 = NOTIFYICONDATAW()
        nid2.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid2.hWnd = hwnd
        nid2.uID = self._uid
        nid2.uVersion = NOTIFYICON_VERSION_4
        shell32.Shell_NotifyIconW(NIM_SETVERSION, ctypes.byref(nid2))

        self._added = True

    def remove(self):
        """从托盘移除图标、销毁窗口。可重复调用。"""
        if self._destroyed:
            return
        self._destroyed = True
        if self._hwnd:
            try:
                if self._added:
                    nid = NOTIFYICONDATAW()
                    nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
                    nid.hWnd = self._hwnd
                    nid.uID = self._uid
                    shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))
            except Exception:
                pass
            try:
                user32.DestroyWindow(self._hwnd)
            except Exception:
                pass
            _REGISTRY.pop(self._hwnd, None)
            self._hwnd = None
        self._added = False

    # ---------- 运行时更新 ----------

    def set_tip(self, text):
        """修改 tooltip。已添加则立即更新托盘。"""
        self._tip = (text or '')[:127]
        if not self._added or not self._hwnd:
            return
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self._hwnd
        nid.uID = self._uid
        nid.uFlags = NIF_TIP
        nid.szTip = self._tip
        shell32.Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(nid))

    def set_menu(self, menu):
        """替换右键菜单（不立即弹，只在下次右键时生效）。"""
        self._menu_items = list(menu or [])
        self._menu_id_to_item = {}

    # ---------- 内部 ----------

    def _show_menu(self):
        if not self._menu_items:
            return
        hmenu = user32.CreatePopupMenu()
        if not hmenu:
            return
        try:
            self._menu_id_to_item = {}
            for item in self._menu_items:
                if item.label is None:
                    user32.AppendMenuW(hmenu, MF_SEPARATOR, 0, None)
                    continue
                flags = MF_STRING
                if not item.enabled:
                    flags |= MF_GRAYED
                user32.AppendMenuW(hmenu, flags, item.id, item.label)
                self._menu_id_to_item[item.id] = item

            pt = wintypes.POINT()
            user32.GetCursorPos(ctypes.byref(pt))

            # 关键：TrackPopupMenu 前 SetForegroundWindow，否则菜单会立即消失
            user32.SetForegroundWindow(self._hwnd)

            cmd = user32.TrackPopupMenu(
                hmenu,
                TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY,
                pt.x, pt.y, 0, self._hwnd, None)
            if cmd:
                self._dispatch_menu_item(cmd)
        finally:
            user32.DestroyMenu(hmenu)

    def _dispatch_menu_item(self, item_id):
        item = self._menu_id_to_item.get(item_id)
        if item is None:
            return
        cb = item.callback
        if cb:
            try:
                cb()
            except Exception:
                pass

    # 上下文管理器支持：with TrayIcon(...) as t: t.show()
    def __enter__(self):
        self.show()
        return self

    def __exit__(self, *exc):
        self.remove()
        return False


# ===================== 自检 =====================

if __name__ == '__main__':
    # 简单自检：建图标 + 菜单 + 消息循环
    import time

    clicks = []

    def on_show():
        clicks.append('show')
        print('[回调] 显示倒计时')

    def on_exit():
        clicks.append('exit')
        print('[回调] 关闭计时器')
        user32.PostQuitMessage(0)

    tray = TrayIcon(
        tip='会务助手 - 倒计时',
        icon_id=IDI_INFORMATION,
        on_double_click=on_show,
        menu=[
            MenuItem('显示倒计时', on_show),
            MenuItem(None),
            MenuItem('关闭计时器', on_exit),
        ],
    )
    try:
        tray.show()
        print('托盘图标已添加；右键弹出菜单 / 双击触发回调')
        print('将运行 8 秒后自动退出')

        # 简易消息循环（不依赖 Tk）
        import ctypes.wintypes as wt

        msg = wt.MSG()
        t_end = time.time() + 8.0
        while time.time() < t_end:
            ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if ret <= 0:
                break
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
    finally:
        tray.remove()
        print('已清理；clicks =', clicks)