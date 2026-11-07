# -*- coding: utf-8 -*-
"""会务视频播放：用系统自带的 Windows Media Player 组件按顺序全屏播放。

支持三种播放后端，自动挑最合适的：

1. VLC —— 解码最全（MP4/MKV/AVI/WMV/RMVB/FLV…），
   用 --qt-fullscreen-screennumber 直接把全屏放到指定显示器
2. Edge / Chrome —— 系统必有，HTML5 解码 MP4(H.264)/WebM/MOV 很稳
3. Windows Media Player 组件 —— 兜底

三种后端都能“按顺序播完整个列表、循环、手动停止”。
"""

import ctypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from ctypes import wintypes

user32 = ctypes.windll.user32
WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND,
                                 wintypes.LPARAM)
MONITORENUMPROC = ctypes.WINFUNCTYPE(
    ctypes.c_bool, wintypes.HMONITOR, wintypes.HDC,
    ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)

SWP_NOZORDER = 0x0004
SWP_SHOWWINDOW = 0x0040
SWP_FRAMECHANGED = 0x0020
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
GWL_STYLE = -16
WS_CHILD = 0x40000000
WS_POPUP = 0x80000000
WS_CAPTION = 0x00C00000
WS_THICKFRAME = 0x00040000
WS_SYSMENU = 0x00080000

VIDEO_EXTS = ('.mp4', '.wmv', '.avi', '.mov', '.mkv', '.flv', '.m4v',
              '.mpg', '.mpeg', '.ts', '.webm', '.rmvb', '.3gp')

# Windows 播放设备（音箱/耳机/投影仪）信息在注册表里的位置
_AUDIO_BASE = r'SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
_AUDIO_NAME_KEY = '{a45c254e-df1c-4efd-8020-67d146a850e0},2'
_AUDIO_ADAPTER_KEY = '{b3f8fa53-0004-438e-9003-51a46e139bfc},6'
_NO_WINDOW = 0x08000000          # CREATE_NO_WINDOW：别闪黑框


class MONITORINFO(ctypes.Structure):
    _fields_ = [('cbSize', wintypes.DWORD),
                ('rcMonitor', wintypes.RECT),
                ('rcWork', wintypes.RECT),
                ('dwFlags', wintypes.DWORD)]


def list_monitors():
    """列出所有显示器：[{index, name, rect, primary}]。"""
    found = []

    def _cb(hmon, _hdc, _rect, _param):
        info = MONITORINFO()
        info.cbSize = ctypes.sizeof(MONITORINFO)
        if user32.GetMonitorInfoW(hmon, ctypes.byref(info)):
            r = info.rcMonitor
            found.append({'rect': (r.left, r.top, r.right, r.bottom),
                          'primary': bool(info.dwFlags & 1)})
        return True

    try:
        user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(_cb), 0)
    except Exception:
        pass
    if not found:
        found.append({'rect': (0, 0, user32.GetSystemMetrics(0),
                               user32.GetSystemMetrics(1)),
                      'primary': True})
    for i, item in enumerate(found):
        r = item['rect']
        item['index'] = i
        item['name'] = '显示器 %d（%d×%d）%s' % (
            i + 1, r[2] - r[0], r[3] - r[1],
            '  主屏' if item['primary'] else '  扩展屏')
    return found


def resolve_monitor_index(name, monitors=None):
    """把下拉框里记的显示器名字对到“当前”显示器列表的下标。

    名字里带分辨率，插拔投影仪之后名字就变了；这时按原来的意思
    （主屏 / 扩展屏）重新挑一台，免得明明选了扩展屏、画面却回到主屏。
    """
    items = monitors if monitors is not None else list_monitors()
    if not items:
        return 0
    name = (name or '').strip()
    tight = ''.join(name.split())
    for item in items:
        if item['name'] == name or ''.join(item['name'].split()) == tight:
            return item['index']
    want_extended = '扩展屏' in name
    for item in items:
        if bool(item['primary']) != want_extended:
            return item['index']
    return items[0]['index']


def find_video_window(pid=None):
    """找 VLC 的视频输出窗口（无界面模式下标题是 VLC video main）。"""
    result = []

    def _cb(hwnd, _param):
        cls = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls, 256)
        if cls.value.lower().startswith('vlc video'):
            if pid is not None:
                owner = wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
                if owner.value != pid:
                    return True
            result.append(hwnd)
        return True

    try:
        user32.EnumWindows(WNDENUMPROC(_cb), 0)
    except Exception:
        pass
    return result[0] if result else None


def window_rect(hwnd):
    """窗口当前矩形；取不到返回 None。"""
    if not hwnd:
        return None
    rect = wintypes.RECT()
    try:
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return None
    except Exception:
        return None
    return (rect.left, rect.top, rect.right, rect.bottom)


GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_CAPTION = 0x00C00000
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000


def _get_long(hwnd, index):
    try:
        func = getattr(user32, 'GetWindowLongPtrW', user32.GetWindowLongW)
        func.restype = ctypes.c_ssize_t
        func.argtypes = [ctypes.c_void_p, ctypes.c_int]
        return int(func(ctypes.c_void_p(hwnd), index))
    except Exception:
        return 0


def _set_long(hwnd, index, value):
    try:
        func = getattr(user32, 'SetWindowLongPtrW', user32.SetWindowLongW)
        func.restype = ctypes.c_ssize_t
        func.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_ssize_t]
        func(ctypes.c_void_p(hwnd), index, ctypes.c_ssize_t(value))
        return True
    except Exception:
        return False


def hide_window_chrome(hwnd):
    """把播放窗口变成“工具窗口”：任务栏不出现、标题栏和标题文字都去掉。

    会场里不能露出播放器的操作栏或任务栏条目，播之前对视频窗口处理一次即可。
    """
    if not hwnd:
        return False
    ok = False
    try:
        style = _get_long(hwnd, GWL_STYLE)
        style &= ~WS_CAPTION            # 去掉标题栏
        ok = _set_long(hwnd, GWL_STYLE, style) or ok
    except Exception:
        pass
    try:
        ex = _get_long(hwnd, GWL_EXSTYLE)
        ex = (ex | WS_EX_TOOLWINDOW) & ~WS_EX_APPWINDOW   # 不占任务栏/Alt+Tab
        ok = _set_long(hwnd, GWL_EXSTYLE, ex) or ok
    except Exception:
        pass
    try:
        user32.SetWindowTextW(ctypes.c_void_p(hwnd), '')
    except Exception:
        pass
    return ok


def window_visible(hwnd):
    if not hwnd:
        return False
    try:
        return bool(user32.IsWindowVisible(ctypes.c_void_p(hwnd)))
    except Exception:
        return False


def preview_rect(ratio=0.34, margin=40):
    """主屏上预览窗口的位置：默认主屏右上角，按 16:9 算大小。"""
    monitors = list_monitors()
    primary = None
    for item in monitors:
        if item['primary']:
            primary = item['rect']
    if primary is None:
        primary = (0, 0, 1024, 768)
    left, top, right, bottom = primary
    width = max(1, right - left)
    height = max(1, bottom - top)
    vid_w = max(320, int(width * ratio))
    vid_h = int(vid_w * 9 / 16.0)
    if vid_h > height * 0.55:              # 屏幕太小就再缩一点
        vid_h = int(height * 0.55)
        vid_w = int(vid_h * 16 / 9.0)
    x = max(left, right - vid_w - margin)
    y = max(top, top + margin)
    return (x, y, x + vid_w, y + vid_h)


def move_window(hwnd, rect, topmost=False):
    """把窗口搬到指定矩形（rect = left, top, right, bottom）。"""
    if not hwnd:
        return False
    left, top, right, bottom = rect
    after = -1 if topmost else 0          # HWND_TOPMOST / HWND_NOTOPMOST
    flags = SWP_SHOWWINDOW | (0 if topmost else SWP_NOZORDER)
    return bool(user32.SetWindowPos(ctypes.c_void_p(hwnd), ctypes.c_void_p(after),
                                    int(left), int(top),
                                    int(right - left), int(bottom - top),
                                    flags))


def next_index(count, current, loop=True):
    """下一个要播的下标；放完了且不循环时返回 None。

    这是“依次播放直到手动停止”的核心逻辑，单独抽出来便于测试。
    """
    if count <= 0:
        return None
    if current < 0:
        return 0
    if current + 1 < count:
        return current + 1
    return 0 if loop else None


def find_vlc():
    """找 VLC 可执行文件。

    顺序：程序目录里自带的 vlc 文件夹（绿色随包，免安装）→ 系统安装的
    VLC（Program Files / 注册表里记的安装位置）。找不到返回 None。
    """
    for base in _app_dirs():
        for rel in (os.path.join('vlc', 'vlc.exe'),
                    os.path.join('VLC', 'vlc.exe'),
                    'vlc.exe'):
            path = os.path.join(base, rel)
            if os.path.isfile(path):
                return path
    roots = [os.environ.get('ProgramFiles', r'C:\Program Files'),
             os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)')]
    for root in roots:
        if not root:
            continue
        path = os.path.join(root, 'VideoLAN', 'VLC', 'vlc.exe')
        if os.path.isfile(path):
            return path
    try:
        import winreg
        keys = (r'SOFTWARE\VideoLAN\VLC',
                r'SOFTWARE\WOW6432Node\VideoLAN\VLC')
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for key in keys:
                try:
                    with winreg.OpenKey(hive, key) as handle:
                        folder = str(winreg.QueryValueEx(handle,
                                                          'InstallDir')[0])
                except OSError:
                    continue
                path = os.path.join(folder, 'vlc.exe')
                if os.path.isfile(path):
                    return path
    except Exception:
        pass
    return None


def _reg_str(key, name):
    try:
        import winreg
        return str(winreg.QueryValueEx(key, name)[0])
    except Exception:
        return ''


# ---------- libvlc：把预览画面直接画进软件窗口 ----------
_libvlc = None
_libvlc_error = ''


def find_vlc_dir():
    """VLC 安装目录（libvlc.dll 所在的地方）。"""
    exe = find_vlc()
    return os.path.dirname(exe) if exe else None


def reset_vlc_cache():
    """装好 VLC 之后清掉缓存，让程序重新去找 libvlc。"""
    global _libvlc, _libvlc_error
    _libvlc = None
    _libvlc_error = ''


def _app_dirs():
    """程序自己所在的目录（打包后是 exe 所在目录），用来找随包的 VLC。"""
    dirs = []
    if getattr(sys, 'frozen', False):
        dirs.append(os.path.dirname(os.path.abspath(sys.executable)))
    else:
        dirs.append(os.path.dirname(os.path.abspath(__file__)))
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass and meipass not in dirs:
        dirs.append(meipass)
    return dirs


def find_bundled_vlc_installer():
    """程序目录（及 soft 子目录）里如果放了 VLC 安装包
    （vlc-...-win64.exe 之类），返回它的路径。"""
    bases = []
    for base in _app_dirs():
        bases.append(base)
        soft = os.path.join(base, 'soft')
        if soft not in bases:
            bases.append(soft)
    for base in bases:
        try:
            names = sorted(os.listdir(base))
        except OSError:
            continue
        for name in names:
            low = name.lower()
            if not low.endswith('.exe') or 'vlc' not in low:
                continue
            if low.startswith('vlc') or 'install' in low or 'setup' in low:
                path = os.path.join(base, name)
                if os.path.isfile(path):
                    return path
    return None


_install_lock = threading.Lock()


def install_vlc():
    """装 VLC：优先跑程序目录里自带的安装包（静默），否则用系统 winget。

    返回 (ok, 说明)。这一步会等安装完成，界面里请放到后台线程里调。
    已有安装任务在跑时直接返回，避免两个安装进程打架。
    """
    if not _install_lock.acquire(blocking=False):
        return False, 'VLC 正在后台安装，请稍候…'
    try:
        return _install_vlc_locked()
    finally:
        _install_lock.release()


def _install_vlc_locked():
    exe = find_bundled_vlc_installer()
    if exe:
        try:
            proc = subprocess.Popen([exe, '/S', '/L=1033'], close_fds=True)
            proc.wait(timeout=900)
        except Exception as exc:
            return False, '运行自带安装包失败：%s' % exc
        if find_vlc():
            return True, '已用随包安装包装好 VLC'
        return False, '安装包已运行，但没检测到 VLC'
    winget = shutil.which('winget')
    if winget:
        try:
            proc = subprocess.Popen(
                [winget, 'install', '-e', '--id', 'VideoLAN.VLC', '--silent',
                 '--accept-package-agreements', '--accept-source-agreements'],
                close_fds=True)
            proc.wait(timeout=1800)
        except Exception as exc:
            return False, 'winget 安装失败：%s' % exc
        if find_vlc():
            return True, '已用 winget 装好 VLC'
        return False, 'winget 已运行，但没检测到 VLC'
    return False, ('没找到随包安装包，这台电脑也没有 winget。\n'
                   '可以手动装一个 VLC（www.videolan.org），'
                   '或把 VLC 整个文件夹复制成程序目录下的 vlc 子文件夹。')


def _vlc_lib():
    """加载 VLC 自带的 libvlc.dll；不可用时返回 None。"""
    global _libvlc, _libvlc_error
    if _libvlc is not None or _libvlc_error:
        return _libvlc
    folder = find_vlc_dir()
    if not folder:
        _libvlc_error = '没找到 VLC'
        return None
    dll = os.path.join(folder, 'libvlc.dll')
    if not os.path.isfile(dll):
        _libvlc_error = '没找到 libvlc.dll'
        return None
    try:
        if hasattr(os, 'add_dll_directory'):
            os.add_dll_directory(folder)
        lib = ctypes.CDLL(dll)
    except OSError as exc:
        _libvlc_error = 'libvlc 加载失败：%s' % exc
        return None
    lib.libvlc_new.restype = ctypes.c_void_p
    lib.libvlc_new.argtypes = [ctypes.c_int, ctypes.POINTER(ctypes.c_char_p)]
    lib.libvlc_release.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_new_path.restype = ctypes.c_void_p
    lib.libvlc_media_new_path.argtypes = [ctypes.c_void_p, ctypes.c_char_p]
    lib.libvlc_media_release.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_new.restype = ctypes.c_void_p
    lib.libvlc_media_player_new.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_set_media.argtypes = [ctypes.c_void_p,
                                                  ctypes.c_void_p]
    lib.libvlc_media_player_set_hwnd.argtypes = [ctypes.c_void_p,
                                                 ctypes.c_void_p]
    lib.libvlc_media_player_play.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_stop.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_release.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_get_state.restype = ctypes.c_int
    lib.libvlc_media_player_get_state.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_pause.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_set_pause.argtypes = [ctypes.c_void_p,
                                                  ctypes.c_int]
    lib.libvlc_media_player_is_playing.restype = ctypes.c_int
    lib.libvlc_media_player_is_playing.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_get_time.restype = ctypes.c_longlong
    lib.libvlc_media_player_get_time.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_get_length.restype = ctypes.c_longlong
    lib.libvlc_media_player_get_length.argtypes = [ctypes.c_void_p]
    lib.libvlc_media_player_set_time.restype = ctypes.c_int
    lib.libvlc_media_player_set_time.argtypes = [ctypes.c_void_p,
                                                 ctypes.c_longlong]
    lib.libvlc_audio_set_volume.argtypes = [ctypes.c_void_p, ctypes.c_int]
    lib.libvlc_audio_set_mute.argtypes = [ctypes.c_void_p, ctypes.c_int]
    lib.libvlc_video_take_snapshot.argtypes = [
        ctypes.c_void_p, ctypes.c_uint, ctypes.c_char_p,
        ctypes.c_uint, ctypes.c_uint]
    lib.libvlc_video_take_snapshot.restype = ctypes.c_int
    _libvlc = lib
    return lib


def grab_frame(video_path, out_png, host_hwnd=None, timeout=10.0):
    """静音放一下这个视频，抓一帧存成 PNG（给“版式预览”当示意图用）。

    画面渲染在 host_hwnd 上（调用方一般给它一个屏幕外的小窗口），
    抓完就停掉，用户看不到、也听不到。
    """
    lib = _vlc_lib()
    if lib is None or not video_path or not os.path.isfile(video_path):
        return False
    try:
        if os.path.exists(out_png):
            os.remove(out_png)
    except OSError:
        pass
    args = [b'--no-audio', b'--no-video-title-show', b'--ignore-config',
            b'--quiet', b'--no-osd']
    argv = (ctypes.c_char_p * len(args))(*args)
    inst = lib.libvlc_new(len(args), argv)
    if not inst:
        return False
    player = lib.libvlc_media_player_new(ctypes.c_void_p(inst))
    ok = False
    try:
        if host_hwnd:
            lib.libvlc_media_player_set_hwnd(ctypes.c_void_p(player),
                                             ctypes.c_void_p(int(host_hwnd)))
        for name in ('libvlc_video_set_mouse_input',
                     'libvlc_video_set_key_input'):
            func = getattr(lib, name, None)
            if func is not None:
                try:
                    func(ctypes.c_void_p(player), 0)
                except Exception:
                    pass
        media = lib.libvlc_media_new_path(ctypes.c_void_p(inst),
                                          video_path.encode('utf-8'))
        if not media:
            return False
        lib.libvlc_media_player_set_media(ctypes.c_void_p(player),
                                          ctypes.c_void_p(media))
        lib.libvlc_media_release(ctypes.c_void_p(media))
        lib.libvlc_media_player_play(ctypes.c_void_p(player))
        deadline = time.time() + timeout
        while time.time() < deadline:
            state = int(lib.libvlc_media_player_get_state(
                ctypes.c_void_p(player)))
            if state == 3:                      # Playing：有画面了，抓一帧
                time.sleep(0.5)
                res = lib.libvlc_video_take_snapshot(
                    ctypes.c_void_p(player), 0, out_png.encode('utf-8'), 0, 0)
                if res == 0:
                    for _ in range(30):
                        if os.path.exists(out_png):
                            ok = True
                            break
                        time.sleep(0.1)
                break
            if state in (6, 7):                 # Ended / Error
                break
            time.sleep(0.2)
    except Exception:
        ok = False
    finally:
        try:
            lib.libvlc_media_player_stop(ctypes.c_void_p(player))
            lib.libvlc_media_player_release(ctypes.c_void_p(player))
        except Exception:
            pass
        try:
            lib.libvlc_release(ctypes.c_void_p(inst))
        except Exception:
            pass
    return ok


class EmbeddedPreview(object):
    """软件内预览：用 libvlc 把静音画面直接画进我们自己的窗口里。

    正片还是独立的 VLC 进程，全屏带声音投到扩展屏；这里借同一个 VLC 的
    库在软件窗口内再画一份静音画面，不额外开窗、也不会挡别的程序。
    """

    ENDED = 6
    PAUSED = 4

    def __init__(self):
        self._lib = None
        self._inst = None
        self._player = None
        self._files = []
        self._idx = 0
        self._loop = True
        self._hwnd = None
        self._host_hwnd = None
        self.last_error = ''

    @property
    def available(self):
        return _vlc_lib() is not None

    @property
    def running(self):
        return self._player is not None

    def describe(self):
        if self.available:
            return '软件内预览可用'
        return '预览不可用：%s' % (_libvlc_error or '未知原因')

    def state(self):
        if self._lib is None or self._player is None:
            return 0
        return int(self._lib.libvlc_media_player_get_state(
            ctypes.c_void_p(self._player)))

    def start(self, hwnd, playlist, index=0, loop=True):
        """把列表交给预览：画面画进 hwnd 这块窗口，静音播放。"""
        self.stop()
        lib = _vlc_lib()
        if lib is None:
            self.last_error = _libvlc_error
            return False
        files = [os.path.abspath(p) for p in playlist]
        if not files:
            self.last_error = '播放列表是空的'
            return False
        args = [b'--no-audio', b'--no-video-title-show', b'--ignore-config',
                b'--no-osd', b'--quiet']
        argv = (ctypes.c_char_p * len(args))(*args)
        inst = lib.libvlc_new(len(args), argv)
        if not inst:
            self.last_error = 'VLC 初始化失败'
            return False
        player = lib.libvlc_media_player_new(ctypes.c_void_p(inst))
        if not player:
            lib.libvlc_release(ctypes.c_void_p(inst))
            self.last_error = '创建预览播放器失败'
            return False
        self._lib, self._inst, self._player = lib, inst, player
        self._files = files
        self._loop = bool(loop)
        self._hwnd = int(hwnd)
        self._host_hwnd = int(hwnd)
        lib.libvlc_media_player_set_hwnd(ctypes.c_void_p(player),
                                         ctypes.c_void_p(self._hwnd))
        # 画面上的鼠标/键盘交给软件自己处理，VLC 不要抢
        for name in ('libvlc_video_set_mouse_input', 'libvlc_video_set_key_input'):
            func = getattr(lib, name, None)
            if func is not None:
                try:
                    func(ctypes.c_void_p(player), 0)
                except Exception:
                    pass
        return self._play(max(0, min(int(index), len(files) - 1)))

    def _play(self, index):
        lib = self._lib
        if lib is None:
            return False
        media = lib.libvlc_media_new_path(
            ctypes.c_void_p(self._inst), self._files[index].encode('utf-8'))
        if not media:
            self.last_error = '打不开文件：%s' % self._files[index]
            return False
        lib.libvlc_media_player_set_media(ctypes.c_void_p(self._player),
                                          ctypes.c_void_p(media))
        lib.libvlc_media_release(ctypes.c_void_p(media))
        self._idx = index
        lib.libvlc_media_player_play(ctypes.c_void_p(self._player))
        return True

    def tick(self):
        """播完一个自动接下个（界面定时调用即可）。"""
        if self._player is None:
            return
        if self.state() == self.ENDED:
            nxt = self._idx + 1
            if nxt >= len(self._files):
                if not self._loop:
                    return
                nxt = 0
            self._play(nxt)

    def set_pause(self, paused):
        """暂停/继续（软件内预览跟着主画面一起暂停）。"""
        if self._lib is None or self._player is None:
            return
        try:
            self._lib.libvlc_media_player_set_pause(
                ctypes.c_void_p(self._player), 1 if paused else 0)
        except Exception:
            pass

    def pause(self):
        self.set_pause(True)

    def resume(self):
        self.set_pause(False)

    @property
    def paused(self):
        if self._lib is None or self._player is None:
            return False
        try:
            return int(self._lib.libvlc_media_player_get_state(
                ctypes.c_void_p(self._player))) == self.PAUSED
        except Exception:
            return False

    def stop(self):
        lib, player, inst = self._lib, self._player, self._inst
        self._lib = self._inst = self._player = None
        self._hwnd = None
        if lib is None:
            return
        try:
            if player:
                lib.libvlc_media_player_stop(ctypes.c_void_p(player))
                lib.libvlc_media_player_release(ctypes.c_void_p(player))
        except Exception:
            pass
        try:
            if inst:
                lib.libvlc_release(ctypes.c_void_p(inst))
        except Exception:
            pass


def list_audio_devices():
    """列出可用的声音输出设备：[{id, name}]。

    只列“正在使用”的设备（插着的音箱/耳机/投影仪），按系统里的名字显示。
    id 是 Windows 的 MMDevice 标识，VLC 的 --mmdevice-audio-device 要的就是它。
    """
    out = []
    try:
        import winreg
    except ImportError:
        return out
    try:
        root = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, _AUDIO_BASE)
    except OSError:
        return out
    index = 0
    while True:
        try:
            guid = winreg.EnumKey(root, index)
        except OSError:
            break
        index += 1
        try:
            device = winreg.OpenKey(root, guid)
            if winreg.QueryValueEx(device, 'DeviceState')[0] != 1:
                continue            # 1=正在使用，其余是拔出/禁用/不存在
            props = winreg.OpenKey(device, 'Properties')
            name = _reg_str(props, _AUDIO_NAME_KEY) or '播放设备'
            adapter = _reg_str(props, _AUDIO_ADAPTER_KEY)
        except OSError:
            continue
        if adapter and adapter.lower() not in name.lower():
            name = '%s（%s）' % (name, adapter)
        out.append({'id': '{0.0.0.00000000}.%s' % guid, 'name': name})
    return out


def default_audio_device_id():
    """系统当前默认播放设备的 id；查不到返回 None（不影响播放）。"""
    if os.name != 'nt':
        return None
    script = (
        "[Windows.Media.Devices.MediaDevice,Windows.Media.Devices,"
        "ContentType=WindowsRuntime]|Out-Null;"
        "[Windows.Media.Devices.MediaDevice]::GetDefaultAudioRenderId("
        "[Windows.Media.Devices.AudioDeviceRole]::Default)")
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-Command', script],
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
            creationflags=_NO_WINDOW, timeout=8)
    except Exception:
        return None
    text = (result.stdout or b'').decode('utf-8', 'replace')
    match = re.search(r'\{[^#{}]+\}\.\{[0-9A-Fa-f-]{36}\}', text)
    return match.group(0) if match else None


_LAST_TONE = []


def play_test_tone(audio_device=None, volume=100):
    """用一小段提示音试听指定的声音输出设备。返回 (ok, 说明)。"""
    import math
    import struct
    import wave
    path = os.path.join(tempfile.gettempdir(), 'meeting_tone_%s.wav'
                        % uuid.uuid4().hex[:8])
    try:
        rate = 44100
        frames = []
        for i in range(int(rate * 1.2)):
            # 首尾各留一点淡入淡出，避免爆音
            fade = min(1.0, i / (rate * 0.05), (rate * 1.2 - i) / (rate * 0.15))
            value = int(0.7 * fade * 32767 * math.sin(
                2 * math.pi * 880.0 * i / rate))
            frames.append(struct.pack('<hh', value, value))
        with wave.open(path, 'wb') as handle:
            handle.setnchannels(2)
            handle.setsampwidth(2)
            handle.setframerate(rate)
            handle.writeframes(b''.join(frames))
    except Exception as exc:
        return False, '生成试听音失败：%s' % exc
    # 上一段试听音用完就删，免得临时目录越堆越多
    while _LAST_TONE:
        old = _LAST_TONE.pop()
        try:
            os.remove(old)
        except OSError:
            pass
    _LAST_TONE.append(path)

    exe = find_vlc()
    gain = max(0.0, min(int(volume), 100)) / 100.0
    if exe:
        args = [exe, '--intf', 'dummy', '--vout=dummy',
                '--no-video-title-show', '--no-qt-privacy-ask',
                '--mmdevice-volume=%.3f' % gain,
                '--waveout-volume=%.3f' % gain,
                '--play-and-exit']
        if audio_device:
            args.append('--mmdevice-audio-device=%s' % audio_device)
        args.append(path)
        try:
            subprocess.Popen(args, close_fds=True)
        except Exception as exc:
            return False, '试听失败：%s' % exc
        return True, '正在试听：%s' % ('默认设备' if not audio_device else '所选设备')
    try:
        import winsound
        winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
    except Exception as exc:
        return False, '试听失败：%s' % exc
    return True, '正在试听（系统默认设备）'


def _to_file_url(path):
    from urllib.parse import quote
    return 'file:///' + quote(os.path.abspath(path).replace('\\', '/'),
                              safe=':/')


class LibVlcPlayer(object):
    """VLC 内嵌播放（推荐）：用 VLC 官方的 libvlc 直接把画面渲染进我们自己的窗口。

    这是 VLC 官方在 .NET/ActiveX/各路开源项目里用的同一条路：
    播放器不再自己开窗、不会挪来挪去，也不会有标题栏/任务栏/操作栏；
    窗口位置、大小、置顶全部由会务助手说了算。
    """

    name = 'VLC内嵌'
    supports_pause = True
    supports_seek = True     # 可以直接跳到任意位置（拖进度条）
    embedded = True          # 画面直接画进我们给的窗口句柄

    NOTHING, OPENING, BUFFERING, PLAYING, PAUSED, STOPPED, ENDED, ERROR = range(8)

    def __init__(self):
        self._lib = None
        self._inst = None
        self._player = None
        self._files = []
        self._idx = -1
        self._loop = True
        self._volume = 100
        self.last_error = ''

    # ---------- 基本信息 ----------
    @property
    def available(self):
        return _vlc_lib() is not None

    def describe(self):
        if self.available:
            return 'VLC 内嵌播放（画面固定，推荐）'
        return 'VLC 不可用：%s' % (_libvlc_error or '未知原因')

    def current_name(self):
        if 0 <= self._idx < len(self._files):
            return os.path.basename(self._files[self._idx])
        return ''

    # ---------- 播放 ----------
    def play(self, playlist, index=0, monitor_index=0, loop=True, volume=100,
             topmost=True, audio_device=None, window_rect=None,
             host_hwnd=None):
        self.stop()
        lib = _vlc_lib()
        if lib is None:
            self.last_error = _libvlc_error or 'VLC 不可用'
            return False, self.last_error
        files = [os.path.abspath(p) for p in playlist]
        if not files:
            return False, '播放列表是空的'
        args = [b'--no-video-title-show', b'--ignore-config', b'--no-osd',
                b'--quiet']
        if audio_device:
            # 指定用哪台声音设备出声（和之前“声音输出”下拉一致）
            args.append(('--mmdevice-audio-device=%s'
                         % audio_device).encode('utf-8'))
        argv = (ctypes.c_char_p * len(args))(*args)
        inst = lib.libvlc_new(len(args), argv)
        if not inst:
            self.last_error = 'VLC 初始化失败'
            return False, self.last_error
        player = lib.libvlc_media_player_new(ctypes.c_void_p(inst))
        if not player:
            lib.libvlc_release(ctypes.c_void_p(inst))
            self.last_error = '创建播放器失败'
            return False, self.last_error
        self._lib, self._inst, self._player = lib, inst, player
        self._files = files
        self._loop = bool(loop)
        self._volume = max(0, min(int(volume), 100))
        self._idx = max(0, min(int(index), len(files) - 1))
        if host_hwnd:
            lib.libvlc_media_player_set_hwnd(ctypes.c_void_p(player),
                                             ctypes.c_void_p(int(host_hwnd)))
        for name in ('libvlc_video_set_mouse_input',
                     'libvlc_video_set_key_input'):
            func = getattr(lib, name, None)
            if func is not None:
                try:
                    func(ctypes.c_void_p(player), 0)
                except Exception:
                    pass
        self.set_volume(self._volume)
        if not self._play_index(self._idx):
            return False, self.last_error or '打不开媒体'
        return True, '正在播放 %d/%d：%s' % (
            self._idx + 1, len(files), self.current_name())

    def _play_index(self, index):
        lib = self._lib
        if lib is None or not (0 <= index < len(self._files)):
            return False
        media = lib.libvlc_media_new_path(
            ctypes.c_void_p(self._inst), self._files[index].encode('utf-8'))
        if not media:
            self.last_error = '打不开文件：%s' % self._files[index]
            return False
        lib.libvlc_media_player_set_media(ctypes.c_void_p(self._player),
                                          ctypes.c_void_p(media))
        lib.libvlc_media_release(ctypes.c_void_p(media))
        self._idx = index
        lib.libvlc_media_player_play(ctypes.c_void_p(self._player))
        return True

    def _advance(self):
        """一个播完自动接下个；到末尾按“循环”决定。"""
        nxt = self._idx + 1
        if nxt >= len(self._files):
            if not self._loop:
                return False
            nxt = 0
        return self._play_index(nxt)

    def tick(self):
        """由界面定时调用：播完自动接着放。"""
        self.state()

    @property
    def paused(self):
        """当前是不是暂停着（界面上“暂停/继续”按钮靠这个判断）。"""
        return self._raw_state() == self.PAUSED

    def stop(self):
        lib, player, inst = self._lib, self._player, self._inst
        self._lib = self._inst = self._player = None
        self._idx = -1
        if lib is None:
            return
        try:
            if player:
                lib.libvlc_media_player_stop(ctypes.c_void_p(player))
                lib.libvlc_media_player_release(ctypes.c_void_p(player))
        except Exception:
            pass
        try:
            if inst:
                lib.libvlc_release(ctypes.c_void_p(inst))
        except Exception:
            pass

    def close(self):
        self.stop()

    # ---------- 状态 / 控制 ----------
    def _raw_state(self):
        if self._lib is None or self._player is None:
            return None
        try:
            return int(self._lib.libvlc_media_player_get_state(
                ctypes.c_void_p(self._player)))
        except Exception:
            return None

    def state(self):
        """1=已结束/已停止，3=播放中（和界面里的约定一致）。"""
        raw = self._raw_state()
        if raw is None:
            return 1
        if raw == self.ENDED:
            return 3 if self._advance() else 1
        if raw in (self.PLAYING, self.PAUSED, self.OPENING, self.BUFFERING):
            return 3
        return 1

    def duration(self):
        if self._lib is None or self._player is None:
            return 0.0
        try:
            length = int(self._lib.libvlc_media_player_get_length(
                ctypes.c_void_p(self._player)))
        except Exception:
            return 0.0
        return max(0.0, length / 1000.0)

    def position(self):
        if self._lib is None or self._player is None:
            return 0.0
        try:
            ms = int(self._lib.libvlc_media_player_get_time(
                ctypes.c_void_p(self._player)))
        except Exception:
            return 0.0
        return max(0.0, ms / 1000.0)

    def seek(self, seconds):
        """跳到第 seconds 秒（拖进度条用）。"""
        if self._lib is None or self._player is None:
            return False
        try:
            ms = int(round(max(0.0, float(seconds)) * 1000.0))
            return int(self._lib.libvlc_media_player_set_time(
                ctypes.c_void_p(self._player), ms)) == 0
        except Exception as exc:
            self.last_error = str(exc)
            return False

    def pause(self):
        if self._lib is None or self._player is None:
            return
        try:
            self._lib.libvlc_media_player_set_pause(
                ctypes.c_void_p(self._player), 1)
        except Exception:
            pass

    def resume(self):
        if self._lib is None or self._player is None:
            return
        try:
            self._lib.libvlc_media_player_set_pause(
                ctypes.c_void_p(self._player), 0)
        except Exception:
            pass

    def set_volume(self, value):
        try:
            self._volume = max(0, min(int(value), 100))
        except (TypeError, ValueError):
            return
        if self._lib is None or self._player is None:
            return
        try:
            self._lib.libvlc_audio_set_volume(ctypes.c_void_p(self._player),
                                              self._volume)
        except Exception:
            pass


class VlcPlayer(object):
    """VLC 后端：一次把整个列表交给 VLC，它自己顺序播放、循环、全屏。"""

    name = 'VLC'
    supports_pause = False
    supports_seek = False    # 独立进程的 VLC 没法从外面跳位置

    def __init__(self):
        self._exe = find_vlc()
        self._proc = None
        self._win_rect = None
        self.last_error = ''

    @property
    def available(self):
        return bool(self._exe)

    def describe(self):
        return 'VLC（解码最全）' if self._exe else 'VLC 未安装'

    def play(self, playlist, index=0, monitor_index=0, loop=True,
             volume=100, topmost=True, audio_device=None, window_rect=None,
             host_hwnd=None):
        if not self._exe:
            return False, '没找到 VLC'
        self.stop()
        # VLC 3 起 --volume 已经不存在（会打印 “option --volume no longer
        # exists” 并忽略），音量要交给具体的声音输出模块，否则音量条没用。
        gain = max(0.0, min(int(volume), 100)) / 100.0
        common = [self._exe, '--no-video-title-show', '--no-qt-privacy-ask',
                  '--mmdevice-volume=%.3f' % gain,
                  '--waveout-volume=%.3f' % gain,
                  '--directx-volume=%.3f' % gain]
        if not window_rect:
            # 没指定框就是整屏：直接用目标显示器的整块区域
            monitors = list_monitors()
            index = max(0, min(int(monitor_index), len(monitors) - 1))
            window_rect = tuple(monitors[index]['rect'])
        monitors = list_monitors()
        index = max(0, min(int(monitor_index), len(monitors) - 1))
        full_rect = tuple(monitors[index]['rect'])
        is_full = tuple(window_rect) == full_rect
        # 一律用无界面模式 + 我们自己摆窗口：
        # 这样不会有 VLC 的标题栏、控制栏、任务栏条目（会场不能露出来）
        args = common + ['--intf', 'dummy', '--no-fullscreen',
                         '--no-embedded-video',
                         '--width=%d' % (window_rect[2] - window_rect[0]),
                         '--height=%d' % (window_rect[3] - window_rect[1]),
                         '--video-x=%d' % window_rect[0],
                         '--video-y=%d' % window_rect[1]]
        if audio_device:
            # 不选的话 VLC 跟系统默认设备走；系统默认选错（比如选了没接
            # 音箱的 USB 声卡）就会“画面在放、一点声音都没有”
            args.append('--mmdevice-audio-device=%s' % audio_device)
        if loop:
            args.append('--loop')
        else:
            args.append('--play-and-exit')
        files = list(playlist)
        if 0 < index < len(files):
            files = files[index:] + files[:index]     # 从选中的开始
        args.extend(files)
        try:
            self._proc = subprocess.Popen(args, close_fds=True)
        except Exception as exc:
            self.last_error = str(exc)
            return False, '启动 VLC 失败：%s' % exc
        self._win_rect = tuple(window_rect) if window_rect else None
        if self._win_rect:
            threading.Thread(target=self._keep_window,
                             args=(self._proc.pid, self._win_rect),
                             daemon=True).start()
            if is_full:
                return True, '正在全屏播放（显示器 %d，无操作栏）' % (index + 1)
            return True, '已用合成画面播放（画面位置：%d,%d）' % (
                window_rect[0], window_rect[1])
        return True, '已交给 VLC 全屏播放（显示器 %d）' % (monitor_index + 1)

    def _keep_window(self, pid, rect):
        """把 VLC 的视频窗口摆到指定矩形上（整屏 / 合成画面框）。

        关键：VLC 刚启动那 1 秒还在给自己的窗口排版，这时候去摆它，
        它随后又会按自己的尺寸重排，两边对拉——看起来就是画面一直在
        左右滑动。所以这里等它排好再摆一次，之后只在“明显跑偏”时才补摆，
        不再每 0.5 秒拽一把。顺便把标题栏/任务栏条目去掉。
        """
        time.sleep(1.2)                 # 等 VLC 把自己的窗口排好版
        misses = 0
        placed = False
        while True:
            proc = self._proc
            if proc is None or proc.pid != pid or proc.poll() is not None:
                return
            if self._win_rect != tuple(rect):
                return
            hwnd = find_video_window(pid)
            if hwnd:
                misses = 0
                cur = window_rect(hwnd)
                if not placed:
                    # 先摆位置，再去标题栏（反过来 VLC 会重排）
                    move_window(hwnd, rect, topmost=True)
                    hide_window_chrome(hwnd)
                    placed = True
                elif cur is not None and (
                        abs(cur[0] - rect[0]) > 40 or abs(cur[1] - rect[1]) > 40
                        or cur[2] - cur[0] < (rect[2] - rect[0]) * 0.6
                        or cur[3] - cur[1] < (rect[3] - rect[1]) * 0.6):
                    # 明显跑掉/缩太小才补一次；小差异不管，免得来回拉
                    move_window(hwnd, rect, topmost=True)
            else:
                misses += 1
                if misses > 40:            # 20 秒还没窗口就别耗着了
                    return
            time.sleep(0.5)

    def stop(self):
        self._win_rect = None
        if self._proc is None:
            return
        try:
            self._proc.terminate()
        except Exception:
            pass
        try:
            self._proc.wait(timeout=3)
        except Exception:
            try:
                self._proc.kill()
            except Exception:
                pass
        self._proc = None

    def state(self):
        if self._proc is None:
            return 1
        return 3 if self._proc.poll() is None else 1

    def close(self):
        self.stop()

    def set_volume(self, value):
        pass                     # VLC 的音量在启动时给定了

    def duration(self):
        return 0.0

    def position(self):
        return 0.0

    def seek(self, seconds):
        return False


class VideoPlayer(object):
    """按 VLC内嵌 → VLC（独立进程）的顺序自动挑后端。

    各后端对外接口一致：
        ok, note = player.play(播放列表, 起始序号, 显示器序号, 循环, 音量)
        player.stop() / player.state() / player.current_name()
    """

    def __init__(self, backend='auto'):
        self._backends = {'libvlc': LibVlcPlayer(), 'vlc': VlcPlayer()}
        self.backend_name = 'auto'
        self.backend = None
        self.set_backend(backend)

    def set_backend(self, name):
        """name: auto / libvlc / vlc。"""
        self.backend_name = name if name in self._backends or name == 'auto' \
            else 'auto'
        if self.backend is not None:
            try:
                self.backend.stop()
            except Exception:
                pass
        if self.backend_name != 'auto':
            self.backend = self._backends[self.backend_name]
            return self.backend
        # 内嵌播放最稳（画面固定、无播放器界面），优先用它
        order = ('libvlc', 'vlc')
        for key in order:
            candidate = self._backends[key]
            if candidate.available:
                self.backend = candidate
                return candidate
        self.backend = self._backends['libvlc']      # 都不可用时给个空壳
        return self.backend

    def describe(self):
        if self.backend_name != 'auto':
            return self.backend.describe()
        return '%s（自动选择）' % self.backend.describe()

    # ---------- 转发 ----------
    @property
    def available(self):
        return bool(self.backend and self.backend.available)

    @property
    def supports_pause(self):
        return bool(self.backend and self.backend.supports_pause)

    @property
    def supports_seek(self):
        """能不能拖进度条跳位置（独立进程的 VLC 不支持）。"""
        return bool(getattr(self.backend, 'supports_seek', False))

    @property
    def last_error(self):
        return getattr(self.backend, 'last_error', '')

    def play(self, playlist, index=0, monitor_index=0, loop=True,
             volume=100, topmost=True, audio_device=None, window_rect=None,
             host_hwnd=None):
        if getattr(self.backend, 'embedded', False) and host_hwnd is None:
            # 内嵌播放必须给窗口句柄；没给就退回外部播放器
            for key in ('vlc',):
                cand = self._backends[key]
                if cand.available:
                    self.backend = cand
                    self.backend_name = key
                    break
        return self.backend.play(playlist, index, monitor_index, loop,
                                 volume, topmost, audio_device, window_rect,
                                 host_hwnd)

    @property
    def embedded(self):
        """当前后端是不是“画进我们自己的窗口”那种。"""
        return bool(getattr(self.backend, 'embedded', False))

    def tick(self):
        tick = getattr(self.backend, 'tick', None)
        if callable(tick):
            try:
                tick()
            except Exception:
                pass

    @property
    def paused(self):
        """当前后端是不是暂停状态（没有这个概念的就返回 False）。"""
        return bool(getattr(self.backend, 'paused', False))

    def stop(self):
        if self.backend:
            self.backend.stop()

    def test_sound(self, audio_device=None, volume=100):
        """试听选定的声音输出设备（放一小段提示音）。"""
        return play_test_tone(audio_device, volume)

    def close(self):
        for player in self._backends.values():
            try:
                player.close()
            except Exception:
                pass

    def state(self):
        return self.backend.state() if self.backend else 1

    def current_name(self):
        if hasattr(self.backend, 'current_name'):
            return self.backend.current_name()
        return ''

    def position(self):
        return self.backend.position() if self.backend else 0.0

    def duration(self):
        return self.backend.duration() if self.backend else 0.0

    def pause(self):
        if hasattr(self.backend, 'pause'):
            self.backend.pause()

    def seek(self, seconds):
        """跳到第 seconds 秒；不支持的后端返回 False。"""
        seek = getattr(self.backend, 'seek', None)
        if not callable(seek):
            return False
        try:
            return bool(seek(seconds))
        except Exception:
            return False

    def resume(self):
        if hasattr(self.backend, 'resume'):
            self.backend.resume()

    def set_volume(self, value):
        if self.backend:
            self.backend.set_volume(value)
