# -*- coding: utf-8 -*-
"""现场守护（会务现场防打扰）。

两件事：
1. 防止系统进入屏保 / 关闭显示器 / 睡眠：
   调用 SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED |
   ES_DISPLAY_REQUIRED)。这是 PowerPoint 放映、视频播放器使用的同一套
   标准接口，不需要管理员权限，也不修改系统设置；本进程退出后系统
   自动恢复原来的电源策略（MSDN 明确 ES_DISPLAY_REQUIRED 同时阻止
   屏幕保护启动）。

2. 自动最小化弹窗广告：
   后台线程定时枚举可见顶层窗口，按“来源进程黑名单 + 标题强特征”
   识别广告弹窗，命中后只做 SW_MINIMIZE（最小化，不杀进程、不关闭、
   不模拟点击），并留下记录可在界面里一键还原。系统窗口、安全软件、
   浏览器、办公/会议软件及本程序自身在保护名单中，永不处理；新窗口
   还有 2.5 秒宽限，避免误伤开机自启闪窗。
"""

import ctypes
import os
import threading
import time
from ctypes import wintypes

# ---------------- SetThreadExecutionState 标志 ----------------
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

# ---------------- ShowWindow 命令 ----------------
SW_MINIMIZE = 6
SW_RESTORE = 9

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

# 扫描间隔 / 新窗口宽限（秒）：宽限期内不动，正常软件的启动闪窗不受影响
_POLL_INTERVAL = 1.5
_GRACE_SECONDS = 2.5

# ---------------- 广告来源进程（小写进程名命中即拦） ----------------
# 只收录“弹窗推送组件”，不收录对应软件的主程序（主程序在保护名单）。
AD_PROCESSES = frozenset((
    'sohunews.exe',        # 搜狐资讯/微门户
    'minipage.exe',        # QQ 腾讯网迷你版
    'qqexternal.exe',      # QQ 外部弹窗/广告组件
    'teniol.exe',          # 腾讯下载/广告组件
    'thunderpush.exe',     # 迅雷资讯推送
    'ksolaunch.exe',       # WPS 热点/广告推送
    'wpscloudsvr.exe',     # WPS 云广告
    '360news.exe',         # 360 新闻弹窗
    '2345spsvc.exe',       # 2345 弹窗服务
    '2345mp_svc.exe',
    '2345mpsvc.exe',
    '2345softmgr.exe',
    'masterhelper.exe',    # 鲁大师弹窗
    'qyfragment.exe',      # 爱奇艺弹窗外挂
    'pinyinup.exe',        # 搜狗输入法升级/广告弹窗
    'ppminisite.exe',      # PPS 迷你资讯
    'baidunetdiskrender.exe',  # 百度网盘营销弹窗
    'dgminisite.exe',      # 驱动精灵迷你弹窗
    'mydriversmini.exe',
    'kwminipage.exe',      # 酷我音乐迷你资讯
    'kgminipage.exe',      # 酷狗迷你页
    'haozipnews.exe',      # 好压资讯
))

# ---------------- 窗口标题强特征（小写包含匹配，命中即拦） ----------------
# 只收“广告味很强、正常工作窗口几乎不会使用”的词，宁漏勿错。
AD_TITLE_KEYWORDS = (
    '今日热点', '资讯快报', '资讯迷你版', '迷你首页', '搜狐资讯',
    '迅雷资讯', '暴风资讯', '好压资讯', '搜狗资讯', '热点资讯',
    '每日推荐', '今日推荐', '壁纸推荐', '猜你喜欢', '为你推荐',
    '优惠券', '内部优惠券',
    '限时秒杀', '秒杀', '限时抢购', '限时特惠', '特价清仓', '九块九',
    '火爆热销', '热卖', '大促', '年中大促', '双11', '双十一', '双12',
    '红包雨', '领红包', '现金红包',
    '免费领取', '0元购', '0元抢', '免费抽', '免费玩',
    '马上下载', '立即下载', '点击下载', '高速下载', '安全下载',
    '点击查看详情',
    '装备回收', '贪玩蓝月', '高爆率', '充值返利', '传奇霸业',
    '棋牌', '电玩城', '彩金', '娱乐城', '威尼斯人', '六合彩',
    '极速放款', '秒下款',
    '美女直播',
)

# ---------------- 保护进程（永不处理） ----------------
# 系统组件
_PROTECT_SYSTEM = (
    'explorer.exe', 'dwm.exe', 'sihost.exe', 'ctfmon.exe',
    'taskhostw.exe', 'runtimebroker.exe', 'applicationframehost.exe',
    'startmenuexperiencehost.exe', 'searchhost.exe',
    'shellexperiencehost.exe', 'systemsettings.exe', 'control.exe',
    'taskmgr.exe', 'lockapp.exe', 'logonui.exe', 'consent.exe',
    'credentialuibroker.exe', 'conhost.exe', 'powershell.exe',
    'cmd.exe', 'windowsterminal.exe', 'mmc.exe', 'regedit.exe',
)
# 安全软件（带自我保护，碰了会报毒；现场电脑上让它们自己安静）
_PROTECT_SECURITY = (
    '360tray.exe', 'zhudongfangyu.exe', '360sd.exe', '360rp.exe',
    '360safe.exe', 'qhsafe.exe',
    'hipstray.exe', 'hipsdaemon.exe', 'usysdiag.exe',
    'msmpeng.exe', 'nissrv.exe', 'securityhealthservice.exe',
    'securityhealthsystray.exe',
    'qqpctray.exe', 'qqpcrtp.exe', 'qqpcleauncher.exe',
    'kxetray.exe', 'kxescore.exe',
    'rstray.exe', 'ravmond.exe', 'rsmain.exe', 'rsagent.exe',
    'baidusd.exe', 'baiduansvr.exe',
)
# 浏览器（网页里的广告归浏览器管，顶层窗口是浏览器主窗不能最小化）
_PROTECT_BROWSER = (
    'chrome.exe', 'msedge.exe', 'firefox.exe', 'iexplore.exe',
    '360se.exe', '360chrome.exe', 'qqbrowser.exe', 'liebao.exe',
    'sogouexplorer.exe',
)
# 办公 / 会议通讯 / 播放器（现场常用的正经软件）
_PROTECT_WORK = (
    'powerpnt.exe', 'winword.exe', 'excel.exe', 'outlook.exe',
    'onenote.exe', 'wps.exe', 'et.exe', 'wpp.exe', 'wpspdf.exe',
    'acrord32.exe', 'acrrobat.exe', 'foxitpdfreader.exe',
    'sumatrapdf.exe',
    'wechat.exe', 'qq.exe', 'tim.exe', 'dingtalk.exe', 'wemeet.exe',
    'tencentmeeting.exe', 'lark.exe', 'feishu.exe', 'zoom.exe',
    'wxwork.exe',
    'vlc.exe', 'potplayermini64.exe', 'potplayer.exe', 'wmplayer.exe',
    'mpc-hc64.exe', 'mpc-be64.exe',
    'notepad.exe', 'calc.exe', 'mspaint.exe', 'snippingtool.exe',
    'screenclippinghost.exe', 'snipaste.exe',
)
PROTECT_PROCESSES = frozenset(
    _PROTECT_SYSTEM + _PROTECT_SECURITY + _PROTECT_BROWSER
    + _PROTECT_WORK)

# 系统窗口类名：即使标题命中也一律不碰
PROTECT_CLASSES = frozenset((
    'Progman', 'Shell_TrayWnd', 'Windows.UI.Core.CoreWindow',
    'Xaml_WindowedPopupClass', 'Button', 'IME',
))


def _load_apis():
    """惰性加载 ctypes 接口（非 Windows 环境下 import 不报错）。"""
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)

    enum_windows = user32.EnumWindows
    enum_windows.argtypes = [ctypes.WINFUNCTYPE(
        wintypes.BOOL, wintypes.HWND, wintypes.LPARAM), wintypes.LPARAM]
    enum_windows.restype = wintypes.BOOL

    is_window_visible = user32.IsWindowVisible
    is_window_visible.argtypes = [wintypes.HWND]
    is_window_visible.restype = wintypes.BOOL

    get_window_text_len = user32.GetWindowTextLengthW
    get_window_text_len.argtypes = [wintypes.HWND]
    get_window_text_len.restype = ctypes.c_int

    get_window_text = user32.GetWindowTextW
    get_window_text.argtypes = [
        wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    get_window_text.restype = ctypes.c_int

    get_class_name = user32.GetClassNameW
    get_class_name.argtypes = [
        wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    get_class_name.restype = ctypes.c_int

    get_pid = user32.GetWindowThreadProcessId
    get_pid.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    get_pid.restype = wintypes.DWORD

    show_window = user32.ShowWindow
    show_window.argtypes = [wintypes.HWND, ctypes.c_int]
    show_window.restype = wintypes.BOOL

    is_window = user32.IsWindow
    is_window.argtypes = [wintypes.HWND]
    is_window.restype = wintypes.BOOL

    open_process = kernel32.OpenProcess
    open_process.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    open_process.restype = wintypes.HANDLE

    close_handle = kernel32.CloseHandle
    close_handle.argtypes = [wintypes.HANDLE]
    close_handle.restype = wintypes.BOOL

    query_name = kernel32.QueryFullProcessImageNameW
    query_name.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD)]
    query_name.restype = wintypes.BOOL

    set_state = kernel32.SetThreadExecutionState
    set_state.argtypes = [wintypes.DWORD]
    set_state.restype = wintypes.DWORD

    return {
        'enum_windows': enum_windows,
        'is_window_visible': is_window_visible,
        'get_window_text_len': get_window_text_len,
        'get_window_text': get_window_text,
        'get_class_name': get_class_name,
        'get_pid': get_pid,
        'show_window': show_window,
        'is_window': is_window,
        'open_process': open_process,
        'close_handle': close_handle,
        'query_name': query_name,
        'set_state': set_state,
    }


class FocusGuard(object):
    """现场守护。一个实例即可，start()/stop() 成对调用。

    on_event：拦截发生时的回调（在后台线程被调用），收到一个 dict：
        {'time': 'HH:MM:SS', 'title': ..., 'process': ..., 'hwnd': int}
    主程序应在回调里用 root.after 把 UI 更新切回主线程。
    """

    def __init__(self, on_event=None, awake_on=True, block_on=True):
        self._apis = None
        self._on_event = on_event
        self._awake_on = bool(awake_on)
        self._block_on = bool(block_on)
        self._lock = threading.Lock()
        self._stop_evt = threading.Event()
        self._thread = None
        self._self_pid = os.getpid()
        # hwnd -> 首次发现时间；记录只在本轮扫描周期内维护
        self._seen = {}
        # 已最小化过的窗口句柄（最小化后仍“可见”，靠它防重复处理）
        self._handled = set()
        # 已处理记录（界面展示/还原用）
        self.history = []
        self._count = 0

    # ---------------- 开关 ----------------
    def set_awake(self, on):
        with self._lock:
            changed = self._awake_on != bool(on)
            self._awake_on = bool(on)
        if changed and self._thread is not None:
            self._apply_awake()

    def set_block(self, on):
        with self._lock:
            self._block_on = bool(on)

    @property
    def awake_on(self):
        return self._awake_on

    @property
    def block_on(self):
        return self._block_on

    @property
    def count(self):
        return self._count

    # ---------------- 生命周期 ----------------
    def start(self):
        if self._thread is not None:
            return
        try:
            self._apis = _load_apis()
        except Exception:
            # 非 Windows / API 不可用：守护静默不工作，不影响主程序
            self._apis = None
            return
        self._stop_evt.clear()
        self._apply_awake()
        self._thread = threading.Thread(
            target=self._run, name='FocusGuard', daemon=True)
        self._thread.start()

    def stop(self):
        self._stop_evt.set()
        t = self._thread
        if t is not None:
            t.join(timeout=2.0)
        self._thread = None
        # 无论是否在常亮，退出时都明确归还系统电源控制权
        self._set_execution_state(ES_CONTINUOUS)

    def _apply_awake(self):
        if self._apis is None:
            return
        with self._lock:
            on = self._awake_on
        if on:
            self._set_execution_state(
                ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)
        else:
            self._set_execution_state(ES_CONTINUOUS)

    def _set_execution_state(self, flags):
        if self._apis is None:
            return
        try:
            self._apis['set_state'](flags)
        except Exception:
            pass

    # ---------------- 后台扫描 ----------------
    def _run(self):
        apis = self._apis
        while not self._stop_evt.wait(_POLL_INTERVAL):
            # 常亮状态需要持续声明（有些电源事件会清除线程执行状态）
            with self._lock:
                awake = self._awake_on
                block = self._block_on
            if awake:
                self._set_execution_state(
                    ES_CONTINUOUS | ES_SYSTEM_REQUIRED
                    | ES_DISPLAY_REQUIRED)
            if not block:
                # 关闭拦截时顺手清掉观察期记录，重新开启后重新计时
                self._seen.clear()
                continue
            try:
                self._scan_once(apis)
            except Exception:
                # 守护线程不能因任何意外挂掉
                continue

    def _scan_once(self, apis):
        now = time.monotonic()
        alive = set()

        def _cb(hwnd, _lparam):
            alive.add(hwnd)
            return True

        apis['enum_windows'](
            ctypes.WINFUNCTYPE(
                wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(_cb), 0)

        # 清掉已销毁窗口的观察记录与已处理记录（hwnd 复用风险也一并消除）
        dead = [h for h in self._seen if h not in alive]
        for h in dead:
            del self._seen[h]
        self._handled &= alive

        for hwnd in alive:
            try:
                if not apis['is_window_visible'](hwnd):
                    continue
                if hwnd in self._handled:
                    continue
                title = self._get_text(apis, hwnd)
                cls = self._get_class(apis, hwnd)
                if not title or cls in PROTECT_CLASSES:
                    continue
                pid = wintypes.DWORD(0)
                apis['get_pid'](hwnd, ctypes.byref(pid))
                if pid.value == self._self_pid:
                    continue
                proc = self._get_process_name(apis, pid.value)
                if proc and proc in PROTECT_PROCESSES:
                    continue
                if not self._looks_like_ad(title, cls, proc):
                    continue
                # 新窗口先观察，宽限期后还在才认定是弹窗
                first = self._seen.get(hwnd)
                if first is None:
                    self._seen[hwnd] = now
                    continue
                if now - first < _GRACE_SECONDS:
                    continue
                self._minimize(hwnd, title, proc)
            except Exception:
                continue

    @staticmethod
    def _looks_like_ad(title, cls, proc):
        """广告判定：进程黑名单优先；其次标题强特征。"""
        if proc and proc in AD_PROCESSES:
            return True
        low = title.lower()
        for kw in AD_TITLE_KEYWORDS:
            if kw in low:
                return True
        return False

    def _minimize(self, hwnd, title, proc):
        try:
            self._apis['show_window'](hwnd, SW_MINIMIZE)
        except Exception:
            return
        self._handled.add(hwnd)
        rec = {
            'time': time.strftime('%H:%M:%S'),
            'title': title[:60],
            'process': proc or '未知进程',
            'hwnd': int(hwnd),
        }
        self.history.append(rec)
        self._count += 1
        if len(self.history) > 200:
            del self.history[:-200]
        if self._on_event is not None:
            try:
                self._on_event(dict(rec))
            except Exception:
                pass

    # ---------------- 工具 ----------------
    @staticmethod
    def _get_text(apis, hwnd):
        try:
            n = apis['get_window_text_len'](hwnd)
            if n <= 0:
                return ''
            buf = ctypes.create_unicode_buffer(n + 2)
            apis['get_window_text'](hwnd, buf, n + 2)
            return buf.value.strip()
        except Exception:
            return ''

    @staticmethod
    def _get_class(apis, hwnd):
        try:
            buf = ctypes.create_unicode_buffer(256)
            apis['get_class_name'](hwnd, buf, 256)
            return buf.value
        except Exception:
            return ''

    @staticmethod
    def _get_process_name(apis, pid):
        try:
            h = apis['open_process'](
                PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if not h:
                return ''
            try:
                size = wintypes.DWORD(1024)
                buf = ctypes.create_unicode_buffer(1024)
                ok = apis['query_name'](h, 0, buf, ctypes.byref(size))
                if ok:
                    return os.path.basename(buf.value).lower()
            finally:
                apis['close_handle'](h)
        except Exception:
            pass
        return ''

    # ---------------- 还原 ----------------
    def restore(self, hwnd):
        """把已最小化的弹窗还原到屏幕上，返回是否成功。"""
        if self._apis is None:
            return False
        try:
            if not self._apis['is_window'](hwnd):
                return False
            self._apis['show_window'](hwnd, SW_RESTORE)
            return True
        except Exception:
            return False

    def clear_history(self):
        self.history.clear()
        self._count = 0
