# -*- coding: utf-8 -*-
"""统一的崩溃防护与运行日志。

四个入口（会务助手 / 座位牌生成器 / 远程打印 / 打印代理）都调用本模块：

1. 未捕获异常写入日志文件，不再直接退出；
2. 图形界面下用弹窗提示用户，并做频率限制，避免错误反复弹窗卡死界面；
3. 线程里的异常单独兜底，子线程出错不会影响主程序。

设计原则：本模块自身绝不能再抛异常。任何一步失败都静默降级。
"""

import os
import sys
import time
import traceback

MAX_LOG_BYTES = 1024 * 1024  # 单个日志文件上限，超过后轮转
_BACKUP_SUFFIX = '.old'

# 弹窗频率限制：避免某个错误每秒触发一次时弹窗刷屏、把程序卡死
_DIALOG_MIN_INTERVAL = 15.0
_DIALOG_MAX_TIMES = 5


def app_data_dir():
    """程序数据目录（用户目录下，不会因为桌面整理而丢日志）。"""
    base = (os.environ.get('APPDATA') or os.environ.get('LOCALAPPDATA')
            or os.path.expanduser('~'))
    return os.path.join(base, '会务助手')


def _program_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


_resolved_log = None


def log_path():
    """日志文件位置：优先用户目录，不可写时退回程序目录、临时目录。"""
    global _resolved_log
    if _resolved_log:
        return _resolved_log
    import tempfile
    candidates = [
        os.path.join(app_data_dir(), '运行日志.txt'),
        os.path.join(_program_dir(), '运行日志.txt'),
        os.path.join(tempfile.gettempdir(), '会务助手_运行日志.txt'),
    ]
    for path in candidates:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, 'a', encoding='utf-8'):
                pass
            _resolved_log = path
            return path
        except Exception:
            continue
    _resolved_log = candidates[-1]
    return _resolved_log


def _rotate(path):
    try:
        if os.path.getsize(path) < MAX_LOG_BYTES:
            return
    except OSError:
        return
    try:
        backup = path + _BACKUP_SUFFIX
        if os.path.exists(backup):
            os.remove(backup)
        os.replace(path, backup)
    except OSError:
        # 轮转失败时退而求其次：清空当前文件，保证还能继续记录
        try:
            with open(path, 'w', encoding='utf-8'):
                pass
        except OSError:
            pass


def write_log(message, path=None):
    """把一行日志追加到文件；失败时静默忽略（日志不能影响主流程）。"""
    path = path or log_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        _rotate(path)
        line = '[%s] %s' % (time.strftime('%Y-%m-%d %H:%M:%S'), message)
        with open(path, 'a', encoding='utf-8') as f:
            f.write(line.rstrip('\n') + '\n')
    except Exception:
        pass


def _message_box(title, text):
    """不依赖 Tk 的弹窗，Tk 窗口还没建好或已损坏时也能提示。"""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, text, title, 0x00040040)
        return True
    except Exception:
        return False


class CrashGuard(object):
    """异常兜底器：记录日志、限频率弹窗。"""

    def __init__(self, app_name='会务助手', gui=True, show_dialog=True):
        self.app_name = app_name
        self.gui = gui
        self.show_dialog = show_dialog and gui
        self._last_dialog = 0.0
        self._dialog_times = 0
        self._seen = set()
        self._handling = False

    # ---------- 对外 ----------
    def handle(self, exc_type, exc_value, exc_tb, where=''):
        """处理一个未捕获异常；返回日志文本（不抛异常）。"""
        if self._handling:  # 处理器自身出错时不再递归
            return ''
        self._handling = True
        try:
            text = self._format(exc_type, exc_value, exc_tb, where)
            write_log(text + '\n' + ''.join(
                traceback.format_exception(exc_type, exc_value, exc_tb)))
            self._maybe_dialog(text)
            self._write_stderr(text)
            return text
        except Exception:
            return ''
        finally:
            self._handling = False

    def install(self, root=None):
        """挂上全局兜底：主线程 / 子线程 / Tk 回调。"""
        def hook(exc_type, exc_value, exc_tb):
            self.handle(exc_type, exc_value, exc_tb)

        try:
            sys.excepthook = hook
        except Exception:
            pass
        try:
            import threading
            threading.excepthook = (
                lambda args: self.handle(args.exc_type, args.exc_value,
                                         args.exc_traceback, '子线程'))
        except Exception:
            pass
        if root is not None:
            try:
                root.report_callback_exception = (
                    lambda exc_type, exc_value, exc_tb: self.handle(
                        exc_type, exc_value, exc_tb, '界面操作'))
            except Exception:
                pass
        return self

    # ---------- 内部 ----------
    def _format(self, exc_type, exc_value, exc_tb, where):
        name = getattr(exc_type, '__name__', str(exc_type))
        detail = str(exc_value).strip() or name
        if len(detail) > 300:
            detail = detail[:300] + '…'
        prefix = ('%s出错' % where) if where else '程序出错'
        return '%s：%s: %s' % (prefix, name, detail)

    def _maybe_dialog(self, text):
        if not self.show_dialog:
            return
        now = time.time()
        if self._dialog_times >= _DIALOG_MAX_TIMES:
            return
        if now - self._last_dialog < _DIALOG_MIN_INTERVAL:
            return
        self._last_dialog = now
        self._dialog_times += 1
        tail = ('\n\n（已记录到日志：%s）' % log_path())
        if self._dialog_times >= _DIALOG_MAX_TIMES:
            tail += '\n本次运行不再弹出此类提示，问题已写入日志。'
        _message_box(
            self.app_name,
            '刚才的操作遇到一个问题，程序已跳过该操作继续运行，'
            '不会影响其他功能。\n\n%s%s' % (text, tail))

    def _write_stderr(self, text):
        try:
            if sys.stderr is not None:
                sys.stderr.write(text + '\n')
                sys.stderr.flush()
        except Exception:
            pass


def install(app_name='会务助手', gui=True, root=None, show_dialog=True):
    """一行接入：安装全局兜底并返回 CrashGuard 实例。"""
    guard = CrashGuard(app_name=app_name, gui=gui,
                       show_dialog=show_dialog)
    guard.install(root=root)
    return guard


def main_guard(app_name='会务助手', gui=True, func=None):
    """包住入口函数：即使初始化阶段就出错，也只提示不闪退。"""
    guard = CrashGuard(app_name=app_name, gui=gui)
    try:
        return func() if func is not None else None
    except SystemExit:
        raise
    except BaseException:
        exc_type, exc_value, exc_tb = sys.exc_info()
        text = guard.handle(exc_type, exc_value, exc_tb, '启动')
        if gui:
            _message_box(
                app_name,
                '程序启动失败，未能正常运行。%s\n\n'
                '详细信息已记录到日志：\n%s'
                % (('\n\n' + text) if text else '', log_path()))
        return None
