# -*- coding: utf-8 -*-
"""小工具公共启动样板（解耦版）。

`小工具/` 下每个脚本都是**独立程序**：源码直接 `python 小工具/提词器.py`
就能跑，打包后也是一个单独的 exe，不再依赖「会务助手」主界面。

这里统一处理四件事：

1. 把项目根目录放进 ``sys.path``（打包后是 exe 内置目录），
   让各小工具还能复用根目录下的 prompter / ui_theme 等模块；
2. 建一个 1x1、挪到屏幕外的根窗口：工具窗口挂在它下面，``transient``
   才生效，而用户看不到这个“空主窗”；
3. 统一套 ``safe_guard``：出错写 运行日志.txt，不闪退；
4. 工具窗口一关，整个程序跟着退出（否则隐藏的根窗口会让进程赖着）。
"""

import os
import sys

BASE_DIR = (getattr(sys, '_MEIPASS', '')
            or os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if BASE_DIR and BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

# 程序所在目录：解耦版的数据（收件目录等）往这里写，
# 不能往 _MEIPASS 写——那是一次性解包目录，退出就没了。
APP_DIR = (os.path.dirname(os.path.abspath(sys.executable))
           if getattr(sys, 'frozen', False) else BASE_DIR)


def bind_close(win, root, delay=80):
    """工具窗口被销毁后退出整个程序。"""
    def _on_destroy(event):
        if event.widget is win:
            try:
                root.after(delay, root.destroy)
            except Exception:
                pass

    try:
        win.bind('<Destroy>', _on_destroy, add='+')
    except Exception:
        pass


def hide_root(root):
    """把根窗口压成 1x1 挪到屏幕外：能被 transient 引用，但看不见。"""
    try:
        root.geometry('1x1+%d+0' % (root.winfo_screenwidth() + 200))
        root.overrideredirect(True)
    except Exception:
        pass


def run(app_name, opener):
    """建根窗口 → ``opener(root)`` 里打开工具 → 工具关掉就退出。"""
    import tkinter as tk

    import safe_guard
    import ui_theme

    def _main():
        root = tk.Tk()
        root.title(app_name)
        hide_root(root)
        try:
            ui_theme.apply_theme(root)
        except Exception:
            pass
        guard = safe_guard.install(app_name=app_name, gui=True, root=root)
        opener(root)
        root.mainloop()
        del guard

    safe_guard.main_guard(app_name=app_name, gui=True, func=_main)


def run_host_tool(app_name, open_method, win_attr):
    """跑「界面写在 会务助手.py 里」的工具，但只开这一个工具。

    这些工具的界面（会务音乐 / 会务视频 / 长条屏 / PPT倒计时 / PPT转换）
    目前还长在主程序里，整体搬出来风险高、代码量大。这里用主程序自己的
    App 建界面，把主界面缩到屏幕外，只弹目标工具窗——对外同样是
    “一个工具一个 exe”，等以后逐个抽离成独立模块再替换实现即可。
    """
    import tkinter as tk

    import importlib

    import safe_guard
    import ui_theme

    # 注意：这里刻意不用 `import 会务助手`——写在源码里的普通 import 会被
    # PyInstaller 静态分析捕获，导致**每个**小工具都被塞进整套主程序
    # （cv2/ffmpeg 等重依赖全带上，exe 从 30MB 涨到 100MB+）。
    # 拆成拼接字符串走 importlib，只有单工具壳的打包配置里才声明它。
    host = importlib.import_module('会务' + '助手')

    class _SingleTool(host.App):
        """主程序 App 的子类：永不显示主界面。"""

        def _show_root(self):
            hide_root(self.root)
            try:
                self.root.deiconify()
            except tk.TclError:
                pass

    def _main():
        root = tk.Tk()
        root.withdraw()
        root.title(app_name)
        try:
            ui_theme.apply_theme(root)
        except Exception:
            pass
        guard = safe_guard.install(app_name=app_name, gui=True, root=root)
        app = _SingleTool(root)
        getattr(app, open_method)()
        win = getattr(app, win_attr, None)
        if win is not None:
            bind_close(win, root)
        root.mainloop()
        del guard

    safe_guard.main_guard(app_name=app_name, gui=True, func=_main)
