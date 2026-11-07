# -*- coding: utf-8 -*-
"""现场抽签排序（独立小工具）：名单洗牌 + 大屏投屏 + 历史导出。

源码运行：python 小工具/现场抽签排序.py
"""

import _boot


def _open(root):
    import tkinter as tk

    import lottery
    import ui_theme

    win = tk.Toplevel(root)
    win.withdraw()          # 先建内容，摆好再显示，避免闪一下
    win.geometry('1080x760')
    try:
        ui_theme.style_window(win)
    except Exception:
        pass

    lottery.LotteryAdmin(win, on_close=lambda: _shutdown(win, root))
    _boot.bind_close(win, root)

    win.update_idletasks()
    win.deiconify()
    win.lift()


def _shutdown(win, root):
    try:
        win.destroy()
    finally:
        try:
            root.destroy()
        except Exception:
            pass


_boot.run('现场抽签排序', _open)
