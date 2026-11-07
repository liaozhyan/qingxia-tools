# -*- coding: utf-8 -*-
"""截图（独立小工具）：区域/全屏/窗口/滚动截图，仿 ShareX。

源码运行：python 小工具/截图.py
"""

import _boot


def _open(root):
    import screenshot

    panel = screenshot.open_panel(root)
    win = getattr(panel, 'win', None)
    if win is not None:
        _boot.bind_close(win, root)


_boot.run('截图', _open)
