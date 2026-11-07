# -*- coding: utf-8 -*-
"""局域网互传（独立小工具）：两台电脑同一 WiFi 下互传文件。

源码运行：python 小工具/局域网互传.py
"""

import _boot


def _open(root):
    import lan_share

    app = lan_share.LanShareApp(root)
    if app.win is not None:
        _boot.bind_close(app.win, root)


_boot.run('局域网互传', _open)
