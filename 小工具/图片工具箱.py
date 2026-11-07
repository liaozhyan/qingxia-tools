# -*- coding: utf-8 -*-
"""图片工具箱（独立小工具）：裁剪缩放旋转、水印、压缩、拼图。

源码运行：python 小工具/图片工具箱.py
"""

import _boot


def _open(root):
    import image_tools

    app = image_tools.ImageToolApp(root)
    if app.win is not None:
        _boot.bind_close(app.win, root)


_boot.run('图片工具箱', _open)
