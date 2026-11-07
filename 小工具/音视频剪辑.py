# -*- coding: utf-8 -*-
"""音视频剪辑（独立小工具）：剪切拼接、提音频、压缩，需要 ffmpeg。

源码运行：python 小工具/音视频剪辑.py
"""

import _boot


def _open(root):
    import av_tools

    app = av_tools.AvToolApp(root)
    if app.win is not None:
        _boot.bind_close(app.win, root)


_boot.run('音视频剪辑', _open)
