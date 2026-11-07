# -*- coding: utf-8 -*-
"""提词器（独立小工具）。

源码运行：python 小工具/提词器.py
"""

import _boot


def _open(root):
    import prompter

    app = prompter.PrompterApp(root)
    if app.win is not None:
        _boot.bind_close(app.win, root)


_boot.run('提词器', _open)
