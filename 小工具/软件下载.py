# -*- coding: utf-8 -*-
"""软件下载（独立小工具）：走系统 winget 一键装常用软件。

源码运行：python 小工具/软件下载.py
（等价于直接跑根目录的 winget_downloader.py）
"""

import _boot  # noqa: F401  仅为让根目录进 sys.path
import safe_guard


def _main():
    import winget_downloader as wd

    app = wd.App()
    app.mainloop()


if __name__ == '__main__':
    safe_guard.main_guard(app_name='软件下载', gui=True, func=_main)
