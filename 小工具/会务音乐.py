# -*- coding: utf-8 -*-
"""会务音乐（独立小工具）：本地暖场音乐 + 在线音乐试听/下载。

源码运行：python 小工具/会务音乐.py

说明：界面目前仍长在主程序里（会务助手.py 的 _open_music），
这里用「单工具模式」打开它——主界面藏在屏幕外，只弹音乐窗口。
"""

import _boot

_boot.run_host_tool('会务音乐', '_open_music', '_music_win')
