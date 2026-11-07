# -*- coding: utf-8 -*-
"""座位牌打印（独立小工具）：批量生成座位牌 / 文件打印 / 云打印。

源码运行：python 小工具/座位牌打印.py
（等价于直接跑根目录的 seat_card.py）

注意：这个工具的界面就是根窗口本身，所以不走 _boot.run 的隐藏根窗口。
"""

import _boot  # noqa: F401  仅为让根目录进 sys.path
import safe_guard


def _main():
    import tkinter as tk

    import seat_card

    root = tk.Tk()
    guard = safe_guard.install(app_name='座位牌打印', gui=True, root=root)
    seat_card.App(root)
    root.mainloop()
    del guard


if __name__ == '__main__':
    safe_guard.main_guard(app_name='座位牌打印', gui=True, func=_main)
