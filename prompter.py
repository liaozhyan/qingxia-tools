# -*- coding: utf-8 -*-
"""内置提词器（参考 QPrompt 核心能力，兼顾极简轻量）。

编辑窗 + 全屏演示窗（+ 可选副屏投屏镜像）：
- 外观：多套深色主题、当前行高亮色带、可选背景图片；
- 富文本：任意选中片段可设 加粗/颜色/字号，演示按格式逐片段渲染；
- 滚动：连续滚动 / 逐行翻页 双模式；
- 信息：已播放计时（顶部大号显示）、倒计时、进度条、章节标题一键跳转；
- 快捷：空格/↑/↓/PgUp/PgDn/数字键/T/Tab/单击/ESC。

字号/字体直接按需新建 tkfont.Font（不用 set_font_override），天然不污染主界面字体。
"""

import ctypes
import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import font as tkfont
from tkinter import ttk

import ui_theme

try:
    import video_player as _video
except Exception:  # pragma: no cover
    _video = None

try:
    from PIL import Image, ImageEnhance, ImageTk
    _PIL = True
except Exception:  # pragma: no cover
    _PIL = False


# ---------- 主题：name -> (背景, 正文, 高亮, 色带) ----------
THEMES = {
    '曜石黑': ('#0c101c', '#ffffff', '#8fd0ff', '#1e2a4a'),
    '墨绿':   ('#0d1f17', '#e8f5ec', '#7ee0a8', '#153325'),
    '深蓝':   ('#0b1230', '#e8ecff', '#9db8ff', '#172448'),
    '暗紫':   ('#1a1030', '#f0eaff', '#c9a8ff', '#2b1a4a'),
}
DIRECTIONS = ('正常',)

TICK_MS = 16                # 连续滚动刷新间隔（约 60fps）
ROW_PAD = 12

FONT_MIN = 24
FONT_MAX = 140
FONT_STEP = 4
FONT_DEFAULT = 72

SPEED_MIN = 1
SPEED_MAX = 12
SPEED_DEFAULT = 7
SPEED_FAC = 0.12           # 速度值 → 每帧滚动像素：1 档约 7px/s（很慢）
DIGIT_FACTOR = 0.6          # 数字键节奏：speed = SPEED_DEFAULT * key * factor
PROGRESS_H = 4              # 底部进度条高度

# ---- 富文本片段格式（支持任意选中文字的加粗/颜色/字号）----
FT_COLORS = {'默认': None, '白': '#ffffff', '黄': '#ffd400',
             '红': '#ff5a5a', '绿': '#7ee0a8', '蓝': '#8fd0ff'}
SIZE_MIN = -3               # 字号档位（相对基准）
SIZE_MAX = 5
SIZE_STEP = 0.12            # 演示时每档字号增幅比例
EDITOR_BASE = 20            # 编辑窗基准字号(px)


def _sz_editor(level):
    return max(10, EDITOR_BASE + level * 4)


def _sz_demo(base, level):
    return max(12, int(base * (1 + SIZE_STEP * level)))


def _row_text(row):
    return (''.join(r['t'] for r in row)) if row else ''


def _ffn(master, size):
    try:
        fam = ui_theme.font_family()
    except Exception:
        fam = 'Microsoft YaHei'
    return tkfont.Font(master, family=fam, size=-size)


def _is_section(row):
    s = row.strip()
    return s.startswith('==') or s.startswith('§') or (
        len(s) >= 3 and set(s) <= set('=-~#'))


_GWL_STYLE = -16
_WS_CAPTION = 0x00C00000
_WS_THICKFRAME = 0x00040000
_SWP_FRAMECHANGED = 0x0020
_SWP_NOZORDER = 0x0004


def _win_borderless(win, left, top, w, h):
    """把受管理的 Toplevel 变成无边框全屏窗（Win32）。

    注意：不能用 overrideredirect(True)——实测部分副屏/虚拟屏上
    overrideredirect 窗口不被 DWM 合成，屏上只剩壁纸。受管理窗口
    剥掉 WS_CAPTION/WS_THICKFRAME 后 DWM 正常合成，无边框等效。
    """
    user32 = ctypes.windll.user32
    hwnd = user32.GetAncestor(win.winfo_id(), 2)      # GA_ROOT
    if not hwnd:
        return
    style = user32.GetWindowLongPtrW(hwnd, _GWL_STYLE)
    user32.SetWindowLongPtrW(
        hwnd, _GWL_STYLE, style & ~(_WS_CAPTION | _WS_THICKFRAME))
    user32.SetWindowPos(hwnd, 0, int(left), int(top), int(w), int(h),
                        _SWP_FRAMECHANGED | _SWP_NOZORDER)


class PrompterApp:
    """提词器：编辑窗 + 全屏演示窗（可投副屏镜像）。master 为主程序根窗口。"""

    def __init__(self, master):
        self.master = master
        self.win = None
        self.demo = None
        self.demo_canvas = None
        self.mirror = None
        self.mirror_canvas = None
        self._after = None
        self._sec_after = None
        self._scroll = 0.0
        self._cnt = 0
        self._paused = False
        self.speed_now = float(SPEED_DEFAULT)
        self.step_index = 0
        self.vis = []
        self.max_disp = 1
        self.row_h = 1
        self._sections = []
        self._bgimg = None
        self._bg_drawn = False
        self._bg_path = ''
        self.theme_name = '曜石黑'
        self.mode = 'smooth'
        self.direction = '正常'
        self.align = 'center'                 # 台词对齐：left / center / right
        self.show_timer = True
        self._t0 = None                       # 演示开始时刻（计时用）
        self.mirror_on = False
        self._monitors = []
        self._mon_idx = 0
        self._mirror_rect = None
        self.font_size = FONT_DEFAULT
        self.start_line = 0
        self._build_editor()

    # ================= 编辑窗 =================
    def _build_editor(self):
        pwin = tk.Toplevel(self.master)
        pwin.title('提词器')
        pwin.configure(bg=ui_theme.PALETTE['bg'])
        _center(pwin, 840, 620)
        pwin.resizable(True, True)
        try:
            ui_theme.style_window(pwin)
        except Exception:
            pass
        pwin.protocol('WM_DELETE_WINDOW', self.close)

        bar = tk.Frame(pwin, bg=ui_theme.PALETTE['bg'])
        bar.pack(fill=tk.X, padx=14, pady=(12, 6))

        def rowfr(pady=(0, 0)):
            """每行独立 Frame：各行网格列宽互不干扰（速度滑块撑宽
            第0行第5列时，不会把下面行的单选框/倒计时强制分散）。"""
            f = tk.Frame(bar, bg=ui_theme.PALETTE['bg'])
            f.pack(fill=tk.X, pady=pady)
            return f

        def lbl(fr, txt):
            return tk.Label(fr, text=txt, font=ui_theme.font(9),
                            bg=ui_theme.PALETTE['bg'],
                            fg=ui_theme.PALETTE['muted'])

        p = ui_theme.PALETTE

        def mbtn(fr, text, cmd, bold=False, accent=False, padx=8):
            """统一样式的扁平按钮：浅色底（主色按钮用实心蓝白字）。"""
            if accent:
                bg, fg, a_bg, a_fg = (p['primary'], '#ffffff',
                                      p['primary_dark'], '#ffffff')
            else:
                bg, fg, a_bg, a_fg = (p['surface_alt'], p['text'],
                                      p['hover'], p['primary_dark'])
            return tk.Button(fr, text=text, command=cmd, relief='flat', bd=0,
                             padx=padx, pady=3,
                             font=ui_theme.font(9, 'bold' if bold else 'normal'),
                             bg=bg, fg=fg, activebackground=a_bg,
                             activeforeground=a_fg,
                             highlightthickness=0)

        # 行0：字号 + 速度（速度滑块撑满剩余宽度，仅限本行）
        r0 = rowfr()
        lbl(r0, '字号').grid(row=0, column=0, padx=(0, 5))
        self._fs_label = tk.Label(r0, text=str(self.font_size),
                                  font=ui_theme.font(9, 'bold'), width=4,
                                  bg=ui_theme.PALETTE['bg'],
                                  fg=ui_theme.PALETTE['primary'])
        self._fs_label.grid(row=0, column=1, padx=(0, 4))
        mbtn(r0, '-', lambda: self._adj_font(-1), bold=True, padx=9)\
            .grid(row=0, column=2)
        mbtn(r0, '+', lambda: self._adj_font(1), bold=True, padx=9)\
            .grid(row=0, column=3, padx=(1, 14))
        lbl(r0, '速度').grid(row=0, column=4, padx=(0, 5))
        self._speed = tk.IntVar(value=SPEED_DEFAULT)
        tk.Scale(r0, from_=SPEED_MIN, to=SPEED_MAX, orient=tk.HORIZONTAL,
                 variable=self._speed, showvalue=False,
                 bg=ui_theme.PALETTE['bg'], highlightthickness=0,
                 fg=ui_theme.PALETTE['text'],
                 troughcolor=ui_theme.PALETTE['border']
                 ).grid(row=0, column=5, sticky='ew', padx=(0, 8))
        self._speed_str = tk.StringVar(value=str(SPEED_DEFAULT))
        self._speed_entry = tk.Entry(r0, textvariable=self._speed_str, width=5,
                                     justify='center', font=ui_theme.font(9))
        self._speed_entry.grid(row=0, column=6, padx=(0, 2))
        self._speed_entry.bind('<Return>', lambda e: self._apply_speed_input())
        self._speed.trace_add('write', lambda *a: self._speed_str.set(
            str(self._speed.get())))
        r0.columnconfigure(5, weight=1)

        # 行1：主题 / 模式（连续、逐行紧挨着）
        r1 = rowfr((6, 0))
        lbl(r1, '主题').grid(row=0, column=0, padx=(0, 5))
        self._theme_cb = ttk.Combobox(
            r1, values=list(THEMES.keys()), width=8, state='readonly')
        self._theme_cb.set(self.theme_name)
        self._theme_cb.grid(row=0, column=1, columnspan=3, sticky='w')
        lbl(r1, '模式').grid(row=0, column=4, padx=(0, 5))
        self._mode_var = tk.StringVar(value='smooth')
        tk.Radiobutton(r1, text='连续', variable=self._mode_var,
                       value='smooth', bg=ui_theme.PALETTE['bg'],
                       fg=ui_theme.PALETTE['text'],
                       activebackground=ui_theme.PALETTE['bg'])\
            .grid(row=0, column=5, sticky='w', padx=(0, 4))
        tk.Radiobutton(r1, text='逐行', variable=self._mode_var, value='step',
                       bg=ui_theme.PALETTE['bg'], fg=ui_theme.PALETTE['text'],
                       activebackground=ui_theme.PALETTE['bg'])\
            .grid(row=0, column=6, sticky='w', padx=(0, 10))
        r1.columnconfigure(7, weight=1)

        # 行2：计时 / 投屏(目标) / 倒计时 / 背景
        r2 = rowfr((4, 0))
        self._timer_var = tk.BooleanVar(value=True)
        tk.Checkbutton(r2, text='计时', variable=self._timer_var,
                       bg=ui_theme.PALETTE['bg'], fg=ui_theme.PALETTE['text'],
                       activebackground=ui_theme.PALETTE['bg'])\
            .grid(row=0, column=0, columnspan=2, sticky='w')
        self._mirror_var = tk.BooleanVar(value=False)
        tk.Checkbutton(r2, text='投屏', variable=self._mirror_var,
                       command=self._toggle_mirror,
                       bg=ui_theme.PALETTE['bg'], fg=ui_theme.PALETTE['text'],
                       activebackground=ui_theme.PALETTE['bg'])\
            .grid(row=0, column=2, sticky='w')
        tk.Label(r2, text='投到:', font=ui_theme.font(9),
                 bg=ui_theme.PALETTE['bg'],
                 fg=ui_theme.PALETTE['muted']).grid(row=0, column=3,
                                                    padx=(0, 4))
        self._mon_cb = ttk.Combobox(r2, width=16, state='readonly')
        self._mon_cb.grid(row=0, column=4, sticky='w', padx=(0, 12))
        lbl(r2, '倒计时(分)').grid(row=0, column=5, padx=(0, 5))
        self._cnt_min = tk.IntVar(value=0)
        tk.Spinbox(r2, from_=0, to=180, width=5,
                   textvariable=self._cnt_min, justify='center')\
            .grid(row=0, column=6, sticky='w', padx=(0, 12))
        self._bg_label = tk.Label(r2, text='背景：无', font=ui_theme.font(9),
                                  bg=ui_theme.PALETTE['bg'],
                                  fg=ui_theme.PALETTE['muted'])
        self._bg_label.grid(row=0, column=7, columnspan=3, sticky='w')
        r2.columnconfigure(7, weight=1)
        # 初始填充一次显示器下拉（默认选第一块扩展屏）
        self.win = pwin
        self._refresh_monitors()

        # 行3：操作按钮（文本格式工具条放到编辑器上面，见下）
        btns = tk.Frame(pwin, bg=ui_theme.PALETTE['bg'])
        btns.pack(fill=tk.X, padx=14, pady=(6, 6))
        self._go_btn = tk.Button(
            btns, text='▶ 开始演示', font=ui_theme.font(10, 'bold'),
            command=self.start, relief='flat', bd=0, padx=22, pady=5,
            bg=ui_theme.PALETTE['primary'], fg='#ffffff',
            activebackground=ui_theme.PALETTE['primary_dark'],
            activeforeground='#ffffff')
        self._go_btn.pack(side=tk.LEFT)
        tk.Button(btns, text='导入文稿…', command=self._import_text,
                  relief='flat', bd=0, padx=12, pady=5,
                  bg=ui_theme.PALETTE['surface_alt'],
                  fg=ui_theme.PALETTE['text']).pack(side=tk.LEFT, padx=(8, 0))
        tk.Button(btns, text='选背景图…', command=self._choose_bg,
                  relief='flat', bd=0, padx=12, pady=5,
                  bg=ui_theme.PALETTE['surface_alt'],
                  fg=ui_theme.PALETTE['text']).pack(side=tk.LEFT, padx=(6, 0))
        tk.Button(btns, text='清空', command=self.clear_text, relief='flat',
                  bd=0, padx=12, pady=5,
                  bg=ui_theme.PALETTE['surface_alt'],
                  fg=ui_theme.PALETTE['muted']).pack(side=tk.LEFT, padx=(6, 0))

        # 精简富文本工具条（紧贴编辑器上方，先选中台词再设置）
        fmt = tk.Frame(pwin, bg=ui_theme.PALETTE['bg'])
        fmt.pack(fill=tk.X, padx=14, pady=(0, 4))

        def fbtn(text, cmd, bold=False, accent=False, padx=7):
            if accent:
                bg, fg, a_bg, a_fg = (p['primary'], '#ffffff',
                                      p['primary_dark'], '#ffffff')
            else:
                bg, fg, a_bg, a_fg = (p['surface_alt'], p['text'],
                                      p['hover'], p['primary_dark'])
            return tk.Button(fmt, text=text, command=cmd, relief='flat', bd=0,
                             padx=padx, pady=2,
                             font=ui_theme.font(9, 'bold' if bold else 'normal'),
                             bg=bg, fg=fg, activebackground=a_bg,
                             activeforeground=a_fg, highlightthickness=0)

        fbtn('B', self._fmt_bold, bold=True).pack(side=tk.LEFT)
        fbtn('斜', lambda: self._fmtstyle('ft_i'), bold=True,
             padx=6).pack(side=tk.LEFT, padx=(0, 0))
        fbtn('U', lambda: self._fmtstyle('ft_u'), bold=True,
             padx=8).pack(side=tk.LEFT, padx=(0, 0))
        fbtn('S', lambda: self._fmtstyle('ft_s'), padx=8)\
            .pack(side=tk.LEFT, padx=(0, 4))
        self._color_cb = ttk.Combobox(fmt, values=list(FT_COLORS.keys()),
                                      width=6, state='readonly')
        self._color_cb.set('默认')
        self._color_cb.pack(side=tk.LEFT, padx=(0, 4))
        self._color_cb.bind('<<ComboboxSelected>>',
                            lambda e: self._fmt_color(self._color_cb.get()))
        fbtn('字号+', lambda: self._fmt_size(+1)).pack(side=tk.LEFT, padx=(0, 4))
        fbtn('字号-', lambda: self._fmt_size(-1)).pack(side=tk.LEFT, padx=(0, 4))
        fbtn('清除格式', self._fmt_clear, accent=True).pack(side=tk.LEFT,
                                                           padx=(0, 8))
        tk.Label(fmt, text='对齐', font=ui_theme.font(8),
                 bg=ui_theme.PALETTE['bg'],
                 fg=ui_theme.PALETTE['muted']).pack(side=tk.LEFT, padx=(0, 2))
        self._align_btns = {}
        for m, label in (('left', '左对齐'), ('center', '居中'),
                         ('right', '右对齐')):
            b = fbtn(label, lambda mm=m: self._set_align(mm), padx=7)
            b.pack(side=tk.LEFT, padx=(1, 1))
            self._align_btns[m] = b
        self._set_align('center')
        tk.Label(fmt, text='先选中台词文字再设格式', font=ui_theme.font(8),
                 bg=ui_theme.PALETTE['bg'],
                 fg=ui_theme.PALETTE['faint']).pack(side=tk.LEFT, padx=(8, 0))

        # 台词文本框
        self._text = tk.Text(pwin, wrap='word', undo=True, relief='flat',
                             bd=0, padx=14, pady=10, font=_ffn(pwin, 20),
                             bg=ui_theme.PALETTE['surface'],
                             fg=ui_theme.PALETTE['text'],
                             insertbackground=ui_theme.PALETTE['primary'],
                             highlightthickness=1,
                             highlightbackground=ui_theme.PALETTE['border'],
                             highlightcolor=ui_theme.PALETTE['primary'])
        self._text.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 12))
        self._configure_ftags()
        self._text.focus_set()

        tk.Label(pwin,
                 text='演示快捷键：空格/单击=暂停·继续   ↑↓=调速·翻行   '
                      'PgUp/PgDn=跳转   1-9=节奏    T=连续/逐行   Tab=下节   '
                      'ESC=退出',
                 font=ui_theme.font(8), fg=ui_theme.PALETTE['faint'],
                 bg=ui_theme.PALETTE['bg']).pack(side=tk.BOTTOM, pady=(0, 6))

        self.win = pwin

    def _adj_font(self, d):
        self.font_size = max(FONT_MIN, min(FONT_MAX,
                                           self.font_size + d * FONT_STEP))
        self._fs_label.config(text=str(self.font_size))

    # ---- 富文本编辑：加粗 / 颜色 / 字号 / 清除 ----
    def _configure_ftags(self):
        fam = ui_theme.font_family()
        try:
            self._text.tag_configure('ft_bold',
                                     font=(fam, EDITOR_BASE, 'bold'))
            self._text.tag_configure('ft_i', font=(fam, EDITOR_BASE, 'italic'))
            self._text.tag_configure('ft_u', underline=True)
            self._text.tag_configure('ft_s', overstrike=True)
            for lvl in range(SIZE_MIN, SIZE_MAX + 1):
                self._text.tag_configure('ft_sz%d' % lvl,
                                         font=(fam, _sz_editor(lvl)))
            for key, col in FT_COLORS.items():
                if col:
                    self._text.tag_configure('ft_c_%s' % key, foreground=col)
        except Exception:
            pass

    def _sel_range(self):
        try:
            r = self._text.tag_ranges('sel')
        except Exception:
            r = None
        if not r:
            messagebox.showinfo('提示', '请先在台词文本中选中要设置的文字。',
                                parent=self.win)
            return None
        return r[0], r[1]

    def _tag_covers(self, tag, idx):
        try:
            rr = self._text.tag_ranges(tag)
            it = iter(rr)
            for a, b in zip(it, it):
                if self._text.compare(a, '<=', idx) and \
                        self._text.compare(b, '>', idx):
                    return True
        except Exception:
            pass
        return False

    def _remove_kind(self, prefix, st, en):
        if prefix == 'ft_sz':
            for lvl in range(SIZE_MIN, SIZE_MAX + 1):
                try:
                    self._text.tag_remove('ft_sz%d' % lvl, st, en)
                except Exception:
                    pass
        else:
            for key in FT_COLORS:
                try:
                    self._text.tag_remove('ft_c_%s' % key, st, en)
                except Exception:
                    pass

    def _fmt_bold(self):
        r = self._sel_range()
        if not r:
            return
        st, en = r
        try:
            rr = self._text.tag_ranges('ft_bold')
            it = iter(rr)
            covers = False
            for a, b in zip(it, it):
                if self._text.compare(b, '>', st) and \
                        self._text.compare(a, '<', en):
                    covers = True
                    break
        except Exception:
            covers = False
        if covers:
            self._text.tag_remove('ft_bold', st, en)
        else:
            self._text.tag_add('ft_bold', st, en)

    def _fmtstyle(self, tag):
        """切换到斜体/下划线/删除线（选区覆盖则该标签翻转）。"""
        r = self._sel_range()
        if not r:
            return
        st, en = r
        try:
            rr = self._text.tag_ranges(tag)
            it = iter(rr)
            covers = False
            for a, b in zip(it, it):
                if self._text.compare(b, '>', st) and \
                        self._text.compare(a, '<', en):
                    covers = True
                    break
        except Exception:
            covers = False
        if covers:
            self._text.tag_remove(tag, st, en)
        else:
            self._text.tag_add(tag, st, en)

    def _set_align(self, mode):
        self.align = mode
        p = ui_theme.PALETTE
        for m, b in getattr(self, '_align_btns', {}).items():
            on = (m == mode)
            b.configure(bg=p['primary'] if on else p['surface_alt'],
                        fg='#ffffff' if on else p['text'])

    def _fmt_color(self, key):
        r = self._sel_range()
        if not r:
            self._color_cb.set('默认')
            return
        st, en = r
        self._remove_kind('ft_c_', st, en)
        if key and key != '默认':
            self._text.tag_add('ft_c_%s' % key, st, en)

    def _fmt_size(self, dv):
        r = self._sel_range()
        if not r:
            return
        st, en = r
        cur = 0
        for lvl in range(SIZE_MAX, SIZE_MIN - 1, -1):
            idx = st
            if self._tag_covers('ft_sz%d' % lvl, idx):
                cur = lvl
                break
        new = max(SIZE_MIN, min(SIZE_MAX, cur + dv))
        self._remove_kind('ft_sz', st, en)
        self._text.tag_add('ft_sz%d' % new, st, en)

    def _fmt_clear(self):
        r = self._sel_range()
        if not r:
            return
        st, en = r
        try:
            self._text.tag_remove('ft_bold', st, en)
        except Exception:
            pass
        self._remove_kind('ft_sz', st, en)
        self._remove_kind('ft_c_', st, en)

    def _import_text(self):
        path = filedialog.askopenfilename(
            parent=self.win, title='导入文稿',
            filetypes=[('文本', '*.txt'), ('所有文件', '*.*')])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                data = f.read()
        except Exception as e:
            messagebox.showerror('错误', '读取失败：%s' % e, parent=self.win)
            return
        self._text.delete('1.0', 'end')
        self._text.insert('1.0', data)

    def _choose_bg(self):
        path = filedialog.askopenfilename(
            parent=self.win, title='选择提词背景图',
            filetypes=[('图片', '*.png *.jpg *.jpeg *.bmp'), ('所有文件', '*.*')])
        if not path:
            return
        if not _PIL:
            messagebox.showwarning('提示', '未安装 Pillow，无法使用背景图。',
                                   parent=self.win)
            return
        self._bg_path = path
        self._bg_label.config(text='背景：' + os.path.basename(path))

    def clear_text(self):
        self._text.delete('1.0', 'end')

    def get_text(self):
        try:
            return self._text.get('1.0', 'end-1c')
        except tk.TclError:
            return ''

    # ================= 开始演示 =================
    def _get_runs(self, fsz):
        """从编辑窗 dump 出带格式的 runs，按行切分。

        每个 run 为 dict: {'t':文本, 'level':字号档, 'bold':bool, 'c':颜色key/None}
        只产出逻辑行（未做宽度换行）。
        """
        try:
            data = self._text.dump('1.0', 'end-1c')
        except Exception:
            data = []
        rows = []
        cur = []
        bold = italic = under = strike = False
        level = 0
        col = None
        for typ, val, _ in data:
            if typ == 'tagon':
                if val == 'ft_bold':
                    bold = True
                elif val == 'ft_i':
                    italic = True
                elif val == 'ft_u':
                    under = True
                elif val == 'ft_s':
                    strike = True
                elif val.startswith('ft_sz'):
                    try:
                        level = int(val[5:])
                    except Exception:
                        level = 0
                elif val.startswith('ft_c_'):
                    col = val[5:]
            elif typ == 'tagoff':
                if val == 'ft_bold':
                    bold = False
                elif val == 'ft_i':
                    italic = False
                elif val == 'ft_u':
                    under = False
                elif val == 'ft_s':
                    strike = False
                elif val.startswith('ft_sz'):
                    level = 0
                elif val.startswith('ft_c_'):
                    col = None
            elif typ == 'text':
                parts = val.split('\n')
                for k, ps in enumerate(parts):
                    if ps:
                        cur.append({'t': ps, 'level': level, 'bold': bold,
                                    'i': italic, 'u': under, 's': strike,
                                    'c': col})
                    if k < len(parts) - 1:
                        rows.append(cur[:])
                        cur = []
        if cur or not rows:
            rows.append(cur[:])
        return rows

    def _fpool(self, px, b, i=False, u=False, s=False):
        key = (px, b, i, u, s)
        f = self._fontpool.get(key)
        if f is None:
            try:
                f = _ffn(self.demo, px)
                f.configure(weight='bold' if b else 'normal',
                            slant='italic' if i else 'roman',
                            underline=u, overstrike=s)
            except Exception:
                f = None
            self._fontpool[key] = f
        return f

    def _wrap_all_rows(self, master, rows, fsz, maxw):
        """把每个逻辑行按实际片段字号逐字换行成视觉行。

        返回 list[ list[run] ]，每 run {'t','f','c'}，f 为已解析的 Font。
        """
        out = []
        # 逐字 width 缓存：同一 (字体, 字符) 只 measure 一次，
        # 避免长稿反复测量（该机 measure ~4ms/次，会造成整机卡死）
        mc = {}

        def _mwidth(f, ch):
            fd = mc.get(id(f))
            if fd is None:
                fd = {}
                mc[id(f)] = fd
            w = fd.get(ch)
            if w is None:
                try:
                    w = f.measure(ch)
                except Exception:
                    w = 1
                fd[ch] = w
            return w

        for row in rows:
            chars = []
            for run in row:
                px = _sz_demo(fsz, run['level'])
                f = self._fpool(px, run['bold'], run.get('i', False),
                                run.get('u', False), run.get('s', False))
                if f is None:
                    continue
                for ch in run['t']:
                    chars.append((ch, f, run['c']))
            wrapped = []
            line_runs = []
            w = 0.0
            for ch, f, c in chars:
                cw = _mwidth(f, ch)
                if line_runs and w + cw > maxw:
                    wrapped.append(line_runs)
                    line_runs = []
                    w = 0.0
                line_runs.append((ch, f, c))
                w += cw
            if line_runs:
                wrapped.append(line_runs)
            # 压缩相邻同格式字符成 run
            for rr in wrapped:
                groups = []
                for ch, f, c in rr:
                    if groups and groups[-1][1] is f and groups[-1][2] == c:
                        groups[-1][0] += ch
                    else:
                        groups.append([ch, f, c])
                out.append([{'t': g[0], 'f': g[1], 'c': g[2]}
                            for g in groups])
        return out

    def start(self):
        if self.demo is not None:
            return

        self.theme_name = self._theme_cb.get()
        self.mode = self._mode_var.get()
        self.direction = '正常'
        try:
            self.show_timer = bool(self._timer_var.get())
        except Exception:
            self.show_timer = True
        try:
            self.mirror_on = bool(self._mirror_var.get())
        except Exception:
            self.mirror_on = False
        self.font_size = max(FONT_MIN, min(FONT_MAX, self.font_size))
        fsz = self.font_size
        self.speed_now = float(self._param_speed())

        sr_rows = self._get_runs(fsz)
        if not any(sr_rows):        # 空稿时 _get_runs 返回 [[]]，不能用 not 判
            messagebox.showinfo('提示', '请先输入台词，再开始演示。',
                                parent=self.win)
            return

        # 投屏优先：勾选投屏且屏幕可用时先开镜像窗，演示本体退居隐藏——
        # 主屏不再被全屏演示窗占据，画面只出现在所选显示器上
        mirror_shown = False
        if self.mirror_on:
            try:
                self._refresh_monitors()
                self._mirror_rect = self._pick_mirror_rect()
                self._open_mirror(0, 0)
                mirror_shown = self.mirror_canvas is not None
            except Exception:
                self.mirror_on = False    # 投屏失败不影响演示本体

        d = tk.Toplevel(self.master)
        try:
            ww = d.winfo_screenwidth()
            wh = d.winfo_screenheight()
        except Exception:
            ww, wh = 1920, 1080
        d.geometry('%dx%d+0+0' % (ww, wh))
        d.overrideredirect(True)
        d.attributes('-topmost', True)
        if mirror_shown:
            d.withdraw()          # 仅副屏出画面：演示窗保持隐藏（播放引擎）
        else:
            d.lift()
            try:
                d.focus_force()
            except Exception:
                pass
        canvas = tk.Canvas(d, highlightthickness=0, bd=0)
        canvas.pack(fill=tk.BOTH, expand=True)

        self.demo = d
        self.demo_canvas = canvas
        self.scroll_w, self.scroll_h = ww, wh

        self._fontpool = {}
        self._fscale_cache = {}     # 缩放字体缓存随场重置
        self.row_h = max(6, int(fsz * 1.45))
        # 排版按“实际出画面的那块屏”的宽度计算：投屏时按镜像屏宽度换行，
        # 同字号铺满副屏；不投屏时按主屏宽度。隐藏的另一面按比例缩放渲染。
        self._wrap_w = self.mirror_w if mirror_shown else ww
        self._relayout(self._wrap_w, sr_rows)

        sl = max(0, min(self.start_line, self.max_disp - 1))
        self._scroll = sl * self.row_h
        self.step_index = sl
        self._paused = False

        self._cnt = max(0, int(self._cnt_min.get())) * 60
        self._t0 = time.time()                # 演示计时起点

        self._prep_bg(d if self._bg_path else None, ww, wh)

        self._bind_demo_keys(d)
        if mirror_shown:
            self._bind_demo_keys(self.mirror, self.mirror_canvas)

        self._render_all()
        self._after = d.after(TICK_MS, self._loop)
        self._sec_after = d.after(1000, self._second_tick)
        self._set_go_btn(True)

    def _relayout(self, target_w, sr_rows=None):
        """按出画面屏幕的宽度重排文本行，并保留当前阅读进度比例。

        投屏开启时 target_w=镜像屏宽：同字号直接铺满副屏；关闭时回主屏宽。
        """
        fsz = self.font_size
        margin = max(40, int(fsz * 0.5))
        maxw = max(120, target_w - margin * 2 - ROW_PAD * 2)
        if sr_rows is None:
            sr_rows = self._get_runs(fsz)
        if not any(sr_rows):
            return
        try:
            frac = self._scroll / float(max(1, (self.max_disp - 1)
                                            * self.row_h))
        except Exception:
            frac = 0.0
        frac = max(0.0, min(1.0, frac))
        self.vis = self._wrap_all_rows(self.demo, sr_rows, fsz, maxw)
        self.max_disp = max(1, len(self.vis))
        self._sections = [i for i, row in enumerate(self.vis)
                          if _is_section(_row_text(row))]
        self._scroll = frac * (self.max_disp - 1) * self.row_h
        self.step_index = max(0, min(int(frac * (self.max_disp - 1) + 0.5),
                                     self.max_disp - 1))

    def _bind_demo_keys(self, w, canvas=None):
        """演示快捷键：演示窗与镜像窗共用同一套控制。"""
        w.bind('<space>', lambda e: self._on_space())
        w.bind('<Up>', lambda e: self._on_up())
        w.bind('<Down>', lambda e: self._on_down())
        w.bind('<Prior>', lambda e: self._on_page(-1))
        w.bind('<Next>', lambda e: self._on_page(1))
        w.bind('<Escape>', lambda e: self.stop())
        w.bind('<Key>', self._on_key)
        w.bind('<Tab>', lambda e: self._jump_to(True))
        w.bind('<Shift-Tab>', lambda e: self._jump_to(False))
        if canvas is not None:
            canvas.bind('<Button-1>', lambda e: self._on_space())

    def _set_go_btn(self, running):
        """演示中把编辑窗按钮切为“停止演示”（镜像模式下主屏没有演示窗，
        这是唯一的停止入口之一）。"""
        try:
            if running:
                self._go_btn.config(text='■ 停止演示', command=self.stop)
            else:
                self._go_btn.config(text='▶ 开始演示', command=self.start)
        except Exception:
            pass

    def _prep_bg(self, d, ww, wh):
        if not d or not _PIL or not self._bg_path:
            self._bgimg = None
            return
        try:
            img = Image.open(self._bg_path).convert('RGB')
            img = img.resize((int(ww), int(wh)))
            img = ImageEnhance.Brightness(img).enhance(0.45)
            self._bgimg = ImageTk.PhotoImage(img)
            self._bg_drawn = False
        except Exception:
            self._bgimg = None

    def _refresh_monitors(self):
        """刷新显示器下拉框；默认选中第一块扩展屏（既有组合则保留选择）。"""
        if _video is None:
            self._monitors = []
            return
        try:
            mons = _video.list_monitors()
        except Exception:
            mons = []
        self._monitors = mons
        try:
            names = [m['name'] for m in mons] or ['无扩展屏']
            self._mon_cb.config(values=names)
            if not mons:
                return
            cur = self._mon_cb.current()
            if cur < 0 or cur >= len(mons):
                for i, m in enumerate(mons):
                    if not m.get('primary'):
                        cur = i
                        break
                else:
                    cur = 0
                self._mon_cb.current(cur)
            self._mon_idx = cur
        except Exception:
            pass

    def _pick_mirror_rect(self):
        """从下拉选中的显示器取 rect；未选任何显示器则自动挑第一块扩展屏。"""
        mons = self._monitors
        if not mons:
            return None
        try:
            cur = self._mon_cb.current()
        except Exception:
            cur = self._mon_idx
        if cur < 0 or cur >= len(mons):
            cur = 0
            for i, m in enumerate(mons):
                if not m.get('primary'):
                    cur = i
                    break
        return mons[cur].get('rect')

    def _open_mirror(self, ww, wh):
        if _video is None:
            messagebox.showwarning('提示', '无法检测副屏（缺少显示模块）。',
                                   parent=self.win)
            self._mirror_var.set(False)
            self.mirror_on = False
            return
        rect = getattr(self, '_mirror_rect', None) or self._pick_mirror_rect()
        if rect is None:
            messagebox.showinfo('提示', '未检测到显示器，跳过投屏。',
                                parent=self.win)
            self._mirror_var.set(False)
            self.mirror_on = False
            return
        if not any(not m.get('primary') for m in self._monitors):
            # 系统里只有主屏：镜像窗只会盖在演示窗上，没有意义
            messagebox.showwarning(
                '提示', '未检测到扩展屏。\n请按 Win+P 选择“扩展”，'
                '把副屏/投影设为扩展显示后，再重新勾选投屏。',
                parent=self.win)
            self._mirror_var.set(False)
            self.mirror_on = False
            return
        left, top, right, bottom = rect
        mw, mh = right - left, bottom - top
        win = tk.Toplevel(self.master)
        # 关键：不能用 overrideredirect(True) 做无边框——本机实测副屏上
        # overrideredirect 窗口不被 DWM 合成（屏上只有壁纸）。改用受管理
        # 窗口先定位显示，再剥 WS_CAPTION/WS_THICKFRAME 样式实现无边框。
        win.withdraw()
        win.geometry('%dx%d+%d+%d' % (mw, mh, left, top))
        mc = tk.Canvas(win, highlightthickness=0, bd=0)
        mc.pack(fill=tk.BOTH, expand=True)
        win.update_idletasks()
        win.deiconify()
        win.attributes('-topmost', True)
        try:
            _win_borderless(win, left, top, mw, mh)
        except Exception:
            pass          # 剥样式失败则带边框投屏，功能不受影响
        win.lift()
        self.mirror = win
        self.mirror_canvas = mc
        self.mirror_w, self.mirror_h = mw, mh

    def _toggle_mirror(self):
        """实时开/关投屏：演示中立即生效；未演示时记住勾选，开始演示时生效。"""
        if self.demo is None:
            return
        if bool(self._mirror_var.get()):
            if self.mirror is not None:
                return
            try:
                self._refresh_monitors()
            except Exception:
                pass
            self._mirror_rect = self._pick_mirror_rect()
            if self._mirror_rect is None:
                self._mirror_var.set(False)   # 没有可用屏幕，回弹勾选
            self._open_mirror(self.scroll_w, self.scroll_h)
            if self.mirror_canvas is not None:
                self._bind_demo_keys(self.mirror, self.mirror_canvas)
                try:
                    # 改按镜像屏宽度重排，同字号铺满副屏（而非缩放副本）
                    self._wrap_w = self.mirror_w
                    self._relayout(self._wrap_w)
                    self._render_all()
                except Exception:
                    pass
                try:
                    if self.demo is not None:
                        self.demo.withdraw()   # 副屏出画面后主屏退居隐藏
                except Exception:
                    pass
        else:
            if self.mirror is not None:
                try:
                    self.mirror.destroy()
                except Exception:
                    pass
            self.mirror = None
            self.mirror_canvas = None
            self._mirror_rect = None
            try:
                # 回到主屏宽度重排，再刷新隐藏面
                self._wrap_w = self.scroll_w
                self._relayout(self._wrap_w)
                self._render_all()
            except Exception:
                pass
            # 没了镜像就要把演示窗放回主屏，否则现场黑屏没得看
            try:
                if self.demo is not None:
                    self.demo.deiconify()
                    self.demo.lift()
                    self.demo.focus_force()
            except Exception:
                pass

    # ================= 滚动 / 翻页 =================
    def _on_space(self):
        if self.mode == 'step':
            if self.step_index < self.max_disp - 1:
                self.step_index += 1
            self._render_all()
        else:
            self._paused = not self._paused

    def _on_up(self):
        if self.mode == 'step':
            if self.step_index > 0:
                self.step_index -= 1
            self._render_all()
        else:
            self.change_speed(+2)

    def _on_down(self):
        if self.mode == 'step':
            if self.step_index < self.max_disp - 1:
                self.step_index += 1
            self._render_all()
        else:
            self.change_speed(-2)

    def _on_page(self, sign):
        if self.mode == 'step':
            self.step_index = max(0, min(self.max_disp - 1,
                                         self.step_index + sign * 5))
            self._render_all()
        else:
            n = int(self.scroll_h / self.row_h * 0.8)
            self._scroll = max(0.0, min((self.max_disp - 1) * self.row_h,
                                        self._scroll + sign * n * self.row_h))

    def _on_key(self, e):
        if e.char in '123456789':
            k = int(e.char)
            self.speed_now = max(SPEED_MIN, min(
                SPEED_MAX, SPEED_DEFAULT * k * DIGIT_FACTOR))
            try:
                self._speed.set(int(round(self.speed_now)))
            except Exception:
                pass
            return 'break'
        if e.char in 'tT':
            self._toggle_mode()
            return 'break'
        return None

    def _toggle_mode(self):
        self.mode = 'step' if self.mode == 'smooth' else 'smooth'
        if self.mode == 'step':
            cur = int(self._scroll // self.row_h)
            self.step_index = max(0, min(self.max_disp - 1, cur))
        self._render_all()

    def _jump_to(self, forward):
        cur = self._cur_pos()
        if forward:
            nxt = [i for i in self._sections if i > cur]
            idx = nxt[0] if nxt else 0
        else:
            prv = [i for i in self._sections if i < cur]
            idx = prv[-1] if prv else (self._sections[-1] if
                                       self._sections else 0)
        if self.mode == 'step':
            self.step_index = idx
        else:
            self._scroll = idx * self.row_h
        self._render_all()

    def _cur_pos(self):
        return self.step_index if self.mode == 'step' \
            else int(self._scroll // self.row_h)

    def change_speed(self, dv):
        self.speed_now = max(SPEED_MIN,
                             min(SPEED_MAX, self.speed_now + dv))
        try:
            self._speed.set(int(round(self.speed_now)))
        except Exception:
            pass

    def _apply_speed_input(self):
        """把数值输入框里的值（回车）解析并夹到合法区间，同步到滑块。"""
        try:
            v = int(float(self._speed_str.get().strip()))
        except Exception:
            v = self._speed.get()
        v = max(SPEED_MIN, min(SPEED_MAX, v))
        self._speed.set(v)
        self.speed_now = float(v)

    def _param_speed(self):
        return max(SPEED_MIN, min(SPEED_MAX, self._speed.get()))

    # ================= 帧循环 =================
    def _loop(self):
        if self.demo is None:
            return
        if self.mode == 'smooth' and not self._paused:
            self._scroll += self.speed_now * SPEED_FAC
            total = (self.max_disp - 1) * self.row_h
            if self._scroll > total:
                self._scroll = total
        self._render_all()
        self._after = self.demo.after(TICK_MS, self._loop)

    def _second_tick(self):
        if self.demo is None:
            return
        if self._cnt > 0:
            self._cnt -= 1
        self._refresh_info()
        self._sec_after = self.demo.after(1000, self._second_tick)

    def _render_all(self):
        if self.demo_canvas is not None:
            self._render_frame(self.demo_canvas, self.scroll_w, self.scroll_h,
                               main=True)
        if self.mirror_canvas is not None:
            self._render_frame(self.mirror_canvas, self.mirror_w,
                               self.mirror_h, main=False)

    def _refresh_info(self):
        if self.demo_canvas is not None:
            try:
                self.demo_canvas.delete('info')
            except tk.TclError:
                pass
            self._draw_info(self.demo_canvas, self.scroll_w, self.scroll_h,
                            *self._palette()[:2])
        if self.mirror_canvas is not None:
            try:
                self.mirror_canvas.delete('info')
            except tk.TclError:
                pass
            self._draw_info(self.mirror_canvas, self.mirror_w, self.mirror_h,
                            *self._palette()[:2])

    # ================= 渲染 =================
    def _palette(self):
        return THEMES[self.theme_name]

    def _render_frame(self, canvas, ww, wh, main=True):
        bg, tc, cc, band = self._palette()
        try:
            canvas.configure(bg=bg)
        except tk.TclError:
            return
        if main and self._bgimg is not None and not self._bg_drawn:
            try:
                canvas.create_image(0, 0, anchor='nw', image=self._bgimg,
                                    tags='bg')
                self._bg_drawn = True
            except Exception:
                pass
        canvas.delete('t')
        canvas.delete('info')

        # 排版宽度是 _wrap_w（出画面那块屏）。与其一致的面按原尺寸绘制，
        # 另一面（隐藏的播放引擎）等比缩放，内容带完全对应。
        s = 1.0
        wrap_w = self._wrap_w or self.scroll_w or ww
        if wrap_w and ww:
            s = ww / float(wrap_w)
            if abs(s - 1.0) < 0.001:
                s = 1.0

        if self.mode == 'step':
            self._draw_step(canvas, ww, wh, tc, cc, band, s)
        else:
            self._draw_smooth(canvas, ww, wh, tc, cc, band, s)
        self._draw_info(canvas, ww, wh, tc, cc)

    def _scaled_font(self, f, s):
        """取字体 f 的 s 倍缩放版本（镜像面专用，带缓存）。"""
        if s == 1.0:
            return f
        cache = getattr(self, '_fscale_cache', None)
        if cache is None:
            cache = self._fscale_cache = {}
        key = (id(f), round(s, 3))
        sf = cache.get(key)
        if sf is None:
            try:
                a = f.actual()
                px = max(4, int(round(abs(a.get('size', 12)) * s)))
                sf = tkfont.Font(self.master, family=a.get('family'),
                                 size=-px, weight=a.get('weight'),
                                 slant=a.get('slant'),
                                 underline=a.get('underline'),
                                 overstrike=a.get('overstrike'))
            except Exception:
                sf = f
            cache[key] = sf
        return sf

    def _draw_smooth(self, canvas, ww, wh, tc, cc, band, s=1.0):
        rh = self.row_h * s
        scroll_s = self._scroll * s
        margin = max(40, int(self.font_size * 0.5)) * s
        midx = ww // 2
        center_y = wh // 2
        # 开场偏置：首行从阅读窗最下格升起（而非停在屏幕顶部），
        # 否则刚开始说话时首行已越过阅读窗、被遮罩盖住看不见
        off = center_y + rh
        # 中央 3 行阅读窗：中心行 + 上下各一行，着色带 + 高亮字
        gi_c = int((center_y + scroll_s - off) // rh)   # 中心行号（可为负）
        hot_lo, hot_hi = gi_c - 1, gi_c + 1

        for gi in range(self.max_disp):
            y = gi * rh - scroll_s + off
            if y + rh < 0 or y > wh:
                continue                       # 屏外行直接跳过
            is_hot = hot_lo <= gi <= hot_hi
            if is_hot:
                canvas.create_rectangle(margin, y - 3, ww - margin,
                                        y + rh + 3, fill=band,
                                        stipple='gray50', outline='',
                                        tags='t')
            row = self.vis[gi]
            total = sum(self._scaled_font(r['f'], s).measure(r['t'])
                        for r in row)
            if self.align == 'right':
                x = ww - margin - total
            elif self.align == 'left':
                x = margin
            else:
                x = midx - total / 2.0
            for r in row:
                fill = FT_COLORS.get(r['c'], '') or (cc if is_hot else tc)
                sf = self._scaled_font(r['f'], s)
                canvas.create_text(x, y, anchor='nw', text=r['t'],
                                   font=sf, fill=fill, tags='t')
                x += sf.measure(r['t'])

        # 阅读窗以上（已读）与以下（未读）盖黑色半透明遮罩
        win_top = hot_lo * rh - scroll_s + off - 3
        win_bot = (hot_hi + 1) * rh - scroll_s + off + 3
        if win_top > 0:
            canvas.create_rectangle(-2, -2, ww + 2, win_top, fill='#000000',
                                    stipple='gray50', outline='', tags='t')
        if 0 < win_bot < wh:
            canvas.create_rectangle(-2, win_bot, ww + 2, wh + 2,
                                    fill='#000000', stipple='gray50',
                                    outline='', tags='t')

    def _draw_step(self, canvas, ww, wh, tc, cc, band, s=1.0):
        row = self.vis[self.step_index]
        midx = ww // 2
        midy = wh // 2
        rh = self.row_h * s
        canvas.create_rectangle(60 * s, midy - rh - 8, ww - 60 * s,
                                midy + rh + 8, fill=band,
                                stipple='gray50', outline='', tags='t')
        total = sum(self._scaled_font(r['f'], s).measure(r['t'])
                    for r in row)
        if self.align == 'right':
            x = ww - 60 * s - total
        elif self.align == 'left':
            x = 60 * s
        else:
            x = midx - total / 2.0
        for r in row:
            fill = FT_COLORS.get(r['c'], '') or cc
            sf = self._scaled_font(r['f'], s)
            canvas.create_text(x, midy, anchor='nw', text=r['t'],
                               font=sf, fill=fill, tags='t')
            x += sf.measure(r['t'])

    def _draw_info(self, canvas, ww, wh, tc, cc):
        bg, _, _, _ = self._palette()
        total_row = max(1, (self.max_disp - 1) * self.row_h)
        pos = self._scroll if self.mode == 'smooth' \
            else self.step_index * self.row_h
        ratio = min(1.0, max(0.0, pos / total_row))
        py = wh - PROGRESS_H - 4
        canvas.create_rectangle(24, py, ww - 24, py + PROGRESS_H,
                                fill=_shade(bg, 0.25), outline='',
                                tags='info')
        canvas.create_rectangle(24, py, 24 + (ww - 48) * ratio, py + PROGRESS_H,
                                fill=_shade(tc, -0.1), outline='',
                                tags='info')

        if self.mode == 'step':
            pos_txt = '第 %d / %d 行' % (self.step_index + 1, self.max_disp)
        else:
            pos_txt = '行 %.0f / %d' % (self._scroll / self.row_h + 1,
                                        self.max_disp)
        canvas.create_text(24, wh - PROGRESS_H - 18, anchor='w',
                           text=pos_txt, font=_ffn(canvas, 20),
                           fill=_shade(tc, -0.05), tags='info')

        seg = ''
        if self._cnt > 0:
            seg = '%02d:%02d' % (self._cnt // 60, self._cnt % 60)
        if self.show_timer:
            # 已播放时长（从开始演示起计时），如 02:05 / 1:02:05
            el = max(0, int(time.time() - self._t0)) if self._t0 else 0
            if el >= 3600:
                el_txt = '%d:%02d:%02d' % (el // 3600, el % 3600 // 60,
                                           el % 60)
            else:
                el_txt = '%02d:%02d' % (el // 60, el % 60)
            seg = (seg + '  ' + el_txt) if seg else el_txt
        if seg:
            canvas.create_text(ww - 24, 12, anchor='ne', text=seg,
                               font=_ffn(canvas, 48),
                               fill=('#ff5a5a' if 0 < self._cnt <= 10 else cc),
                               tags='info')

    # ================= 停止 / 关闭 =================
    def stop(self):
        d = self.demo
        self.demo = None
        if self._after is not None and d is not None:
            try:
                d.after_cancel(self._after)
            except Exception:
                pass
            self._after = None
        if self._sec_after is not None and d is not None:
            try:
                d.after_cancel(self._sec_after)
            except Exception:
                pass
            self._sec_after = None
        if self.vis and self.row_h:
            total = (self.max_disp - 1) * self.row_h
            if self.mode == 'step':
                # 逐行模式位置在 step_index（_scroll 不随后行更新）
                at_end = self.step_index >= self.max_disp - 1
                cur_row = self.step_index
            else:
                at_end = self._scroll >= total
                cur_row = int(self._scroll // self.row_h)
            if at_end:
                # 播放完了：下次开始演示从头播放
                self.start_line = 0
            else:
                # 没播完：记住中断行（留 2 行上下文），下次接着播
                self.start_line = min(self.max_disp - 1, max(0, cur_row - 2))
        else:
            self.start_line = 0
        for w in (d, self.mirror):
            if w is None:
                continue
            try:
                w.unbind('<space>')
                w.unbind('<Up>')
                w.unbind('<Down>')
                w.unbind('<Prior>')
                w.unbind('<Next>')
                w.unbind('<Escape>')
                w.unbind('<Key>')
                w.unbind('<Tab>')
                w.unbind('<Shift-Tab>')
                try:
                    w.winfo_children()[0].unbind('<Button-1>')
                except Exception:
                    pass
                w.destroy()
            except Exception:
                pass
        self.demo_canvas = None
        self.mirror = None
        self.mirror_canvas = None
        self._set_go_btn(False)

    def close(self, force=False):
        # X 关闭前二次确认：现场提词中断风险高；force 为主程序退出时免确认
        if not force and self.win is not None:
            try:
                if not ui_theme.confirm_close(self.win, '提词器'):
                    return
            except Exception:
                pass
        self.stop()
        if self.win is not None:
            try:
                self.win.destroy()
            except Exception:
                pass
            self.win = None


def _shade(color, f):
    """颜色微调亮度：f>0 提亮，f<0 压暗，值域 [-1,1]。"""
    try:
        c = color.lstrip('#')
        out = []
        for i in (0, 2, 4):
            v = int(c[i:i + 2], 16)
            v = int(max(0, min(255, v + f * 255)))
            out.append('%02x' % v)
        return '#' + ''.join(out)
    except Exception:
        return color


def _center(win, width, height):
    try:
        w = win.winfo_screenwidth()
        h = win.winfo_screenheight()
    except Exception:
        w, h = 1280, 800
    x = max(0, (w - width) // 2)
    y = max(0, (h - height) // 2)
    win.geometry('%dx%d+%d+%d' % (int(width), int(height), x, y))