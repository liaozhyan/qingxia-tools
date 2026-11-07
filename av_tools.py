# -*- coding: utf-8 -*-
"""音视频剪辑：会场临时剪辑（ffmpeg 命令行封装，界面零第三方依赖）。

功能：
- 打开视频/音频文件，显示时长、分辨率、音轨信息
- 区间剪切：滑块定起止，无损模式秒出（切点对齐关键帧）/ 精确模式重编码
- 导出区分视频/音频两栏：提取音频（m4a/mp3/wav）、压缩（画质三档）
- 拼接：多文件按顺序合并（同格式 -c copy 秒拼；失败自动转重编码）
- 所有导出都有实时进度条（解析 ffmpeg -progress 输出）

ffmpeg 二进制：优先 imageio_ffmpeg 自带（pip install imageio-ffmpeg），
其次系统 PATH；都没有时给出明确提示。
"""

import os
import re
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import ui_theme
import video_player
from file_inbox import unique_path

_NO_WINDOW = 0x08000000          # CREATE_NO_WINDOW：不弹 ffmpeg 黑窗
_AUDIO_EXT = ('.mp3', '.m4a', '.aac', '.wav', '.wma', '.flac', '.ogg')
_VIDEO_EXT = ('.mp4', '.mov', '.mkv', '.avi', '.ts', '.webm', '.flv')
_MEDIA_EXT = _VIDEO_EXT + _AUDIO_EXT
_CRFS = {'高': 22, '中': 27, '低': 32}


def _center(win, w, h):
    try:
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        win.geometry('%dx%d+%d+%d' % (w, h, max(0, (sw - w) // 2),
                                      max(0, (sh - h) // 2 - 40)))
    except Exception:
        pass


def find_ffmpeg():
    """定位 ffmpeg.exe：imageio_ffmpeg 自带 → 系统 PATH → 常见安装位置。"""
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        pass
    try:
        r = subprocess.run(['where', 'ffmpeg'], capture_output=True,
                           text=True, creationflags=_NO_WINDOW)
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.strip().splitlines()[0].strip()
    except Exception:
        pass
    for pth in (r'C:\ffmpeg\bin\ffmpeg.exe',
                r'C:\Program Files\ffmpeg\bin\ffmpeg.exe'):
        if os.path.exists(pth):
            return pth
    return None


def probe(exe, path):
    """用 ffmpeg -i 探测：返回 (时长秒, 有视频, 有音频, 分辨率)。"""
    r = subprocess.run([exe, '-hide_banner', '-i', path],
                       capture_output=True, text=True, encoding='utf-8',
                       errors='replace', creationflags=_NO_WINDOW)
    err = r.stderr or ''
    dur = 0.0
    m = re.search(r'Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)', err)
    if m:
        dur = int(m.group(1)) * 3600 + int(m.group(2)) * 60 + float(m.group(3))
    has_v = has_a = False
    res = ''
    for line in err.splitlines():
        if 'Stream #' not in line:
            continue
        if ': Video:' in line:
            has_v = True
            rm = re.search(r', (\d{2,5})x(\d{2,5})', line)
            if rm and not res:
                res = '%sx%s' % (rm.group(1), rm.group(2))
        elif ': Audio:' in line:
            has_a = True
    return dur, has_v, has_a, res


def _fmt(t):
    t = max(0.0, float(t))
    h = int(t // 3600)
    m = int(t % 3600 // 60)
    s = t % 60
    return ('%d:%02d:%04.1f' % (h, m, s)) if h else ('%02d:%04.1f' % (m, s))


def _parse_time(txt):
    """'1:23.5' / '83.5' → 秒；解析失败返回 None。"""
    txt = str(txt).strip()
    if not txt:
        return None
    try:
        parts = txt.split(':')
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
        if len(parts) == 2:
            return int(parts[0]) * 60 + float(parts[1])
        return float(txt)
    except Exception:
        return None


def _time_entry(fr, var, maxv=None):
    """数字输入框样式的时间输入：右侧内嵌竖排上下小箭头，0.1s 步进微调。"""
    p = ui_theme.PALETTE
    # 1px 外边框（聚焦时轻微变色，不用蓝色实心按钮）
    wrap = tk.Frame(fr, bg=p['border'], bd=0)
    wrap.pack(side='left')
    inner = tk.Frame(wrap, bg=p['surface'])
    inner.pack(fill='both', expand=True, padx=1, pady=1)
    ent = tk.Entry(inner, textvariable=var, width=8, justify='center',
                   font=ui_theme.font(9), bd=0, highlightthickness=0,
                   bg=p['surface'], fg=p['text'],
                   insertbackground=p['text'])
    ent.pack(side='left', padx=(5, 3), pady=2)

    def nudge(step):
        v = _parse_time(var.get())
        cur = 0.0 if v is None else v
        cur = round(cur + step, 1)
        if maxv is not None:
            cur = max(0.0, min(maxv, cur))
        var.set(_fmt(cur))

    # 输入区与箭头之间的细分隔线
    tk.Frame(inner, bg=p['border'], width=1).pack(side='left', fill='y',
                                                  pady=2)
    # 右侧竖排上下箭头（仿 <input type="number">：浅灰箭头、悬停淡灰底）
    spin = tk.Frame(inner, bg=p['surface'])
    spin.pack(side='left', fill='y', padx=(2, 2))
    spiners = []
    for s, txt, pos in ((0.1, '▲', 'top'), (-0.1, '▼', 'bottom')):
        lbl = tk.Label(spin, text=txt, font=ui_theme.font(6),
                       bg=p['surface'], fg=p['faint'],
                       padx=3, pady=0, cursor='arrow', bd=0)
        lbl.pack(side=pos, fill='x')
        lbl.bind('<Enter>', lambda e, w=lbl:
                 w.configure(bg=p['surface_alt'], fg=p['muted']))
        lbl.bind('<Leave>', lambda e, w=lbl:
                 w.configure(bg=p['surface'], fg=p['faint']))
        lbl.bind('<Button-1>', lambda e, s=s: nudge(s))
        spiners.append(lbl)

    ent.bind('<FocusIn>', lambda e: wrap.configure(bg=p['border_hover']))
    ent.bind('<FocusOut>', lambda e: wrap.configure(bg=p['border']))

    # 鼠标在输入框上滚轮也可以 0.1s 微调
    def wheel(e):
        nudge(0.1 if e.delta > 0 else -0.1)
        return 'break'

    ent.bind('<MouseWheel>', wheel)
    for w in spiners:
        w.bind('<MouseWheel>', wheel)
    return ent


def _run_ffmpeg(cmd, total, progress):
    """执行 ffmpeg 命令并解析 -progress 输出回调进度。

    progress(0.0~1.0) 由工作线程调用；total<=0 时不回调。
    返回 (ok, 错误尾部文字)。
    """
    proc = subprocess.Popen(cmd + ['-progress', 'pipe:1', '-nostats'],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            text=True, encoding='utf-8', errors='replace',
                            creationflags=_NO_WINDOW)
    errbuf = []

    def drain():
        try:
            for line in proc.stderr:
                errbuf.append(line)
        except Exception:
            pass

    threading.Thread(target=drain, daemon=True).start()
    for line in proc.stdout:
        if progress is None or total <= 0:
            continue
        if line.startswith('out_time_ms=') or line.startswith('out_time_us='):
            # ffmpeg 的 out_time_ms/out_time_us 单位都是微秒
            try:
                us = int(line.split('=', 1)[1].strip())
            except ValueError:
                continue
            progress(min(0.99, max(0.0, us / 1e6 / total)))
        elif line.startswith('progress=end'):
            progress(1.0)
    proc.wait()
    try:
        proc.stdout.close()
    except Exception:
        pass
    err = ''.join(errbuf)
    if proc.returncode != 0:
        tail = err.strip().splitlines()[-3:]
        return False, '\n'.join(tail) or ('ffmpeg 返回码 %s' % proc.returncode)
    if not os.path.exists(cmd[-1]):
        return False, '输出文件未生成'
    return True, ''


def cut(exe, src, dst, start, end, precise, progress=None):
    """区间剪切。无损：-c copy 秒出（切点对齐关键帧）；精确：重编码。"""
    dur = max(0.1, end - start)
    cmd = [exe, '-y', '-ss', _fmt(start), '-i', src, '-t', _fmt(dur)]
    if precise:
        cmd += ['-c:v', 'libx264', '-preset', 'veryfast', '-crf', '23',
                '-c:a', 'aac', '-b:a', '192k']
    else:
        cmd += ['-c', 'copy']
    cmd.append(dst)
    return _run_ffmpeg(cmd, dur, progress)


def extract_audio(exe, src, dst, progress=None, total=0.0):
    """提取/转换音频轨。"""
    ext = os.path.splitext(dst)[1].lower()
    codec = {'.m4a': ['-c:a', 'aac', '-b:a', '192k'],
             '.mp3': ['-c:a', 'libmp3lame', '-q:a', '2'],
             '.wav': ['-c:a', 'pcm_s16le']}.get(ext, ['-c:a', 'aac'])
    cmd = [exe, '-y', '-i', src, '-vn'] + codec + [dst]
    return _run_ffmpeg(cmd, total, progress)


def compress(exe, src, dst, crf, progress=None, total=0.0):
    """压缩：视频 libx264+CRF；纯音频降码率。"""
    cmd = [exe, '-y', '-i', src]
    if os.path.splitext(src)[1].lower() in _VIDEO_EXT:
        cmd += ['-c:v', 'libx264', '-preset', 'veryfast', '-crf', str(crf),
                '-c:a', 'aac', '-b:a', '128k']
    else:
        cmd += ['-c:a', 'aac', '-b:a', '96k']
    cmd.append(dst)
    return _run_ffmpeg(cmd, total, progress)


def replace_audio(exe, video, audio, dst, progress=None, total=0.0,
                  start=0.0, end=None, offset=0.0, volume=1.0):
    """替换视频音轨：保留画面，换成新音频（可手动调整）。

    start/end：新音频只取 [start,end> 片段；offset：接入视频的偏移秒；
    volume：音量系数(1.0=原样)。输出长度以视频时长为准。
    """
    cmd = [exe, '-y', '-i', video]
    if start > 0 or (end is not None):
        cmd += ['-ss', _fmt(start)]
        if end is not None:
            cmd += ['-t', _fmt(max(0.1, end - start))]
    cmd += ['-i', audio]
    wants_filter = (volume != 1.0) or (offset > 0)
    if wants_filter:
        fc, cur = [], '[1:a]'
        if volume != 1.0:
            fc.append('[1:a]volume=%.3f[a1]' % volume)
            cur = '[a1]'
        if offset > 0:
            fc.append('%sadelay=%d:all=1[a2]' % (cur, int(round(offset * 1000))))
            cur = '[a2]'
        fc.append('%sapad[a]' % cur)
        outmap = '[a]'
        cmd += ['-filter_complex', ';'.join(fc), '-map', '0:v', '-map', outmap]
    else:
        cmd += ['-map', '0:v', '-map', '1:a']
    if total > 0:
        cmd += ['-t', _fmt(total)]
    cmd += ['-c:v', 'copy', '-c:a', 'aac', '-b:a', '192k', dst]
    return _run_ffmpeg(cmd, total, progress)


def concat(exe, files, dst):
    """按列表顺序拼接。先尝试 -c copy（同格式秒拼），失败转重编码。

    拼接总时长未知，进度条走不确定模式（界面层处理），此处不回调。
    """
    lst = os.path.join(os.path.dirname(dst), '_concat_list.txt')
    with open(lst, 'w', encoding='utf-8') as f:
        for pth in files:
            safe = os.path.abspath(pth).replace("'", "'\\''")
            f.write("file '%s'\n" % safe)
    try:
        base = [exe, '-y', '-f', 'concat', '-safe', '0', '-i', lst]
        r = subprocess.run(base + ['-c', 'copy', dst], capture_output=True,
                           text=True, encoding='utf-8', errors='replace',
                           creationflags=_NO_WINDOW)
        if r.returncode != 0 or not os.path.exists(dst):
            r = subprocess.run(base + ['-c:v', 'libx264', '-preset',
                                       'veryfast', '-crf', '23', '-c:a',
                                       'aac', '-b:a', '192k', dst],
                               capture_output=True, text=True,
                               encoding='utf-8', errors='replace',
                               creationflags=_NO_WINDOW)
            if r.returncode != 0 or not os.path.exists(dst):
                tail = (r.stderr or '').strip().splitlines()[-3:]
                raise RuntimeError('\n'.join(tail) or 'ffmpeg 返回码')
    finally:
        try:
            os.remove(lst)
        except Exception:
            pass


class RangeScale(tk.Canvas):
    """单条轨道、可拖动两端的区间滑块。

    绑两个 DoubleVar（单位：秒/值域由 set_range 决定）。拖动首个/结束柄把值写回
    var_lo/var_hi，控制器通过 var 的 trace 同步刷新文本、剪切输出等。
    普通箭头光标；圆角扁平均由当前 Tk 风格决定。
    """

    _PAD = 10
    _H = 30

    def __init__(self, parent, var_lo, var_hi,
                 trough, fill, handle, **kw):
        self.on_change = kw.pop('on_change', None)
        self.on_release = kw.pop('on_release', None)
        kw.setdefault('height', self._H)
        kw.setdefault('highlightthickness', 0)
        kw.setdefault('bd', 0)
        super().__init__(parent, bg=parent['bg'], **kw)
        self.var_lo = var_lo
        self.var_hi = var_hi
        self._d = 1.0
        self._drag = None
        self._last = 'lo'          # 最后操作的手柄，供方向键微调
        self.trough = trough
        self.fill = fill
        self.handle = handle
        self.bind('<Configure>', lambda e: self.redraw())
        self.bind('<ButtonPress-1>', self._press)
        self.bind('<B1-Motion>', self._drag_mv)
        self.bind('<ButtonRelease-1>', self._release)
        self.bind('<Left>', lambda e: self._nudge(-0.1))
        self.bind('<Right>', lambda e: self._nudge(0.1))
        var_lo.trace_add('write', lambda *a: self.redraw())
        var_hi.trace_add('write', lambda *a: self.redraw())
        self.after_idle(self.redraw)

    def set_range(self, dur):
        self._d = max(0.1, float(dur))
        self.redraw()

    def _px(self, v):
        w = self.winfo_width()
        return self._PAD + (v / self._d) * max(1, w - 2 * self._PAD)

    def _val(self, x):
        w = self.winfo_width()
        return self._d * max(0.0, min(1.0,
                                      (x - self._PAD) / max(1, w - 2 * self._PAD)))

    def redraw(self):
        self.delete('all')
        w = self.winfo_width()
        if w < 30:
            return                    # 宽度未就绪，等 Configure 再重绘（不忙等）
        lo = self.var_lo.get()
        hi = self.var_hi.get()
        if hi < lo:
            lo, hi = hi, lo
        x0, x1 = self._px(lo), self._px(hi)
        # 轨道 + 选中区间
        self.create_rectangle(self._PAD, self._H // 2 - 3,
                              w - self._PAD, self._H // 2 + 3,
                              fill=self.trough, outline='')
        self.create_rectangle(x0, self._H // 2 - 3, x1, self._H // 2 + 3,
                              fill=self.fill, outline='')
        # 两个手柄
        for xv in (x0, x1):
            self.create_rectangle(xv - 5, 4, xv + 5, self._H - 4,
                                  fill=self.handle, outline='')

    def _press(self, e):
        lo, hi = self.var_lo.get(), self.var_hi.get()
        dlo, dhi = abs(e.x - self._px(lo)), abs(e.x - self._px(hi))
        self._drag = 'lo' if dlo <= dhi else 'hi'
        self._last = self._drag
        self.focus_set()           # 让左右方向键能落到本控件

    def _drag_mv(self, e):
        if self._drag is None:
            return
        v = self._val(e.x)
        if self._drag == 'lo':
            if v > self.var_hi.get() - 0.1:
                v = self.var_hi.get() - 0.1
            self.var_lo.set(round(max(0.0, v), 1))
        else:
            if v < self.var_lo.get() + 0.1:
                v = self.var_lo.get() + 0.1
            self.var_hi.set(round(min(self._d, v), 1))
        if self.on_change is not None:
            self.on_change(v)

    def _release(self, e):
        self._drag = None
        if self.on_release is not None:
            self.on_release()

    def _nudge(self, step):
        """方向键微调最后操作的手柄（0.1 秒）。"""
        which = 'lo' if self._last == 'lo' else 'hi'
        var = self.var_lo if which == 'lo' else self.var_hi
        if which == 'lo':
            v = min(max(0.0, var.get() + step), self.var_hi.get() - 0.1)
        else:
            v = max(min(self._d, var.get() + step), self.var_lo.get() + 0.1)
        var.set(round(v, 1))
        if self.on_change is not None:
            self.on_change(var.get())


class AvToolApp(object):
    """音视频剪辑工具窗。master 为主程序根窗口。"""

    def __init__(self, master):
        self.master = master
        self.win = None
        self.exe = find_ffmpeg()
        self.path = ''
        self.duration = 0.0
        self.has_video = False
        self.has_audio = False
        self._busy = False
        self._pv = None            # LibVlcPlayer 预览引擎（懒初始化）
        self._pv_drag = False      # 用户正在拖进度条时不回写位置
        self._pv_failed = False
        self._pv_tick_id = None
        self._pv_hold = False      # 等真正 PLAYING 再暂停在第一帧
        self._pv_hold_n = 0
        self._build()

    # ================================================== 界面搭建
    def _build(self):
        p = ui_theme.PALETTE
        win = tk.Toplevel(self.master)
        win.title('音视频剪辑')
        self.win = win
        win.withdraw()                      # 先隐藏，摆好位置再显示
        win.configure(bg=p['bg'])
        try:
            ui_theme.style_window(win)
        except Exception:
            pass
        win.geometry('820x800')
        win.minsize(760, 730)
        _center(win, 820, 800)
        win.protocol('WM_DELETE_WINDOW', self.close)

        top = tk.Frame(win, bg=p['bg'])
        top.pack(fill='x', padx=14, pady=(12, 4))
        ui_theme.FlatButton(top, '打开文件…', kind='primary',
                            command=self._pick).pack(side='left')
        ui_theme.FlatButton(top, '拼接文件…',
                            command=self._concat_dlg).pack(side='left',
                                                           padx=(8, 0))
        info = tk.Frame(top, bg=p['bg'])
        info.pack(side='right')
        self._file_lbl = tk.Label(info, text='未打开文件（支持 mp4/mov/mkv/'
                                  'avi/mp3/m4a/wav 等）',
                                  font=ui_theme.font(10, 'bold'),
                                  bg=p['bg'],
                                  fg=p['text'] if self.path else p['faint'],
                                  anchor='w')
        self._file_lbl.pack(anchor='e')
        self._meta_lbl = tk.Label(info, text='', font=ui_theme.font(9),
                                  bg=p['bg'], fg=p['muted'], anchor='w')
        self._meta_lbl.pack(anchor='e')

        # ================= 预览区（视频画面 / 音频提示）=================
        self._pv_frame = tk.Frame(win, bg='#000000', height=250)
        self._pv_frame.pack(fill='x', padx=14, pady=(2, 0))
        self._pv_frame.pack_propagate(False)
        self._pv_hwnd = None
        self._audio_lbl = tk.Label(self._pv_frame, text='', font=ui_theme.
                                   font(11, 'bold'), bg='#000000',
                                   fg='#5a6b8c')
        self._audio_lbl.place(relx=0.5, rely=0.5, anchor='center')

        # ================= 预览控制条 =================
        pc = tk.Frame(win, bg=p['bg'])
        pc.pack(fill='x', padx=14, pady=(6, 2))
        self._btn_play = ui_theme.FlatButton(pc, '播放', kind='primary',
                                             command=self._pv_toggle)
        self._btn_play.pack(side='left')
        ui_theme.FlatButton(pc, '停止', command=self._pv_stop)\
            .pack(side='left', padx=(6, 10))
        self._pv_pos_var = tk.DoubleVar(value=0.0)
        self._pv_pos = ttk.Scale(pc, from_=0.0, to=1.0,
                                 variable=self._pv_pos_var)
        self._pv_pos.pack(side='left', fill='x', expand=True)
        self._pv_pos.bind('<ButtonPress-1>',
                          lambda e: setattr(self, '_pv_drag', True))
        self._pv_pos.bind('<ButtonRelease-1>', self._pv_user_seek)
        self._pv_time_lbl = tk.Label(pc, text='00:00.0 / 00:00.0',
                                     font=ui_theme.font(9), width=15,
                                     bg=p['bg'], fg=p['text'])
        self._pv_time_lbl.pack(side='left', padx=(8, 6))
        tk.Label(pc, text='音量', font=ui_theme.font(8), bg=p['bg'],
                 fg=p['faint']).pack(side='left')
        self._pv_vol = ttk.Scale(pc, from_=0, to=100, value=80,
                                 command=self._pv_volume)
        self._pv_vol.config(length=70)
        self._pv_vol.pack(side='left', padx=(4, 0))

        pc2 = tk.Frame(win, bg=p['bg'])
        pc2.pack(fill='x', padx=14, pady=(0, 4))
        self._pv_hint = tk.Label(pc2, text='', font=ui_theme.font(8),
                                 bg=p['bg'], fg=p['faint'], anchor='e')
        self._pv_hint.pack(side='right')

        # ================= 主体：两张并排卡片 =================
        body = tk.Frame(win, bg=p['bg'])
        body.pack(fill='both', expand=True, padx=14, pady=(6, 4))
        body.columnconfigure(0, weight=3, uniform='cols')
        body.columnconfigure(1, weight=2, uniform='cols')
        body.rowconfigure(0, weight=1)

        def card(parent, title, col):
            fr = tk.Frame(parent, bg=p['surface'], highlightbackground=
                          p['border'], highlightthickness=1)
            fr.grid(row=0, column=col, sticky='nsew',
                    padx=(0, 8) if col == 0 else (8, 0))
            title_lbl = tk.Label(fr, text=title, font=ui_theme.font(10, 'bold'),
                                 bg=p['surface'], fg=p['text'], anchor='w')
            title_lbl.pack(fill='x', padx=12, pady=(10, 6))
            inner = tk.Frame(fr, bg=p['surface'])
            inner.pack(fill='both', expand=True, padx=12, pady=(0, 10))
            return fr, inner, title_lbl

        # —— 左卡：区间剪切 ——
        _, left, _left_title = card(body, '区间剪切', 0)
        left.columnconfigure(0, weight=1)
        self._s_var = tk.DoubleVar(value=0.0)
        self._e_var = tk.DoubleVar(value=1.0)
        self._s_txt = tk.StringVar(value='00:00.0')
        self._e_txt = tk.StringVar(value='00:00.0')
        # 单条区间滑块（拖动两端）
        self._range_sc = RangeScale(left, self._s_var, self._e_var,
                                    trough=p['border'], fill=p['primary'],
                                    handle=p['text'],
                                    on_change=self._rs_seek,
                                    on_release=self._rs_release)
        self._range_sc.grid(row=0, column=0, columnspan=3, sticky='ew',
                            pady=(2, 6))
        self._s_sc = self._e_sc = self._range_sc
        # 起止时间输入
        trow = tk.Frame(left, bg=p['surface'])
        trow.grid(row=1, column=0, columnspan=3, sticky='ew', pady=(0, 4))
        tk.Label(trow, text='开始', font=ui_theme.font(9), bg=p['surface'],
                 fg=p['text']).pack(side='left')
        self._s_ent = _time_entry(trow, self._s_txt)
        tk.Label(trow, text='结束', font=ui_theme.font(9), bg=p['surface'],
                 fg=p['text']).pack(side='left', padx=(18, 0))
        self._e_ent = _time_entry(trow, self._e_txt)
        self._loop_sel = tk.BooleanVar(value=False)
        tk.Checkbutton(trow, text='循环选中区间', variable=self._loop_sel,
                       bg=p['surface'], fg=p['text'], font=ui_theme.font(8),
                       activebackground=p['surface']).pack(side='left')
        # 设起止（跟随预览播放位置）
        self._s_var.trace_add('write', lambda *a: self._sync_time('s'))
        self._e_var.trace_add('write', lambda *a: self._sync_time('e'))
        self._s_ent.bind('<Return>', lambda e: self._apply_time('s'))
        self._e_ent.bind('<Return>', lambda e: self._apply_time('e'))
        mode_fr = tk.Frame(left, bg=p['surface'])
        mode_fr.grid(row=2, column=0, columnspan=3, sticky='w', padx=2,
                     pady=(6, 4))
        self._mode_var = tk.StringVar(value='copy')
        tk.Radiobutton(mode_fr, text='无损快（切点对齐关键帧）',
                       variable=self._mode_var, value='copy',
                       bg=p['surface'], fg=p['text'],
                       activebackground=p['surface']).pack(side='left')
        tk.Radiobutton(mode_fr, text='精确（重编码）',
                       variable=self._mode_var, value='precise',
                       bg=p['surface'], fg=p['text'],
                       activebackground=p['surface']).pack(side='left',
                                                          padx=(12, 0))
        btn_fr = tk.Frame(left, bg=p['surface'])
        btn_fr.grid(row=3, column=0, columnspan=3, sticky='e', pady=(8, 2))
        self._btn_cut = ui_theme.FlatButton(btn_fr, '剪切导出', kind='primary',
                                            command=self._do_cut)
        self._btn_cut.pack(side='right')

        # —— 右卡：导出 ——
        self._export_lf, right, self._right_title = card(body, '', 1)
        right.columnconfigure(1, weight=1)
        right.grid_columnconfigure(1, weight=1)
        tk.Label(right, text='提取音频', font=ui_theme.font(9),
                 bg=p['surface'], fg=p['text'])\
            .grid(row=0, column=0, padx=(0, 8), pady=6, sticky='w')
        self._a_fmt = ttk.Combobox(right, values=('m4a', 'mp3', 'wav'),
                                   width=5, state='readonly')
        self._a_fmt.set('m4a')
        self._a_fmt.grid(row=0, column=1, sticky='w')
        self._btn_audio = ui_theme.FlatButton(right, '导出音频',
                                              command=self._do_audio)
        self._btn_audio.grid(row=0, column=2, padx=(10, 0))
        tk.Label(right, text='压缩', font=ui_theme.font(9),
                 bg=p['surface'], fg=p['text'])\
            .grid(row=1, column=0, padx=(0, 8), pady=6, sticky='w')
        self._q_fmt = ttk.Combobox(right, values=('高', '中', '低'),
                                   width=5, state='readonly')
        self._q_fmt.set('中')
        self._q_fmt.grid(row=1, column=1, sticky='w')
        self._btn_zip = ui_theme.FlatButton(right, '压缩导出',
                                            command=self._do_zip)
        self._btn_zip.grid(row=1, column=2, padx=(10, 0))
        self._a_hint = tk.Label(right, text='', font=ui_theme.font(8),
                                bg=p['surface'], fg=p['faint'], anchor='w',
                                justify='left', wraplength=200)
        self._a_hint.grid(row=2, column=0, columnspan=3, sticky='ew',
                          padx=2, pady=(4, 0))

        # —— 内联卡片：替换音频（不弹框，直接显示，和区间剪切联动）——
        swap_card = tk.Frame(body, bg=p['surface'],
                             highlightbackground=p['border'],
                             highlightthickness=1)
        swap_card.grid(row=1, column=0, columnspan=2, sticky='nsew',
                       padx=(0, 8) if False else 0, pady=(6, 0))
        tk.Label(swap_card, text='替换音频', font=ui_theme.font(10, 'bold'),
                 bg=p['surface'], fg=p['text'], anchor='w')\
            .pack(fill='x', padx=12, pady=(10, 4))
        sw1 = tk.Frame(swap_card, bg=p['surface'])
        sw1.pack(fill='x', padx=12, pady=2)
        self._btn_swap_pick = ui_theme.FlatButton(
            sw1, '选择新音频…', command=self._do_swap)
        self._btn_swap_pick.pack(side='left')
        self._swap_path = tk.StringVar(value=' 未选择新音频')
        tk.Label(sw1, textvariable=self._swap_path, font=ui_theme.font(9),
                 bg=p['surface'], fg=p['faint'], anchor='w')\
            .pack(side='left', padx=(8, 0))
        sw2 = tk.Frame(swap_card, bg=p['surface'])
        sw2.pack(fill='x', padx=12, pady=4)
        self._swap_as = tk.StringVar(value='00:00.0')
        self._swap_ae = tk.StringVar(value='')
        self._swap_off = tk.StringVar(value='0')
        self._swap_vol = tk.StringVar(value='100')
        for txt, var in (('新音频开始', self._swap_as),
                         ('新音频结束', self._swap_ae)):
            tk.Label(sw2, text=txt, font=ui_theme.font(8), bg=p['surface'],
                     fg=p['faint']).pack(side='left')
            _time_entry(sw2, var)
            tk.Label(sw2, text='  ', bg=p['surface']).pack(side='left')
        for txt, var in (('偏移(秒)', self._swap_off),
                         ('音量%', self._swap_vol)):
            tk.Label(sw2, text=txt, font=ui_theme.font(8), bg=p['surface'],
                     fg=p['faint']).pack(side='left')
            tk.Entry(sw2, width=8, justify='center', textvariable=var,
                     font=ui_theme.font(9)).pack(side='left', padx=(2, 10))
        self._btn_swap = ui_theme.FlatButton(sw2, '执行替换', kind='primary',
                                             command=self._run_swap_exec)
        self._btn_swap.pack(side='left')

        # ================= 底栏：进度 + 状态 =================
        bottom = tk.Frame(win, bg=p['bg'])
        bottom.pack(fill='x', padx=14, pady=(0, 12))
        self._bar = ttk.Progressbar(bottom, mode='determinate', maximum=1.0,
                                    length=200)
        self._bar.pack(side='left', fill='x', expand=True)
        self._pct_lbl = tk.Label(bottom, text='', font=ui_theme.font(9),
                                 width=5, bg=p['bg'], fg=p['text'])
        self._pct_lbl.pack(side='left', padx=(8, 0))
        tk.Frame(bottom, width=14, bg=p['bg']).pack(side='left')
        self._status = tk.Label(bottom, text='就绪', font=ui_theme.font(9),
                                bg=p['bg'], fg=p['muted'], anchor='w')
        self._status.pack(side='left', fill='x', expand=True)

        if not self.exe:
            self._set_status('未找到 ffmpeg：请先 pip install imageio-ffmpeg',
                             True)
            for b in (self._btn_cut, self._btn_audio, self._btn_zip,
                      self._btn_swap):
                b.set_enabled(False)

        win.deiconify()
        win.update_idletasks()

    # ================================================== 文件与信息
    def _pick(self):
        if self._busy:
            return
        path = filedialog.askopenfilename(
            parent=self.win, title='选择视频或音频文件',
            filetypes=[('媒体文件', '*.mp4 *.mov *.mkv *.avi *.ts *.webm '
                        '*.flv *.mp3 *.m4a *.aac *.wav *.wma *.flac'),
                       ('所有文件', '*.*')])
        if not path:
            return
        self.load_file(path)

    def load_file(self, path):
        if not self.exe:
            messagebox.showwarning(
                '音视频剪辑', '未找到 ffmpeg，无法读取/处理文件。\n'
                '请先运行：pip install imageio-ffmpeg，装好后重新打开本窗口。',
                parent=self.win)
            return
        try:
            dur, hv, ha, res = probe(self.exe, path)
        except Exception as e:
            messagebox.showerror('音视频剪辑', '读取文件失败：%s' % e,
                                 parent=self.win)
            return
        if dur <= 0:
            messagebox.showwarning('音视频剪辑', '无法识别该文件的时长',
                                   parent=self.win)
            return
        self.path = path
        self.duration = dur
        self.has_video = hv
        self.has_audio = ha
        name = os.path.basename(path)
        self._file_lbl.config(text=name, fg=ui_theme.PALETTE['text'])
        meta = '时长 %s' % _fmt(dur)
        if hv:
            meta += ' · %s' % (res or '未知分辨率')
        meta += ' · 音轨' if ha else ' · 无音轨'
        self._meta_lbl.config(text=meta)
        self._s_var.set(0.0)
        self._e_var.set(round(dur, 1))
        self._range_sc.set_range(max(0.1, dur))
        # 导出栏按文件类型分栏显示
        kind = '导出（视频）' if hv else '导出（音频）'
        self._right_title.config(text=' %s ' % kind)
        hint = '' if hv else '纯音频：压缩按降码率处理，提取音频=转换格式'
        self._a_hint.config(text=hint)
        self._set_status('已打开：%s' % name)
        self._pv_load(autoplay=True)

    def _set_status(self, txt, warn=False):
        self._status.config(
            text=txt, fg=ui_theme.PALETTE['danger'] if warn else
            ui_theme.PALETTE['muted'])

    # ================================================== 实时预览
    def _pv_init(self):
        """懒初始化 libvlc 内嵌播放器（需要本机有 VLC）。"""
        if self._pv is not None:
            return True
        if self._pv_failed:
            return False
        try:
            pv = video_player.LibVlcPlayer()
            if not pv.available:
                self._pv_failed = True
                self._pv_hint.config(
                    text='预览需要 VLC：程序启动时会自动安装，装好后重新打开本窗口')
                return False
            self._pv = pv
            return True
        except Exception as e:
            self._pv_failed = True
            self._pv_hint.config(text='预览不可用：%s' % e)
            return False

    def _pv_load(self, autoplay=False):
        """把当前文件装进预览。

        视频画进预览框；纯音频不初始化内嵌播放器（避免无画面播放卡死），
        仅显示文件信息，区间剪切与导出照常可用。
        """
        self._pv_stop()
        if not self.path:
            return
        if not self.has_video:
            self._pv_frame.config(height=90)
            self._audio_lbl.config(
                text='♪  %s（时长 %s，拖动区间即可剪切/导出）'
                     % (os.path.basename(self.path), _fmt(self.duration)))
            return
        if not self._pv_init():
            return
        self._pv_frame.config(height=240)
        self._audio_lbl.config(text='')
        try:
            self.win.update_idletasks()
            host = self._pv_frame.winfo_id() if self.has_video else None
            ok, msg = self._pv.play([self.path], loop=False,
                                    volume=int(float(self._pv_vol.get())),
                                    host_hwnd=host)
            if not ok:
                self._pv_hint.config(text='预览失败：%s' % msg)
                return
            self._pv_hint.config(text='')
            if not autoplay:
                # VLC 在 OPENING 阶段收 pause 会被吞掉，交给 tick 等
                # 真正 PLAYING 后再暂停在第一帧
                self._pv_hold = True
                self._pv_hold_n = 0
        except Exception as e:
            self._pv_hint.config(text='预览失败：%s' % e)
            return
        if self._pv_tick_id is None:
            self._pv_tick_id = self.win.after(200, self._pv_tick)

    def _pv_toggle(self):
        self._pv_hold = False
        pv = self._pv
        if pv is None:
            if self.path:
                self._pv_load(autoplay=True)
            return
        if pv.paused:
            pv.resume()
        elif pv.state() == 1:             # 播完了 → 从头再来
            pv.seek(0.0)
            pv.resume()
        else:
            pv.pause()

    def _pv_stop(self):
        pv = self._pv
        if pv is not None:
            try:
                pv.stop()
            except Exception:
                pass
        self._pv = None
        self._pv_drag = False
        self._pv_hold = False
        try:
            self._btn_play.configure(text='播放')
            self._pv_pos_var.set(0.0)
            self._pv_time_lbl.config(
                text='00:00.0 / %s' % _fmt(self.duration))
        except tk.TclError:
            pass

    def _pv_user_seek(self, event=None):
        self._pv_drag = False
        pv = self._pv
        if pv is None or self.duration <= 0:
            return
        dur = pv.duration() or self.duration
        pv.seek(self._pv_pos_var.get() * dur)

    def _pv_volume(self, val):
        if self._pv is not None:
            try:
                self._pv.set_volume(int(float(val)))
            except Exception:
                pass

    def _pv_point(self):
        """当前真实播放位置（秒）：优先取 VLC 播放时刻，兜底用播放条比例。"""
        if self._pv is not None:
            try:
                pos = self._pv.position()
                if pos is not None and pos > 0:
                    return pos
            except Exception:
                pass
        if self.duration > 0:
            return self._pv_pos_var.get() * self.duration
        return 0.0

    def _rs_seek(self, seconds):
        """拖区间滑块时，预览跳到被拖动端点对应的时间。"""
        if self._pv is None or not self.path:
            return
        try:
            self._pv.seek(seconds)
        except Exception:
            pass

    def _rs_release(self):
        """拖完区间滑块后暂停预览，便于方向键精确定位。"""
        if self._pv is not None:
            try:
                self._pv.pause()
            except Exception:
                pass

    def _pv_tick(self):
        """每 200ms 刷新：位置条、时间文字、播放/暂停按钮、选区循环。"""
        self._pv_tick_id = None
        win = self.win
        if win is None:
            return
        try:
            alive = bool(win.winfo_exists())
        except tk.TclError:
            return
        pv = self._pv
        if pv is not None:
            pos = pv.position()
            dur = pv.duration() or self.duration
            if not self._pv_drag and dur > 0:
                self._pv_pos_var.set(min(1.0, pos / dur))
            self._pv_time_lbl.config(
                text='%s / %s' % (_fmt(pos), _fmt(dur)))
            if self._pv_hold:
                # 加载后停在第一帧：OPENING 阶段 pause 会被吞，重试到成功
                self._pv_hold_n += 1
                if pv.paused:
                    pv.seek(0.05)
                    self._pv_hold = False
                elif pv.state() == 3 and self._pv_hold_n > 2:
                    pv.pause()
                elif self._pv_hold_n > 60:      # 12s 还没起来就放弃
                    self._pv_hold = False
            elif pv.paused:
                self._btn_play.configure(text='播放')
            else:
                self._btn_play.configure(text='暂停')
            # 循环播放选中区间：播过“结束”就跳回“开始”
            if self._loop_sel.get() and dur > 0 and not pv.paused \
                    and not self._pv_hold:
                s, e = self._s_var.get(), self._e_var.get()
                if e - s > 0.5 and pos >= e - 0.05:
                    pv.seek(s + 0.02)
        try:
            self._pv_tick_id = win.after(200, self._pv_tick)
        except tk.TclError:
            self._pv_tick_id = None

    # ================================================== 进度显示
    def _progress(self, frac):
        """工作线程回调 → 回到主线程刷新进度条。"""
        def apply():
            try:
                self._bar.config(mode='determinate')
                self._bar.stop()
                self._bar_var_set(frac)
                self._pct_lbl.config(text='%d%%' % round(frac * 100))
            except tk.TclError:
                pass
        try:
            self.win.after(0, apply)
        except RuntimeError:
            pass

    def _bar_var_set(self, frac):
        self._bar['value'] = max(0.0, min(1.0, frac))

    def _bar_indeterminate(self):
        def apply():
            try:
                self._bar.config(mode='indeterminate')
                self._bar_var_set(0.0)
                self._pct_lbl.config(text='')
                self._bar.start(12)
            except tk.TclError:
                pass
        try:
            self.win.after(0, apply)
        except RuntimeError:
            pass

    def _bar_reset(self):
        def apply():
            try:
                self._bar.stop()
                self._bar.config(mode='determinate')
                self._bar_var_set(0.0)
                self._pct_lbl.config(text='')
            except tk.TclError:
                pass
        try:
            self.win.after(0, apply)
        except RuntimeError:
            pass

    # ================================================== 时间同步
    def _sync_time(self, which):
        var = self._s_var if which == 's' else self._e_var
        txt = self._s_txt if which == 's' else self._e_txt
        txt.set(_fmt(var.get()))

    def _apply_time(self, which):
        ent = self._s_ent if which == 's' else self._e_ent
        var = self._s_var if which == 's' else self._e_var
        v = _parse_time(ent.get())
        if v is None:
            self._sync_time(which)
            return
        v = max(0.0, min(self.duration, v))
        var.set(round(v, 1))

    def _range(self):
        s = self._s_var.get()
        e = self._e_var.get()
        if e - s < 0.2:
            messagebox.showwarning('音视频剪辑', '结束时间需大于开始时间',
                                   parent=self.win)
            return None
        return s, e

    def _out_path(self, suffix, ext):
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path \
            else '输出'
        folder = os.path.dirname(self.path) if self.path else os.getcwd()
        return unique_path(folder, base + suffix + ext)

    # ================================================== 后台执行
    def _run_bg(self, fn, done_txt, indeterminate=False):
        """线程执行 fn() → (输出路径, 错误)；期间禁用按钮并显示进度。"""
        self._busy = True
        self._set_status('处理中…')
        for b in (self._btn_cut, self._btn_audio, self._btn_zip,
                  self._btn_swap):
            b.set_enabled(False)
        if indeterminate:
            self._bar_indeterminate()

        def work():
            try:
                out = fn()
                err = ''
            except Exception as e:
                out, err = None, str(e)
            self.win.after(0, lambda: self._done(out, err, done_txt))

        threading.Thread(target=work, daemon=True).start()

    def _done(self, out, err, done_txt):
        self._busy = False
        self._bar_reset()
        for b in (self._btn_cut, self._btn_audio, self._btn_zip,
                  self._btn_swap):
            b.set_enabled(True)
        if err:
            self._set_status('失败：%s' % err.splitlines()[-1][:120], True)
            messagebox.showerror('音视频剪辑', '处理失败：\n%s' % err,
                                 parent=self.win)
        else:
            self._set_status(done_txt % out)
            if messagebox.askyesno('音视频剪辑', '%s\n\n打开所在文件夹？'
                                   % (done_txt % out), parent=self.win):
                try:
                    os.startfile(os.path.dirname(out))
                except Exception:
                    pass

    def _need_file(self):
        if not self.path:
            messagebox.showinfo('音视频剪辑', '请先打开一个视频/音频文件',
                                parent=self.win)
            return False
        return True

    def _do_cut(self):
        if self._busy or not self._need_file() or not self.exe:
            return
        rg = self._range()
        if not rg:
            return
        precise = self._mode_var.get() == 'precise'
        ext = os.path.splitext(self.path)[1].lower() or '.mp4'
        dst = self._out_path('_剪辑', ext)
        total = rg[1] - rg[0]

        def fn():
            ok, err = cut(self.exe, self.path, dst, rg[0], rg[1], precise,
                          progress=self._progress)
            if not ok:
                raise RuntimeError(err)
            return dst

        self._run_bg(fn, '已导出：%s')

    def _do_audio(self):
        if self._busy or not self._need_file() or not self.exe:
            return
        ext = '.' + self._a_fmt.get()
        dst = self._out_path('_音频', ext)
        total = self.duration

        def fn():
            ok, err = extract_audio(self.exe, self.path, dst,
                                    progress=self._progress, total=total)
            if not ok:
                raise RuntimeError(err)
            return dst

        self._run_bg(fn, '已导出：%s')

    def _do_zip(self):
        if self._busy or not self._need_file() or not self.exe:
            return
        crf = _CRFS.get(self._q_fmt.get(), 27)
        dst = self._out_path('_压缩', os.path.splitext(self.path)[1] or '.mp4')
        total = self.duration

        def fn():
            ok, err = compress(self.exe, self.path, dst, crf,
                               progress=self._progress, total=total)
            if not ok:
                raise RuntimeError(err)
            return dst

        self._run_bg(fn, '已导出：%s')

    def _do_swap(self):
        """选择要替换成的新音频：文件选好后填入内联替换区，自动设结束默认。"""
        if self._busy or not self.exe:
            return
        if not self.has_video:
            messagebox.showwarning('音视频剪辑', '替换音频需要先打开视频文件',
                                   parent=self.win)
            return
        audio = filedialog.askopenfilename(
            parent=self.win, title='选择要替换成的新音频',
            filetypes=[('音频', '*.mp3 *.m4a *.aac *.wav *.flac *.ogg *.wma'),
                       ('所有文件', '*.*')])
        if not audio:
            return
        self._swap_audio = audio
        self._swap_path.set('  %s' % os.path.basename(audio))
        try:
            adur, _, _, _ = probe(self.exe, audio)
            if adur:
                self._swap_ae.set(_fmt(adur))
        except Exception:
            pass

    def _run_swap_exec(self):
        """读内联替换区参数并执行替换。"""
        audio = getattr(self, '_swap_audio', '')
        if not audio:
            messagebox.showinfo('音视频剪辑', '请先点「选择新音频…」选一段音频',
                                parent=self.win)
            return
        if self._busy or not self.exe:
            return
        try:
            asv = _parse_time(self._swap_as.get()) or 0.0
            ae_val = self._swap_ae.get().strip()
            aev = _parse_time(ae_val) if ae_val else None
            off = float(self._swap_off.get() or 0)
            vol = float(self._swap_vol.get() or 100)
        except ValueError:
            messagebox.showwarning('音视频剪辑', '数值格式不对，请检查',
                                   parent=self.win)
            return
        self._run_swap(audio, asv, aev, max(0.0, off), max(0, vol) / 100.0)

    def _run_swap(self, audio, start, end, offset, volume):
        dst = self._out_path('_替换音频', '.mp4')
        total = self.duration

        def fn():
            ok, err = replace_audio(self.exe, self.path, audio, dst,
                                    progress=self._progress, total=total,
                                    start=start, end=end, offset=offset,
                                    volume=volume)
            if not ok:
                raise RuntimeError(err)
            return dst

        self._run_bg(fn, '已导出：%s')

    # ================================================== 拼接对话框
    def _concat_dlg(self):
        if self._busy:
            return
        if not self.exe:
            messagebox.showwarning('音视频剪辑', '未找到 ffmpeg',
                                   parent=self.win)
            return
        files = filedialog.askopenfilenames(
            parent=self.win, title='选择要拼接的文件（可多选，按列表顺序合并）',
            filetypes=[('媒体文件', '*.mp4 *.mov *.mkv *.avi *.ts *.webm '
                        '*.flv *.mp3 *.m4a *.aac *.wav *.wma *.flac'),
                       ('所有文件', '*.*')])
        if len(files) < 2:
            if files:
                messagebox.showinfo('音视频剪辑', '请至少选择两个文件',
                                    parent=self.win)
            return
        dlg = tk.Toplevel(self.win)
        dlg.title('拼接顺序')
        dlg.configure(bg=ui_theme.PALETTE['bg'])
        dlg.transient(self.win)
        _center(dlg, 560, 360)
        dlg.grab_set()

        tk.Label(dlg, text='从上到下依次拼接，可调整顺序：',
                 font=ui_theme.font(9), bg=ui_theme.PALETTE['bg'],
                 fg=ui_theme.PALETTE['muted']).pack(anchor='w', padx=12,
                                                    pady=(10, 4))
        lb = tk.Listbox(dlg, font=ui_theme.font(9), bg=ui_theme.PALETTE[
            'surface'], fg=ui_theme.PALETTE['text'],
            highlightthickness=1, selectbackground=ui_theme.PALETTE['primary'])
        lb.pack(fill='both', expand=True, padx=12)
        for f in files:
            lb.insert('end', os.path.basename(f))

        btns = tk.Frame(dlg, bg=ui_theme.PALETTE['bg'])
        btns.pack(fill='x', padx=12, pady=(6, 0))

        def move(delta):
            sel = lb.curselection()
            if not sel:
                return
            i = sel[0]
            j = i + delta
            if 0 <= j < lb.size():
                txt = lb.get(i)
                lb.delete(i)
                lb.insert(j, txt)
                lb.selection_set(j)

        ui_theme.FlatButton(btns, '上移', command=lambda: move(-1))\
            .pack(side='left')
        ui_theme.FlatButton(btns, '下移', command=lambda: move(1))\
            .pack(side='left', padx=(6, 0))

        def go():
            ordered = [files[i] for i in range(lb.size())]
            folder = os.path.dirname(ordered[0])
            ext = os.path.splitext(ordered[0])[1].lower() or '.mp4'
            dst = unique_path(folder, '拼接结果' + ext)
            dlg.destroy()
            self._set_status('拼接中…')

            def fn():
                concat(self.exe, ordered, dst)
                return dst

            # 拼接总时长未知 → 进度条走不确定模式
            self._run_bg(fn, '已导出：%s', indeterminate=True)

        ui_theme.FlatButton(btns, '开始拼接', kind='primary', command=go)\
            .pack(side='right')

    # ================================================== 关闭
    def close(self, force=False):
        self._pv_stop()
        try:
            win = self.win
            if win is not None:
                if self._pv_tick_id is not None:
                    try:
                        win.after_cancel(self._pv_tick_id)
                    except Exception:
                        pass
                win.destroy()
        except Exception:
            pass
        self.win = None
