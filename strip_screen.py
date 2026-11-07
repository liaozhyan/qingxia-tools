# -*- coding: utf-8 -*-
"""长条形显示屏合成投屏。

现场常见问题：显示屏是长条（比如 3840×1080），直接放 16:9 的 PPT 会被拉扁。
这里的做法是——

    左边固定图 | 中间等比放 PPT（不拉伸） | 右边固定图

先用 PIL 把整屏画面合成好，再整屏投到目标显示器上；翻页只重绘中间那块。
"""

import os
import tkinter as tk

import ui_theme


def compute_layout(width, height, left_ratio=0.15, right_ratio=0.15,
                   margin_ratio=0.0):
    """按比例算出左/中/右三块区域，返回字典。"""
    width = max(2, int(width))
    height = max(2, int(height))
    left_ratio = max(0.0, min(float(left_ratio), 0.45))
    right_ratio = max(0.0, min(float(right_ratio), 0.45))
    left_w = int(round(width * left_ratio))
    right_w = int(round(width * right_ratio))
    mid_w = max(2, width - left_w - right_w)
    margin = int(round(height * max(0.0, min(float(margin_ratio), 0.3))))
    return {
        'left': (0, 0, left_w, height),
        'middle': (left_w, margin, left_w + mid_w, height - margin),
        'right': (width - right_w, 0, width, height),
        'size': (width, height),
    }


def fit_inside(size, box):
    """等比缩放到 box 内（不变形），返回 (宽, 高)。"""
    box_w = max(1, box[2] - box[0])
    box_h = max(1, box[3] - box[1])
    src_w, src_h = max(1, size[0]), max(1, size[1])
    scale = min(box_w / float(src_w), box_h / float(src_h))
    return max(1, int(round(src_w * scale))), max(1, int(round(src_h * scale)))


def cover_crop(image, box_w, box_h):
    """把图按“铺满并居中裁剪”缩放（用于两侧固定图）。"""
    if image.width < 1 or image.height < 1:
        return image
    target_ratio = float(box_w) / max(1, box_h)
    src_ratio = float(image.width) / max(1, image.height)
    if src_ratio > target_ratio:
        crop_w = max(1, int(round(image.height * target_ratio)))
        left = max(0, (image.width - crop_w) // 2)
        image = image.crop((left, 0, left + crop_w, image.height))
    else:
        crop_h = max(1, int(round(image.width / target_ratio)))
        top = max(0, (image.height - crop_h) // 2)
        image = image.crop((0, top, image.width, top + crop_h))
    return image.resize((max(1, box_w), max(1, box_h)))


def load_image(path, max_pixels=40000000):
    """打开图片并按需缩小（防止超大素材把内存吃满）。"""
    from PIL import Image
    with Image.open(path) as raw:
        try:
            raw.draft('RGB', (3840, 2160))
        except Exception:
            pass
        img = raw.convert('RGB')
    if img.width * img.height > max_pixels:
        scale = (max_pixels / float(img.width * img.height)) ** 0.5
        img = img.resize((max(1, int(img.width * scale)),
                          max(1, int(img.height * scale))))
    return img


class Compositor(object):
    """把两侧固定图 + 中间 PPT 画面合成成一整屏。"""

    def __init__(self, width, height, left_ratio=0.15, right_ratio=0.15,
                 margin_ratio=0.0, background='#0b1a33',
                 left_image=None, right_image=None, show_divider=True,
                 background_image=None, slide_box=None):
        self.layout = compute_layout(width, height, left_ratio, right_ratio,
                                     margin_ratio)
        self.width = max(2, int(width))
        self.height = max(2, int(height))
        self.background = background
        self.left_image = left_image
        self.right_image = right_image
        self.show_divider = show_divider
        # 自定义版式：整屏背景图 + 任意位置的 PPT 框（比例 0~1）
        self.background_image = background_image
        self.slide_box = slide_box
        self._base = None

    def slide_rect(self):
        """PPT 那块画面在整屏里的像素位置 (x, y, w, h)。"""
        if self.slide_box:
            x, y, w, h = self.slide_box
            return (int(round(self.width * x)), int(round(self.height * y)),
                    max(2, int(round(self.width * w))),
                    max(2, int(round(self.height * h))))
        box = self.layout['middle']
        return (box[0], box[1], max(2, box[2] - box[0]),
                max(2, box[3] - box[1]))

    def _build_base(self):
        """两侧图和底图只算一次（内容相对固定）。"""
        from PIL import Image, ImageDraw
        width, height = self.layout['size']
        if self.background_image is not None:
            # 自定义版式：整屏铺满用户上传的背景图（居中裁切，不拉伸变形）
            self._base = cover_crop(self.background_image, width, height)
            return self._base
        base = Image.new('RGB', (width, height), self.background)
        for key, source in (('left', self.left_image),
                            ('right', self.right_image)):
            box = self.layout[key]
            box_w, box_h = box[2] - box[0], box[3] - box[1]
            if box_w <= 0 or box_h <= 0:
                continue
            if source is not None:
                base.paste(cover_crop(source, box_w, box_h), (box[0], box[1]))
            else:
                shade = ui_theme.mix(self.background, '#000000', 0.30) \
                    if key == 'left' else ui_theme.mix(self.background,
                                                       '#ffffff', 0.06)
                base.paste(Image.new('RGB', (box_w, box_h), shade),
                           (box[0], box[1]))
        if self.show_divider:
            draw = ImageDraw.Draw(base)
            for key in ('left', 'right'):
                box = self.layout[key]
                x = box[2] if key == 'left' else box[0]
                draw.line([(x, 0), (x, height)], fill='#2b3f66', width=2)
        self._base = base
        return base

    def invalidate(self):
        """换了背景图/比例后调用，下一帧重新合成。"""
        self._base = None

    def render(self, slide_image=None):
        """合成一整屏；slide_image 为 None 时只有两侧内容。"""
        base = self._base if self._base is not None else self._build_base()
        if slide_image is None:
            return base
        frame = base.copy()
        x0, y0, box_w, box_h = self.slide_rect()
        fitted = fit_inside(slide_image.size, (x0, y0, x0 + box_w, y0 + box_h))
        img = slide_image.resize(fitted)
        x = x0 + (box_w - fitted[0]) // 2
        y = y0 + (box_h - fitted[1]) // 2
        frame.paste(img, (x, y))
        return frame


class ComposeScreen(object):
    """大屏合成窗口：背景图铺满整屏，视频嵌在指定的框里。

    关键点：视频是这块窗口的**子窗口**（VLC 直接画进这个 Frame），
    所以 z 序天然正确——不存在"背景盖住视频"或"视频掉到后面"的情况，
    也不需要在两个窗口之间抢层级/追着重画。
    """

    def __init__(self, master, rect):
        self.rect = tuple(rect)
        self.width = max(320, self.rect[2] - self.rect[0])
        self.height = max(200, self.rect[3] - self.rect[1])
        self._bg_photo = None
        self.win = tk.Toplevel(master)
        self.win.overrideredirect(True)          # 无标题栏、不占任务栏
        self.win.geometry('%dx%d+%d+%d' % (self.width, self.height,
                                           self.rect[0], self.rect[1]))
        try:
            self.win.attributes('-topmost', True)
        except tk.TclError:
            pass
        self.win.configure(bg='#000000')
        self.canvas = tk.Canvas(self.win, width=self.width,
                                height=self.height, highlightthickness=0,
                                bd=0, bg='#000000')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        # 视频落在这个 Frame 里（VLC 直接渲染进它的窗口句柄）
        self.host = tk.Frame(self.win, bg='#000000')
        self.host.place(x=0, y=0, width=self.width, height=self.height)
        self._repaint_jobs = []
        self._keep_job = None
        ui_theme.present_window(self.win)
        for delay in (200, 600, 1200, 2200, 3600):
            try:
                self._repaint_jobs.append(self.win.after(delay, self._paint))
            except tk.TclError:
                break
        self._keep_visible()

    # ---------- 基本 ----------
    def is_open(self):
        try:
            return bool(self.win.winfo_exists())
        except tk.TclError:
            return False

    def hwnd(self):
        """VLC 渲染用的窗口句柄（视频框那个 Frame）。"""
        try:
            self.host.update_idletasks()
            return self.host.winfo_id()
        except tk.TclError:
            return None

    def _paint(self):
        try:
            self.win.lift()
            self.win.attributes('-topmost', True)
            ui_theme.force_repaint(self.win)
        except tk.TclError:
            pass

    def _keep_visible(self):
        """大屏窗口要一直亮着（扩展屏经常不发重画消息）。"""
        try:
            if not self.win.winfo_exists():
                self._keep_job = None
                return
            self._paint()
            self._keep_job = self.win.after(2500, self._keep_visible)
        except tk.TclError:
            self._keep_job = None

    # ---------- 版式 ----------
    def show_background(self, pil_image):
        """背景图按“铺满裁切”画到整屏；传 None 就是纯黑底。"""
        from PIL import ImageTk
        self.canvas.delete('all')
        self._bg_photo = None
        if pil_image is None:
            self.canvas.create_rectangle(0, 0, self.width, self.height,
                                         fill='#000000', outline='')
            self._paint()
            return
        try:
            img = pil_image
            if img.size != (self.width, self.height):
                img = cover_crop(img, self.width, self.height)
            self._bg_photo = ImageTk.PhotoImage(img)
            self.canvas.create_image(0, 0, image=self._bg_photo, anchor='nw')
        except Exception:
            self.canvas.create_rectangle(0, 0, self.width, self.height,
                                         fill='#000000', outline='')
        self._paint()

    def set_box(self, box):
        """视频框位置（整屏比例 x, y, w, h）。"""
        x, y, w, h = box
        left = max(0, min(int(round(self.width * x)), self.width - 2))
        top = max(0, min(int(round(self.height * y)), self.height - 2))
        box_w = max(2, min(int(round(self.width * w)), self.width - left))
        box_h = max(2, min(int(round(self.height * h)), self.height - top))
        try:
            self.host.place(x=left, y=top, width=box_w, height=box_h)
            self.host.lift()
        except tk.TclError:
            pass

    def close(self):
        for job in getattr(self, '_repaint_jobs', []):
            try:
                self.win.after_cancel(job)
            except tk.TclError:
                pass
        self._repaint_jobs = []
        if getattr(self, '_keep_job', None) is not None:
            try:
                self.win.after_cancel(self._keep_job)
            except tk.TclError:
                pass
            self._keep_job = None
        try:
            self.win.destroy()
        except tk.TclError:
            pass


class StripScreen(object):
    """把合成好的整屏画面全屏显示在指定显示器上。"""

    def __init__(self, master, rect, on_close=None,
                 on_next=None, on_prev=None, on_toggle_auto=None,
                 topmost=True, keep_lift=True):
        self.rect = rect
        self.on_close = on_close
        self.on_next = on_next
        self.on_prev = on_prev
        self.on_toggle_auto = on_toggle_auto
        self.topmost = bool(topmost)
        self.keep_lift = bool(keep_lift)
        self.width = max(320, rect[2] - rect[0])
        self.height = max(200, rect[3] - rect[1])
        self._photo = None

        self.win = tk.Toplevel(master)
        self.win.overrideredirect(True)
        self.win.geometry('%dx%d+%d+%d'
                          % (self.width, self.height, rect[0], rect[1]))
        if self.topmost:
            ui_theme.present_window(self.win)
        else:
            # 背景窗口不置顶（要让视频盖在上面），但同样需要抬一下 z 序：
            # 系统只认这一下才会把画面真正显示到扩展屏上
            try:
                self.win.lift()
            except tk.TclError:
                pass
            ui_theme.force_repaint(self.win)
        self.win.configure(bg='#000000')
        self.canvas = tk.Canvas(self.win, width=self.width,
                                height=self.height, highlightthickness=0,
                                bd=0, bg='#000000')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.create_text(self.width / 2, self.height / 2,
                                text='正在准备画面…', fill='#ffffff',
                                font=ui_theme.font(18, 'bold'))
        self.canvas.create_text(self.width - 24, 26,
                                text='ESC 退出　↑ ↓ 翻页　空格 自动播放',
                                anchor='e', fill='#8fa4c8',
                                font=ui_theme.font(9))

        self.win.bind('<Escape>', self._on_escape)
        self.canvas.bind('<Escape>', self._on_escape)
        for widget in (self.win, self.canvas):
            # 上下键（含翻页键）是主用翻页键，左右键也保留
            for seq in ('<Down>', '<Next>', '<Right>'):
                widget.bind(seq, lambda e: self._call(self.on_next))
            for seq in ('<Up>', '<Prior>', '<Left>'):
                widget.bind(seq, lambda e: self._call(self.on_prev))
            widget.bind('<space>', lambda e: self._call(self.on_toggle_auto))
            widget.bind('<Button-1>', lambda e: self._focus())
        # 无边框窗口直接摆到扩展屏上时，系统有时不发重画消息，大屏就是黑的
        self._repaint_jobs = []
        for delay in (200, 500, 900, 1400, 2000, 2800, 3600, 4500, 6000, 8000):
            try:
                self._repaint_jobs.append(self.win.after(
                    delay, self._paint))
            except tk.TclError:
                break
        self._keep_job = None
        self._keep_visible()
        self._focus()

    def _focus(self):
        try:
            self.win.focus_force()
        except tk.TclError:
            pass
        try:
            self.canvas.focus_set()
        except tk.TclError:
            pass

    def _paint(self, lift=None):
        """重画一次，并抬一下 z 序。

        系统有时不给扩展屏上的无边框窗口发重画消息，抬一次才会显示；
        keep_lift=False 的背景窗口只在创建时抬一次，之后只重画，
        避免它反复抢到视频窗口上面去。
        """
        if lift is None:
            lift = self.keep_lift
        try:
            if lift:
                self.win.lift()
                # 只有“该抬”的窗口才重新置顶；否则每次重绘都会把自己
                # 又顶到最前面，把压在上面的视频窗口盖住（背景窗口就出过这问题）
                if self.topmost:
                    self.win.attributes('-topmost', True)
        except tk.TclError:
            return
        ui_theme.force_repaint(self.win)

    def _keep_visible(self):
        """投屏窗口一直留在扩展屏最前面：定时补置顶 + 重绘，防止大屏发黑。"""
        try:
            if not self.win.winfo_exists():
                self._keep_job = None
                return
            if self.keep_lift:
                self.win.lift()
            self._paint()
            self._keep_job = self.win.after(3000, self._keep_visible)
        except tk.TclError:
            self._keep_job = None

    def _call(self, func):
        if callable(func):
            try:
                func()
            except Exception:
                pass
        return 'break'

    def _on_escape(self, _event=None):
        self.close()
        return 'break'

    def is_open(self):
        try:
            return bool(self.win.winfo_exists())
        except tk.TclError:
            return False

    def move_to(self, rect):
        """换到另一块屏：尺寸一样就只挪位置，不重建窗口。"""
        old_w = max(320, self.rect[2] - self.rect[0])
        old_h = max(200, self.rect[3] - self.rect[1])
        new_w = max(320, rect[2] - rect[0])
        new_h = max(200, rect[3] - rect[1])
        if (new_w, new_h) != (old_w, old_h):
            return False
        self.rect = rect
        try:
            self.win.geometry('+%d+%d' % (rect[0], rect[1]))
            ui_theme.force_repaint(self.win)
            return True
        except tk.TclError:
            return False

    def close(self, notify=True):
        for job in getattr(self, '_repaint_jobs', []):
            try:
                self.win.after_cancel(job)
            except tk.TclError:
                pass
        self._repaint_jobs = []
        if getattr(self, '_keep_job', None) is not None:
            try:
                self.win.after_cancel(self._keep_job)
            except tk.TclError:
                pass
            self._keep_job = None
        try:
            self.win.destroy()
        except tk.TclError:
            pass
        if notify and self.on_close is not None:
            try:
                self.on_close()
            except Exception:
                pass

    def show(self, pil_image):
        """显示一张合成好的画面（按屏幕分辨率 1:1，不做拉伸）。"""
        from PIL import ImageTk
        if not self.is_open():
            return False
        try:
            if pil_image.size != (self.width, self.height):
                pil_image = pil_image.resize((self.width, self.height))
            photo = ImageTk.PhotoImage(pil_image)
            self._photo = photo          # 必须留引用，否则图会消失
            self.canvas.delete('all')
            self.canvas.create_image(0, 0, image=photo, anchor='nw')
            try:
                self.win.update_idletasks()
            except tk.TclError:
                pass
            self._paint()
            return True
        except Exception:
            return False
