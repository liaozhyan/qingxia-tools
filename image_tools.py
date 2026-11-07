# -*- coding: utf-8 -*-
"""图片工具箱：会场临时修图（Pillow + OpenCV，界面零第三方依赖）。

布局：
- 顶栏：打开图片 / 打开文件夹（批量队列缩略图）/ 拼接图片 / 撤销 / 重置 / 按住对比原图
- 左列：区域工具（直接改底图，可撤销）：裁剪 / 透视矫正 / 马赛克 / 修复涂抹
- 中间：深色预览画布（框选、角点、涂抹都画在这里）+ 状态栏
- 右列：参数面板（实时预览）：校色 / 增强 / 滤镜·水印 / 装饰 / 导出

处理管线：区域工具结果（底图）→ process_params() 参数化调整 → 预览/导出。
预览在降采样图上实时计算，导出时对原图全分辨率重算，保证画质。
OpenCV 缺失时自动降级：去噪/磨皮/自动增强/修复不可用，其余功能不受影响。
"""

import io
import math
import os
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox
from tkinter import ttk

from PIL import (Image, ImageChops, ImageColor, ImageDraw, ImageEnhance,
                 ImageFilter, ImageFont, ImageOps, ImageTk)

import ui_theme
from file_inbox import ensure_folder, unique_path

try:
    import cv2
except Exception:
    cv2 = None
try:
    import numpy as np
except Exception:
    np = None

FORMATS = ('原格式', 'JPEG (.jpg)', 'PNG (.png)', 'WebP (.webp)')
_FORMAT_EXT = {'JPEG (.jpg)': '.jpg', 'PNG (.png)': '.png',
               'WebP (.webp)': '.webp'}
FILTERS = ('无', '黑白', '怀旧', '冷色', '暖色', '胶片')
_CANVAS_BG = '#151b29'

DEFAULTS = {
    'auto': False, 'denoise': 0, 'skin': 0, 'sharpen': 0, 'temperature': 0,
    'brightness': 0, 'contrast': 0, 'saturation': 0, 'filter': '无',
    'text': '', 'text_pos': '下中', 'text_margin': 5, 'text_size': 0,
    'text_color': '#ffffff', 'text_stroke': True, 'text_xy': None,
    'text_opacity': 100,
    'watermark': '', 'wm_size': 0, 'wm_alpha': 0.35, 'wm_tile': True,
    'round_r': 0, 'border_w': 0, 'border_color': '#ffffff',
}

# 添加文字的九宫格位置 → (纵向 t/m/b, 横向 l/c/r)
_TEXT_POS = {
    '左上': ('t', 'l'), '上中': ('t', 'c'), '右上': ('t', 'r'),
    '左中': ('m', 'l'), '居中': ('m', 'c'), '右中': ('m', 'r'),
    '左下': ('b', 'l'), '下中': ('b', 'c'), '右下': ('b', 'r'),
}


# ---------------------------------------------------------------- 基础工具

def _font(size):
    """水印用的 PIL 字体：中文首选微软雅黑，找不到再退回。"""
    candidates = (r'C:\Windows\Fonts\msyh.ttc',
                  r'C:\Windows\Fonts\simhei.ttf')
    try:
        for c in candidates:
            if os.path.exists(c):
                return ImageFont.truetype(c, size)
        return ImageFont.truetype(r'C:\Windows\Fonts\arial.ttf', size)
    except Exception:
        return ImageFont.load_default()


def apply_watermark(img, text, size=0, alpha=0.35, tile=True):
    """给图片加半透明文字水印。返回新图。"""
    if not text:
        return img
    if img.mode != 'RGBA':
        img = img.convert('RGBA')
    w, h = img.size
    font = _font(size or max(24, int(min(w, h) * 0.06)))
    overlay = Image.new('RGBA', img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    fw, fh = draw.textbbox((0, 0), text, font=font)[2:]
    step_y = max(int(fh * 1.8), 40)
    step_x = max(int(fw * 1.4), 40)
    a = max(0, int(255 * min(1.0, alpha)))
    if tile:
        y = step_y // 2 - fh // 2
        while y < h:
            x = -fw // 2
            while x < w:
                draw.text((x, y), text, font=font, fill=(255, 255, 255, a))
                x += step_x
            y += step_y
    else:
        draw.text((w - fw - 20, h - fh - 20), text, font=font,
                  fill=(255, 255, 255, a))
    return Image.alpha_composite(img, overlay)


def _as_rgb(img):
    """RGBA/P 等模式压平到白底 RGB。"""
    if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in
                                      img.info):
        bg = Image.new('RGB', img.size, (255, 255, 255))
        rgba = img.convert('RGBA')
        bg.paste(rgba, mask=rgba.split()[-1])
        return bg
    return img.convert('RGB')


def _to_int(v, default=0):
    try:
        return int(float(str(v).strip()))
    except Exception:
        return default


# ------------------------------------------------------------ 处理算法层

def _pil2bgr(img):
    return cv2.cvtColor(np.asarray(img.convert('RGB')), cv2.COLOR_RGB2BGR)


def _bgr2pil(arr):
    return Image.fromarray(cv2.cvtColor(arr, cv2.COLOR_BGR2RGB))


def _denoise(img, strength):
    """非局部均值去噪（保色彩），strength 1~10。"""
    bgr = _pil2bgr(img)
    return _bgr2pil(cv2.fastNlMeansDenoisingColored(
        bgr, None, strength, strength, 7, 21))


def _skin(img, level):
    """磨皮：双边滤波（保边缘平滑皮肤），level 0~100。"""
    bgr = _pil2bgr(img)
    sigma = 5 + level * 0.15
    out = cv2.bilateralFilter(bgr, 9, sigma, 9)
    if level > 60:          # 高强度再补一遍，效果更接近美颜
        out = cv2.bilateralFilter(out, 9, sigma, 9)
    return _bgr2pil(out)


def _clahe(img):
    """自动增强：LAB 亮度通道 CLAHE 自适应直方图均衡。"""
    bgr = _pil2bgr(img)
    lab = cv2.cvtColor(bgr, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    l = cv2.createCLAHE(2.0, (8, 8)).apply(l)
    return _bgr2pil(cv2.cvtColor(cv2.merge((l, a, b)), cv2.COLOR_LAB2BGR))


def _temperature(img, t):
    """色温：t>0 暖（加红减蓝），t<0 冷。"""
    if np is None or not t:
        return img
    k = max(-100, min(100, int(t))) / 100.0
    arr = np.asarray(img.convert('RGB')).astype(np.int16)
    arr[..., 0] = np.clip(arr[..., 0] + k * 30, 0, 255)
    arr[..., 2] = np.clip(arr[..., 2] - k * 30, 0, 255)
    return Image.fromarray(arr.astype('uint8'), 'RGB')


def _blend_solid(img, rgb, alpha):
    solid = Image.new('RGB', img.size, rgb)
    return Image.blend(img, solid, alpha)


def _apply_filter(img, name):
    """滤镜：黑白 / 怀旧 / 冷色 / 暖色 / 胶片。"""
    if not name or name == '无':
        return img
    rgb = img.convert('RGB')
    if name == '黑白':
        return ImageEnhance.Contrast(
            ImageOps.grayscale(rgb).convert('RGB')).enhance(1.15)
    if name == '怀旧':
        m = (0.393, 0.769, 0.189, 0,
             0.349, 0.686, 0.168, 0,
             0.272, 0.534, 0.131, 0)
        return rgb.convert('RGB', m)
    if name == '冷色':
        return _blend_solid(rgb, (205, 228, 255), 0.14)
    if name == '暖色':
        return _blend_solid(rgb, (255, 213, 150), 0.14)
    if name == '胶片':
        out = ImageEnhance.Contrast(rgb).enhance(1.12)
        out = ImageEnhance.Color(out).enhance(0.90)
        return _blend_solid(out, (255, 224, 196), 0.08)
    return img


def _pos_xy(pos, W, H, tw, th, m):
    row, col = _TEXT_POS.get(pos, ('b', 'c'))
    x = m if col == 'l' else (W - m - tw if col == 'r' else (W - tw) // 2)
    y = m if row == 't' else (H - m - th if row == 'b' else (H - th) // 2)
    return max(0, x), max(0, y)


def _draw_text(img, text, pos='下中', margin=5, size=0, color='#ffffff',
               stroke=True, center=None, opacity=100):
    """在图片上添加可见文字（支持 \\n 换行），默认白色字 + 深色描边。

    center=(fx, fy) 为文字中心点占图宽高的比例（0~1），用于自由拖动定位；
    为 None 时按 pos 九宫格 + margin 定位。opacity 0~100，小于 100 时半透明。
    """
    W, H = img.size
    # 固定字号按「宽 1000px 基准」换算，保证预览与导出比例一致
    if size:
        font_px = max(8, int(int(size) * W / 1000.0))
    else:
        font_px = max(24, int(W * 0.055))
    font = _font(font_px)
    try:
        fill = ImageColor.getrgb(color or '#ffffff')
    except Exception:
        fill = (255, 255, 255)
    sw = max(1, font_px // 16) if stroke else 0
    op = max(0, min(100, int(opacity)))
    if op < 100:                       # 半透明：画到独立图层再合成
        layer = Image.new('RGBA', img.size, (0, 0, 0, 0))
        d = ImageDraw.Draw(layer)
        fill = fill + (int(255 * op / 100),)
        stroke_fill = (0, 0, 0, int(255 * op / 100))
    else:
        layer = None
        d = ImageDraw.Draw(img)
        stroke_fill = (0, 0, 0)
    multi = '\n' in text
    try:
        if multi:
            bbox = d.multiline_textbbox((0, 0), text, font=font, spacing=8,
                                        stroke_width=sw)
        else:
            bbox = d.textbbox((0, 0), text, font=font, stroke_width=sw)
    except Exception:
        bbox = d.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    if center is not None:
        fx = max(0.0, min(1.0, float(center[0])))
        fy = max(0.0, min(1.0, float(center[1])))
        x = max(0, min(int(fx * W - tw / 2), max(0, W - tw)))
        y = max(0, min(int(fy * H - th / 2), max(0, H - th)))
    else:
        m = int(min(W, H) * max(0, min(50, int(margin))) / 100.0)
        x, y = _pos_xy(pos, W, H, tw, th, m)
    xy = (x - bbox[0], y - bbox[1])
    if multi:
        d.multiline_text(xy, text, font=font, fill=fill, spacing=8,
                         stroke_width=sw, stroke_fill=stroke_fill)
    else:
        d.text(xy, text, font=font, fill=fill, stroke_width=sw,
               stroke_fill=stroke_fill)
    if layer is not None:
        out = Image.alpha_composite(img.convert('RGBA'), layer)
        return out.convert('RGB') if img.mode == 'RGB' else out
    return img


def _round_corners(img, radius):
    """圆角：返回带透明角的 RGBA 图。"""
    r = min(int(radius), img.width // 2, img.height // 2)
    if r <= 0:
        return img
    rgba = img.convert('RGBA')
    mask = Image.new('L', rgba.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        (0, 0, rgba.width - 1, rgba.height - 1), radius=r, fill=255)
    alpha = ImageChops.multiply(rgba.getchannel('A'), mask)
    rgba.putalpha(alpha)
    return rgba


def _add_border(img, width, color):
    """加边框：圆角透明角会由边框色补齐。"""
    w = int(width)
    if w <= 0:
        return img
    try:
        rgb = ImageColor.getrgb(color or '#ffffff')
    except Exception:
        rgb = (255, 255, 255)
    out = Image.new('RGB', (img.width + 2 * w, img.height + 2 * w), rgb)
    if img.mode == 'RGBA':
        out.paste(img, (w, w), img)
    else:
        out.paste(img.convert('RGB'), (w, w))
    return out


def _mosaic_region(img, rect, block):
    """把 rect=(x1,y1,x2,y2) 区域打马赛克（块大小 block px）。"""
    x1, y1, x2, y2 = rect
    x1 = max(0, min(int(x1), img.width - 1))
    y1 = max(0, min(int(y1), img.height - 1))
    x2 = max(x1 + 1, min(int(x2), img.width))
    y2 = max(y1 + 1, min(int(y2), img.height))
    block = max(2, int(block))
    region = img.crop((x1, y1, x2, y2))
    w, h = region.size
    small = region.resize((max(1, w // block), max(1, h // block)),
                          Image.NEAREST)
    out = img.copy()
    out.paste(small.resize((w, h), Image.NEAREST), (x1, y1))
    return out


def _warp_quad(img, quad):
    """透视矫正：quad=[TL,TR,BR,BL]，拉正为矩形（PIL QUAD 变换）。"""
    tl, tr, br, bl = [(float(x), float(y)) for x, y in quad]

    def d(a, b):
        return math.hypot(a[0] - b[0], a[1] - b[1])

    w = max(2, int(max(d(tl, tr), d(bl, br))))
    h = max(2, int(max(d(tl, bl), d(tr, br))))
    # QUAD 数据顺序：NW, SW, SE, NE
    data = (tl[0], tl[1], bl[0], bl[1], br[0], br[1], tr[0], tr[1])
    return img.transform((w, h), Image.QUAD, data, Image.BICUBIC)


def _inpaint(img, mask):
    """图像修复：mask 'L' 模式，255=要去除的区域（需 OpenCV）。"""
    if cv2 is None or np is None:
        return img
    bgr = _pil2bgr(img)
    m = np.asarray(mask.convert('L'))
    if m.shape[:2] != bgr.shape[:2]:
        return img
    return _bgr2pil(cv2.inpaint(bgr, m, 5, cv2.INPAINT_TELEA))


def process_params(base, p):
    """参数化调整管线（非破坏）：底图 + settings → 新 PIL Image。"""
    im = _as_rgb(base)
    if cv2 is not None and np is not None:
        dn = int(p.get('denoise') or 0)
        if dn > 0:
            im = _denoise(im, dn)
        sk = int(p.get('skin') or 0)
        if sk > 0:
            im = _skin(im, sk)
        if p.get('auto'):
            im = _clahe(im)
    t = int(p.get('temperature') or 0)
    if t:
        im = _temperature(im, t)
    b = int(p.get('brightness') or 0)
    if b:
        im = ImageEnhance.Brightness(im).enhance(1 + b / 100.0)
    c = int(p.get('contrast') or 0)
    if c:
        im = ImageEnhance.Contrast(im).enhance(1 + c / 100.0)
    s = int(p.get('saturation') or 0)
    if s:
        im = ImageEnhance.Color(im).enhance(1 + s / 100.0)
    im = _apply_filter(im, p.get('filter') or '无')
    sh = int(p.get('sharpen') or 0)
    if sh > 0:
        im = im.filter(ImageFilter.UnsharpMask(radius=2, percent=sh,
                                               threshold=3))
    tx = (p.get('text') or '').strip()
    if tx:
        center = p.get('text_xy') if p.get('text_pos') == '自定义' else None
        im = _draw_text(im, tx,
                        pos=p.get('text_pos') or '下中',
                        margin=int(p.get('text_margin') or 5),
                        size=int(p.get('text_size') or 0),
                        color=p.get('text_color') or '#ffffff',
                        stroke=bool(p.get('text_stroke')),
                        center=center,
                        opacity=int(p.get('text_opacity') or 100))
    wm = (p.get('watermark') or '').strip()
    if wm:
        im = apply_watermark(im, wm,
                             size=int(p.get('wm_size') or 0),
                             alpha=float(p.get('wm_alpha') or 0.35),
                             tile=bool(p.get('wm_tile')))
    rr = int(p.get('round_r') or 0)
    if rr > 0:
        im = _round_corners(im, rr)
    bw = int(p.get('border_w') or 0)
    if bw > 0:
        im = _add_border(im, bw, p.get('border_color') or '#ffffff')
    return im


# ------------------------------------------------------------ 编码 / 导出

def _encode(img, ext, quality, exif=None):
    """把图编码成字节串（ext 形如 '.jpg'）。"""
    ext = ext.lower()
    kw = {}
    if exif and ext in ('.jpg', '.jpeg', '.webp', '.tif', '.tiff'):
        kw['exif'] = exif
    if ext in ('.jpg', '.jpeg'):
        buf = io.BytesIO()
        if quality:
            kw['quality'] = int(quality)
        _as_rgb(img).save(buf, format='JPEG', **kw)
        return buf.getvalue()
    if ext == '.png':
        buf = io.BytesIO()
        img.save(buf, format='PNG', **kw)
        return buf.getvalue()
    if ext == '.webp':
        buf = io.BytesIO()
        if quality:
            kw['quality'] = int(quality)
        img.save(buf, format='WEBP', **kw)
        return buf.getvalue()
    buf = io.BytesIO()
    img.save(buf, **kw)
    return buf.getvalue()


def _save_image(img, dest, ext, quality=85, exif=None, target_kb=0):
    """保存；target_kb>0 时先扫质量再逐步缩图逼近目标大小。返回字节数。"""
    target = max(0, int(target_kb)) * 1024
    if target and ext.lower() not in ('.png',):
        data = None
        for q in range(max(quality, 12), 11, -8):
            data = _encode(img, ext, q, exif)
            if len(data) <= target:
                break
        if len(data) > target:
            cur = img
            for _ in range(8):
                cur = cur.resize(
                    (max(2, int(cur.width * 0.85)),
                     max(2, int(cur.height * 0.85))), Image.LANCZOS)
                for q in (70, 55, 40, 30):
                    data = _encode(cur, ext, q, exif)
                    if len(data) <= target:
                        break
                if len(data) <= target:
                    break
        with open(dest, 'wb') as f:
            f.write(data)
        return len(data)
    if target and ext.lower() == '.png':
        cur = img
        data = _encode(cur, ext, 0, exif)
        for _ in range(8):
            if len(data) <= target:
                break
            cur = cur.resize((max(2, int(cur.width * 0.85)),
                              max(2, int(cur.height * 0.85))), Image.LANCZOS)
            data = _encode(cur, ext, 0, exif)
        with open(dest, 'wb') as f:
            f.write(data)
        return len(data)
    data = _encode(img, ext, quality, exif)
    with open(dest, 'wb') as f:
        f.write(data)
    return len(data)


def stitch_images(paths, horizontal=True, gap=12):
    """拼接多张图：横向对齐最小高度 / 纵向对齐最小宽度，白底。"""
    imgs = []
    for p in paths:
        im = Image.open(p)
        try:
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        imgs.append(im.convert('RGB'))
    if len(imgs) < 2:
        return None
    gap = max(0, int(gap))
    if horizontal:
        h = min(im.height for im in imgs)
        scaled = [im.resize((max(1, int(im.width * h / im.height)), h),
                            Image.LANCZOS) for im in imgs]
        W = sum(im.width for im in scaled) + gap * (len(scaled) - 1)
        out = Image.new('RGB', (W, h), (255, 255, 255))
        x = 0
        for im in scaled:
            out.paste(im, (x, 0))
            x += im.width + gap
    else:
        w = min(im.width for im in imgs)
        scaled = [im.resize((w, max(1, int(im.height * w / im.width))),
                            Image.LANCZOS) for im in imgs]
        H = sum(im.height for im in scaled) + gap * (len(scaled) - 1)
        out = Image.new('RGB', (w, H), (255, 255, 255))
        y = 0
        for im in scaled:
            out.paste(im, (0, y))
            y += im.height + gap
    return out


# ------------------------------------------------------------ 界面控件

class SliderRow(tk.Frame):
    """一行滑杆：名称 + SeekBar + 数值。底色自动跟随所在面板。"""

    def __init__(self, master, text, from_, to, value=0, command=None,
                 label_w=7):
        try:
            bg = master['bg']
        except Exception:
            bg = ui_theme.PALETTE['bg']
        tk.Frame.__init__(self, master, bg=bg)
        self._cmd = command
        tk.Label(self, text=text, font=ui_theme.font(9), bg=bg,
                 fg=ui_theme.PALETTE['muted'], width=label_w,
                 anchor='w').pack(side='left')
        self._val = tk.Label(self, font=ui_theme.font(9), bg=bg,
                             fg=ui_theme.PALETTE['text'], width=4,
                             anchor='e')
        self._val.pack(side='right')
        self.bar = ui_theme.SeekBar(self, from_=from_, to=to, value=value,
                                    command=self._on, height=20)
        self.bar.pack(side='left', fill='x', expand=True, padx=(4, 6))
        self._show(value)

    def _show(self, v):
        try:
            self._val.configure(text='%d' % int(round(float(v))))
        except Exception:
            pass

    def _on(self, v):
        self._show(v)
        if callable(self._cmd):
            self._cmd(int(round(float(v))))

    def get(self):
        return int(round(self.bar.get()))

    def set(self, v):
        self.bar.set(v)
        self._show(v)


def _flat_btn(parent, text, cmd, kind='normal'):
    """工具面板里的小号扁平按钮。"""
    p = ui_theme.PALETTE
    if kind == 'primary':
        bg, fg, ab = p['primary'], '#ffffff', p['primary_dark']
    else:
        bg, fg, ab = p['surface_alt'], p['text'], p['hover']
    return tk.Button(parent, text=text, command=cmd, relief='flat', bd=0,
                     padx=10, pady=4, bg=bg, fg=fg, activebackground=ab,
                     activeforeground=fg, font=ui_theme.font(9),
                     cursor='hand2', disabledforeground=p['faint'])


def _center(win, w, h):
    try:
        win.update_idletasks()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        win.geometry('+%d+%d' % (max(0, (sw - w) // 2),
                                 max(0, (sh - h) // 2)))
    except Exception:
        pass


# ------------------------------------------------------------ 主窗口

class ImageToolApp(object):
    """图片工具箱窗口。master 为主程序根窗口。"""

    def __init__(self, master):
        self.master = master
        self.win = None

        # 文件队列与图源
        self._files = []            # 批量队列（路径）
        self._keep = {}             # 缩略图 PhotoImage 引用（防回收）
        self._path = None           # 当前文件
        self._exif = None           # 当前文件原始 EXIF
        self._original = None       # 载入的原始图（EXIF 已转正）
        self._base = None           # 底图（区域工具结果）
        self._base_small = None     # 底图降采样（预览用）
        self._small_ratio = 1.0     # base / small 比例
        self._undo = []             # 区域工具撤销栈

        # 预览
        self._img = None            # 当前处理结果（small 尺寸）
        self._photo = None
        self._raw_photo = None      # 对比原图用
        self._compare_on = False
        self._view_off = (0, 0)
        self._canvas_size = (0, 0)
        self._pending_resize = False
        self._img_no_text = None    # 文字拖动用的「无文字」缓存
        self._no_text_key = None
        self._work_ver = 0          # 底图/预览图版本号（缓存失效用）

        # 工具状态
        self._tool = None
        self._drag = None           # 拖动中状态
        self._rect = None           # 裁剪/马赛克框（base 坐标）
        self._quad = None           # 透视四点（base 坐标 TL,TR,BR,BL）
        self._strokes = []          # 修复笔迹 [(x,y,r)]（base 坐标）
        self._tool_preview = None   # 马赛克/修复的实时预览小图

        # 参数（实时，非破坏）
        self.p = dict(DEFAULTS)
        self._after_id = None
        self._text_after = None
        self._resize_after = None
        self._resetters = []

        self._build()

    # ================================================== 界面搭建
    def _build(self):
        p = ui_theme.PALETTE
        win = tk.Toplevel(self.master)
        win.title('图片工具箱')
        self.win = win
        win.withdraw()                      # 先隐藏，摆好位置再显示
        win.configure(bg=p['bg'])
        try:
            ui_theme.style_window(win)
        except Exception:
            pass
        win.geometry('1180x740')
        win.minsize(1020, 660)
        _center(win, 1180, 740)
        win.protocol('WM_DELETE_WINDOW', self.close)
        win.bind('<Escape>', lambda e: self._cancel_tool())

        self._build_top(win)
        self._build_body(win)

        win.deiconify()
        win.update_idletasks()

    def _build_top(self, win):
        p = ui_theme.PALETTE
        top = tk.Frame(win, bg=p['bg'])
        top.pack(fill='x', padx=14, pady=(12, 6))

        ui_theme.FlatButton(top, '打开图片…', kind='primary',
                            command=self._pick).pack(side='left')
        ui_theme.FlatButton(top, '打开文件夹…',
                            command=self._pick_folder).pack(side='left',
                                                            padx=(8, 0))
        ui_theme.FlatButton(top, '拼接图片…',
                            command=self._stitch).pack(side='left', padx=(8, 0))
        self._btn_undo = ui_theme.FlatButton(top, '撤销',
                                             command=self._undo_last)
        self._btn_undo.pack(side='left', padx=(16, 0))
        self._btn_undo.set_enabled(False)
        self._btn_reset = ui_theme.FlatButton(top, '重置此图',
                                              command=self._reset_current)
        self._btn_reset.pack(side='left', padx=(8, 0))
        self._btn_reset.set_enabled(False)

        self._count_lbl = tk.Label(top, text='', font=ui_theme.font(9),
                                   bg=p['bg'], fg=p['muted'])
        self._count_lbl.pack(side='right', padx=(0, 10))
        self._cmp_btn = tk.Button(top, text='按住看原图', relief='flat', bd=0,
                                  padx=12, pady=5, bg=p['surface_alt'],
                                  fg=p['text'], activebackground=p['hover'],
                                  font=ui_theme.font(9), cursor='hand2',
                                  state='disabled',
                                  disabledforeground=p['faint'])
        self._cmp_btn.pack(side='right')
        self._cmp_btn.bind('<Button-1>', lambda e: self._compare_show())
        self._cmp_btn.bind('<ButtonRelease-1>', lambda e: self._compare_hide())

        # 批量队列缩略图条
        self._strip = tk.Canvas(win, height=66, bg=p['bg'],
                                highlightthickness=0, xscrollincrement=40)
        self._strip.bind('<Enter>', lambda e: self._strip.bind_all(
            '<MouseWheel>', self._strip_wheel))
        self._strip.bind('<Leave>', lambda e: self._strip.unbind_all(
            '<MouseWheel>'))

    def _strip_wheel(self, e):
        try:
            self._strip.xview_scroll(-int(e.delta / 120) * 3, 'units')
        except Exception:
            pass

    def _build_body(self, win):
        p = ui_theme.PALETTE
        body = tk.Frame(win, bg=p['bg'])
        body.pack(fill='both', expand=True, padx=14, pady=(0, 12))

        # 左列：区域工具 + 导出按钮
        self._build_left(body)
        # 右列：参数面板
        self._build_right(body)
        # 中间：画布 + 状态栏
        center = tk.Frame(body, bg=p['bg'])
        center.pack(side='left', fill='both', expand=True)
        self._canvas = tk.Canvas(center, bg=_CANVAS_BG, highlightthickness=0,
                                 cursor='')
        self._canvas.pack(fill='both', expand=True)
        self._canvas.bind('<Configure>', self._on_canvas_config)
        self._canvas.bind('<Button-1>', self._cv_press)
        self._canvas.bind('<B1-Motion>', self._cv_motion)
        self._canvas.bind('<ButtonRelease-1>', self._cv_release)

        status = tk.Frame(center, bg=p['bg'])
        status.pack(fill='x', pady=(6, 0))
        self._info_lbl = tk.Label(status, text='未载入图片',
                                  font=ui_theme.font(8), bg=p['bg'],
                                  fg=p['muted'], anchor='w')
        self._info_lbl.pack(side='left')
        self._hint_lbl = tk.Label(status, text='', font=ui_theme.font(8),
                                  bg=p['bg'], fg=p['faint'], anchor='e')
        self._hint_lbl.pack(side='right')

        self._redraw_photo()

    def _build_left(self, body):
        p = ui_theme.PALETTE
        col = tk.Frame(body, bg=p['bg'], width=186)
        col.pack(side='left', fill='y', padx=(0, 10))
        col.pack_propagate(False)

        tk.Label(col, text='区域工具', font=ui_theme.font(9, 'bold'),
                 bg=p['bg'], fg=p['muted'], anchor='w').pack(
            fill='x', pady=(2, 4))

        self._tool_btns = {}
        for key, name in (('crop', '裁剪'), ('persp', '透视矫正'),
                          ('mosaic', '马赛克'), ('inpaint', '修复涂抹')):
            b = tk.Button(col, text=name, anchor='w', relief='flat', bd=0,
                          padx=12, pady=6, bg=p['surface'], fg=p['text'],
                          activebackground=p['hover'],
                          activeforeground=p['text'],
                          font=ui_theme.font(9), cursor='hand2',
                          command=lambda k=key: self._toggle_tool(k))
            b.pack(fill='x', pady=1)
            self._tool_btns[key] = b
        if cv2 is None:
            self._tool_btns['inpaint'].configure(
                state='disabled', disabledforeground=p['faint'])

        self._tool_panel = tk.Frame(col, bg=p['bg'])
        self._tool_panel.pack(fill='x', pady=(8, 0))
        self._tool_panel.columnconfigure(0, weight=1)
        self._tp = {}
        self._build_tp_crop()
        self._build_tp_persp()
        self._build_tp_mosaic()
        self._build_tp_inpaint()
        for f in self._tp.values():
            f.grid(row=0, column=0, sticky='nsew')
            f.grid_remove()

        tk.Frame(col, bg=p['bg']).pack(fill='both', expand=True)

        tk.Label(col, text='导出', font=ui_theme.font(9, 'bold'),
                 bg=p['bg'], fg=p['muted'], anchor='w').pack(
            fill='x', pady=(2, 4))
        self._btn_save = ui_theme.FlatButton(col, '另存为…', kind='primary',
                                             command=self._save_as)
        self._btn_save.pack(fill='x')
        self._btn_save.set_enabled(False)
        self._btn_batch = ui_theme.FlatButton(col, '批量导出…',
                                              command=self._export_batch)
        self._btn_batch.pack(fill='x', pady=(6, 0))
        self._btn_batch.set_enabled(False)

    def _tp_frame(self, key, hint):
        """每个工具的参数小面板（带提示文字）。"""
        p = ui_theme.PALETTE
        f = tk.Frame(self._tool_panel, bg=p['surface'],
                     highlightbackground=p['border'], highlightthickness=1)
        tk.Label(f, text=hint, font=ui_theme.font(8), bg=p['surface'],
                 fg=p['faint'], wraplength=158, justify='left').pack(
            fill='x', padx=8, pady=(8, 4))
        self._tp[key] = f
        return f

    def _tp_btns(self, f, pairs):
        p = ui_theme.PALETTE
        row = tk.Frame(f, bg=p['surface'])
        row.pack(fill='x', padx=8, pady=(4, 8))
        for i, (text, cmd, kind) in enumerate(pairs):
            _flat_btn(row, text, cmd, kind).pack(
                side='left' if i == 0 else 'right')

    def _build_tp_crop(self):
        f = self._tp_frame('crop', '在预览图上按住拖出一个框，再点应用。')
        self._tp_btns(f, [('应用裁剪', self._apply_tool, 'primary'),
                          ('取消', self._cancel_tool, 'normal')])

    def _build_tp_persp(self):
        f = self._tp_frame('persp', '拖动四个角点对准要矫正的区域，'
                                    '应用后拉正为矩形。')
        self._tp_btns(f, [('应用矫正', self._apply_tool, 'primary'),
                          ('取消', self._cancel_tool, 'normal')])
        row = tk.Frame(f, bg=ui_theme.PALETTE['surface'])
        row.pack(fill='x', padx=8, pady=(0, 8))
        _flat_btn(row, '重置角点', self._persp_reset).pack(side='left')

    def _build_tp_mosaic(self):
        f = self._tp_frame('mosaic', '在预览图上拖出要打码的区域，'
                                     '松开即见效果。')
        self._mosaic_slider = SliderRow(f, '块大小', 4, 64, 20,
                                        command=lambda v: self._mosaic_redo())
        self._mosaic_slider.pack(fill='x', padx=8)
        self._tp_btns(f, [('应用马赛克', self._apply_tool, 'primary'),
                          ('取消', self._cancel_tool, 'normal')])

    def _build_tp_inpaint(self):
        f = self._tp_frame('inpaint', '涂抹要去除的路人 / 水印 / 杂物，'
                                      '松开自动修复预览。')
        self._brush_slider = SliderRow(f, '笔刷', 8, 120, 36, command=None)
        self._brush_slider.pack(fill='x', padx=8)
        self._tp_btns(f, [('应用修复', self._apply_tool, 'primary'),
                          ('取消', self._cancel_tool, 'normal')])
        row = tk.Frame(f, bg=ui_theme.PALETTE['surface'])
        row.pack(fill='x', padx=8, pady=(0, 8))
        _flat_btn(row, '清除笔迹', self._strokes_clear).pack(side='left')

    def _build_right(self, body):
        p = ui_theme.PALETTE
        right = tk.Frame(body, bg=p['bg'], width=306)
        right.pack(side='right', fill='y', padx=(0, 10))
        right.pack_propagate(False)

        canvas = tk.Canvas(right, bg=p['bg'], highlightthickness=0,
                           yscrollincrement=40)
        vs = ttk.Scrollbar(right, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vs.set)
        vs.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        inner = tk.Frame(canvas, bg=p['bg'])
        canvas.create_window((0, 0), window=inner, anchor='nw', tags='inner')
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfigure('inner', width=e.width))
        canvas.bind('<Enter>', lambda e: canvas.bind_all(
            '<MouseWheel>', lambda ev: canvas.yview_scroll(
                -int(ev.delta / 120) * 2, 'units')))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all('<MouseWheel>'))

        # ---- 校色
        sec = ttk.Labelframe(inner, text=' 校色 ')
        sec.pack(fill='x', pady=(0, 8))
        self._add_slider(sec, '亮度', -100, 100, 'brightness')
        self._add_slider(sec, '对比度', -100, 100, 'contrast')
        self._add_slider(sec, '饱和度', -100, 100, 'saturation')
        self._add_slider(sec, '色温', -100, 100, 'temperature')

        # ---- 增强
        sec = ttk.Labelframe(inner, text=' 增强 ')
        sec.pack(fill='x', pady=(0, 8))
        var_auto = tk.BooleanVar(value=False)
        self._resetters.append(lambda: var_auto.set(False))
        ttk.Checkbutton(sec, text='自动增强（弱光/灰蒙救回）',
                        variable=var_auto,
                        command=lambda: self._setp(
                            'auto', var_auto.get())).pack(
            fill='x', padx=6, pady=(4, 2))
        self._add_slider(sec, '去噪', 0, 10, 'denoise')
        self._add_slider(sec, '磨皮', 0, 100, 'skin')
        self._add_slider(sec, '锐化', 0, 200, 'sharpen')

        # ---- 添加文字
        sec = ttk.Labelframe(inner, text=' 添加文字 ')
        sec.pack(fill='x', pady=(0, 8))
        self._text_var = tk.StringVar()
        self._text_var.trace_add(
            'write', self._text_trace('text', self._text_var, str))
        self._resetters.append(lambda: self._text_var.set(''))
        rowf = tk.Frame(sec, bg=p['bg'])
        rowf.pack(fill='x', padx=6, pady=(6, 2))
        tk.Label(rowf, text='文字内容', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted'], width=7, anchor='w').pack(side='left')
        ttk.Entry(rowf, textvariable=self._text_var).pack(
            side='left', fill='x', expand=True)

        rowf = tk.Frame(sec, bg=p['bg'])
        rowf.pack(fill='x', padx=6, pady=(2, 2))
        tk.Label(rowf, text='位置', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted'], width=7, anchor='w').pack(side='left')
        self._text_pos_cb = ttk.Combobox(
            rowf, values=list(_TEXT_POS) + ['自定义'],
            state='readonly', width=10)
        self._text_pos_cb.set(DEFAULTS['text_pos'])
        self._text_pos_cb.pack(side='left')
        self._text_pos_cb.bind('<<ComboboxSelected>>',
                               lambda e: self._text_pos_pick())
        self._resetters.append(
            lambda: self._text_pos_cb.set(DEFAULTS['text_pos']))

        self._add_slider(sec, '边距%', 0, 30, 'text_margin')
        self._add_slider(sec, '字号(0自动)', 0, 300, 'text_size')
        self._add_slider(sec, '不透明度%', 30, 100, 'text_opacity')

        rowf = tk.Frame(sec, bg=p['bg'])
        rowf.pack(fill='x', padx=6, pady=(2, 6))
        tk.Label(rowf, text='颜色', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted'], width=7, anchor='w').pack(side='left')
        self._text_color_btn = tk.Button(rowf, text='      ', relief='flat',
                                         bd=1, bg=DEFAULTS['text_color'],
                                         activebackground=p['hover'],
                                         cursor='hand2', width=6,
                                         command=self._pick_text_color)
        self._text_color_btn.pack(side='left')
        self._var_text_stroke = tk.BooleanVar(value=True)
        self._resetters.append(lambda: self._var_text_stroke.set(True))
        ttk.Checkbutton(rowf, text='深色描边', variable=self._var_text_stroke,
                        command=lambda: self._setp(
                            'text_stroke',
                            self._var_text_stroke.get())).pack(
            side='left', padx=(10, 0))

        # ---- 滤镜与水印
        sec = ttk.Labelframe(inner, text=' 滤镜 · 水印 ')
        sec.pack(fill='x', pady=(0, 8))
        rowf = tk.Frame(sec, bg=p['bg'])
        rowf.pack(fill='x', padx=6, pady=(6, 2))
        tk.Label(rowf, text='滤镜', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted'], width=7, anchor='w').pack(side='left')
        self._filter_cb = ttk.Combobox(rowf, values=list(FILTERS),
                                       state='readonly', width=10)
        self._filter_cb.set('无')
        self._filter_cb.pack(side='left')
        self._filter_cb.bind('<<ComboboxSelected>>',
                             lambda e: self._setp('filter',
                                                  self._filter_cb.get()))
        self._resetters.append(lambda: self._filter_cb.set('无'))

        self._wm_text_var = tk.StringVar()
        self._wm_text_var.trace_add(
            'write', self._text_trace('watermark', self._wm_text_var, str))
        self._resetters.append(lambda: self._wm_text_var.set(''))
        rowf = tk.Frame(sec, bg=p['bg'])
        rowf.pack(fill='x', padx=6, pady=(6, 2))
        tk.Label(rowf, text='水印文字', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted'], width=7, anchor='w').pack(side='left')
        ttk.Entry(rowf, textvariable=self._wm_text_var).pack(
            side='left', fill='x', expand=True)

        self._wm_size_var = tk.StringVar(value='0')
        self._wm_size_var.trace_add(
            'write', self._text_trace('wm_size', self._wm_size_var,
                                      lambda s: _to_int(s, 0)))
        self._resetters.append(lambda: self._wm_size_var.set('0'))
        rowf = tk.Frame(sec, bg=p['bg'])
        rowf.pack(fill='x', padx=6, pady=(2, 2))
        tk.Label(rowf, text='字号(0自动)', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted'], width=7, anchor='w').pack(side='left')
        ttk.Entry(rowf, textvariable=self._wm_size_var, width=10).pack(
            side='left')

        self._add_slider(sec, '透明度', 0, 100, 'wm_alpha', scale=0.01,
                         init=35)
        var_tile = tk.BooleanVar(value=True)
        self._resetters.append(lambda: var_tile.set(True))
        ttk.Checkbutton(sec, text='平铺水印', variable=var_tile,
                        command=lambda: self._setp('wm_tile',
                                                   var_tile.get())).pack(
            fill='x', padx=6, pady=(0, 6))

        # ---- 装饰
        sec = ttk.Labelframe(inner, text=' 装饰 ')
        sec.pack(fill='x', pady=(0, 8))
        self._add_slider(sec, '圆角', 0, 200, 'round_r')
        self._add_slider(sec, '边框', 0, 60, 'border_w')
        rowf = tk.Frame(sec, bg=p['bg'])
        rowf.pack(fill='x', padx=6, pady=(2, 6))
        tk.Label(rowf, text='边框颜色', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted'], width=7, anchor='w').pack(side='left')
        self._border_btn = tk.Button(rowf, text='      ', relief='flat', bd=1,
                                     bg=self.p['border_color'],
                                     activebackground=p['hover'],
                                     cursor='hand2', width=6,
                                     command=self._pick_border_color)
        self._border_btn.pack(side='left')

        # ---- 导出参数
        sec = ttk.Labelframe(inner, text=' 导出设置 ')
        sec.pack(fill='x', pady=(0, 8))
        rowf = tk.Frame(sec, bg=p['bg'])
        rowf.pack(fill='x', padx=6, pady=(6, 2))
        tk.Label(rowf, text='格式', font=ui_theme.font(9), bg=p['bg'],
                 fg=p['muted'], width=7, anchor='w').pack(side='left')
        self._fmt_var = tk.StringVar(value='原格式')
        ttk.Combobox(rowf, values=list(FORMATS), state='readonly',
                     textvariable=self._fmt_var, width=10).pack(side='left')

        self._quality_slider = SliderRow(sec, '质量', 10, 100, 85,
                                         command=lambda v: None)
        self._quality_slider.pack(fill='x', padx=6, pady=(4, 0))

        self._var_exif = tk.BooleanVar(value=True)
        ttk.Checkbutton(sec, text='清除 EXIF（拍摄位置等隐私）',
                        variable=self._var_exif).pack(
            fill='x', padx=6, pady=(4, 2))

        rowf = tk.Frame(sec, bg=p['bg'])
        rowf.pack(fill='x', padx=6, pady=(2, 2))
        self._var_target = tk.BooleanVar(value=False)
        ttk.Checkbutton(rowf, text='压缩到', variable=self._var_target).pack(
            side='left')
        self._target_var = tk.StringVar(value='500')
        ttk.Entry(rowf, textvariable=self._target_var, width=6).pack(
            side='left', padx=(4, 2))
        tk.Label(rowf, text='KB（JPG/WebP 效果好）', font=ui_theme.font(8),
                 bg=p['bg'], fg=p['faint']).pack(side='left')

        tk.Label(sec, text='批量导出会把右侧参数套到队列里的每张原图；'
                           '裁剪等区域工具只作用于当前图。',
                 font=ui_theme.font(8), bg=p['bg'], fg=p['faint'],
                 wraplength=280, justify='left').pack(
            fill='x', padx=6, pady=(2, 6))

    def _add_slider(self, parent, text, lo, hi, key, scale=1.0, init=None):
        dv = DEFAULTS[key] if init is None else init
        row = SliderRow(parent, text, lo, hi, dv,
                        command=lambda v, k=key: self._setp(k, v * scale))
        row.pack(fill='x', padx=6, pady=2)
        self._resetters.append(lambda r=row, v=dv: r.set(v))
        return row

    def _text_trace(self, key, var, cast):
        def handler(*_a):
            if self._text_after:
                try:
                    self.win.after_cancel(self._text_after)
                except Exception:
                    pass
            try:
                self._text_after = self.win.after(
                    350, lambda: self._setp(key, cast(var.get())))
            except Exception:
                pass
        return handler

    def _pick_border_color(self):
        try:
            rgb, hexv = colorchooser.askcolor(
                color=self.p.get('border_color', '#ffffff'),
                parent=self.win, title='边框颜色')
        except Exception:
            return
        if hexv:
            self._setp('border_color', hexv)
            self._border_btn.configure(bg=hexv)

    def _pick_text_color(self):
        try:
            rgb, hexv = colorchooser.askcolor(
                color=self.p.get('text_color', '#ffffff'),
                parent=self.win, title='文字颜色')
        except Exception:
            return
        if hexv:
            self._setp('text_color', hexv)
            self._text_color_btn.configure(bg=hexv)

    def _text_pos_pick(self):
        v = self._text_pos_cb.get()
        if v != '自定义':
            self.p['text_xy'] = None     # 预设位置清除自由坐标
        self._setp('text_pos', v)

    # ---- 文字自由拖动（无工具时按住画面即可拖）
    def _bump_work(self):
        """底图/工具预览变化后，让「无文字」缓存失效。"""
        self._work_ver += 1

    def _ensure_no_text_cache(self):
        sig = tuple((k, self.p[k]) for k in sorted(self.p)
                    if k not in ('text', 'text_pos', 'text_xy'))
        key = (self._work_ver, sig)
        if self._img_no_text is not None and self._no_text_key == key:
            return
        work = self._tool_preview if self._tool_preview is not None \
            else self._base_small
        if work is None:
            self._img_no_text = None
            return
        pp = dict(self.p)
        pp['text'] = ''
        try:
            self._img_no_text = process_params(work, pp)
            self._no_text_key = key
        except Exception:
            self._img_no_text = None

    def _text_drag_start(self, e):
        self._ensure_no_text_cache()
        if self._img_no_text is None:
            return
        self.p['text_pos'] = '自定义'
        try:
            self._text_pos_cb.set('自定义')   # 不触发选中回调
        except Exception:
            pass
        self._drag = {'text': True}
        self._text_drag_move(e)

    def _text_drag_move(self, e):
        if self._img_no_text is None or self._base is None:
            return
        bx, by = self._v2b(e.x, e.y)
        bw, bh = self._base.size
        self.p['text_xy'] = (max(0.0, min(1.0, bx / float(bw))),
                             max(0.0, min(1.0, by / float(bh))))
        im = self._img_no_text.copy()
        try:
            _draw_text(im, self.p['text'],
                       size=int(self.p.get('text_size') or 0),
                       color=self.p.get('text_color') or '#ffffff',
                       stroke=bool(self.p.get('text_stroke')),
                       center=self.p['text_xy'],
                       opacity=int(self.p.get('text_opacity') or 100))
        except Exception:
            return
        self._img = im
        self._redraw_photo()

    # ================================================== 参数 / 预览
    def _setp(self, key, value):
        self.p[key] = value
        self._schedule()

    def _schedule(self, delay=280):
        if self._after_id:
            try:
                self.win.after_cancel(self._after_id)
            except Exception:
                pass
        self._after_id = self.win.after(delay, self._refresh_preview)

    def _refresh_preview(self):
        self._after_id = None
        if self._base is None or not self.win:
            return
        im = self._tool_preview if self._tool_preview is not None \
            else self._base_small
        if im is None:
            return
        pp = self.p
        if self._tool:
            # 区域工具编辑期间不套边框/圆角，避免坐标错位
            pp = dict(self.p)
            pp['border_w'] = 0
            pp['round_r'] = 0
        try:
            self._img = process_params(im, pp)
        except Exception as exc:
            self._info_lbl.configure(text='处理出错：%s' % exc)
            return
        if not self._compare_on:
            self._redraw_photo()
        self._update_status()
        if not self._tool:
            self._hint_lbl.configure(
                text='按住画面可直接拖动文字'
                if (self.p.get('text') or '').strip() else '')

    def _redraw_photo(self):
        cv = self._canvas
        try:
            cv.delete('pic', 'ov')
        except Exception:
            return
        im = self._img
        cw = max(1, cv.winfo_width())
        ch = max(1, cv.winfo_height())
        if im is None:
            cv.create_text(cw // 2, ch // 2 - 20, tags='pic',
                           text='打开一张图片开始', fill='#5d6a85',
                           font=ui_theme.font(13, 'bold'))
            cv.create_text(cw // 2, ch // 2 + 14, tags='pic',
                           text='支持 PNG / JPG / WebP / BMP，'
                                '可多选或整个文件夹批量载入',
                           fill='#42506e', font=ui_theme.font(9))
            return
        photo = ImageTk.PhotoImage(im)
        self._photo = photo
        ox = max(0, (cw - im.width) // 2)
        oy = max(0, (ch - im.height) // 2)
        self._view_off = (ox, oy)
        cv.create_image(ox, oy, anchor='nw', image=photo, tags='pic')
        self._redraw_overlay()

    def _update_status(self):
        if self._base is None:
            self._info_lbl.configure(text='未载入图片')
            return
        try:
            size_mb = os.path.getsize(self._path) / 1048576.0
            size_txt = '%.2f MB' % size_mb if size_mb >= 1 \
                else '%.0f KB' % (size_mb * 1024)
        except Exception:
            size_txt = ''
        out = self._img
        out_txt = '输出 %d×%d' % (out.width, out.height) if out else ''
        self._info_lbl.configure(
            text='%s · 原图 %d×%d · %s · %s' % (
                os.path.basename(self._path or ''), self._base.width,
                self._base.height, size_txt, out_txt))

    def _update_actions(self):
        has = self._base is not None
        self._btn_undo.set_enabled(bool(self._undo))
        self._btn_reset.set_enabled(has)
        self._btn_save.set_enabled(has)
        self._btn_batch.set_enabled(bool(self._files))
        self._cmp_btn.configure(state='normal' if has else 'disabled')

    # ================================================== 载入 / 队列
    def _pick(self):
        paths = filedialog.askopenfilenames(
            parent=self.win, title='选择图片',
            filetypes=[('图片', '*.png *.jpg *.jpeg *.bmp *.webp *.gif'),
                       ('所有文件', '*.*')])
        if not paths:
            return
        self._files = list(paths)
        self._refresh_thumbs()
        self._load(paths[0])

    def _pick_folder(self):
        folder = filedialog.askdirectory(parent=self.win,
                                         title='选择文件夹（批量）')
        if not folder:
            return
        exts = ('.png', '.jpg', '.jpeg', '.bmp', '.webp', '.gif')
        self._files = [os.path.join(folder, n) for n in os.listdir(folder)
                       if n.lower().endswith(exts)]
        if not self._files:
            messagebox.showinfo('图片工具箱', '文件夹里没有图片', parent=self.win)
            return
        self._refresh_thumbs()
        self._load(self._files[0])

    def _load(self, path):
        try:
            img = Image.open(path)
            img.load()
            try:
                img = ImageOps.exif_transpose(img)
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror('图片工具箱', '无法打开图片：%s' % exc,
                                 parent=self.win)
            return
        self._cancel_tool()
        self._path = path
        self._exif = img.info.get('exif')
        self._original = img.copy()
        self._base = img.copy()
        self._undo = []
        self._compare_on = False
        self._reset_params()
        self._rebuild_small()
        self._refresh_preview()
        self._update_actions()
        self._highlight_thumb()

    def _reset_params(self):
        self.p = dict(DEFAULTS)
        for fn in self._resetters:
            try:
                fn()
            except Exception:
                pass

    def _rebuild_small(self):
        """按画布尺寸生成底图降采样（预览提速）。"""
        if self._base is None:
            self._base_small = None
            return
        cv = self._canvas
        cw = max(240, cv.winfo_width() - 24)
        ch = max(180, cv.winfo_height() - 24)
        bw, bh = self._base.size
        scale = min(cw / float(bw), ch / float(bh), 1.0)
        sw, sh = max(1, int(bw * scale)), max(1, int(bh * scale))
        if (sw, sh) == (bw, bh):
            self._base_small = self._base.copy()
        else:
            self._base_small = self._base.resize((sw, sh), Image.LANCZOS)
        self._small_ratio = bw / float(sw)
        self._bump_work()

    def _on_canvas_config(self, event):
        size = (event.width, event.height)
        if abs(size[0] - self._canvas_size[0]) < 24 and \
                abs(size[1] - self._canvas_size[1]) < 24:
            return
        self._canvas_size = size
        if self._base is None:
            self._redraw_photo()
            return
        if self._tool:                      # 工具进行中先不动，收尾再重算
            self._pending_resize = True
            return
        if self._resize_after:
            try:
                self.win.after_cancel(self._resize_after)
            except Exception:
                pass
        self._resize_after = self.win.after(
            180, self._do_pending_resize)

    def _do_pending_resize(self):
        self._resize_after = None
        if self._base is None or self._tool:
            return
        self._rebuild_small()
        self._refresh_preview()

    # ---- 缩略图条
    def _refresh_thumbs(self):
        strip = self._strip
        strip.delete('all')
        if not self._files:
            strip.pack_forget()
            self._count_lbl.configure(text='')
            return
        strip.pack(fill='x', padx=14, pady=(0, 6))
        self._count_lbl.configure(text='共 %d 张' % len(self._files))
        self._keep = {p: ph for p, ph in self._keep.items()
                      if p in self._files}
        x = 4
        for i, path in enumerate(self._files):
            tag = 'f%d' % i
            if path not in self._keep or self._keep[path] is None:
                try:
                    im = Image.open(path)
                    im.thumbnail((44, 44))
                    if im.mode not in ('RGB', 'RGBA'):
                        im = im.convert('RGBA')
                    self._keep[path] = ImageTk.PhotoImage(im)
                except Exception:
                    self._keep[path] = None
            photo = self._keep[path]
            if photo:
                strip.create_image(x, 2, anchor='nw', image=photo, tags=tag)
            strip.create_rectangle(x - 2, 0, x + 48, 48, tags=(tag, 'tile'),
                                   outline=ui_theme.PALETTE['border'],
                                   width=1)
            name = os.path.basename(path)
            if len(name) > 9:
                name = name[:8] + '…'
            strip.create_text(x + 23, 52, text=name, tags=tag,
                              fill=ui_theme.PALETTE['muted'],
                              font=ui_theme.font(7))
            strip.tag_bind(tag, '<Button-1>',
                           lambda e, pth=path: self._load(pth))
            x += 54
        strip.configure(scrollregion=(0, 0, x, 66))

    def _highlight_thumb(self):
        strip = self._strip
        strip.delete('hl')
        if self._path not in self._files:
            return
        i = self._files.index(self._path)
        x = 4 + i * 54
        strip.create_rectangle(x - 3, -1, x + 49, 49, tags='hl',
                               outline=ui_theme.PALETTE['primary'], width=2)

    # ================================================== 区域工具
    def _toggle_tool(self, key):
        if self._base is None:
            return
        if self._tool == key:
            self._cancel_tool()
        else:
            self._set_tool(key)

    def _set_tool(self, key):
        self._cancel_tool()
        self._tool = key
        p = ui_theme.PALETTE
        for k, b in self._tool_btns.items():
            if k == key:
                b.configure(bg=p['primary'], fg='#ffffff',
                            activebackground=p['primary_dark'])
            else:
                b.configure(bg=p['surface'], fg=p['text'],
                            activebackground=p['hover'])
        self._tp[key].grid()
        self._canvas.configure(cursor='crosshair')
        bw, bh = self._base.size
        if key == 'persp':
            self._quad = [(0.0, 0.0), (float(bw), 0.0),
                          (float(bw), float(bh)), (0.0, float(bh))]
        self._hint_lbl.configure(text={'crop': '拖拽画裁剪框，可反复重画',
                                       'persp': '拖四角对准要拉正的区域',
                                       'mosaic': '拖拽画打码区域，可重复',
                                       'inpaint': '涂抹去除区域，松开自动预览'
                                       }.get(key, ''))
        self._redraw_overlay()

    def _cancel_tool(self):
        self._tool = None
        self._drag = None
        self._rect = None
        self._quad = None
        self._strokes = []
        if self._tool_preview is not None:
            self._tool_preview = None
            self._bump_work()
        p = ui_theme.PALETTE
        for b in self._tool_btns.values():
            b.configure(bg=p['surface'], fg=p['text'],
                        activebackground=p['hover'])
        for f in self._tp.values():
            f.grid_remove()
        if self.win:
            self._canvas.configure(cursor='')
            self._hint_lbl.configure(text='')
            self._redraw_photo()
            if self._pending_resize:
                self._pending_resize = False
                self._rebuild_small()
                self._refresh_preview()

    def _close_tool_commit(self):
        """工具应用成功后的收尾。"""
        self._tool = None
        self._drag = None
        self._rect = None
        self._quad = None
        self._strokes = []
        self._tool_preview = None
        p = ui_theme.PALETTE
        for b in self._tool_btns.values():
            b.configure(bg=p['surface'], fg=p['text'],
                        activebackground=p['hover'])
        for f in self._tp.values():
            f.grid_remove()
        self._canvas.configure(cursor='')
        self._hint_lbl.configure(text='')
        self._rebuild_small()
        self._refresh_preview()
        if self._pending_resize:
            self._pending_resize = False
            self._rebuild_small()
            self._refresh_preview()

    def _push_undo(self):
        if self._base is None:
            return
        self._undo.append(self._base.copy())
        if len(self._undo) > 3:
            self._undo.pop(0)
        self._update_actions()

    def _undo_last(self):
        if self._tool:
            self._cancel_tool()
        if not self._undo:
            return
        self._base = self._undo.pop()
        self._rebuild_small()
        self._refresh_preview()
        self._update_actions()

    def _reset_current(self):
        if self._original is None:
            return
        if self._undo and not messagebox.askyesno(
                '图片工具箱', '重置会丢弃已做的裁剪等区域操作，确定？',
                parent=self.win):
            return
        if self._tool:
            self._cancel_tool()
        self._base = self._original.copy()
        self._undo = []
        self._reset_params()
        self._rebuild_small()
        self._refresh_preview()
        self._update_actions()

    # ---- 画布事件
    def _v2b(self, x, y):
        r = self._small_ratio or 1.0
        ox, oy = self._view_off
        return ((x - ox) * r, (y - oy) * r)

    def _b2v(self, bx, by):
        r = self._small_ratio or 1.0
        ox, oy = self._view_off
        return (bx / r + ox, by / r + oy)

    def _cv_press(self, e):
        if self._compare_on or self._base is None:
            return
        if not self._tool:
            if (self.p.get('text') or '').strip():
                self._text_drag_start(e)
            return
        bx, by = self._v2b(e.x, e.y)
        bw, bh = self._base.size
        bx = max(0, min(bx, bw))
        by = max(0, min(by, bh))
        if self._tool in ('crop', 'mosaic'):
            self._drag = {'x0': bx, 'y0': by}
            self._rect = (bx, by, bx, by)
        elif self._tool == 'inpaint':
            r = self._brush_slider.get()
            self._drag = {'last': (bx, by)}
            self._strokes.append((bx, by, r))
        elif self._tool == 'persp':
            vx, vy = e.x, e.y
            best, bi = 18 ** 2, -1
            for i, (qx, qy) in enumerate(self._quad):
                px, py = self._b2v(qx, qy)
                d = (px - vx) ** 2 + (py - vy) ** 2
                if d < best:
                    best, bi = d, i
            if bi >= 0:
                self._drag = {'handle': bi}

    def _cv_motion(self, e):
        if self._compare_on or self._drag is None:
            return
        if self._drag.get('text'):
            self._text_drag_move(e)
            return
        bx, by = self._v2b(e.x, e.y)
        bw, bh = self._base.size
        bx = max(0, min(bx, bw))
        by = max(0, min(by, bh))
        if self._tool in ('crop', 'mosaic'):
            self._rect = (self._drag['x0'], self._drag['y0'], bx, by)
            self._redraw_overlay()
        elif self._tool == 'persp':
            self._quad[self._drag['handle']] = (bx, by)
            self._redraw_overlay()
        elif self._tool == 'inpaint':
            lx, ly = self._drag['last']
            r = self._brush_slider.get()
            step = max(2.0, r * 0.3)
            if math.hypot(bx - lx, by - ly) >= step:
                self._strokes.append((bx, by, r))
                self._drag['last'] = (bx, by)
                self._redraw_overlay()

    def _cv_release(self, e):
        if self._compare_on or self._drag is None:
            return
        if self._drag.get('text'):
            self._drag = None
            self._refresh_preview()     # 走一次正式管线，保持一致
            return
        tool = self._tool
        self._drag = None
        if tool in ('crop',):
            self._redraw_overlay()
        elif tool == 'mosaic':
            if self._rect and self._rect_valid(self._rect):
                self._mosaic_redo()
            self._redraw_overlay()
        elif tool == 'inpaint':
            self._inpaint_preview()
        elif tool == 'persp':
            self._redraw_overlay()

    @staticmethod
    def _rect_valid(rect):
        x1, y1, x2, y2 = rect
        return abs(x2 - x1) >= 4 and abs(y2 - y1) >= 4

    def _norm_rect(self, rect):
        x1, y1, x2, y2 = rect
        return (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))

    def _mosaic_redo(self, *_a):
        """马赛克实时预览（作用于小图）。"""
        if self._tool != 'mosaic' or not self._rect \
                or not self._rect_valid(self._rect):
            return
        r = self._small_ratio or 1.0
        x1, y1, x2, y2 = self._norm_rect(self._rect)
        block = self._mosaic_slider.get()
        rect_s = (x1 / r, y1 / r, x2 / r, y2 / r)
        try:
            self._tool_preview = _mosaic_region(self._base_small, rect_s,
                                                max(2, block / r))
            self._bump_work()
            self._schedule(0)
        except Exception:
            pass

    def _inpaint_preview(self):
        if self._tool != 'inpaint' or not self._strokes:
            return
        if cv2 is None or self._base_small is None:
            return
        r = self._small_ratio or 1.0
        mask = Image.new('L', self._base_small.size, 0)
        d = ImageDraw.Draw(mask)
        for (sx, sy, sr) in self._strokes:
            d.ellipse((sx / r - sr / r, sy / r - sr / r,
                       sx / r + sr / r, sy / r + sr / r), fill=255)
        try:
            self._tool_preview = _inpaint(self._base_small, mask)
            self._bump_work()
            self._schedule(0)
        except Exception:
            pass

    def _strokes_clear(self):
        self._strokes = []
        self._tool_preview = None
        self._bump_work()
        self._refresh_preview()

    def _persp_reset(self):
        if self._base is None:
            return
        bw, bh = self._base.size
        self._quad = [(0.0, 0.0), (float(bw), 0.0),
                      (float(bw), float(bh)), (0.0, float(bh))]
        self._redraw_overlay()

    def _apply_tool(self):
        if self._base is None or not self._tool:
            return
        tool = self._tool
        if tool == 'crop':
            if not self._rect or not self._rect_valid(self._rect):
                messagebox.showinfo('图片工具箱', '先在图上拖出裁剪框',
                                    parent=self.win)
                return
            x1, y1, x2, y2 = self._norm_rect(self._rect)
            rect = (max(0, int(round(x1))), max(0, int(round(y1))),
                    min(self._base.width, int(round(x2))),
                    min(self._base.height, int(round(y2))))
            if rect[2] - rect[0] < 8 or rect[3] - rect[1] < 8:
                messagebox.showinfo('图片工具箱', '裁剪区域太小', parent=self.win)
                return
            self._push_undo()
            self._base = self._base.crop(rect)
        elif tool == 'persp':
            if not self._quad:
                return
            bw, bh = self._base.size
            for (qx, qy) in self._quad:
                if qx < -bw * 0.1 or qx > bw * 1.1 or qy < -bh * 0.1 \
                        or qy > bh * 1.1:
                    messagebox.showinfo('图片工具箱', '角点拖得太远了',
                                        parent=self.win)
                    return
            self._push_undo()
            self._base = _warp_quad(self._base, self._quad)
        elif tool == 'mosaic':
            if not self._rect or not self._rect_valid(self._rect):
                messagebox.showinfo('图片工具箱', '先在图上拖出打码区域',
                                    parent=self.win)
                return
            self._push_undo()
            self._base = _mosaic_region(self._base,
                                        self._norm_rect(self._rect),
                                        self._mosaic_slider.get())
        elif tool == 'inpaint':
            if not self._strokes:
                messagebox.showinfo('图片工具箱', '先涂抹要去除的区域',
                                    parent=self.win)
                return
            if cv2 is None:
                return
            mask = Image.new('L', self._base.size, 0)
            d = ImageDraw.Draw(mask)
            for (sx, sy, sr) in self._strokes:
                d.ellipse((sx - sr, sy - sr, sx + sr, sy + sr), fill=255)
            self._push_undo()
            self._base = _inpaint(self._base, mask)
        self._close_tool_commit()
        self._update_actions()

    # ---- 覆盖层绘制
    def _redraw_overlay(self):
        cv = self._canvas
        try:
            cv.delete('ov')
        except Exception:
            return
        if self._compare_on or self._base is None or not self._tool:
            return
        p = ui_theme.PALETTE
        cw = cv.winfo_width()
        ch = cv.winfo_height()
        if self._tool in ('crop', 'mosaic') and self._rect:
            vx1, vy1 = self._b2v(self._rect[0], self._rect[1])
            vx2, vy2 = self._b2v(self._rect[2], self._rect[3])
            rx1, rx2 = min(vx1, vx2), max(vx1, vx2)
            ry1, ry2 = min(vy1, vy2), max(vy1, vy2)
            if self._tool == 'crop':
                for quad in ((0, 0, cw, ry1), (0, ry2, cw, ch),
                             (0, ry1, rx1, ry2), (rx2, ry1, cw, ry2)):
                    cv.create_rectangle(*quad, fill='#000000',
                                        stipple='gray50', outline='',
                                        tags='ov')
            cv.create_rectangle(rx1, ry1, rx2, ry2, outline='#4da3ff',
                                width=2, tags='ov')
            if self._tool == 'crop':
                bw = abs(self._rect[2] - self._rect[0])
                bh = abs(self._rect[3] - self._rect[1])
                cv.create_rectangle(rx2 - 64, ry2 + 4, rx2, ry2 + 22,
                                    fill='#000000', outline='', tags='ov')
                cv.create_text(rx2 - 32, ry2 + 13, text='%d×%d'
                               % (int(bw), int(bh)), fill='#ffffff',
                               font=ui_theme.font(8), tags='ov')
        elif self._tool == 'persp' and self._quad:
            pts = []
            for (qx, qy) in self._quad:
                pts.extend(self._b2v(qx, qy))
            cv.create_polygon(*pts, fill='', outline='#4da3ff', width=2,
                              tags='ov')
            for (qx, qy) in self._quad:
                px, py = self._b2v(qx, qy)
                cv.create_rectangle(px - 6, py - 6, px + 6, py + 6,
                                    fill='#ffffff', outline=p['primary'],
                                    width=2, tags='ov')
        elif self._tool == 'inpaint' and self._strokes:
            r = self._small_ratio or 1.0
            for (sx, sy, sr) in self._strokes:
                px, py = self._b2v(sx, sy)
                pr = sr / r
                cv.create_oval(px - pr, py - pr, px + pr, py + pr,
                               fill='#ff7a45', stipple='gray50',
                               outline='', tags='ov')

    # ---- 对比原图
    def _compare_show(self):
        if self._original is None or self._compare_on:
            return
        self._compare_on = True
        cv = self._canvas
        cv.delete('ov')
        im = self._original.copy()
        cw = max(1, cv.winfo_width() - 20)
        ch = max(1, cv.winfo_height() - 20)
        im.thumbnail((cw, ch), Image.LANCZOS)
        self._raw_photo = ImageTk.PhotoImage(im)
        ox = max(0, (cv.winfo_width() - im.width) // 2)
        oy = max(0, (cv.winfo_height() - im.height) // 2)
        cv.delete('pic')
        cv.create_image(ox, oy, anchor='nw', image=self._raw_photo,
                        tags='pic')
        cv.create_text(ox + 8, oy + 8, text='原图', anchor='nw',
                       fill='#ffffff', font=ui_theme.font(9, 'bold'),
                       tags='pic')

    def _compare_hide(self):
        if not self._compare_on:
            return
        self._compare_on = False
        self._redraw_photo()

    # ================================================== 拼接
    def _stitch(self):
        paths = filedialog.askopenfilenames(
            parent=self.win, title='选择要拼接的图片（2 张以上）',
            filetypes=[('图片', '*.png *.jpg *.jpeg *.bmp *.webp'),
                       ('所有文件', '*.*')])
        if len(paths) < 2:
            if paths:
                messagebox.showinfo('图片工具箱', '至少选择 2 张图片',
                                    parent=self.win)
            return
        p = ui_theme.PALETTE
        dlg = tk.Toplevel(self.win)
        dlg.title('拼接图片')
        dlg.configure(bg=p['bg'])
        dlg.transient(self.win)
        dlg.resizable(False, False)
        dlg.protocol('WM_DELETE_WINDOW', dlg.destroy)
        # 按钮栏先放底部，防止被上面的内容挤出窗口
        btns = tk.Frame(dlg, bg=p['bg'])
        btns.pack(side='bottom', fill='x', padx=12, pady=10)

        box = tk.Frame(dlg, bg=p['bg'])
        box.pack(fill='both', expand=True, padx=14, pady=(14, 6))
        var_h = tk.BooleanVar(value=True)
        ttk.Radiobutton(box, text='横向拼接（左右排）', value=True,
                        variable=var_h).pack(anchor='w')
        ttk.Radiobutton(box, text='纵向拼接（上下排）', value=False,
                        variable=var_h).pack(anchor='w', pady=(2, 6))
        rowf = tk.Frame(box, bg=p['bg'])
        rowf.pack(fill='x')
        tk.Label(rowf, text='图片间距(px)', font=ui_theme.font(9),
                 bg=p['bg'], fg=p['muted']).pack(side='left')
        gap_var = tk.StringVar(value='12')
        ttk.Entry(rowf, textvariable=gap_var, width=6).pack(side='left',
                                                            padx=(6, 0))
        tk.Label(box, text='按较小的一边对齐缩放，白底填充。',
                 font=ui_theme.font(8), bg=p['bg'],
                 fg=p['faint']).pack(anchor='w', pady=(8, 0))

        def do_stitch():
            out = stitch_images(paths, horizontal=var_h.get(),
                                gap=_to_int(gap_var.get(), 12))
            if out is None:
                return
            dest = filedialog.asksaveasfilename(
                parent=dlg, title='保存拼接图', defaultextension='.png',
                initialfile='拼接图.png',
                filetypes=[('PNG', '*.png'), ('JPEG', '*.jpg')])
            if not dest:
                return
            try:
                if dest.lower().endswith(('.jpg', '.jpeg')):
                    out.save(dest, 'JPEG', quality=92)
                else:
                    out.save(dest, 'PNG')
            except Exception as exc:
                messagebox.showerror('图片工具箱', '保存失败：%s' % exc,
                                     parent=dlg)
                return
            dlg.destroy()
            messagebox.showinfo('图片工具箱', '已保存：\n%s' % dest,
                                parent=self.win)

        _flat_btn(btns, '生成并保存…', do_stitch, 'primary').pack(side='right')
        _flat_btn(btns, '取消', dlg.destroy).pack(side='right', padx=(0, 8))
        dlg.update_idletasks()
        try:
            x = self.win.winfo_rootx() + (self.win.winfo_width()
                                          - dlg.winfo_width()) // 2
            y = self.win.winfo_rooty() + (self.win.winfo_height()
                                          - dlg.winfo_height()) // 2
            dlg.geometry('+%d+%d' % (max(0, x), max(0, y)))
        except Exception:
            pass

    # ================================================== 保存 / 导出
    def _export_path(self, src, ext):
        stem = os.path.splitext(os.path.basename(src))[0]
        folder = ensure_folder(os.path.join(
            os.path.expanduser('~'), 'Documents', '会务图片输出'))
        return unique_path(folder, stem + ext), folder

    def _fmt_ext(self, src):
        fmt = self._fmt_var.get()
        if fmt == '原格式':
            ext = (os.path.splitext(src)[1] or '.png').lower()
            if ext not in ('.jpg', '.jpeg', '.png', '.webp'):
                ext = '.png'
            return fmt, ext
        return fmt, _FORMAT_EXT.get(fmt, '.png')

    def _target_kb(self):
        if not self._var_target.get():
            return 0
        return max(0, _to_int(self._target_var.get(), 0))

    def _save_as(self):
        if self._base is None:
            messagebox.showinfo('图片工具箱', '请先选择图片', parent=self.win)
            return
        fmt, ext = self._fmt_ext(self._path or 'image.png')
        stem = os.path.splitext(os.path.basename(self._path or 'image'))[0]
        dest = filedialog.asksaveasfilename(
            parent=self.win, defaultextension=ext, initialfile=stem + ext,
            filetypes=[('图片', '*%s' % ext)])
        if not dest:
            return
        try:
            self.win.configure(cursor='watch')
            self.win.update_idletasks()
            img = process_params(self._base, self.p)
            exif = None if self._var_exif.get() else self._exif
            kb = _save_image(img, dest, ext,
                             quality=self._quality_slider.get(),
                             exif=exif, target_kb=self._target_kb())
            messagebox.showinfo(
                '图片工具箱', '已保存：\n%s\n（%.0f KB）'
                % (dest, kb / 1024.0), parent=self.win)
        except Exception as exc:
            messagebox.showerror('图片工具箱', '保存失败：%s' % exc,
                                 parent=self.win)
        finally:
            self.win.configure(cursor='')

    def _export_batch(self):
        if not self._files:
            messagebox.showinfo('图片工具箱', '尚未选择图片', parent=self.win)
            return
        self.win.configure(cursor='watch')
        ok = 0
        last_err = None
        try:
            self.win.update_idletasks()
            quality = self._quality_slider.get()
            target = self._target_kb()
            folder = None
            for src in self._files:
                try:
                    img = Image.open(src)
                    try:
                        img = ImageOps.exif_transpose(img)
                    except Exception:
                        pass
                    img = process_params(img, self.p)
                    _, ext = self._fmt_ext(src)
                    exif = None
                    if not self._var_exif.get():
                        exif = img.info.get('exif') or \
                            Image.open(src).info.get('exif')
                    dest, folder = self._export_path(src, ext)
                    _save_image(img, dest, ext, quality=quality, exif=exif,
                                target_kb=target)
                    ok += 1
                except Exception as exc:
                    last_err = exc
                    continue
        finally:
            self.win.configure(cursor='')
        if ok:
            messagebox.showinfo(
                '图片工具箱', '已导出 %d/%d 张到：\n%s'
                % (ok, len(self._files), folder or ''), parent=self.win)
        else:
            messagebox.showerror('图片工具箱', '导出失败：%s' % (last_err or '未知错误'),
                                 parent=self.win)

    # ================================================== 关闭
    def close(self, force=False):
        for aid in (self._after_id, self._text_after, self._resize_after):
            if aid:
                try:
                    self.win.after_cancel(aid)
                except Exception:
                    pass
        try:
            self.win.unbind_all('<MouseWheel>')
        except Exception:
            pass
        try:
            self.win.destroy()
        except Exception:
            pass
        self.win = None
