# -*- coding: utf-8 -*-
import os
import json
import re
import shutil
import sys
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont, ImageTk

import safe_guard
import ui_theme

try:
    import winreg
except ImportError:
    winreg = None


def get_font_dir():
    return os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')


CARD_CM = 19.9  # 卡面在 Word 里的打印尺寸（厘米）
CENTER_SHIFT_CM = 0.5  # 正反向文字各向中间折线靠拢的厘米数（中间空白共缩小 1cm）

# ---- 资源上限：防止极端输入把内存撑爆导致程序崩溃 ----
MAX_NAMES = 2000          # 一次最多生成/打印多少张座位牌
MAX_NAME_LEN = 40         # 单个姓名的最大字符数
MAX_CANVAS = 2048         # 卡面边长（像素）上限
MAX_UPLOAD_BYTES = 180 * 1024 * 1024   # 远程打印上传压缩包上限

# 图片解码上限：超过约 8000 万像素的图片直接拒绝，避免解码时内存暴涨
MAX_IMAGE_PIXELS = 80 * 1000 * 1000
try:
    Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS
except Exception:
    pass

if getattr(sys, 'frozen', False):
    # 打包成 exe 后，内置文件在临时解压目录里
    _BASE_DIR = sys._MEIPASS
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BUNDLED_FONT_DIR = os.path.join(_BASE_DIR, 'fonts')
# 软件自带的字体：显示名 -> 内置文件名
BUNDLED_FONTS = {
    '华文中宋': '华文中宋.ttf',
    '华文行楷': '华文行楷.ttf',
}


_FONT_FILE_MAP = {
    '黑体': 'simhei.ttf', '楷体': 'simkai.ttf', '宋体': 'simsun.ttc',
    '仿宋': 'simfang.ttf', '微软雅黑': 'msyh.ttc', '微软雅黑 Bold': 'msyhbd.ttc',
    '等线': 'deng.ttf', '华文仿宋': 'stfangsong.ttf', '华文楷体': 'stkaiti.ttf',
    '华文中宋': 'stzhongsong.ttf', '华文宋体': 'stsongs.ttf', '华文细黑': 'sthupo.ttf',
    '华文新魏': 'stxinqi.ttf',
    '华文行楷': 'stxingka.ttf', '华文琥珀': 'sthuapo.ttf', '华文彩云': 'stcaiyun.ttf',
    '方正舒体': 'fzshut.ttf', '方正姚体': 'fzyao.ttf',
    '新宋体': 'nsimsun.ttc', '幼圆': 'simyou.ttf',
    'Arial': 'arial.ttf', 'Times New Roman': 'times.ttf',
    'Impact': 'impact.ttf', 'Comic Sans MS': 'comic.ttf', 'Georgia': 'georgia.ttf',
}


def _registry_fonts():
    """从 Windows 注册表读取所有已安装字体（含用户单独安装的字体）。"""
    fonts = {}
    if winreg is None:
        return fonts
    subkeys = [
        (winreg.HKEY_LOCAL_MACHINE,
         r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'),
        (winreg.HKEY_LOCAL_MACHINE,
         r'SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion\Fonts'),
        (winreg.HKEY_CURRENT_USER,
         r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'),
    ]
    for hive, key in subkeys:
        try:
            k = winreg.OpenKey(hive, key)
        except OSError:
            continue
        try:
            i = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(k, i)
                except OSError:
                    break
                i += 1
                if not isinstance(value, str) or not value:
                    continue
                path = value
                if not os.path.isabs(path):
                    path = os.path.join(get_font_dir(), value)
                if not os.path.exists(path):
                    continue
                # 去掉“(TrueType)/(OpenType)”这类后缀，只留字体名
                clean = re.sub(
                    r'\s*\((TrueType|OpenType|Type 1|All res|VGA res)\)\s*$',
                    '', name).strip()
                if clean:
                    fonts[clean] = path
        finally:
            winreg.CloseKey(k)
    return fonts


_FONTS_CACHE = None


def _installed_fonts():
    global _FONTS_CACHE
    if _FONTS_CACHE is None:
        _FONTS_CACHE = _registry_fonts()
    return _FONTS_CACHE


def user_fonts_dir():
    """Windows 用户字体目录（无需管理员权限即可安装字体）。"""
    return os.path.join(os.environ.get('LOCALAPPDATA', ''),
                        'Microsoft', 'Windows', 'Fonts')


def ensure_bundled_fonts():
    """若系统没有安装内置的华文中宋/华文行楷，自动复制并注册到用户字体目录。"""
    installed = _installed_fonts()
    need_install = []
    for display_name, fname in BUNDLED_FONTS.items():
        if display_name in installed:
            continue
        src = os.path.join(BUNDLED_FONT_DIR, fname)
        if os.path.exists(src):
            need_install.append((display_name, src, fname))
    for display_name, src, fname in need_install:
        try:
            fdir = user_fonts_dir()
            os.makedirs(fdir, exist_ok=True)
            dst = os.path.join(fdir, fname)
            if not os.path.exists(dst):
                shutil.copy2(src, dst)
            if winreg is not None:
                key = (r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
                k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0,
                                   winreg.KEY_SET_VALUE)
                winreg.SetValueEx(k, f'{display_name} (TrueType)', 0,
                                  winreg.REG_SZ, dst)
                winreg.CloseKey(k)
        except Exception:
            continue  # 安装失败不影响使用，后面仍可按内置字体文件渲染
    global _FONTS_CACHE
    _FONTS_CACHE = _registry_fonts()


def list_system_fonts():
    fonts = []
    seen = set()
    for display_name in _FONT_FILE_MAP:
        path = os.path.join(get_font_dir(), _FONT_FILE_MAP[display_name])
        if os.path.exists(path):
            fonts.append(display_name)
            seen.add(display_name)
    # 其它电脑上已安装的字体（以注册表为准，包含用户字体）
    for name in sorted(k for k in _installed_fonts() if k not in seen):
        fonts.append(name)
    # 软件自带字体（即使未安装成功也显示，保证可用）
    for display_name, fname in BUNDLED_FONTS.items():
        if display_name not in seen and \
                os.path.exists(os.path.join(BUNDLED_FONT_DIR, fname)):
            fonts.append(display_name)
            seen.add(display_name)
    if not fonts:
        # 注册表读不到时退回到按文件扫描
        import glob
        for ext in ('*.ttf', '*.ttc', '*.otf'):
            for f in glob.glob(os.path.join(get_font_dir(), ext)):
                fonts.append(os.path.splitext(os.path.basename(f))[0])
    return fonts


def get_font_path(font_name):
    if font_name in _FONT_FILE_MAP:
        path = os.path.join(get_font_dir(), _FONT_FILE_MAP[font_name])
        if os.path.exists(path):
            return path
    installed = _installed_fonts()
    if font_name in installed and os.path.exists(installed[font_name]):
        return installed[font_name]
    fname = BUNDLED_FONTS.get(font_name)
    if fname:
        path = os.path.join(BUNDLED_FONT_DIR, fname)
        if os.path.exists(path):
            return path
    fdir = get_font_dir()
    import glob
    for ext in ('*.ttf', '*.ttc', '*.otf'):
        for f in glob.glob(os.path.join(fdir, ext)):
            bn = os.path.splitext(os.path.basename(f))[0].lower()
            if font_name.lower().replace(' ', '') in bn or bn in font_name.lower().replace(' ', ''):
                return f
    return os.path.join(fdir, 'simhei.ttf')


def _draw_dashed_rect(img, box, dash=22, gap=13,
                      fill=(217, 217, 217), width=3):
    """画模板里的卡面虚线框。"""
    d = ImageDraw.Draw(img)
    x0, y0, x1, y1 = box

    def dash_line(a, b):
        dx, dy = b[0] - a[0], b[1] - a[1]
        length = (dx * dx + dy * dy) ** 0.5
        pos = 0
        while pos < length:
            end = min(pos + dash, length)
            d.line([(a[0] + dx * pos / length, a[1] + dy * pos / length),
                    (a[0] + dx * end / length, a[1] + dy * end / length)],
                   fill=fill, width=width)
            pos = end + gap

    dash_line((x0, y0), (x1, y0))
    dash_line((x1, y0), (x1, y1))
    dash_line((x1, y1), (x0, y1))
    dash_line((x0, y1), (x0, y0))


def generate(name, font_path, canvas_size=1024, font_size=280,
             show_border=False):
    """生成一张座位牌图片（RGB）。调用方用完请 img.close() 释放内存。"""
    name = name.strip()
    if not name:
        return None
    try:
        canvas_size = int(canvas_size)
    except (TypeError, ValueError):
        canvas_size = 1024
    canvas_size = max(200, min(canvas_size, MAX_CANVAS))
    try:
        font_size = int(font_size)
    except (TypeError, ValueError):
        font_size = 280
    font_size = max(8, min(font_size, canvas_size))
    n = len(name)
    row_h = canvas_size // 2
    center_shift = int(round(canvas_size * CENTER_SHIFT_CM / CARD_CM))
    margin_x = int(canvas_size * 0.08)  # 与模板一致：文字约占卡面宽 84%
    usable_w = canvas_size - 2 * margin_x

    font = ImageFont.truetype(font_path, int(font_size))

    char_imgs = []
    tmp_size = int(font_size) + 8
    for ch in name:
        tmp = Image.new('RGBA', (tmp_size, tmp_size), (0, 0, 0, 0))
        try:
            td = ImageDraw.Draw(tmp)
            bbox = font.getbbox(ch)
            td.text((4 - bbox[0], 4 - bbox[1]), ch, font=font,
                    fill=(0, 0, 0, 255))
            cb = tmp.getbbox()
            if cb:
                cropped = tmp.crop(cb)
                tmp.close()
                tmp = cropped
            char_imgs.append(tmp)
        except Exception:
            tmp.close()
            raise

    total_w = sum(im.size[0] for im in char_imgs)
    if n > 1 and total_w > usable_w:
        # 文字多到放不下时才横向压扁不换行；
        # 文字少时保持字形原宽，靠字间距两端对齐
        scale = usable_w / total_w
        squeezed = []
        for im in char_imgs:
            new_im = im.resize(
                (max(1, int(im.size[0] * scale)), im.size[1]),
                Image.LANCZOS)
            squeezed.append(new_im)
            im.close()
        char_imgs = squeezed

    img = None
    try:
        img = Image.new('RGB', (canvas_size, canvas_size), (255, 255, 255))

        for row_idx in range(2):
            oy = row_idx * row_h
            row_img = Image.new('RGBA', (canvas_size, row_h),
                                (255, 255, 255, 0))
            try:
                total_rendered = sum(im.size[0] for im in char_imgs)
                if n == 1:
                    x = (canvas_size - char_imgs[0].size[0]) // 2
                    gap = 0
                else:
                    gap = max(4, int((usable_w - total_rendered) / (n - 1)))
                    x = margin_x

                for im in char_imgs:
                    y = (row_h - im.size[1]) // 2 - center_shift
                    row_img.paste(im, (x, y), im)
                    x += im.size[0] + gap

                if row_idx == 0:
                    row_img = row_img.rotate(180, expand=False)

                img.paste(row_img, (0, oy), row_img)
            finally:
                row_img.close()

        if show_border:
            m = int(canvas_size * 0.02)
            _draw_dashed_rect(img, (m, m, canvas_size - m, canvas_size - m))
    except Exception:
        if img is not None:
            img.close()
        raise
    finally:
        for im in char_imgs:
            try:
                im.close()
            except Exception:
                pass

    return img


def parse_names(text):
    """把批量输入的文字拆成姓名列表。

    分隔符：换行、逗号、顿号、分号。空格不是分隔符——
    名字里的空格（含全角空格、Tab）一律删掉，如“张   三”→“张三”。
    """
    names = []
    for line in text.splitlines():
        for part in re.split(r'[,，、;；]+', line):
            part = re.sub(r'[\s\u3000]+', '', part)
            if part:
                names.append(part)
    return names


def validate_names(names):
    """检查姓名列表，返回 (可用列表, 提示文字)；超长/超量在这里拦住。"""
    if not names:
        return [], '请先输入姓名（每行一个）。'
    too_long = [n for n in names if len(n) > MAX_NAME_LEN]
    if too_long:
        sample = '、'.join(too_long[:3])
        return [], ('有 %d 个名字太长（超过 %d 个字）：%s…\n'
                    '请检查是否把整段文字当成了姓名。'
                    % (len(too_long), MAX_NAME_LEN, sample))
    if len(names) > MAX_NAMES:
        return [], ('一次最多处理 %d 个名字，当前 %d 个。\n'
                    '请分批操作，避免占用过多内存。'
                    % (MAX_NAMES, len(names)))
    return names, ''


def iter_cards(names, font_path, canvas_size=1024, font_size=280,
               show_border=False, on_progress=None, should_stop=None):
    """逐个生成座位牌图片，用完自动释放上一张，内存占用始终只有一张图。

    生成器用法：for img in iter_cards(...): 处理 img
    on_progress(序号, 总数, 姓名) 可用于显示进度。
    """
    total = len(names)
    for i, nm in enumerate(names, 1):
        if should_stop is not None and should_stop():
            return
        if on_progress is not None:
            try:
                on_progress(i, total, nm)
            except Exception:
                pass
        img = generate(nm, font_path, canvas_size=canvas_size,
                       font_size=font_size, show_border=show_border)
        if img is None:
            continue
        try:
            yield img
        finally:
            # 生成器继续迭代或提前关闭时都会释放，避免图片堆积
            img.close()


def build_docx_from_images(images, out_path):
    """把多张座位牌图片合成为一个 docx，每张单独一页、A4 页面居中。

    images 可以是列表，也可以是生成器：边写边释放，几百张也不会撑爆内存。
    返回实际写入的页数。
    """
    from io import BytesIO
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    for margin in ('left_margin', 'right_margin', 'top_margin', 'bottom_margin'):
        setattr(section, margin, Cm(0.6))

    # 页面内容垂直居中
    v_align = OxmlElement('w:vAlign')
    v_align.set(qn('w:val'), 'center')
    section._sectPr.append(v_align)

    card = Cm(CARD_CM)
    count = 0
    for img in images:
        if count:
            br = doc.add_paragraph()
            br.add_run().add_break(WD_BREAK.PAGE)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        buf = BytesIO()
        try:
            img.save(buf, format='PNG')
            buf.seek(0)
            p.add_run().add_picture(buf, width=card, height=card)
        finally:
            buf.close()
        count += 1
    if count == 0:
        raise ValueError('没有可写入的座位牌内容')
    doc.save(out_path)
    return count


def _show_print_dialog(pd):
    """调用系统标准打印对话框；用户确认后返回 True，取消返回 False。"""
    import ctypes
    from ctypes import wintypes
    comdlg32 = ctypes.windll.comdlg32
    comdlg32.PrintDlgW.argtypes = [ctypes.POINTER(type(pd))]
    comdlg32.PrintDlgW.restype = wintypes.BOOL
    return bool(comdlg32.PrintDlgW(ctypes.byref(pd)))


def _default_devmode(printer_name):
    """读取打印机当前默认的 DEVMODE 原始字节；失败返回 None。"""
    import ctypes
    from ctypes import wintypes
    winspool = ctypes.WinDLL('winspool.drv')
    winspool.OpenPrinterW.argtypes = [
        wintypes.LPCWSTR, ctypes.POINTER(wintypes.HANDLE), ctypes.c_void_p]
    winspool.OpenPrinterW.restype = wintypes.BOOL
    winspool.DocumentPropertiesW.argtypes = [
        wintypes.HWND, wintypes.HANDLE, wintypes.LPCWSTR,
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD]
    winspool.DocumentPropertiesW.restype = ctypes.c_long
    winspool.ClosePrinter.argtypes = [wintypes.HANDLE]
    winspool.ClosePrinter.restype = wintypes.BOOL

    hprinter = wintypes.HANDLE()
    if not winspool.OpenPrinterW(printer_name, ctypes.byref(hprinter), None):
        return None
    try:
        size = winspool.DocumentPropertiesW(None, hprinter, printer_name,
                                            None, None, 0)
        if size <= 0:
            return None
        buf = ctypes.create_string_buffer(size)
        ret = winspool.DocumentPropertiesW(
            None, hprinter, printer_name,
            ctypes.cast(buf, ctypes.c_void_p), None, 1)
        if ret < 0:
            return None
        return buf.raw[:size]
    finally:
        winspool.ClosePrinter(hprinter)


def _devmode_simplex(raw):
    """把 DEVMODE 改为单面打印（dmDuplex = DM_DUPLEX_SIMPLEX）。"""
    import struct
    if len(raw) < 96:
        return raw
    data = bytearray(raw)
    fields = struct.unpack_from('<I', data, 72)[0]
    fields |= 0x1000  # DM_DUPLEX
    struct.pack_into('<I', data, 72, fields)
    struct.pack_into('<H', data, 94, 1)  # DM_DUPLEX_SIMPLEX
    return bytes(data)


def print_images_direct(images, doc_title='座位牌', owner_hwnd=0,
                        on_progress=None):
    """不依赖 Word/WPS：弹 Windows 标准打印对话框，直接把图片打印到 A4 纸。

    用系统的 PrintDlgW 公共对话框（不依赖 MFC），在没有任何 Office 软件的
    电脑上也能一键打印；默认单面（可在打印窗口的“属性/首选项”里改双面），
    用户取消时返回 False。

    images 可以是列表，也可以是生成器：逐张绘制、逐张释放，
    几百页也不会因为内存不足而崩溃。
    """
    import ctypes
    from ctypes import wintypes
    import win32con
    import win32print
    import win32ui
    from PIL import ImageWin

    class PRINTDLGW(ctypes.Structure):
        _fields_ = [
            ('lStructSize', wintypes.DWORD),
            ('hwndOwner', wintypes.HWND),
            ('hDevMode', wintypes.HGLOBAL),
            ('hDevNames', wintypes.HGLOBAL),
            ('hDC', wintypes.HDC),
            ('Flags', wintypes.DWORD),
            ('nFromPage', wintypes.WORD),
            ('nToPage', wintypes.WORD),
            ('nMinPage', wintypes.WORD),
            ('nMaxPage', wintypes.WORD),
            ('nCopies', wintypes.WORD),
            ('hInstance', wintypes.HINSTANCE),
            ('lCustData', wintypes.LPARAM),
            ('lpfnPrintHook', ctypes.c_void_p),
            ('lpfnSetupHook', ctypes.c_void_p),
            ('lpPrintTemplateName', wintypes.LPCWSTR),
            ('lpSetupTemplateName', wintypes.LPCWSTR),
            ('hPrintTemplate', wintypes.HGLOBAL),
            ('hSetupTemplate', wintypes.HGLOBAL),
        ]

    PD_RETURNDC = 0x0100
    PD_NOSELECTION = 0x0004
    PD_NOPAGENUMS = 0x0008
    PD_USEDEVMODECOPIESANDCOLLATE = 0x00040000

    pd = PRINTDLGW()
    pd.lStructSize = ctypes.sizeof(PRINTDLGW)
    pd.hwndOwner = owner_hwnd
    pd.Flags = (PD_RETURNDC | PD_NOSELECTION | PD_NOPAGENUMS |
                PD_USEDEVMODECOPIESANDCOLLATE)
    pd.nFromPage = 1
    pd.nToPage = 0xFFFF
    pd.nMinPage = 1
    pd.nMaxPage = 0xFFFF
    pd.nCopies = 1

    # 默认单面打印：取打印机默认 DEVMODE，把 dmDuplex 改成单面后传给对话框
    kernel32 = ctypes.windll.kernel32
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalFree.argtypes = [ctypes.c_void_p]
    kernel32.GlobalFree.restype = ctypes.c_void_p
    hdev = None
    try:
        printer_name = win32print.GetDefaultPrinter()
        raw = _default_devmode(printer_name)
        if raw:
            raw = _devmode_simplex(raw)
            hdev = kernel32.GlobalAlloc(0x0002, len(raw))
            if hdev:
                ptr = kernel32.GlobalLock(hdev)
                if ptr:
                    ctypes.memmove(ptr, raw, len(raw))
                    kernel32.GlobalUnlock(hdev)
                else:
                    kernel32.GlobalFree(hdev)
                    hdev = None
    except Exception:
        hdev = None
    pd.hDevMode = hdev

    if not _show_print_dialog(pd):
        if pd.hDevMode:
            kernel32.GlobalFree(pd.hDevMode)
        return False  # 用户取消（或对话框打开失败）
    hdc = pd.hDC
    if pd.hDevMode:
        kernel32.GlobalFree(pd.hDevMode)  # 设备上下文已建好，释放 DEVMODE

    printer_dc = None
    try:
        printer_dc = win32ui.CreateDCFromHandle(hdc)
        # 打印区域对应 A4（21.0×29.7cm），卡面 19.9cm 等比缩放后居中
        off_x = printer_dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
        off_y = printer_dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)
        pw = printer_dc.GetDeviceCaps(win32con.HORZRES)
        ph = printer_dc.GetDeviceCaps(win32con.VERTRES)
        card = int(min(pw * CARD_CM / 21.0, ph * CARD_CM / 29.7))
        x = off_x + (pw - card) // 2
        y = off_y + (ph - card) // 2
        box = (x, y, x + card, y + card)

        printer_dc.StartDoc(doc_title)
        try:
            for i, img in enumerate(images, 1):
                if on_progress is not None:
                    try:
                        on_progress(i)
                    except Exception:
                        pass
                printer_dc.StartPage()
                ImageWin.Dib(img).draw(printer_dc.GetHandleOutput(), box)
                printer_dc.EndPage()
        finally:
            printer_dc.EndDoc()
    finally:
        if printer_dc is not None:
            try:
                printer_dc.DeleteDC()
            except Exception:
                pass
        else:
            ctypes.windll.gdi32.DeleteDC(hdc)
    return True


# 虚拟打印机（出不了纸，只生成文件）：检测默认设备时排除
VIRTUAL_PRINTER_HINTS = (
    'microsoft print to pdf',
    'microsoft xps document writer',
    'microsoft print to onenote',
    'onenote',
    '发送到 onenote',
    'fax',
    'wps pdf',
    '另存为',
)


def get_default_physical_printer():
    """返回系统默认打印机名。

    没有安装打印机，或默认设备是“打印成 PDF / XPS / OneNote / 传真”
    这类虚拟设备时返回 None——它们出不了纸，“直接打印”对它们无意义。
    pywin32 不可用时同样返回 None。
    """
    try:
        import win32print
    except Exception:
        return None
    try:
        name = win32print.GetDefaultPrinter()
    except Exception:
        return None
    if not name:
        return None
    low = name.lower()
    for hint in VIRTUAL_PRINTER_HINTS:
        if hint in low:
            return None
    return name


def print_image_silent(img, doc_title='打印图片'):
    """不弹任何对话框，把单张 PIL 图片等比居中直送默认实体打印机。

    供“二维码/海报一键打印”使用：已连打印机时点一下就出纸。
    返回 (ok, info)：
      ok=True  -> info 为实际打印机名（任务已提交到打印队列）；
      ok=False -> info 为可直接展示给用户的原因文案。
    """
    try:
        import win32con
        import win32print
        import win32ui
        from PIL import ImageWin
    except Exception:
        return (False,
                '打印组件（pywin32）不可用，请先用“保存二维码图片”'
                '保存后手动打印。')

    printer = get_default_physical_printer()
    if printer is None:
        try:
            default = win32print.GetDefaultPrinter()
        except Exception:
            default = ''
        if default:
            return (False,
                    '默认设备“%s”是虚拟打印机（打印成文件，出不了纸）。\n'
                    '请先在 Windows 设置里把实体打印机设为默认，再点打印。'
                    % default)
        return False, '没有检测到打印机。\n请先连接并安装打印机，或先用“保存二维码图片”。'

    # ImageWin.Dib 需要 RGB；调用方给的是海报（RGB），这里兜底转换
    if img.mode != 'RGB':
        img = img.convert('RGB')

    dc = None
    try:
        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(printer)
        off_x = dc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
        off_y = dc.GetDeviceCaps(win32con.PHYSICALOFFSETY)
        pw = dc.GetDeviceCaps(win32con.HORZRES)
        ph = dc.GetDeviceCaps(win32con.VERTRES)
        iw, ih = img.size
        scale = min(pw / float(iw), ph / float(ih))
        dw, dh = int(iw * scale), int(ih * scale)
        x = off_x + (pw - dw) // 2
        y = off_y + (ph - dh) // 2
        box = (x, y, x + dw, y + dh)
        dc.StartDoc(doc_title)
        try:
            dc.StartPage()
            ImageWin.Dib(img).draw(dc.GetHandleOutput(), box)
            dc.EndPage()
        finally:
            dc.EndDoc()
    except Exception as e:
        return False, '发送打印任务失败：%s' % e
    finally:
        if dc is not None:
            try:
                dc.DeleteDC()
            except Exception:
                pass
    return True, printer


def app_dir():
    """程序所在目录（打包成 exe 后为 exe 所在目录）。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def remote_config_dir():
    """远程打印配置文件目录：%APPDATA%\\远程打印。

    放在用户目录里与程序位置解耦，软件放在桌面时配置不会混在桌面，
    桌面整理或误删文件不影响已保存的服务器地址和令牌。
    """
    base = os.environ.get('APPDATA')
    if base:
        return os.path.join(base, '远程打印')
    return app_dir()


# 内置默认服务器与令牌：没保存过配置文件时直接用这两项，
# 第一次使用无需先点“远程打印设置”保存。
DEFAULT_REMOTE_SERVER = 'http://print.51zixue.top'
DEFAULT_REMOTE_TOKEN = 'dsdsssserwerwtrrtewewewererwre'


def remote_config_path():
    return os.path.join(remote_config_dir(), '云打印配置.json')


def legacy_remote_config_path():
    """旧版本放在程序目录里的配置文件（用于自动迁移）。"""
    return os.path.join(app_dir(), '云打印配置.json')


def default_remote_config():
    return {'server': DEFAULT_REMOTE_SERVER,
            'token': DEFAULT_REMOTE_TOKEN}


def load_remote_config():
    """返回实际使用的远程打印配置。

    没有配置文件时直接返回内置默认服务器和令牌，保证开箱即用；
    有保存文件时，文件里的值会覆盖默认值。程序目录里的旧配置
    会自动迁移到用户目录。
    """
    cfg = default_remote_config()
    path = remote_config_path()
    if not os.path.exists(path):
        legacy = legacy_remote_config_path()
        if os.path.exists(legacy):
            try:
                os.makedirs(remote_config_dir(), exist_ok=True)
                shutil.move(legacy, path)
            except Exception:
                path = legacy
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict):
            for key in ('server', 'token'):
                value = data.get(key)
                if isinstance(value, str) and value.strip():
                    cfg[key] = value.strip()
    except Exception:
        pass
    return cfg


def save_remote_config(cfg):
    path = remote_config_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def upload_print_job(server, token, zip_bytes, job_name, timeout=300):
    """把 zip 上传到云打印服务器，返回任务号。"""
    import urllib.error
    import urllib.parse
    import urllib.request
    url = server.rstrip('/') + '/api/jobs'
    req = urllib.request.Request(url, data=zip_bytes, method='POST')
    req.add_header('X-Token', token)
    req.add_header('X-Job-Name', urllib.parse.quote(job_name))
    req.add_header('Content-Type', 'application/zip')
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        # 把服务器返回的中文错误（如“上传空间已满…”）显示给用户
        try:
            body = json.loads(e.read().decode('utf-8'))
            msg = body.get('error') or body.get('message') or ''
        except Exception:
            msg = ''
        if msg:
            raise RuntimeError(msg) from None
        raise
    return data['id']


def upload_print_job_file(server, token, zip_path, job_name, timeout=600):
    """把磁盘上的 zip 文件流式上传（不整个读进内存），返回任务号。"""
    import urllib.error
    import urllib.parse
    import urllib.request
    url = server.rstrip('/') + '/api/jobs'
    size = os.path.getsize(zip_path)
    try:
        with open(zip_path, 'rb') as f:
            # 注意：data 必须在加请求头之前设好——urllib 里给 data 赋值
            # 会清空已设置的请求头，导致 Content-Length 丢失
            req = urllib.request.Request(url, data=f, method='POST')
            req.add_header('X-Token', token)
            req.add_header('X-Job-Name', urllib.parse.quote(job_name))
            req.add_header('Content-Type', 'application/zip')
            # 明确给出长度，urllib 才会分块读文件发送，不一次性读进内存
            req.add_header('Content-Length', str(size))
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        try:
            body = json.loads(e.read().decode('utf-8'))
            msg = body.get('error') or body.get('message') or ''
        except Exception:
            msg = ''
        if msg:
            raise RuntimeError(msg) from None
        raise
    return data['id']


def check_upload_size(paths):
    """上传前检查总体积；过大的文件直接拒绝，避免打包时内存耗尽。"""
    total = 0
    for p in paths:
        try:
            total += os.path.getsize(p)
        except OSError:
            continue
        if total > MAX_UPLOAD_BYTES:
            raise RuntimeError(
                '所选文件太大（超过 %d MB），请分批上传。'
                % (MAX_UPLOAD_BYTES // (1024 * 1024)))
    return total


def zip_files(paths, dest_path, on_progress=None):
    """把多个本地文件打包到磁盘上的 zip（不占用大块内存）。"""
    import zipfile
    # 低压缩级别：速度快、内存占用固定，图片/PDF 本来就压不动
    with zipfile.ZipFile(dest_path, 'w', zipfile.ZIP_DEFLATED,
                         compresslevel=1) as zf:
        items = sorted(paths)
        for i, p in enumerate(items, 1):
            if on_progress is not None:
                try:
                    on_progress(i, len(items), os.path.basename(p))
                except Exception:
                    pass
            ext = os.path.splitext(p)[1].lower()
            zf.write(p, '%03d%s' % (i, ext))
    return dest_path


class App:
    def __init__(self, root):
        self.root = root
        root.title("远程打印")
        root.geometry("540x700")
        root.resizable(False, False)
        ui_theme.style_window(root)
        try:
            icon_path = os.path.join(_BASE_DIR, '远程打印.ico')
            if os.path.exists(icon_path):
                root.iconbitmap(icon_path)
        except Exception:
            pass
        self._build()

    def _build(self):
        f = ttk.Frame(self.root, padding=(16, 12))
        f.pack(fill=tk.BOTH, expand=True)

        ff = ttk.Frame(f)
        ff.pack(fill=tk.X)
        ttk.Label(ff, text="字体：", font=("", 10)).pack(side=tk.LEFT)
        ensure_bundled_fonts()
        try:
            self._sys_fonts = list_system_fonts()
        except Exception:
            self._sys_fonts = ['黑体']
        default_font = self._sys_fonts[0] if self._sys_fonts else '黑体'
        for fnt in self._sys_fonts:
            if fnt == '华文行楷':
                default_font = fnt
                break
        else:
            for fnt in self._sys_fonts:
                if '黑' in fnt or 'Hei' in fnt or 'simhei' in fnt.lower():
                    default_font = fnt
                    break
        self.font_var = tk.StringVar(value=default_font)
        cb = ttk.Combobox(ff, textvariable=self.font_var,
                          values=self._sys_fonts, width=24)
        cb.pack(side=tk.LEFT, padx=(6, 0))
        cb.bind("<KeyRelease>", self._on_font_key)
        self.font_combo = cb
        self.font_var.trace_add("write", self._on_font_text)

        sf = ttk.Frame(f)
        sf.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(sf, text="字号：", font=("", 10)).pack(side=tk.LEFT)
        self.font_size_var = tk.IntVar(value=280)
        sb = ttk.Spinbox(sf, from_=40, to=500, width=6,
                         textvariable=self.font_size_var, font=("", 11))
        sb.pack(side=tk.LEFT, padx=(6, 4))
        ttk.Label(sf, text="px", font=("", 9)).pack(side=tk.LEFT)
        self.font_size_var.trace_add("write", lambda *a: self._preview())
        ttk.Label(sf, text="尺寸：", font=("", 10)).pack(side=tk.LEFT,
                                                          padx=(18, 0))
        self.size_var = tk.IntVar(value=1024)
        slider = ttk.Scale(sf, from_=400, to=2048, orient=tk.HORIZONTAL,
                           variable=self.size_var, length=150)
        slider.pack(side=tk.LEFT, padx=(6, 4))
        slider.configure(command=lambda v: self._set_size(int(float(v))))
        self.size_label = ttk.Label(sf, text="1024×1024", font=("", 9))
        self.size_label.pack(side=tk.LEFT)
        self.size_var.trace_add("write", lambda *a: self._preview())

        cf = ttk.Frame(f)
        cf.pack(fill=tk.X, pady=(8, 6))
        self.border_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(cf, text="虚线边框",
                        variable=self.border_var,
                        command=self._preview).pack(side=tk.LEFT)

        ttk.Label(f, text="姓名列表（每行一个，也支持逗号、顿号分隔；名字中的空格会自动删除）：",
                  font=("", 10)).pack(anchor=tk.W)
        tf = ttk.Frame(f)
        tf.pack(fill=tk.X, pady=(6, 10))
        self.names_text = tk.Text(
            tf, height=8, font=("", 12), wrap=tk.CHAR,
            relief=tk.FLAT, bd=0, padx=8, pady=6,
            bg=ui_theme.PALETTE['surface'], fg=ui_theme.PALETTE['text'],
            insertbackground=ui_theme.PALETTE['primary'],
            highlightthickness=1,
            highlightbackground=ui_theme.PALETTE['border'],
            highlightcolor=ui_theme.PALETTE['primary'])
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL,
                            command=self.names_text.yview)
        self.names_text.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.names_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.names_text.bind("<<Modified>>", self._on_text_modified)
        self.names_text.edit_modified(False)
        self.names_text.focus()

        self.preview = tk.Label(f, bg=ui_theme.PALETTE['surface_alt'],
                                relief=tk.FLAT, bd=0,
                                highlightthickness=1,
                                highlightbackground=ui_theme.PALETTE['border'],
                                anchor=tk.CENTER, height=9)
        self.preview.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        bf = ttk.Frame(f)
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="一键打印（本楼层）", style='Accent.TButton',
                   command=self._print).pack(side=tk.LEFT)
        ttk.Button(bf, text="保存座位牌到电脑",
                   style='Ghost.TButton',
                   command=self._batch).pack(side=tk.LEFT, padx=(8, 0))

        # 远程打印：独立区块，浅蓝色背景与本地操作区分
        soft = ui_theme.PALETTE['primary_soft']
        rf = tk.Frame(f, bg=soft,
                      highlightbackground=ui_theme.PALETTE['border'],
                      highlightthickness=1, bd=0,
                      padx=12, pady=6)
        rf.pack(fill=tk.X, pady=(6, 0))
        tk.Label(rf, text="远程打印", font=ui_theme.font(11, "bold"),
                 bg=soft, fg=ui_theme.PALETTE['primary']).pack(anchor=tk.W)
        rfb = tk.Frame(rf, bg=soft)
        rfb.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(rfb, text="远程打印设置", style='Ghost.TButton',
                   command=self._cloud_settings).pack(side=tk.LEFT)
        ttk.Button(rfb, text="座位牌远程打印", style='Accent.TButton',
                   command=self._cloud_print).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(rfb, text="文件打印（远程）", style='Ghost.TButton',
                   command=self._upload_file).pack(side=tk.LEFT, padx=(8, 0))
        self.cloud_status = tk.Label(
            rf, text=self._cloud_status_text(), fg=ui_theme.PALETTE['muted'],
            bg=soft, font=ui_theme.font(9))
        self.cloud_status.pack(anchor=tk.W, pady=(6, 0))

        self.preview.config(text="预览：显示姓名列表中的第一个名字")

    def _set_size(self, v):
        self.size_var.set(v)
        self.size_label.config(text=f"{v}×{v}")

    def _on_font_key(self, event=None):
        """输入时按文字过滤字体列表，并展开下拉框显示结果。"""
        text = self.font_var.get().strip()
        if text:
            matched = [f for f in self._sys_fonts
                       if text.lower() in f.lower()]
        else:
            matched = self._sys_fonts
        self.font_combo['values'] = matched
        if text and text in self._sys_fonts:
            return  # 已经是完整字体名，不自动展开
        if getattr(event, 'keysym', None) in ('Down', 'Up', 'Return',
                                              'Escape', 'Tab'):
            return  # 方向键/回车等用来选择，不重复展开
        self.font_combo.event_generate('<Down>')

    def _on_font_text(self, *a):
        text = self.font_var.get().strip()
        if text and text not in self._sys_fonts:
            return  # 正在输入搜索词，等选到完整字体名再刷新预览
        self._preview()

    def _on_text_modified(self, event=None):
        if self.names_text.edit_modified():
            self.names_text.edit_modified(False)
            self._preview()

    def _preview(self):
        names = parse_names(self.names_text.get("1.0", tk.END))
        if not names:
            self.preview.config(image="", text="预览：显示姓名列表中的第一个名字")
            self.preview.image = None
            return
        old = getattr(self.preview, 'image', None)
        img = None
        try:
            fp = self._get_font()
            if not fp:
                raise OSError("未找到字体")
            img = generate(names[0], fp,
                           canvas_size=self.size_var.get(),
                           font_size=self.font_size_var.get(),
                           show_border=self.border_var.get())
        except Exception:
            img = None
        if img is None:
            self.preview.config(image="", text="预览：显示姓名列表中的第一个名字")
            self.preview.image = None
            return
        try:
            # 直接在这张图上做缩略图（它只是预览用，不需要保留原图）
            img.thumbnail((170, 170))
            photo = ImageTk.PhotoImage(img)
        finally:
            img.close()
        self.preview.config(image=photo, text="")
        self.preview.image = photo
        del old  # 旧预览图交给垃圾回收，避免反复预览堆积内存

    def _get_font(self):
        p = get_font_path(self.font_var.get())
        return p if p else get_font_path("黑体")

    def _batch(self):
        names = parse_names(self.names_text.get("1.0", tk.END))
        if not names:
            messagebox.showwarning("提示", "请先输入姓名（每行一个）")
            return
        names, problem = validate_names(names)
        if problem:
            messagebox.showwarning("提示", problem)
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialfile=f"座位牌批量_{len(names)}人.docx")
        if not path:
            return
        self.root.configure(cursor='watch')
        self.root.update_idletasks()
        try:
            fp = self._get_font()
            size = self.size_var.get()
            fs = self.font_size_var.get()
            border = self.border_var.get()
            # 逐张生成、逐张写入，内存里始终只有一张图
            count = build_docx_from_images(
                iter_cards(names, fp, canvas_size=size, font_size=fs,
                           show_border=border),
                path)
        except Exception as e:
            messagebox.showerror("错误", f"生成失败：\n{e}")
            return
        finally:
            self.root.configure(cursor='')
        messagebox.showinfo("完成", f"已生成 {count} 页座位牌：\n{path}")


    def _print(self):
        """一键打印：弹系统打印窗口，默认单面，把图片直接打印到 A4 纸。"""
        names = parse_names(self.names_text.get("1.0", tk.END))
        if not names:
            messagebox.showwarning("提示", "请先输入姓名（每行一个）")
            return
        names, problem = validate_names(names)
        if problem:
            messagebox.showwarning("提示", problem)
            return
        self.root.configure(cursor='watch')
        self.root.update_idletasks()
        progress = {'count': 0}

        def _progress(index):
            progress['count'] = index
            self._print_progress(index)

        try:
            fp = self._get_font()
            size = self.size_var.get()
            fs = self.font_size_var.get()
            border = self.border_var.get()
            cards = iter_cards(names, fp, canvas_size=size, font_size=fs,
                               show_border=border)
            try:
                hwnd = int(self.root.winfo_id())
            except Exception:
                hwnd = 0
            # 生成器逐张出图、逐张打印，避免几百张图同时驻留内存
            printed = print_images_direct(cards, owner_hwnd=hwnd,
                                          on_progress=_progress)
        except Exception as e:
            messagebox.showerror("错误", f"打印失败：\n{e}")
            return
        finally:
            cards = None
            self.root.configure(cursor='')
            self.preview.config(text="预览：显示姓名列表中的第一个名字",
                                image="")
            self.preview.image = None
        if printed and not progress['count']:
            messagebox.showwarning("提示", "没有可打印的内容")
        elif printed:
            messagebox.showinfo(
                "完成", "打印任务已提交，共 %d 页。" % progress['count'])

    def _print_progress(self, index):
        """打印过程中刷新一下界面，让操作者知道程序没有卡死。"""
        try:
            self.preview.config(text="正在打印第 %d 张…" % index, image="")
            self.root.update_idletasks()
        except tk.TclError:
            pass


    def _cloud_status_text(self):
        cfg = load_remote_config()
        if os.path.exists(remote_config_path()):
            return '远程打印：已配置 %s' % cfg['server']
        return '远程打印：默认服务器 %s（换服务器才需点“远程打印设置”）' % cfg['server']

    def _cloud_settings(self):
        """远程打印设置对话框：填写服务器地址和令牌。"""
        cfg = load_remote_config()
        win = tk.Toplevel(self.root)
        win.title("远程打印设置")
        win.resizable(False, False)
        f = ttk.Frame(win, padding=14)
        f.pack()

        ttk.Label(f, text="服务器地址：").grid(
            row=0, column=0, sticky=tk.W)
        server_var = tk.StringVar(
            value=cfg['server'])
        ttk.Entry(f, textvariable=server_var, width=38).grid(
            row=0, column=1, pady=4)
        ttk.Label(f, text="打印令牌：").grid(
            row=1, column=0, sticky=tk.W)
        token_var = tk.StringVar(
            value=cfg['token'])
        ttk.Entry(f, textvariable=token_var, width=38).grid(
            row=1, column=1, pady=4)
        ttk.Label(f, text="不修改直接关闭即可，默认服务器仍会生效；"
                          "只有更换服务器时才需要改完点“保存”。",
                  foreground="gray", font=("", 9)).grid(
            row=2, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        def save():
            cfg['server'] = server_var.get().strip().rstrip('/')
            cfg['token'] = token_var.get().strip()
            if not cfg['server'] or not cfg['token']:
                messagebox.showwarning("提示", "服务器地址和令牌不能为空")
                return
            try:
                save_remote_config(cfg)
            except Exception as e:
                messagebox.showerror("错误", f"保存设置失败：\n{e}",
                                     parent=win)
                return
            self.cloud_status.config(text=self._cloud_status_text())
            win.destroy()
            messagebox.showinfo(
                "已保存",
                "远程打印设置已保存\n配置文件：%s" % remote_config_path())

        btns = ttk.Frame(f)
        btns.grid(row=3, column=0, columnspan=2, sticky=tk.E, pady=(12, 0))
        ttk.Button(btns, text="保存", command=save).pack(side=tk.LEFT)
        ttk.Button(btns, text="取消", command=win.destroy).pack(
            side=tk.LEFT, padx=(8, 0))
        win.transient(self.root)
        win.grab_set()
        self.root.wait_window(win)

    def _cloud_print(self):
        """远程打印：把座位牌图片打包上传到云服务器，由打印电脑自动打印。"""
        cfg = load_remote_config()
        names = parse_names(self.names_text.get("1.0", tk.END))
        if not names:
            messagebox.showwarning("提示", "请先输入姓名（每行一个）")
            return
        names, problem = validate_names(names)
        if problem:
            messagebox.showwarning("提示", problem)
            return
        self.root.configure(cursor='watch')
        self.root.update_idletasks()
        tmp_zip = None
        try:
            fp = self._get_font()
            size = self.size_var.get()
            fs = self.font_size_var.get()
            border = self.border_var.get()
            import io
            import tempfile
            import zipfile
            # 压缩包写到临时文件，图片逐张生成逐张释放，内存不随页数增长
            fd, tmp_zip = tempfile.mkstemp(prefix='座位牌_', suffix='.zip')
            os.close(fd)
            cards = iter_cards(names, fp, canvas_size=size, font_size=fs,
                               show_border=border)
            count = 0
            with zipfile.ZipFile(tmp_zip, 'w', zipfile.ZIP_DEFLATED,
                                 compresslevel=1) as zf:
                for img in cards:
                    b = io.BytesIO()
                    try:
                        img.save(b, format='PNG')
                        zf.writestr('%03d.png' % (count + 1), b.getvalue())
                    finally:
                        b.close()
                    count += 1
            if count == 0:
                messagebox.showwarning("提示", "没有可打印的内容")
                return

            # 直接流式上传磁盘上的压缩包，不再整包读进内存
            job_id = upload_print_job_file(
                cfg['server'], cfg['token'], tmp_zip,
                '座位牌_%d页_%s' % (count,
                                   time.strftime('%H%M%S')))
            messagebox.showinfo(
                "已提交",
                "已提交到远程打印队列：\n任务号：%s\n共 %d 页\n"
                "打印电脑会自动完成打印。" % (job_id, count))
        except Exception as e:
            messagebox.showerror("错误", f"远程打印提交失败：\n{e}")
        finally:
            self.root.configure(cursor='')
            if tmp_zip:
                try:
                    os.remove(tmp_zip)
                except OSError:
                    pass


    def _upload_file(self):
        """上传文件：选择本地 PDF、Word 或图片，打包上传到远程打印队列。"""
        cfg = load_remote_config()
        paths = filedialog.askopenfilenames(
            title="选择要上传的文件（PDF、Word 或图片，可多选）",
            filetypes=[
                ("支持的文件", "*.pdf *.doc *.docx "
                               "*.png *.jpg *.jpeg *.bmp *.gif *.webp"),
                ("PDF 文档", "*.pdf"),
                ("Word 文档", "*.doc *.docx"),
                ("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"),
                ("所有文件", "*.*"),
            ])
        if not paths:
            return
        allowed = {'.pdf', '.doc', '.docx',
                   '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp'}
        bad = [p for p in paths
               if os.path.splitext(p)[1].lower() not in allowed]
        if bad:
            messagebox.showwarning(
                "提示",
                "只允许上传 PDF、Word 和图片文件，请去掉：\n" +
                "\n".join(os.path.basename(b) for b in bad))
            return
        try:
            check_upload_size(paths)
        except Exception as e:
            messagebox.showwarning("提示", str(e))
            return
        self.root.configure(cursor='watch')
        self.root.update_idletasks()
        tmp_zip = None
        try:
            import tempfile
            fd, tmp_zip = tempfile.mkstemp(prefix='上传文件_', suffix='.zip')
            os.close(fd)
            zip_files(paths, tmp_zip)
            if os.path.getsize(tmp_zip) > MAX_UPLOAD_BYTES:
                messagebox.showwarning(
                    "提示", "打包后仍然超过 %d MB，请分批上传。"
                    % (MAX_UPLOAD_BYTES // (1024 * 1024)))
                return
            job_name = '上传文件_%d个_%s' % (
                len(paths), time.strftime('%H%M%S'))
            job_id = upload_print_job_file(
                cfg['server'], cfg['token'], tmp_zip, job_name)
            messagebox.showinfo(
                "已提交",
                "已提交到远程打印队列：\n文件：共 %d 个\n任务号：%s\n"
                "打印电脑会自动完成打印。" % (len(paths), job_id))
        except Exception as e:
            messagebox.showerror("错误", f"上传失败：\n{e}")
        finally:
            self.root.configure(cursor='')
            if tmp_zip:
                try:
                    os.remove(tmp_zip)
                except OSError:
                    pass


if __name__ == "__main__":
    def _run():
        root = tk.Tk()
        guard = safe_guard.install(app_name='远程打印', gui=True, root=root)
        App(root)
        root.mainloop()
        del guard

    safe_guard.main_guard(app_name='远程打印', gui=True, func=_run)
