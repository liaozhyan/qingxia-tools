# -*- coding: utf-8 -*-
"""PPT 转换工具（基于 WPS 演示组件）。

两个功能：
1. convert_ppt_to_images：把 PPT 每页导出为一张 PNG 图片；
2. convert_ppt_to_16x9：把非常规比例的 PPT 先导出为图片，
   再重新生成 16:9 的 PPT，图片拉伸铺满整页，播放即可全屏。
"""

import os
import shutil
import sys
import tempfile
import time
import json
import uuid


# WPS 演示 / Microsoft PowerPoint 的 COM ProgID
WPS_PROGIDS = ['KWPP.Application', 'PowerPoint.Application']

# 16:9 幻灯片尺寸（单位：磅）
SLIDE_W = 960.0
SLIDE_H = 540.0

# 导出图片宽度；高度按源 PPT 页面比例自动计算
EXPORT_WIDTH = 1920

SUPPORTED_EXTS = {'.ppt', '.pptx'}
_TEMPLATE_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'templet')
WELCOME_BACKGROUND_DIR = os.path.join(_TEMPLATE_DIR, 'backgrounds')
TEMPLATE_LIBRARY_DIR = os.path.join(_TEMPLATE_DIR, 'library')
TEMPLATE_MANIFEST = os.path.join(TEMPLATE_LIBRARY_DIR, 'library.json')
BACKGROUND_IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.bmp', '.webp'}

# 图片解码上限：约 8000 万像素。超过就拒绝，避免解码时内存爆掉。
MAX_IMAGE_PIXELS = 80 * 1000 * 1000
MAX_BACKGROUND_BYTES = 50 * 1024 * 1024


def _pil_image_module():
    """导入 PIL.Image 并设置解码像素上限（超过上限直接报错而不是吃爆内存）。"""
    from PIL import Image
    try:
        # 设成上限的一半：PIL 在超过 2 倍该值时才硬报错，
        # 这样刚好让超过 MAX_IMAGE_PIXELS 的图片直接报错而不是解码占内存
        Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS // 2
    except Exception:
        pass
    return Image


def check_image_file(path, max_bytes=MAX_BACKGROUND_BYTES):
    """上传前检查图片体积和分辨率，不合格时抛 PptConvertError。"""
    try:
        size = os.path.getsize(path)
    except OSError as e:
        raise PptConvertError('读取图片失败：%s' % e) from None
    if size > max_bytes:
        raise PptConvertError('图片文件太大（%.1f MB），请压缩到 %d MB 以内。'
                              % (size / 1048576.0, max_bytes // 1048576))
    Image = _pil_image_module()
    try:
        with Image.open(path) as im:
            w, h = im.size
    except Exception as e:
        if type(e).__name__ == 'DecompressionBombError':
            raise PptConvertError(
                '图片分辨率过高，请先缩小到 %d 万像素以内再上传。'
                % (MAX_IMAGE_PIXELS // 10000)) from None
        raise PptConvertError('无法识别这张图片：%s' % e) from None
    if w <= 0 or h <= 0:
        raise PptConvertError('图片尺寸无效。')
    if w * h > MAX_IMAGE_PIXELS:
        raise PptConvertError(
            '图片分辨率过高（%.0f 万像素），请先缩小到 %d 万像素以内。'
            % (w * h / 10000.0, MAX_IMAGE_PIXELS // 10000))
    return w, h


class PptConvertError(RuntimeError):
    """转换失败（含本机没有 WPS 的情况）。"""


class WpsSession:
    """上下文管理器：负责启动/退出 WPS 演示组件。"""

    def __enter__(self):
        try:
            import win32com.client
        except Exception as e:
            raise PptConvertError(
                '缺少 pywin32 组件：%s' % e) from None
        self.app = None
        last_err = None
        for progid in WPS_PROGIDS:
            try:
                self.app = win32com.client.DispatchEx(progid)
                break
            except Exception as e:
                last_err = e
        if self.app is None:
            raise PptConvertError(
                '无法启动 WPS 演示或 Microsoft PowerPoint，'
                '请先安装 WPS Office 或 Microsoft Office。\n%s'
                % (last_err or '')) from None
        try:
            self.app.Visible = False
        except Exception:
            pass
        try:
            self.app.DisplayAlerts = 0
        except Exception:
            pass
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            self.app.Quit()
        except Exception:
            pass
        return False


def _ensure_input(path):
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise PptConvertError('找不到文件：\n%s' % path)
    ext = os.path.splitext(path)[1].lower()
    if ext not in SUPPORTED_EXTS:
        raise PptConvertError('仅支持 .ppt / .pptx 文件。')
    return path


def _page_size_pts(pres):
    ps = pres.PageSetup
    w = float(ps.SlideWidth)
    h = float(ps.SlideHeight)
    if w <= 0 or h <= 0:
        raise PptConvertError('无法读取 PPT 页面尺寸。')
    return w, h


def _image_size(w, h, target_width=EXPORT_WIDTH):
    height = int(round(target_width * h / w))
    return target_width, height


def _report(on_progress, current, total, msg=None):
    if on_progress:
        on_progress(current, total, msg or '')


def open_presentation(app, path):
    """只读打开 PPT，避免改到原文件。"""
    try:
        return app.Presentations.Open(path, True, False, False)
    except Exception as e:
        raise PptConvertError('打开 PPT 失败：\n%s\n%s' % (path, e)) from None


def _text_width(line):
    """估算一行文字的显示宽度（以汉字宽度为单位）。"""
    width = 0.0
    for ch in line:
        # 中文/全角字符按一个汉字宽，半角字符约半个
        width += 1.0 if ord(ch) > 255 else 0.55
    return max(width, 1.0)


def _text_frame_metrics(shape):
    tf = shape.TextFrame
    try:
        ml = float(tf.MarginLeft or 0)
        mr = float(tf.MarginRight or 0)
        mt = float(tf.MarginTop or 0)
        mb = float(tf.MarginBottom or 0)
    except Exception:
        ml = mr = mt = mb = 7.2
    usable_w = max(float(shape.Width) - ml - mr, 50.0)
    usable_h = max(float(shape.Height) - mt - mb, 20.0)
    return usable_w, usable_h


def auto_font_size(shape, lines, base_size=24,
                   min_size=8, max_size=120):
    """根据文本框尺寸估算字号（磅）。

    以模板原字号为上限：内容能放下就保持原字号，
    内容过多时自动缩小到能放下的字号。
    """
    usable_w, usable_h = _text_frame_metrics(shape)
    if not lines:
        lines = ['']
    longest = max(_text_width(line) for line in lines)
    # 中文字号约等于磅值，行高按 1.3 倍估算
    size_by_width = usable_w / longest
    size_by_height = usable_h / (len(lines) * 1.3)
    fit = min(max_size, size_by_width, size_by_height)
    return max(min_size, min(int(fit), int(base_size)))


def list_welcome_backgrounds():
    """读取背景图库，返回 [名称, 原图, 缩略图] 列表。"""
    manifest = os.path.join(WELCOME_BACKGROUND_DIR, 'library.json')
    if not os.path.isfile(manifest):
        return []
    try:
        with open(manifest, 'r', encoding='utf-8') as f:
            items = json.load(f)
    except Exception:
        return []
    result = []
    for item in items:
        fname = item.get('file', '')
        full = os.path.join(WELCOME_BACKGROUND_DIR, fname)
        thumb = os.path.join(WELCOME_BACKGROUND_DIR, 'thumbs', fname)
        if os.path.isfile(full):
            result.append({
                'name': item.get('name', os.path.splitext(fname)[0]),
                'image': full,
                'thumb': thumb if os.path.isfile(thumb) else full,
            })
    return result


def presentation_app_available():
    """检测本机是否有 WPS 演示或 Microsoft PowerPoint 组件。"""
    try:
        import winreg
    except ImportError:
        return False
    for progid in WPS_PROGIDS:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
                return True
        except OSError:
            continue
    return False


def _read_template_manifest():
    if not os.path.isfile(TEMPLATE_MANIFEST):
        return []
    try:
        with open(TEMPLATE_MANIFEST, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_template_manifest(items):
    os.makedirs(TEMPLATE_LIBRARY_DIR, exist_ok=True)
    with open(TEMPLATE_MANIFEST, 'w', encoding='utf-8') as f:
        json.dump(items, f, ensure_ascii=False, indent=2)


def list_welcome_templates():
    """返回模板库清单（含模板文件路径与可替换文本框信息）。"""
    result = []
    for item in _read_template_manifest():
        path = os.path.join(TEMPLATE_LIBRARY_DIR, item.get('file', ''))
        if not os.path.isfile(path):
            continue
        result.append(dict(item, path=path))
    return result


def template_text_shapes(template_path):
    """读取模板 PPT 里所有文本框，供导入时选择替换目标。"""
    template_path = os.path.abspath(template_path)
    if not os.path.isfile(template_path):
        raise PptConvertError('找不到模板文件：\n%s' % template_path)
    shapes = []
    with WpsSession() as session:
        pres = session.app.Presentations.Open(
            template_path, True, False, False)
        try:
            for si in range(1, pres.Slides.Count + 1):
                slide = pres.Slides(si)
                for idx in range(1, slide.Shapes.Count + 1):
                    try:
                        shp = slide.Shapes(idx)
                        if not shp.HasTextFrame:
                            continue
                        tf = shp.TextFrame
                        if not tf.HasText:
                            text = ''
                        else:
                            text = tf.TextRange.Text.replace(
                                '\r', ' ').replace('\n', ' ').strip()
                        shapes.append({
                            'slide': si,
                            'index': idx,
                            'name': getattr(shp, 'Name', ''),
                            'text': text[:60],
                        })
                    except Exception:
                        continue
        finally:
            try:
                pres.Close()
            except Exception:
                pass
    return shapes


def add_welcome_template(src, display_name, slide_index, shape_index):
    """把用户提供的 PPT 加入模板库，返回模板信息。"""
    src = os.path.abspath(src)
    ext = os.path.splitext(src)[1].lower()
    if ext not in SUPPORTED_EXTS:
        raise PptConvertError('模板只支持 .ppt / .pptx 文件。')
    if not display_name or not display_name.strip():
        display_name = os.path.splitext(os.path.basename(src))[0]
    os.makedirs(TEMPLATE_LIBRARY_DIR, exist_ok=True)
    template_id = uuid.uuid4().hex[:12]
    dst_name = template_id + ext
    dst = os.path.join(TEMPLATE_LIBRARY_DIR, dst_name)
    try:
        shutil.copy2(src, dst)
    except Exception as e:
        raise PptConvertError('保存模板失败：%s' % e) from None

    items = _read_template_manifest()
    item = {
        'id': template_id,
        'name': display_name.strip(),
        'file': dst_name,
        'slide': int(slide_index),
        'shape_index': int(shape_index),
        'created': time.strftime('%Y-%m-%d %H:%M:%S'),
    }
    items.append(item)
    _save_template_manifest(items)
    return dict(item, path=dst)


def delete_welcome_template(template_id):
    """从模板库删除指定模板，成功返回 True。"""
    items = _read_template_manifest()
    remain = [x for x in items if x.get('id') != template_id]
    if len(remain) == len(items):
        return False
    for item in items:
        if item.get('id') == template_id:
            try:
                os.remove(os.path.join(
                    TEMPLATE_LIBRARY_DIR, item.get('file', '')))
            except OSError:
                pass
            break
    _save_template_manifest(remain)
    return True


def _resolve_background_file(item):
    """条目对应的图片文件：优先 file 字段，再退回 image / 按 id 找。"""
    folder = WELCOME_BACKGROUND_DIR
    fname = str(item.get('file') or '')
    if fname:
        full = os.path.join(folder, fname)
        if os.path.isfile(full):
            return fname, full
    image = str(item.get('image') or '')
    if image:
        full = image if os.path.isabs(image) else os.path.join(folder, image)
        if os.path.isfile(full):
            return os.path.basename(full), full
    bg_id = str(item.get('id') or '')
    if bg_id:
        for name in os.listdir(folder) if os.path.isdir(folder) else []:
            if name.startswith(bg_id + '.') and not name.endswith('.json'):
                full = os.path.join(folder, name)
                if os.path.isfile(full):
                    return name, full
    return '', ''


def list_uploaded_backgrounds():
    """读取用户上传的背景图库。"""
    manifest = os.path.join(WELCOME_BACKGROUND_DIR, 'backgrounds.json')
    if not os.path.isfile(manifest):
        return []
    try:
        with open(manifest, 'r', encoding='utf-8') as f:
            items = json.load(f)
    except Exception:
        return []
    result = []
    for item in items if isinstance(items, list) else []:
        # 清单里 file 字段丢了也能从目录里找回来（不让整个图库变空）
        fname, full = _resolve_background_file(item)
        if full:
            result.append({
                'id': item.get('id', ''),
                'name': item.get('name', os.path.splitext(fname)[0]),
                'image': full,
                'created': item.get('created', ''),
                'file': fname,
            })
    return result


def add_uploaded_background(src, display_name=None):
    """把图片复制进背景库，返回条目信息。"""
    src = os.path.abspath(src)
    ext = os.path.splitext(src)[1].lower()
    if ext not in BACKGROUND_IMAGE_EXTS:
        raise PptConvertError('背景图片支持 PNG/JPG/BMP/WebP。')
    # 上传时就拦住超大图：坏图一旦进库，之后每次预览都会拖慢甚至崩程序
    check_image_file(src)
    os.makedirs(WELCOME_BACKGROUND_DIR, exist_ok=True)
    if not display_name or not display_name.strip():
        display_name = os.path.splitext(os.path.basename(src))[0]
    bg_id = uuid.uuid4().hex[:12]
    dst_name = bg_id + ext
    dst = os.path.join(WELCOME_BACKGROUND_DIR, dst_name)
    try:
        shutil.copy2(src, dst)
    except Exception as e:
        raise PptConvertError('保存背景图片失败：%s' % e) from None

    manifest = os.path.join(WELCOME_BACKGROUND_DIR, 'backgrounds.json')
    items = []
    if os.path.isfile(manifest):
        try:
            with open(manifest, 'r', encoding='utf-8') as f:
                items = json.load(f)
        except Exception:
            items = []
    item = {
        'id': bg_id,
        'name': display_name.strip(),
        'file': dst_name,
        'created': time.strftime('%Y-%m-%d %H:%M:%S'),
    }
    items.append(item)
    with open(manifest, 'w', encoding='utf-8') as f:
        json.dump(items, f, ensure_ascii=False, indent=2)
    return dict(item, image=dst)


def delete_uploaded_background(bg_id):
    """从背景库删除图片，成功返回 True。

    删掉的图会先挪进 backgrounds/_deleted/ 里留一份（不是直接抹掉），
    万一删错了还能找回来。
    """
    manifest = os.path.join(WELCOME_BACKGROUND_DIR, 'backgrounds.json')
    # 注意：这里要按“清单原文”删，只去掉选中的那一条。
    # 不能用 list_uploaded_backgrounds() 的结果回写——那样一旦清单里
    # 有读不出来的条目，就会把整个图库一起写没（出现过这个事故）。
    try:
        with open(manifest, 'r', encoding='utf-8') as f:
            raw = json.load(f)
    except Exception:
        raw = []
    if not isinstance(raw, list):
        raw = []
    item = next((x for x in raw if str(x.get('id') or '') == str(bg_id)),
                None)
    if item is None:
        return False
    fname, full = _resolve_background_file(item)
    try:
        keep_dir = os.path.join(WELCOME_BACKGROUND_DIR, '_deleted')
        os.makedirs(keep_dir, exist_ok=True)
        stamp = time.strftime('%Y%m%d_%H%M%S')
        shutil.move(full, os.path.join(keep_dir, '%s_%s'
                                       % (stamp, os.path.basename(full))))
    except Exception:
        try:
            os.remove(full)
        except OSError:
            pass
    remain = [x for x in raw if str(x.get('id') or '') != str(bg_id)]
    with open(manifest, 'w', encoding='utf-8') as f:
        json.dump(remain, f, ensure_ascii=False, indent=2)
    return True


def rename_uploaded_background(bg_id, new_name):
    """改背景库里的名字（图片文件不动），成功返回 True。"""
    new_name = (new_name or '').strip()
    if not bg_id or not new_name:
        return False
    manifest = os.path.join(WELCOME_BACKGROUND_DIR, 'backgrounds.json')
    if not os.path.isfile(manifest):
        return False
    try:
        with open(manifest, 'r', encoding='utf-8') as f:
            items = json.load(f)
    except Exception:
        return False
    if not isinstance(items, list):
        return False
    hit = False
    for item in items:
        if item.get('id') == bg_id:
            item['name'] = new_name
            hit = True
    if not hit:
        return False
    try:
        with open(manifest, 'w', encoding='utf-8') as f:
            json.dump(items, f, ensure_ascii=False, indent=2)
    except Exception:
        return False
    return True


def _apply_background_to_slides(pres, bg_path):
    """把背景图片铺满每页并放到最底层。"""
    bg_path = os.path.abspath(bg_path)
    if not os.path.isfile(bg_path):
        raise PptConvertError('找不到背景图片：\n%s' % bg_path)
    try:
        w = float(pres.PageSetup.SlideWidth)
        h = float(pres.PageSetup.SlideHeight)
    except Exception:
        w = 720.0
        h = 405.0
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)
        pic = slide.Shapes.AddPicture(
            bg_path, False, True, 0, 0, w, h)
        try:
            pic.ZOrder(1)  # 1 = 置于底层
        except Exception:
            pass


def _apply_solid_color_to_slides(pres, color):
    """把每页背景设为纯色；color 形如 #RRGGBB。"""
    color = color.strip().lstrip('#')
    if len(color) != 6:
        raise PptConvertError('颜色格式不正确：%s' % color)
    try:
        r = int(color[0:2], 16)
        g = int(color[2:4], 16)
        b = int(color[4:6], 16)
    except ValueError:
        raise PptConvertError('颜色格式不正确：%s' % color) from None
    ole_color = r | (g << 8) | (b << 16)  # COM 使用 BGR
    for i in range(1, pres.Slides.Count + 1):
        try:
            bg = pres.Slides(i).Background
            bg.Fill.Solid()
            bg.Fill.ForeColor.RGB = ole_color
        except Exception as e:
            raise PptConvertError('设置纯色背景失败：%s' % e) from None


def generate_welcome_ppt(content, out_path=None, font_size=None,
                         on_progress=None, template_id=None,
                         background_path=None, background_color=None):
    """套用模板库里的模板生成欢迎语 PPT。

    只替换导入模板时指定的那个文本框，其余主体保持不动；
    content 中的换行会原样保留；font_size 为空时按内容自动调字号；
    background_path / background_color 为可选背景覆盖。
    """
    entry = next(
        (t for t in list_welcome_templates()
         if t.get('id') == template_id), None)
    if entry is None:
        raise PptConvertError(
            '模板库中找不到该模板，请重新导入。')
    template_path = entry['path']
    content = (content or '').replace('\r\n', '\n').replace('\r', '\n')
    if not content.strip():
        raise PptConvertError('请输入欢迎语内容。')

    if out_path:
        out_path = os.path.abspath(out_path)
    else:
        out_path = os.path.join(
            os.path.expanduser('~'),
            '欢迎语_%s.pptx' % time.strftime('%Y%m%d_%H%M%S'))
    ext = os.path.splitext(out_path)[1].lower()
    save_format = 24 if ext == '.pptx' else (
        1 if ext == '.ppt' else None)
    if save_format is None:
        raise PptConvertError('输出文件只能用 .ppt 或 .pptx 后缀。')
    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except Exception as e:
            raise PptConvertError(
                '无法覆盖输出文件：\n%s\n%s' % (out_path, e)) from None
    try:
        shutil.copy2(template_path, out_path)
    except Exception as e:
        raise PptConvertError('复制模板失败：%s' % e) from None

    used_size = font_size
    try:
        with WpsSession() as session:
            pres = session.app.Presentations.Open(
                out_path, False, False, False)
            try:
                if background_path:
                    _report(on_progress, 1, 2, '正在设置背景…')
                    _apply_background_to_slides(pres, background_path)
                elif background_color:
                    _report(on_progress, 1, 2, '正在设置背景颜色…')
                    _apply_solid_color_to_slides(pres, background_color)
                slide_no = int(entry.get('slide', 1))
                shape_no = int(entry.get('shape_index', 1))
                if slide_no > pres.Slides.Count:
                    raise PptConvertError('模板页数已变化，请重新导入模板。')
                slide = pres.Slides(slide_no)
                if shape_no > slide.Shapes.Count:
                    raise PptConvertError(
                        '模板文本框已变化，请重新导入模板。')
                shape = slide.Shapes(shape_no)
                lines = content.split('\n')
                tr = shape.TextFrame.TextRange
                try:
                    color = tr.Font.Color.RGB
                except Exception:
                    color = None
                try:
                    font_name = tr.Font.Name or '黑体'
                except Exception:
                    font_name = '黑体'
                try:
                    bold = tr.Font.Bold
                except Exception:
                    bold = None
                try:
                    alignment = tr.ParagraphFormat.Alignment
                except Exception:
                    alignment = None
                if font_size is None:
                    try:
                        base_size = float(
                            tr.Font.Size)
                    except Exception:
                        base_size = 24
                    used_size = auto_font_size(shape, lines, base_size)
                    _report(on_progress, 1, 2,
                            '自动字号：%d 磅' % used_size)
                tr.Text = content
                tr.Font.Name = font_name
                if bold is not None:
                    tr.Font.Bold = bold
                tr.Font.Size = float(used_size)
                if color is not None:
                    try:
                        tr.Font.Color.RGB = color
                    except Exception:
                        pass
                if alignment is not None:
                    try:
                        tr.ParagraphFormat.Alignment = alignment
                    except Exception:
                        pass
                shape.TextFrame.WordWrap = True
                try:
                    shape.TextFrame.AutoSize = 2  # 文本自动适应文本框
                except Exception:
                    pass
                _report(on_progress, 2, 2, '正在保存文件…')
                pres.SaveAs(out_path, save_format)
            finally:
                try:
                    pres.Close()
                except Exception:
                    pass
    except PptConvertError:
        raise
    except Exception as e:
        raise PptConvertError('生成欢迎语 PPT 失败：%s' % e) from None
    return out_path, used_size


def _color_rgb(hex_color):
    hex_color = (hex_color or '').strip().lstrip('#')
    if len(hex_color) != 6:
        return 0, 0, 0
    try:
        return (int(hex_color[0:2], 16),
                int(hex_color[2:4], 16),
                int(hex_color[4:6], 16))
    except ValueError:
        return 0, 0, 0


def _text_color_for(background_image=None, background_color=None):
    """按背景深浅自动返回 0xFFFFFF 白或 0x000000 黑。"""
    try:
        if background_image and os.path.isfile(background_image):
            from PIL import Image, ImageStat
            with Image.open(background_image) as raw:
                # draft：大 JPEG 在解码阶段就缩小，省时间也省内存
                try:
                    raw.draft('L', (64, 64))
                except Exception:
                    pass
                img = raw.convert('L')
            try:
                img.thumbnail((64, 64))
                mean = ImageStat.Stat(img).mean[0]
            finally:
                img.close()
        else:
            r, g, b = _color_rgb(background_color)
            mean = 0.299 * r + 0.587 * g + 0.114 * b
    except Exception:
        mean = 0
    return 0x000000 if mean > 165 else 0xFFFFFF


# 背景图预览缓存：同一张图反复预览（每次打字都会刷新）不再重复解码
_BG_CACHE_LIMIT = 6
_bg_cache = {}


def _cover_image(path, width, height):
    """把背景图按“铺满并居中裁剪”的规则缩放成 width×height。

    先在原图上裁出目标比例，再缩放：即使遇到超长/超宽图片，
    中间结果也不会膨胀成几亿像素。
    """
    key = None
    try:
        st = os.stat(path)
        key = (os.path.abspath(path), int(st.st_mtime), st.st_size,
               width, height)
        cached = _bg_cache.get(key)
        if cached is not None:
            return cached
    except OSError:
        key = None

    Image = _pil_image_module()
    with Image.open(path) as raw:
        try:
            raw.draft('RGB', (max(width, 1) * 2, max(height, 1) * 2))
        except Exception:
            pass
        src = raw.convert('RGB')
    try:
        target_ratio = float(width) / float(height)
        src_ratio = float(src.width) / float(src.height or 1)
        if src_ratio > target_ratio:
            crop_w = max(1, int(round(src.height * target_ratio)))
            left = max(0, (src.width - crop_w) // 2)
            box = (left, 0, left + crop_w, src.height)
        else:
            crop_h = max(1, int(round(src.width / target_ratio)))
            top = max(0, (src.height - crop_h) // 2)
            box = (0, top, src.width, top + crop_h)
        img = src.crop(box).resize((width, height), Image.LANCZOS)
    finally:
        src.close()

    if key is not None:
        if len(_bg_cache) >= _BG_CACHE_LIMIT:
            _bg_cache.clear()
        _bg_cache[key] = img
    return img


def _welcome_font_size_pts(content, font_size=None):
    """按版心与内容行数计算字号（磅），字体不会超过 48。"""
    if font_size:
        return int(font_size)
    lines = (content or '').split('\n')
    usable_w = SLIDE_W - 80 - 48
    usable_h = SLIDE_H - 40 - 24
    longest = max(_text_width(line) for line in lines)
    fit_w = usable_w / longest
    fit_h = usable_h / (len(lines) * 1.25)
    return max(8, min(48, int(fit_w), int(fit_h)))


_FONT_CACHE = {}
_FONT_CACHE_LIMIT = 6


def _load_cn_font(pixel_size):
    cached = _FONT_CACHE.get(pixel_size)
    if cached is not None:
        return cached
    candidates = [
        os.path.join(os.environ.get('WINDIR', 'C:\\Windows'),
                     'Fonts', name)
        for name in ('msyh.ttc', 'msyhbd.ttc', 'simhei.ttf',
                     'simsun.ttc')
    ]
    candidates.append(os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'fonts', '华文中宋.ttf'))
    font = None
    for path in candidates:
        if os.path.isfile(path):
            try:
                from PIL import ImageFont
                font = ImageFont.truetype(path, pixel_size)
                break
            except Exception:
                continue
    if font is None:
        from PIL import ImageFont
        font = ImageFont.load_default()
    if len(_FONT_CACHE) >= _FONT_CACHE_LIMIT:
        _FONT_CACHE.clear()
    _FONT_CACHE[pixel_size] = font
    return font


def render_welcome_preview(content, background_image=None,
                           background_color=None, font_size=None,
                           width=640, height=360):
    """按实际输出同样的规则渲染预览图，返回 (PIL Image, 字号磅值)。"""
    from PIL import Image, ImageDraw
    used_size = _welcome_font_size_pts(content, font_size)
    if background_image and os.path.isfile(background_image):
        # 取缓存好的背景（同一张图不重复解码），copy 一份再画字
        img = _cover_image(background_image, width, height).copy()
    else:
        r, g, b = _color_rgb(background_color or '#ffffff')
        img = Image.new('RGB', (width, height), (r, g, b))

    text_ole = _text_color_for(background_image, background_color)
    color = ((text_ole >> 16) & 255, (text_ole >> 8) & 255,
             text_ole & 255)
    draw = ImageDraw.Draw(img)
    scale = width / 960.0
    font_px = max(10, int(round(used_size * scale)))
    font = _load_cn_font(font_px)
    draw.multiline_text(
        (width / 2, height / 2),
        content,
        font=font,
        fill=color,
        align='center',
        anchor='mm',
        spacing=int(font_px * 0.28))
    return img, used_size


def generate_welcome_slide(content, out_path=None, font_size=None,
                           background_image=None, background_color=None,
                           on_progress=None):
    """渲染一页 16:9 效果图并放入新 PPT，文字保证完全居中。"""
    content = (content or '').replace('\r\n', '\n').replace('\r', '\n')
    if not content.strip():
        raise PptConvertError('请输入欢迎语内容。')
    if not out_path:
        out_path = os.path.join(
            os.path.expanduser('~'),
            '欢迎语_%s.pptx' % time.strftime('%Y%m%d_%H%M%S'))
    out_path = os.path.abspath(out_path)
    ext = os.path.splitext(out_path)[1].lower()
    save_format = 24 if ext == '.pptx' else (
        1 if ext == '.ppt' else None)
    if save_format is None:
        raise PptConvertError('输出文件只能用 .ppt 或 .pptx 后缀。')
    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except Exception as e:
            raise PptConvertError(
                '无法覆盖输出文件：\n%s\n%s' % (out_path, e)) from None

    _report(on_progress, 1, 3, '正在渲染欢迎页…')
    used_size = _welcome_font_size_pts(content, font_size)
    try:
        img, _ = render_welcome_preview(
            content, background_image=background_image,
            background_color=background_color, font_size=used_size,
            width=1920, height=1080)
    except Exception as e:
        raise PptConvertError('背景图片处理失败：%s' % e) from None

    tmp_dir = tempfile.mkdtemp(prefix='welcome_ppt_')
    tmp_png = os.path.join(tmp_dir, 'welcome.png')
    try:
        img.save(tmp_png, 'PNG')
    finally:
        img.close()  # 1920×1080 的图，早点释放，别等到函数结束
    try:
        with WpsSession() as session:
            pres = session.app.Presentations.Add()
            try:
                pres.PageSetup.SlideWidth = SLIDE_W
                pres.PageSetup.SlideHeight = SLIDE_H
                slide = pres.Slides.Add(1, 12)
                slide.Shapes.AddPicture(
                    tmp_png, False, True, 0, 0, SLIDE_W, SLIDE_H)
                _report(on_progress, 3, 3, '正在保存文件…')
                pres.SaveAs(out_path, save_format)
            finally:
                try:
                    pres.Close()
                except Exception:
                    pass
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return out_path, used_size


def convert_ppt_to_images(src, out_dir=None, width=EXPORT_WIDTH,
                          on_progress=None):
    """把 PPT 每页导出为一张 PNG，返回 (输出目录, 页数, 图片列表)。"""
    src = _ensure_input(src)
    if out_dir is None:
        stem = os.path.splitext(os.path.basename(src))[0]
        out_dir = os.path.join(os.path.dirname(src), stem + '_图片')
    os.makedirs(out_dir, exist_ok=True)

    stem = os.path.splitext(os.path.basename(src))[0]
    with WpsSession() as session:
        pres = open_presentation(session.app, src)
        try:
            count = pres.Slides.Count
            if count <= 0:
                raise PptConvertError('PPT 里没有任何幻灯片。')
            w, h = _page_size_pts(pres)
            img_w, img_h = _image_size(w, h, width)
            digits = len(str(count))
            paths = []
            for i in range(1, count + 1):
                png = os.path.join(
                    out_dir, '%s_%0*d.png' % (stem, digits, i))
                _report(on_progress, i, count,
                        '正在导出第 %d/%d 页…' % (i, count))
                try:
                    pres.Slides(i).Export(png, 'PNG', img_w, img_h)
                except Exception as e:
                    raise PptConvertError(
                        '第 %d 页导出失败：%s' % (i, e)) from None
                paths.append(png)
        finally:
            try:
                pres.Close()
            except Exception:
                pass
    return out_dir, count, paths


def export_slide_thumbnail(src, slide_index=1, width=800):
    """把 PPT 的某一页导出为一张 PNG，返回该 PNG 绝对路径（用于模板预览）。"""
    src = _ensure_input(src)
    with WpsSession() as session:
        pres = open_presentation(session.app, src)
        try:
            total = pres.Slides.Count
            if total <= 0:
                raise PptConvertError('PPT 里没有任何幻灯片。')
            idx = min(max(int(slide_index), 1), total)
            w, h = _page_size_pts(pres)
            img_w, img_h = _image_size(w, h, width)
            png = os.path.join(
                tempfile.gettempdir(),
                'welcome_preview_%s_%d.png' % (uuid.uuid4().hex[:8], idx))
            pres.Slides(idx).Export(png, 'PNG', img_w, img_h)
        finally:
            try:
                pres.Close()
            except Exception:
                pass
    return png


class SlideShowPlayer(object):
    """原样放映 PPT：用 WPS 演示 / PowerPoint 自己放，动画、切换、内嵌视频音频都在。

    放映窗口可以摆到指定矩形上（长条屏中间那块），所以左右两侧的固定图
    仍然留在背景上。翻页优先走 COM（View.Next/Previous/GotoSlide），
    万一这台机器不认，就退化成给放映窗口发 PgDn/PgUp。
    """

    def __init__(self):
        self._app = None
        self._pres = None
        self._show = None
        self.rect = None
        self.last_error = ''

    # ---------- 打开 / 摆位 ----------
    def start(self, path, rect=None, topmost=True):
        """开始放映；给了 rect 就按“窗口放映”摆到那块矩形里。"""
        self.close()
        try:
            path = _ensure_input(path)
        except Exception as exc:
            self.last_error = str(exc)
            return False
        try:
            import win32com.client
        except Exception as exc:
            self.last_error = '缺少 pywin32 组件：%s' % exc
            return False
        last = None
        for progid in WPS_PROGIDS:
            try:
                self._app = win32com.client.DispatchEx(progid)
                break
            except Exception as exc:
                last = exc
        if self._app is None:
            self.last_error = '打不开 WPS 演示 / PowerPoint：%s' % (last or '')
            return False
        try:
            try:
                self._app.Visible = True
            except Exception:
                pass
            try:
                self._app.DisplayAlerts = 0
            except Exception:
                pass
            for with_window in (False, True):
                try:
                    self._pres = self._app.Presentations.Open(
                        path, True, False, with_window)
                    break
                except Exception as exc:
                    self._pres = None
                    last = exc
            if self._pres is None:
                raise PptConvertError('打不开这份 PPT：%s' % (last or ''))
            settings = self._pres.SlideShowSettings
            if rect is not None:
                try:
                    settings.ShowType = 2      # ppShowTypeWindow：窗口放映
                except Exception:
                    pass
            self._show = settings.Run()
            self.rect = tuple(rect) if rect else None
            time.sleep(0.8)
            if self.rect:
                self.place(self.rect, topmost=topmost)
                time.sleep(0.4)
                self.place(self.rect, topmost=topmost)
            return True
        except Exception as exc:
            self.last_error = str(exc)
            self.close()
            return False

    def place(self, rect, topmost=True):
        """把放映窗口摆到指定矩形（屏幕坐标 left/top/right/bottom）。"""
        if not rect:
            return False
        left, top, right, bottom = [int(v) for v in rect]
        width = max(64, right - left)
        height = max(64, bottom - top)
        self.rect = (left, top, left + width, top + height)
        ok = False
        try:
            for attr, value in (('Left', left), ('Top', top),
                                ('Width', width), ('Height', height)):
                try:
                    setattr(self._show, attr, value)
                    ok = True
                except Exception:
                    pass
        except Exception:
            pass
        hwnd = self.hwnd()
        if hwnd:
            try:
                import video_player
                video_player.move_window(hwnd, self.rect, topmost=topmost)
                ok = True
            except Exception as exc:
                self.last_error = str(exc)
        return ok

    def hwnd(self):
        """放映窗口句柄：先问 COM，再按窗口类名找。"""
        try:
            hwnd = int(self._show.HWND)
            if hwnd:
                return hwnd
        except Exception:
            pass
        return self._find_window()

    def _find_window(self):
        """兜底：找那个放映窗口（类名里带 screen/slide/ppt，且不是本程序）。"""
        try:
            import win32gui
            import win32process
        except Exception:
            return None
        me = os.getpid()
        hits = []

        def _cb(hwnd, _param):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                cls = win32gui.GetClassName(hwnd).lower()
                if not any(k in cls for k in ('screen', 'slide', 'wpp',
                                              'presentation')):
                    return True
                _t, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == me:
                    return True
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                if right - left < 200 or bottom - top < 150:
                    return True
                hits.append(((right - left) * (bottom - top), hwnd))
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(_cb, None)
        except Exception:
            pass
        return max(hits)[1] if hits else None

    # ---------- 放映控制 ----------
    def _view(self):
        try:
            return self._show.View
        except Exception:
            return None

    def next(self):
        return self._step('next')

    def prev(self):
        return self._step('prev')

    def _step(self, which):
        view = self._view()
        if view is not None:
            try:
                if which == 'next':
                    view.Next()
                else:
                    view.Previous()
                return True
            except Exception as exc:
                self.last_error = str(exc)
        hwnd = self.hwnd()
        if not hwnd:
            return False
        try:
            import win32api
            import win32con
            import win32gui
            win32gui.SetForegroundWindow(hwnd)
            vk = win32con.VK_NEXT if which == 'next' else win32con.VK_PRIOR
            win32api.keybd_event(vk, 0, 0, 0)
            win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
            return True
        except Exception as exc:
            self.last_error = str(exc)
            return False

    def goto(self, index):
        """跳到第 index 页（从 1 开始）。"""
        view = self._view()
        if view is None:
            return False
        try:
            view.GotoSlide(max(1, int(index)))
            return True
        except Exception as exc:
            self.last_error = str(exc)
            return False

    def position(self):
        """当前放映到第几页（从 1 开始）；读不到返回 0。"""
        try:
            return int(self._show.View.CurrentShowPosition)
        except Exception:
            return 0

    def close(self):
        show, pres, app = self._show, self._pres, self._app
        self._show = self._pres = self._app = None
        self.rect = None
        try:
            if show is not None:
                show.View.Exit()
        except Exception:
            pass
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass


def _normalize_output(src, out_path):
    """处理输出路径，保证有 .pptx / .ppt 后缀。"""
    if out_path:
        out_path = os.path.abspath(out_path)
    else:
        stem = os.path.splitext(os.path.basename(src))[0]
        out_path = os.path.join(os.path.dirname(src), stem + '_16x9.pptx')
    ext = os.path.splitext(out_path)[1].lower()
    if ext == '.ppt':
        return out_path, 1  # ppSaveAsPresentation
    if ext == '.pptx':
        return out_path, 24  # ppSaveAsOpenXMLPresentation
    raise PptConvertError('输出文件只能用 .ppt 或 .pptx 后缀。')


def convert_ppt_to_16x9(src, out_path=None, width=EXPORT_WIDTH,
                        on_progress=None):
    """生成 16:9 全屏 PPT：每页放一张拉伸铺满整页的原 PPT 图片。"""
    src = _ensure_input(src)
    out_path, save_format = _normalize_output(src, out_path)
    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    tmp_dir = tempfile.mkdtemp(prefix='ppt169_')
    with WpsSession() as session:
        pres = open_presentation(session.app, src)
        try:
            count = pres.Slides.Count
            if count <= 0:
                raise PptConvertError('PPT 里没有任何幻灯片。')
            w, h = _page_size_pts(pres)
            img_w, img_h = _image_size(w, h, width)
            pngs = []
            total_steps = count * 2 + 2
            for i in range(1, count + 1):
                png = os.path.join(tmp_dir, '%03d.png' % i)
                _report(on_progress, i, total_steps,
                        '正在生成第 %d/%d 页图片…' % (i, count))
                try:
                    pres.Slides(i).Export(png, 'PNG', img_w, img_h)
                except Exception as e:
                    raise PptConvertError(
                        '第 %d 页导出失败：%s' % (i, e)) from None
                pngs.append(png)
        finally:
            try:
                pres.Close()
            except Exception:
                pass

        try:
            _report(on_progress, count, total_steps,
                    '正在创建 16:9 新 PPT…')
            newp = session.app.Presentations.Add()
            newp.PageSetup.SlideWidth = SLIDE_W
            newp.PageSetup.SlideHeight = SLIDE_H
            for i, png in enumerate(pngs, 1):
                _report(on_progress, count + i, total_steps,
                        '正在写入第 %d/%d 页…' % (i, count))
                slide = newp.Slides.Add(i, 12)  # 12 = 空白版式
                slide.Shapes.AddPicture(
                    png, False, True, 0, 0, SLIDE_W, SLIDE_H)

            if os.path.exists(out_path):
                try:
                    os.remove(out_path)
                except Exception as e:
                    raise PptConvertError(
                        '无法覆盖已存在的输出文件：\n%s\n%s'
                        % (out_path, e)) from None
            try:
                _report(on_progress, total_steps, total_steps,
                        '正在保存文件…')
                newp.SaveAs(out_path, save_format)
            except Exception as e:
                raise PptConvertError('保存新 PPT 失败：\n%s' % e) from None
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
    return out_path, count
